require('dotenv').config();
const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');
const iconv = require('iconv-lite');
const jschardet = require('jschardet');
const Anthropic = require('@anthropic-ai/sdk');

const app = express();
const PORT = 4001;
const DIST_WARNING_KM = 30; // これを超える距離のマッチングは「距離が遠い」として画面上で警告表示する
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

// 「所定フォルダ」: 担当者から共有されたExcel/CSVをここに置いておけば、ブラウザでファイルを都度選ばなくても
// 取込み画面から一覧表示→選択して取込みできる(uploads/フォルダをそのまま使う)
const UPLOADS_DIR = path.join(__dirname, 'uploads');
fs.mkdirSync(UPLOADS_DIR, { recursive: true });
const INCOMING_FILE_EXTENSIONS = ['.xlsx', '.xls', '.csv', '.txt'];

// 備考・LINEメッセージのAI解析(任意機能)。.envにANTHROPIC_API_KEYが設定されていない場合はnullのままで、
// 関連エンドポイントは「AI未設定」を返す(既存の正規表現ベースの解析は影響を受けない)
const anthropic = process.env.ANTHROPIC_API_KEY ? new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY }) : null;
const AI_MODEL = 'claude-haiku-4-5-20251001'; // 解析用途のため、速く安価なモデルを使う

// 自動マッチングのスコアリング重み(調整しやすいようファイル冒頭に定数化しておく)
const SCORE_WEIGHT_DISTANCE_KM = -1;     // 距離1kmごとの減点
const SCORE_STRONG_MATCH_BONUS = 50;     // 希望店舗/希望エリアが一致する場合のボーナス
// 相性はNG(除外)〜1〜2〜3〜4〜5〜OKの7段階。NGを除く各段階に1〜6の重みを割り当て、重み×この値をボーナスにする
// (1点=6, ... 5点=30, OK=36。数字が大きい/OKほど優先度が上がる)
const SCORE_PREFERENCE_LEVEL_UNIT = 6;
const PREFERENCE_LEVEL_WEIGHT = { '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, 'OK': 6 };
const SCORE_EXPERIENCE_PER_VISIT = 5;    // その店舗での過去派遣1回あたりのボーナス(店舗の勝手を知っている)
const SCORE_EXPERIENCE_MAX_VISITS = 5;   // 店舗経験ボーナスの上限回数
const SCORE_AREA_EXPERIENCE_PER_VISIT = 2; // 同エリアでの過去派遣1回あたりのボーナス(配達エリアの知見)
const SCORE_AREA_EXPERIENCE_MAX_VISITS = 5; // エリア経験ボーナスの上限回数

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const db = new sqlite3.Database(path.join(__dirname, 'matching.db'));

// 複数プロセスから同時にこのDBファイルが開かれても書き込みが失われにくいようにする。
// WALモードは読み取りと書き込みが競合しにくく、busy_timeoutは他プロセスが書き込み中でも
// すぐにエラーにせず一定時間リトライしてから諦めるようにする(誤って二重起動してしまった場合の保険)
db.run('PRAGMA journal_mode = WAL');
db.run('PRAGMA busy_timeout = 5000');

db.serialize(() => {
  // 個人事業主ドライバーのマスタ(自宅住所を持たせ、依頼店舗までの距離をマッチング時に計算する)
  db.run(`
    CREATE TABLE IF NOT EXISTS drivers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      phone TEXT,
      vehicle_type TEXT,
      home_address TEXT,
      home_lat REAL,
      home_lng REAL,
      notes TEXT,
      created_at TEXT
    )
  `);

  // ドライバーの希望シフト(楽シフからのエクスポートを取り込む想定。現時点では手入力/汎用CSVで代用)
  db.run(`
    CREATE TABLE IF NOT EXISTS driver_availability (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      driver_id INTEGER NOT NULL,
      desired_date TEXT NOT NULL,
      desired_area TEXT,
      desired_store TEXT,
      time_start TEXT,
      time_end TEXT,
      requests TEXT,
      created_at TEXT,
      FOREIGN KEY (driver_id) REFERENCES drivers(id)
    )
  `);

  // 店舗からの人員要請(現時点ではExcelを見て手入力/汎用CSVで代用)
  db.run(`
    CREATE TABLE IF NOT EXISTS store_requests (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      store_name TEXT NOT NULL,
      area TEXT,
      address TEXT,
      lat REAL,
      lng REAL,
      request_date TEXT NOT NULL,
      time_start TEXT,
      time_end TEXT,
      required_count INTEGER DEFAULT 1,
      requests TEXT,
      created_at TEXT
    )
  `);

  // 自動マッチングの結果(候補として作成。距離が遠い場合は is_far_warning を立てる)
  db.run(`
    CREATE TABLE IF NOT EXISTS matches (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      store_request_id INTEGER NOT NULL,
      driver_id INTEGER NOT NULL,
      match_date TEXT NOT NULL,
      distance_km REAL,
      is_far_warning INTEGER DEFAULT 0,
      status TEXT DEFAULT '候補',
      created_at TEXT,
      FOREIGN KEY (store_request_id) REFERENCES store_requests(id),
      FOREIGN KEY (driver_id) REFERENCES drivers(id)
    )
  `);

  // 店舗マスタ(店舗依頼は店舗名の自由入力だが、裏側でここに名寄せして集約する。
  // これにより「この店舗が好き/苦手」という相性や、過去の派遣履歴を店舗単位で紐付けられる)
  db.run(`
    CREATE TABLE IF NOT EXISTS stores (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL UNIQUE,
      area TEXT,
      address TEXT,
      lat REAL,
      lng REAL,
      notes TEXT,
      created_at TEXT
    )
  `);

  // ドライバーごとの店舗との相性。NG〜1〜2〜3〜4〜5〜OKの7段階の一つの尺度で管理する
  // (NGは自動マッチングの候補から除外し、数字が大きい/OKほど優先度を上げる。未設定の店舗は「履歴なし」として扱う)
  db.run(`
    CREATE TABLE IF NOT EXISTS driver_store_preferences (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      driver_id INTEGER NOT NULL,
      store_id INTEGER NOT NULL,
      preference TEXT NOT NULL CHECK (preference IN ('NG', '1', '2', '3', '4', '5', 'OK')),
      notes TEXT,
      created_at TEXT,
      UNIQUE(driver_id, store_id),
      FOREIGN KEY (driver_id) REFERENCES drivers(id),
      FOREIGN KEY (store_id) REFERENCES stores(id)
    )
  `);

  // 派遣が実際に完了した履歴(店舗の勝手や配達エリアの知見の蓄積として、マッチングの優先度づけに使う)
  db.run(`
    CREATE TABLE IF NOT EXISTS dispatch_history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      driver_id INTEGER NOT NULL,
      store_id INTEGER,
      match_id INTEGER,
      work_date TEXT NOT NULL,
      notes TEXT,
      created_at TEXT,
      FOREIGN KEY (driver_id) REFERENCES drivers(id),
      FOREIGN KEY (store_id) REFERENCES stores(id),
      FOREIGN KEY (match_id) REFERENCES matches(id)
    )
  `);

  // アプリ全体の設定(key-value)。マッチングの優先順位階層など、画面から変更できる設定をここに保存する
  db.run(`
    CREATE TABLE IF NOT EXISTS app_settings (
      key TEXT PRIMARY KEY,
      value TEXT
    )
  `);

  // 固定希望店舗が変更されるたびに記録しておく(手動設定か、どの自動反映によるものかを追跡できるようにする)
  db.run(`
    CREATE TABLE IF NOT EXISTS fixed_store_change_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      driver_id INTEGER NOT NULL,
      old_store_id INTEGER,
      new_store_id INTEGER,
      source TEXT NOT NULL,
      created_at TEXT,
      FOREIGN KEY (driver_id) REFERENCES drivers(id)
    )
  `);

  // 既存テーブルへのカラム追加(運用中のDBを壊さないよう、無ければ追加する形で行う)
  ensureColumn('store_requests', 'store_id', 'INTEGER REFERENCES stores(id)');
  ensureColumn('matches', 'score', 'REAL');
  ensureColumn('matches', 'preference_flag', 'TEXT');
  ensureColumn('matches', 'experience_count', 'INTEGER DEFAULT 0');
  ensureColumn('matches', 'absence_reason', 'TEXT');
  ensureColumn('matches', 'replaced_by_match_id', 'INTEGER');
  ensureColumn('stores', 'manager_name', 'TEXT'); // 店舗担当者(店長)
  ensureColumn('stores', 'sv_name', 'TEXT');      // SV(スーパーバイザー)
  ensureColumn('stores', 'store_code', 'TEXT');   // 店番(店舗を一意に表す社内コード。現時点では未入力で運用)
  ensureColumn('drivers', 'driver_code', 'TEXT');          // 社員コード(ドライバーを一意に表す社内コード)
  ensureColumn('drivers', 'first_contract_date', 'TEXT');  // 初回委託日
  ensureColumn('drivers', 'status', 'TEXT');               // ステータス(稼働中/休止/契約終了)
  ensureColumn('drivers', 'email', 'TEXT');                // メールアドレス
  ensureColumn('drivers', 'insurance_info', 'TEXT');        // 保険加入状況(貨物保険等のメモ)
  ensureColumn('drivers', 'company_name', 'TEXT');          // 会社名(法人として契約している個人事業主向け。個人の場合は空欄)
  ensureColumn('drivers', 'fixed_store_id', 'INTEGER REFERENCES stores(id)'); // 固定希望店舗(希望シフトで店舗未入力の日の初期値として使う)
  // エリア固定: 店舗を1つに固定するのではなく、曜日ごとの決まった時間帯+複数の候補店舗群(想定デポ)の中から
  // 優先的に割り当てる仕組み。area_fixed_patternは{mon:"10:00-22:00", ...}形式のJSON文字列、
  // area_fixed_store_idsは候補店舗idの配列のJSON文字列
  ensureColumn('drivers', 'area_fixed_enabled', 'INTEGER DEFAULT 0');
  ensureColumn('drivers', 'area_fixed_pattern', 'TEXT');
  ensureColumn('drivers', 'area_fixed_store_ids', 'TEXT');
  ensureColumn('store_requests', 'archived_month', 'TEXT');     // 月次クローズで「YYYY-MM」を入れ、作業画面から隠す(データは消さない)
  ensureColumn('driver_availability', 'archived_month', 'TEXT');
  ensureColumn('matches', 'archived_month', 'TEXT');
  migratePreferenceScale(); // 相性を好き/NGの2択からNG〜1〜5〜OKの7段階スケールに移行する(旧DB向け)
});

// driver_store_preferences の preference が旧スキーマ(good/ng の2択)のままなら、
// NG〜1〜5〜OKの7段階スケールに移行する(SQLiteはCHECK制約を直接変更できないためテーブルを作り直す)
function migratePreferenceScale() {
  db.get(`SELECT sql FROM sqlite_master WHERE type='table' AND name='driver_store_preferences'`, [], (err, row) => {
    if (err || !row || !row.sql.includes("'good'")) return; // 新スキーマ済み、またはテーブル無し
    db.serialize(() => {
      db.run(`ALTER TABLE driver_store_preferences RENAME TO driver_store_preferences_old_migration`);
      db.run(`
        CREATE TABLE driver_store_preferences (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          driver_id INTEGER NOT NULL,
          store_id INTEGER NOT NULL,
          preference TEXT NOT NULL CHECK (preference IN ('NG', '1', '2', '3', '4', '5', 'OK')),
          notes TEXT,
          created_at TEXT,
          UNIQUE(driver_id, store_id),
          FOREIGN KEY (driver_id) REFERENCES drivers(id),
          FOREIGN KEY (store_id) REFERENCES stores(id)
        )
      `);
      db.run(`
        INSERT INTO driver_store_preferences (id, driver_id, store_id, preference, notes, created_at)
        SELECT id, driver_id, store_id,
               CASE preference WHEN 'good' THEN 'OK' WHEN 'ng' THEN 'NG' ELSE preference END,
               notes, created_at
        FROM driver_store_preferences_old_migration
      `);
      db.run(`DROP TABLE driver_store_preferences_old_migration`);
      // matches側に残っている旧表記も新表記に合わせておく
      db.run(`UPDATE matches SET preference_flag = 'OK' WHERE preference_flag = 'good'`);
      db.run(`UPDATE matches SET preference_flag = 'NG' WHERE preference_flag = 'ng'`, err2 => {
        if (err2) console.error('相性スケールの移行に失敗しました:', err2);
        else console.log('相性の評価スケールをNG〜1〜5〜OKの7段階に移行しました');
      });
    });
  });
}

// 指定したカラムがテーブルに無ければ ALTER TABLE で追加する(何度サーバーを再起動しても安全)
function ensureColumn(table, column, ddlType) {
  db.all(`PRAGMA table_info(${table})`, [], (err, cols) => {
    if (err) { console.error(`スキーマ確認に失敗しました(${table}.${column}):`, err); return; }
    if (cols.some(c => c.name === column)) return;
    db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${ddlType}`, err2 => {
      if (err2) console.error(`カラム追加に失敗しました(${table}.${column}):`, err2);
    });
  });
}

function dbAll(sql, params = []) {
  return new Promise((resolve, reject) => db.all(sql, params, (err, rows) => err ? reject(err) : resolve(rows)));
}
function dbGet(sql, params = []) {
  return new Promise((resolve, reject) => db.get(sql, params, (err, row) => err ? reject(err) : resolve(row)));
}
function dbRun(sql, params = []) {
  return new Promise((resolve, reject) => db.run(sql, params, function (err) { err ? reject(err) : resolve(this); }));
}

// 国土地理院(GSI)の住所検索API(無料・キー不要)で住所→緯度経度を取得する
async function geocodeAddress(address) {
  if (!address) return null;
  try {
    const url = `https://msearch.gsi.go.jp/address-search/AddressSearch?q=${encodeURIComponent(address)}`;
    const r = await fetch(url);
    if (!r.ok) return null;
    const data = await r.json();
    if (!Array.isArray(data) || data.length === 0) return null;
    return { lat: data[0].geometry.coordinates[1], lng: data[0].geometry.coordinates[0] };
  } catch (e) {
    return null;
  }
}

// 店舗名の表記ゆれ(「class」接頭辞、「【Tax-Free】」等の装飾)を取り除いて正規化する
function normalizeStoreName(store_name) {
  return (store_name || '')
    .replace(/^class/i, '')
    .replace(/【[^】]*】/g, '')
    .trim();
}

// 店舗名(自由入力)を店舗マスタに名寄せする。既存店舗が見つかればそのidを返し、
// 住所/緯度経度が未設定であれば補完する。見つからなければ新規に登録する
async function resolveStoreId({ store_name, area, address, lat, lng }) {
  // 取込み元によって店舗名に「class」接頭辞や「【Tax-Free】」表記が付いたり付かなかったりするため、
  // 正規化してから既存の店舗マスタと照合する(そうしないと同じ店舗が表記違いで重複登録されてしまう)
  const name = normalizeStoreName(store_name);
  if (!name) return null;
  const now = new Date().toISOString();
  const existing = await dbGet('SELECT * FROM stores WHERE name = ?', [name]);
  if (existing) {
    if (!existing.address && address) {
      await dbRun('UPDATE stores SET area = COALESCE(NULLIF(area, \'\'), ?), address = ?, lat = ?, lng = ? WHERE id = ?', [area || '', address, lat, lng, existing.id]);
    }
    return existing.id;
  }
  const result = await dbRun(
    'INSERT INTO stores (name, area, address, lat, lng, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)',
    [name, area || '', address || '', lat, lng, '', now]
  );
  return result.lastID;
}

// 店舗マスタに登録済みの住所・緯度経度・エリアを取得する(店舗依頼側で住所が未入力のときのフォールバックに使う)
async function getStoreMasterInfo(store_id) {
  if (!store_id) return null;
  return dbGet('SELECT area, address, lat, lng FROM stores WHERE id = ?', [store_id]);
}

// 固定希望店舗の変更を記録する(手動設定か、どの自動反映によるものかを後から追跡できるようにする)
async function logFixedStoreChange(driver_id, old_store_id, new_store_id, source) {
  await dbRun(
    'INSERT INTO fixed_store_change_log (driver_id, old_store_id, new_store_id, source, created_at) VALUES (?, ?, ?, ?, ?)',
    [driver_id, old_store_id || null, new_store_id || null, source, new Date().toISOString()]
  );
}

// 「確定」「完了」のマッチングを見て、1ヶ月間ずっと同じ店舗に割り当てられているドライバーがいれば、
// そのドライバーマスタの「固定希望店舗」に自動反映する(2日以上、かつ全ての確定/完了マッチングが
// 同じ店舗の場合のみ対象。既に同じ店舗が設定済みなら何もしない)。確定・完了操作のたびに呼び出す想定。
// エリア固定が有効なドライバーは対象外(たまたま数日同じ店舗が続いただけで固定希望店舗が設定されてしまうと、
// 本来は複数の候補店舗を柔軟に回る設計のエリア固定が、1店舗に固定される優先階層に上書きされてしまうため)
async function applyAutoFixedStoreFromMatches() {
  const rows = await dbAll(`
    SELECT m.driver_id, s.store_id, d.fixed_store_id
    FROM matches m
    JOIN store_requests s ON s.id = m.store_request_id
    JOIN drivers d ON d.id = m.driver_id
    WHERE m.archived_month IS NULL AND m.status IN ('確定', '完了') AND d.area_fixed_enabled = 0
  `);
  const byDriver = new Map(); // driver_id -> { storeIds: Set, count, currentFixedStoreId }
  for (const r of rows) {
    const entry = byDriver.get(r.driver_id) || { storeIds: new Set(), count: 0, currentFixedStoreId: r.fixed_store_id };
    entry.storeIds.add(r.store_id);
    entry.count++;
    byDriver.set(r.driver_id, entry);
  }

  let updated = 0;
  for (const [driver_id, entry] of byDriver) {
    if (entry.count < 2 || entry.storeIds.size !== 1) continue; // 同じ店舗が2件以上続いている場合のみ対象
    const onlyStoreId = [...entry.storeIds][0];
    if (entry.currentFixedStoreId === onlyStoreId) continue; // 既に同じ設定なら何もしない
    await dbRun('UPDATE drivers SET fixed_store_id = ? WHERE id = ?', [onlyStoreId, driver_id]);
    await logFixedStoreChange(driver_id, entry.currentFixedStoreId, onlyStoreId, 'auto_pattern');
    updated++;
  }
  return updated;
}

// 週次パターン取込みで、曜日/時間の列が読み取れず「要確認」になった行の備考を、AI(Claude)で試しに解釈する(試験的機能)。
// 自動登録はせず、コーディネーターが確認しやすいようエラーメッセージにAIの解釈候補を添えるだけに留める。
// 未設定時やAI呼び出し失敗時はnullを返し、呼び出し側は従来通りの「要確認」メッセージのみ表示する
async function interpretNoteWithAI(noteText, year, month) {
  if (!anthropic || !noteText) return null;
  try {
    const response = await anthropic.messages.create({
      model: AI_MODEL,
      max_tokens: 500,
      tools: [{
        name: 'interpret_note',
        description: '店舗からの依頼スプレッドシートの備考欄から、具体的な依頼内容(日付・時間帯・人数)を読み取る',
        input_schema: {
          type: 'object',
          properties: {
            understood: { type: 'boolean', description: '具体的な日付や時間帯の依頼として読み取れたかどうか' },
            desired_date: { type: 'string', description: `YYYY-MM-DD形式。年は${year}年、月の記載がなければ${month}月として補う。読み取れなければ空文字` },
            time_start: { type: 'string', description: 'HH:MM形式。読み取れなければ空文字' },
            time_end: { type: 'string', description: 'HH:MM形式。読み取れなければ空文字' },
            required_count: { type: 'number', description: '必要人数。読み取れなければ1' }
          },
          required: ['understood', 'desired_date', 'time_start', 'time_end', 'required_count']
        }
      }],
      tool_choice: { type: 'tool', name: 'interpret_note' },
      messages: [{ role: 'user', content: `店舗からの依頼スプレッドシートの備考欄です。具体的な依頼内容を読み取ってください:\n「${noteText}」` }]
    });
    const toolUse = response.content.find(c => c.type === 'tool_use');
    if (!toolUse || !toolUse.input.understood) return null;
    return toolUse.input;
  } catch (e) {
    return null; // AI解析に失敗しても取込み自体は止めない(従来通りの要確認メッセージのみ表示する)
  }
}

// 「10:00-22:00」「10〜22」「10時〜22時」のような自由な書き方の時間帯を{start, end}(HH:MM)に変換する。
// 読み取れなければnull(エリア固定の曜日パターン入力欄で使う)
function parseTimeRangeText(text) {
  const m = String(text || '').match(/(\d{1,2})[:時]?(\d{2})?\s*[-－‐–—〜～~ー―−]\s*(\d{1,2})[:時]?(\d{2})?/);
  if (!m) return null;
  return {
    start: `${m[1].padStart(2, '0')}:${(m[2] || '00').padStart(2, '0')}`,
    end: `${m[3].padStart(2, '0')}:${(m[4] || '00').padStart(2, '0')}`
  };
}

// JSの Date.getDay()(0=日,1=月,...,6=土)を、エリア固定パターンのキー(mon/tue/...)に変換する
const WEEKDAY_KEY_BY_JSDAY = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
function weekdayKeyOf(dateStr) {
  return WEEKDAY_KEY_BY_JSDAY[new Date(dateStr + 'T00:00:00').getDay()];
}

// 候補者(driver_id, home_lat, home_lng, desired_store, desired_area を持つオブジェクト)を
// 店舗依頼に対してスコアリングする(自動マッチングと欠勤時の代替候補探しの両方で使う共通ロジック)
function scoreCandidate(a, store, preferenceByDriverStore, storeVisitCount, areaVisitCount) {
  const hasCoords = a.home_lat != null && a.home_lng != null && store.lat != null && store.lng != null;
  const distance = hasCoords ? haversineKm(a.home_lat, a.home_lng, store.lat, store.lng) : null;
  const isStrongMatch = (a.desired_store && a.desired_store === store.store_name) ||
    (a.desired_area && store.area && a.desired_area === store.area);
  const preference = preferenceByDriverStore.get(`${a.driver_id}:${store.store_id}`) || null;
  const experienceCount = Math.min(storeVisitCount.get(`${a.driver_id}:${store.store_id}`) || 0, SCORE_EXPERIENCE_MAX_VISITS);
  const areaExperienceCount = Math.min(areaVisitCount.get(`${a.driver_id}:${store.area}`) || 0, SCORE_AREA_EXPERIENCE_MAX_VISITS);

  let score = 0;
  score += distance != null ? distance * SCORE_WEIGHT_DISTANCE_KM : -9999; // 距離不明は最低評価にして末尾に回す
  if (isStrongMatch) score += SCORE_STRONG_MATCH_BONUS;
  if (preference && PREFERENCE_LEVEL_WEIGHT[preference]) score += PREFERENCE_LEVEL_WEIGHT[preference] * SCORE_PREFERENCE_LEVEL_UNIT;
  score += experienceCount * SCORE_EXPERIENCE_PER_VISIT;
  score += areaExperienceCount * SCORE_AREA_EXPERIENCE_PER_VISIT;

  return { distance, score, preference, experienceCount };
}

// マッチングの優先順位(階層)のカタログ。/api/matches/run は、画面(マッチング設定)で有効化・
// 並び替えされた階層の順に、店舗の処理順に関係なく全店舗を通して先に確保してから、
// 残りをスコア順で埋める(2段階処理)。test(a, store, ctx)がtrueを返す候補がその階層の対象。
// ここに新しい階層を追加すれば、画面側で有効化・並び替えできるようになる
const TIER_CATALOG = {
  fixed_store: {
    label: '固定希望店舗',
    description: '1ヶ月間ずっと同じ店舗に割り当てられた等でドライバーマスタの「固定希望店舗」に設定されている店舗を最優先する',
    test: (a, store) => a.fixed_store_id === store.store_id
  },
  desired_store_match: {
    label: '希望シフトで指定した店舗',
    description: 'その日の希望シフトで本人が明示的に指定した店舗(固定希望店舗からの自動補完ではなく、本人が入力したもの)を優先する',
    test: (a, store) => !!a.desired_store_explicit && a.desired_store_explicit === store.store_name
  },
  desired_area_match: {
    label: '希望エリアが一致',
    description: 'その日の希望シフトで指定した希望エリアと、店舗のエリアが一致する場合を優先する',
    test: (a, store) => !!a.desired_area && !!store.area && a.desired_area === store.area
  },
  area_fixed: {
    label: 'エリア固定の候補店舗',
    description: 'エリア固定が有効なドライバーについて、その候補店舗(想定デポ)リストに含まれる店舗を優先する',
    test: (a, store, ctx) => (ctx.areaFixedStoreIdsByDriverId.get(a.driver_id) || []).includes(store.store_id)
  },
  preference_ok: {
    label: '店舗相性が「OK」',
    description: '店舗相性マスタで「OK」に設定されている店舗を優先する',
    test: (a, store, ctx) => ctx.preferenceByDriverStore.get(`${a.driver_id}:${store.store_id}`) === 'OK'
  },
  preference_high: {
    label: '店舗相性が高評価(4・5・OK)',
    description: '店舗相性マスタで4・5・OKのいずれかに設定されている店舗を優先する',
    test: (a, store, ctx) => ['4', '5', 'OK'].includes(ctx.preferenceByDriverStore.get(`${a.driver_id}:${store.store_id}`))
  },
  experience_store: {
    label: 'その店舗への派遣経験あり',
    description: '派遣履歴にその店舗への実績が1回以上ある場合を優先する(店舗の勝手を知っている)',
    test: (a, store, ctx) => (ctx.storeVisitCount.get(`${a.driver_id}:${store.store_id}`) || 0) > 0
  },
  experience_area: {
    label: 'そのエリアへの派遣経験あり',
    description: '派遣履歴にその店舗のエリアへの実績が1回以上ある場合を優先する(配達エリアの土地勘がある)',
    test: (a, store, ctx) => (ctx.areaVisitCount.get(`${a.driver_id}:${store.area}`) || 0) > 0
  },
  near_distance: {
    label: '自宅から近い(10km以内)',
    description: '自宅住所から店舗までの距離が10km以内の場合を優先する(※ドライバーの自宅住所が登録されていないと機能しません)',
    test: (a, store) => a.home_lat != null && a.home_lng != null && store.lat != null && store.lng != null &&
      haversineKm(a.home_lat, a.home_lng, store.lat, store.lng) <= 10
  },
};
const DEFAULT_ENABLED_TIER_KEYS = ['fixed_store']; // 画面でまだ設定したことが無い場合の初期値(今までの挙動を維持する)

// app_settingsに保存されている、有効化・並び替え済みの優先順位階層のキー配列を返す
// (カタログに無いキーが混ざっていた場合は無視する。保存が無ければ初期値を返す)
async function getEnabledTierKeys() {
  const row = await dbGet('SELECT value FROM app_settings WHERE key = ?', ['priority_tiers']);
  if (!row) return DEFAULT_ENABLED_TIER_KEYS;
  try {
    const keys = JSON.parse(row.value);
    const filtered = Array.isArray(keys) ? keys.filter(k => TIER_CATALOG[k]) : null;
    return (filtered && filtered.length > 0) ? filtered : DEFAULT_ENABLED_TIER_KEYS;
  } catch (e) {
    return DEFAULT_ENABLED_TIER_KEYS;
  }
}

// ドライバー×店舗の相性マップ、店舗単位・エリア単位の過去派遣回数集計をまとめて用意する
// (自動マッチングと代替候補探しの両方で使う共通の準備処理)
async function loadScoringContext() {
  const preferences = await dbAll('SELECT driver_id, store_id, preference FROM driver_store_preferences');
  const preferenceByDriverStore = new Map(preferences.map(p => [`${p.driver_id}:${p.store_id}`, p.preference]));

  const historyRows = await dbAll(`
    SELECT h.driver_id, h.store_id, s.area
    FROM dispatch_history h
    LEFT JOIN stores s ON s.id = h.store_id
  `);
  const storeVisitCount = new Map(); // "driver_id:store_id" -> 回数
  const areaVisitCount = new Map();  // "driver_id:area" -> 回数
  for (const h of historyRows) {
    if (h.store_id != null) {
      const key = `${h.driver_id}:${h.store_id}`;
      storeVisitCount.set(key, (storeVisitCount.get(key) || 0) + 1);
    }
    if (h.area) {
      const key = `${h.driver_id}:${h.area}`;
      areaVisitCount.set(key, (areaVisitCount.get(key) || 0) + 1);
    }
  }
  return { preferenceByDriverStore, storeVisitCount, areaVisitCount };
}

// 2点間の距離(km)をざっくり計算する(緯度経度からの直線距離。実際の道のりより短くなる点に注意)
function haversineKm(lat1, lng1, lat2, lng2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// シンプルなCSVパーサー(RFC4180ベース。ダブルクォートで囲まれたフィールド内のカンマ・改行・""に対応)。
// xlsxライブラリにCSVの文字列をそのまま渡すと、"2026-09-05"のような日付っぽい文字列を
// 勝手に数値(日付シリアル値)と誤解釈し、"9/5/26"のような意図しない表記に化けることがあるため、
// CSVは型推測をしないこの自前パーサーで読む(値は常に元の文字列のまま扱う)
function parseCsvText(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else { inQuotes = false; }
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ',') {
      row.push(field); field = '';
    } else if (c === '\n') {
      row.push(field); rows.push(row); row = []; field = '';
    } else if (c === '\r') {
      // 改行はLFで処理するのでCRは無視する
    } else {
      field += c;
    }
  }
  if (field !== '' || row.length > 0) { row.push(field); rows.push(row); }

  if (rows.length === 0) return [];
  const headers = rows[0].map(h => h.trim());
  return rows.slice(1)
    .filter(r => r.some(v => v.trim() !== ''))
    .map(r => {
      const obj = {};
      headers.forEach((h, idx) => { obj[h] = (r[idx] !== undefined ? r[idx] : '').trim(); });
      return obj;
    });
}

// アップロードされたファイル(.csv/.xlsx/.xls)を行オブジェクトの配列にする。
// CSVはExcelから保存すると文字コードがShift-JISになっていることが多いため、jschardetで判定してから変換する。
// .xlsx/.xlsはバイナリ形式なのでxlsxライブラリに渡す(こちらはcellDatesで正しく日付セルを扱える)
function parseUploadedSpreadsheet(buffer, filename) {
  const ext = (filename || '').toLowerCase();
  if (ext.endsWith('.csv') || ext.endsWith('.txt')) {
    const detected = jschardet.detect(buffer) || {};
    let text;
    try {
      text = iconv.decode(buffer, detected.encoding || 'UTF-8');
    } catch (e) {
      text = buffer.toString('utf8');
    }
    return parseCsvText(text);
  }
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
}

// 「氏名×日付」のワイド形式(1行目=日付見出し、1列目=氏名、セル=勤務時間 or 休み)のシフト表を
// 見出し行+各行を配列のまま返す(通常のparseUploadedSpreadsheetは列名をキーにしたオブジェクトを返すため、
// 「1(火)」のような日付見出しを順序どおり扱いたいこちらの用途には配列のままの形が必要)
// sheetNameを指定すればそのシートを、未指定(またはファイル内に無い名前)なら先頭シートを読む
function parseWideSpreadsheet(buffer, filename, sheetName) {
  const ext = (filename || '').toLowerCase();
  if (ext.endsWith('.csv') || ext.endsWith('.txt')) {
    const detected = jschardet.detect(buffer) || {};
    let text;
    try {
      text = iconv.decode(buffer, detected.encoding || 'UTF-8');
    } catch (e) {
      text = buffer.toString('utf8');
    }
    const objRows = parseCsvText(text);
    if (objRows.length === 0) return [];
    const keys = Object.keys(objRows[0]);
    return [keys, ...objRows.map(o => keys.map(k => o[k]))];
  }
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const targetName = (sheetName && workbook.SheetNames.includes(sheetName)) ? sheetName : workbook.SheetNames[0];
  const sheet = workbook.Sheets[targetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
}

// 列名は実際のExcel/楽シフのエクスポート形式が分かり次第、ここに実際の見出し名を追記して合わせていく想定。
// 現時点ではよくありそうな見出し名をいくつか候補として登録しておく
function pickField(row, aliases) {
  const keys = Object.keys(row);
  for (const alias of aliases) {
    const foundKey = keys.find(k => k.trim() === alias);
    if (foundKey !== undefined) return String(row[foundKey] ?? '').trim();
  }
  return '';
}

const STORE_ALIASES = {
  store_name: ['店舗名', '店舗', 'ストア名'],
  area: ['エリア', '地域'],
  address: ['住所', '所在地'],
  request_date: ['依頼日', '日付', '勤務日', '対象日'],
  time_start: ['開始時刻', '開始時間', '出勤時間'],
  time_end: ['終了時刻', '終了時間', '退勤時間'],
  required_count: ['必要人数', '人数'],
  requests: ['要望', '備考', 'メモ']
};

const AVAILABILITY_ALIASES = {
  driver_name: ['ドライバー名', '氏名', '名前'],
  desired_date: ['希望日', '勤務日', '対象日', '日付'],
  desired_area: ['希望エリア', 'エリア'],
  desired_store: ['希望店舗', '店舗名', '店舗'],
  time_start: ['開始時刻', '開始時間', '出勤時間'],
  time_end: ['終了時刻', '終了時間', '退勤時間'],
  requests: ['要望', '備考', 'メモ']
};

const HISTORY_ALIASES = {
  driver_name: ['ドライバー名', '氏名', '名前'],
  store_name: ['店舗名', '店舗', 'ストア名'],
  work_date: ['実施日', '派遣日', '勤務日', '日付'],
  notes: ['要望', '備考', 'メモ']
};

const DRIVER_MASTER_ALIASES = {
  driver_code: ['社員コード', 'ドライバーコード', '社員番号', 'driver_code'],
  name: ['氏名', '名前', 'ドライバー名', 'name'],
  company_name: ['会社名', '法人名', '屋号', 'company_name'],
  fixed_store_name: ['固定希望店舗', '固定店舗', 'fixed_store_name'],
  phone: ['電話番号', '電話', 'TEL', 'phone'],
  vehicle_type: ['車両種別', '車両', '車種', 'vehicle_type'],
  home_address: ['お住まい住所', '自宅住所', '住所', 'home_address'],
  first_contract_date: ['初回委託日', '委託開始日', '契約開始日', 'first_contract_date'],
  status: ['ステータス', '状態', 'status'],
  email: ['メールアドレス', 'メール', 'email'],
  insurance_info: ['保険加入状況', '保険', '貨物保険', 'insurance_info'],
  notes: ['備考', 'メモ', 'notes']
};

const STORE_MASTER_ALIASES = {
  store_name: ['店舗名', '店舗', 'ストア名', 'store_name', 'name'],
  store_code: ['店番', '店舗コード', 'store_code'],
  area: ['エリア', '地域', 'area'],
  address: ['住所', '所在地', 'address'],
  manager_name: ['店長', '店舗担当者', '担当者', 'manager_name'],
  sv_name: ['SV', 'スーパーバイザー', 'sv_name'],
  notes: ['備考', 'メモ', 'notes']
};

// Excelの日付セルは "2026-09-01" "2026/9/1" "2026年9月1日" など様々な形で来うるため、YYYY-MM-DDに揃える。
// 想定外の形式はそのまま返す(手動で直してもらう前提)
function normalizeDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})日?$/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return s;
}

// ===== ドライバーマスタ =====
app.get('/api/drivers', async (req, res) => {
  const rows = await dbAll(`
    SELECT d.*, fs.name AS fixed_store_name
    FROM drivers d
    LEFT JOIN stores fs ON fs.id = d.fixed_store_id
    ORDER BY d.id DESC
  `);
  res.json({ success: true, drivers: rows });
});

app.post('/api/drivers', async (req, res) => {
  const { id, name, phone, vehicle_type, home_address, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes,
          area_fixed_enabled, area_fixed_pattern, area_fixed_store_ids } = req.body;
  if (!name) return res.status(400).json({ success: false, message: '氏名は必須です' });

  const existing = id ? await dbGet('SELECT home_address, home_lat, home_lng, area_fixed_enabled, area_fixed_pattern, area_fixed_store_ids, fixed_store_id FROM drivers WHERE id = ?', [id]) : null;

  // 住所が変わった場合のみ再ジオコーディングする(毎回叩くと無駄なため)
  let lat = null, lng = null;
  if (home_address) {
    if (existing && existing.home_address === home_address && existing.home_lat != null) {
      lat = existing.home_lat; lng = existing.home_lng;
    } else {
      const geo = await geocodeAddress(home_address);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }
  }

  // エリア固定の設定はリクエストに含まれていれば更新し、含まれていなければ(一括編集など、この項目を
  // 扱わない呼び出し元から送られてきた場合)既存の値をそのまま維持する(意図せず消してしまわないように)
  const finalAreaFixedEnabled = area_fixed_enabled !== undefined ? (area_fixed_enabled ? 1 : 0) : (existing ? existing.area_fixed_enabled : 0);
  const finalAreaFixedPattern = area_fixed_pattern !== undefined ? JSON.stringify(area_fixed_pattern || {}) : (existing ? existing.area_fixed_pattern : '{}');
  const finalAreaFixedStoreIds = area_fixed_store_ids !== undefined ? JSON.stringify(area_fixed_store_ids || []) : (existing ? existing.area_fixed_store_ids : '[]');

  const now = new Date().toISOString();
  if (id) {
    const newFixedStoreId = fixed_store_id || null;
    if (existing && existing.fixed_store_id !== newFixedStoreId) {
      await logFixedStoreChange(id, existing.fixed_store_id, newFixedStoreId, 'manual');
    }
    await dbRun(
      `UPDATE drivers SET name=?, phone=?, vehicle_type=?, home_address=?, home_lat=?, home_lng=?, driver_code=?, company_name=?, fixed_store_id=?, first_contract_date=?, status=?, email=?, insurance_info=?, notes=?, area_fixed_enabled=?, area_fixed_pattern=?, area_fixed_store_ids=? WHERE id=?`,
      [name, phone || '', vehicle_type || '', home_address || '', lat, lng, driver_code || null, company_name || '', newFixedStoreId, first_contract_date || '', status || '', email || '', insurance_info || '', notes || '', finalAreaFixedEnabled, finalAreaFixedPattern, finalAreaFixedStoreIds, id]
    );
    return res.json({ success: true, id });
  }
  const result = await dbRun(
    `INSERT INTO drivers (name, phone, vehicle_type, home_address, home_lat, home_lng, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes, area_fixed_enabled, area_fixed_pattern, area_fixed_store_ids, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [name, phone || '', vehicle_type || '', home_address || '', lat, lng, driver_code || null, company_name || '', fixed_store_id || null, first_contract_date || '', status || '', email || '', insurance_info || '', notes || '', finalAreaFixedEnabled, finalAreaFixedPattern, finalAreaFixedStoreIds, now]
  );
  if (fixed_store_id) await logFixedStoreChange(result.lastID, null, fixed_store_id, 'manual');
  res.json({ success: true, id: result.lastID });
});

// 固定希望店舗の変更履歴を返す(手動設定か、どの自動反映によるものかを確認できるようにする)
app.get('/api/logs/fixed-store-changes', async (req, res) => {
  const rows = await dbAll(`
    SELECT l.*, d.name AS driver_name, os.name AS old_store_name, ns.name AS new_store_name
    FROM fixed_store_change_log l
    JOIN drivers d ON d.id = l.driver_id
    LEFT JOIN stores os ON os.id = l.old_store_id
    LEFT JOIN stores ns ON ns.id = l.new_store_id
    ORDER BY l.id DESC
    LIMIT 200
  `);
  res.json({ success: true, logs: rows });
});

// ドライバーマスタをCSV/Excelでまとめて取込む(初期データ投入用。社員コードが一致すればそれで、無ければ氏名で既存ドライバーと照合し上書き更新する)
app.post('/api/drivers/import', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, message: 'ファイルが必要です' });
  let rows;
  try {
    rows = parseUploadedSpreadsheet(req.file.buffer, req.file.originalname);
  } catch (e) {
    return res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }

  const now = new Date().toISOString();
  let imported = 0;
  const errors = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const name = pickField(row, DRIVER_MASTER_ALIASES.name);
    if (!name) { errors.push(`${i + 2}行目: 氏名が読み取れませんでした`); continue; }

    const driver_code = pickField(row, DRIVER_MASTER_ALIASES.driver_code) || null;
    const company_name = pickField(row, DRIVER_MASTER_ALIASES.company_name);
    const phone = pickField(row, DRIVER_MASTER_ALIASES.phone);
    const vehicle_type = pickField(row, DRIVER_MASTER_ALIASES.vehicle_type);
    const home_address = pickField(row, DRIVER_MASTER_ALIASES.home_address);
    const first_contract_date = normalizeDateStr(pickField(row, DRIVER_MASTER_ALIASES.first_contract_date));
    const status = pickField(row, DRIVER_MASTER_ALIASES.status);
    const email = pickField(row, DRIVER_MASTER_ALIASES.email);
    const insurance_info = pickField(row, DRIVER_MASTER_ALIASES.insurance_info);
    const notes = pickField(row, DRIVER_MASTER_ALIASES.notes);
    const fixed_store_name = pickField(row, DRIVER_MASTER_ALIASES.fixed_store_name);
    const fixed_store_id = fixed_store_name ? await resolveStoreId({ store_name: fixed_store_name }) : null;

    let lat = null, lng = null;
    if (home_address) {
      const geo = await geocodeAddress(home_address);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }

    // 社員コードが一致すればそれを優先、無ければ氏名で既存ドライバーと照合する(店舗マスタと同様の名寄せ方針)
    const existing = driver_code
      ? await dbGet('SELECT id FROM drivers WHERE driver_code = ?', [driver_code])
      : await dbGet('SELECT id FROM drivers WHERE name = ?', [name]);

    if (existing) {
      await dbRun(
        `UPDATE drivers SET name=?, phone=?, vehicle_type=?, home_address=?, home_lat=?, home_lng=?, driver_code=?, company_name=?, fixed_store_id=?, first_contract_date=?, status=?, email=?, insurance_info=?, notes=? WHERE id=?`,
        [name, phone, vehicle_type, home_address, lat, lng, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes, existing.id]
      );
    } else {
      await dbRun(
        `INSERT INTO drivers (name, phone, vehicle_type, home_address, home_lat, home_lng, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [name, phone, vehicle_type, home_address, lat, lng, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes, now]
      );
    }
    imported++;
  }

  res.json({ success: true, imported, total: rows.length, errors });
});

app.delete('/api/drivers/:id', async (req, res) => {
  await dbRun('DELETE FROM drivers WHERE id = ?', [req.params.id]);
  // 過去の月としてアーカイブ済みの希望シフトは、ドライバーがマスタから削除されても履歴として残す
  // (archived_month IS NULLの、現在進行中の希望シフトのみ削除する)
  await dbRun('DELETE FROM driver_availability WHERE driver_id = ? AND archived_month IS NULL', [req.params.id]);
  res.json({ success: true });
});

// ===== ドライバーの希望シフト =====
app.get('/api/driver-availability', async (req, res) => {
  const rows = await dbAll(`
    SELECT a.*, d.name AS driver_name, d.phone AS driver_phone
    FROM driver_availability a JOIN drivers d ON d.id = a.driver_id
    WHERE a.archived_month IS NULL
    ORDER BY a.desired_date DESC, a.id DESC
  `);
  res.json({ success: true, availability: rows });
});

app.post('/api/driver-availability', async (req, res) => {
  const { id, driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests } = req.body;
  if (!driver_id || !desired_date) return res.status(400).json({ success: false, message: 'driver_id, desired_date は必須です' });
  const now = new Date().toISOString();
  if (id) {
    await dbRun(
      `UPDATE driver_availability SET driver_id=?, desired_date=?, desired_area=?, desired_store=?, time_start=?, time_end=?, requests=? WHERE id=?`,
      [driver_id, desired_date, desired_area || '', desired_store || '', time_start || '', time_end || '', requests || '', id]
    );
    return res.json({ success: true, id });
  }
  // 新規作成時、希望店舗/エリアが未入力なら固定希望店舗を初期値として使う(一覧上も実態のマッチング挙動と一致させる)
  let finalArea = desired_area || '', finalStore = desired_store || '';
  if (!finalArea && !finalStore) {
    const fixedStore = await dbGet(
      'SELECT fs.name, fs.area FROM drivers d LEFT JOIN stores fs ON fs.id = d.fixed_store_id WHERE d.id = ?',
      [driver_id]
    );
    if (fixedStore && fixedStore.name) { finalStore = fixedStore.name; finalArea = fixedStore.area || ''; }
  }
  const result = await dbRun(
    `INSERT INTO driver_availability (driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
    [driver_id, desired_date, finalArea, finalStore, time_start || '', time_end || '', requests || '', now]
  );
  res.json({ success: true, id: result.lastID });
});

app.delete('/api/driver-availability/:id', async (req, res) => {
  await dbRun('DELETE FROM driver_availability WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// 楽シフのエクスポート(CSV/Excel)を取り込む想定。列名は実際の形式が分かり次第AVAILABILITY_ALIASESに追記する。
// ドライバー名はドライバーマスタと名前で一致させるため、先にドライバー登録が済んでいる必要がある
app.post('/api/driver-availability/import', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, message: 'ファイルが必要です' });
  let rows;
  try {
    rows = parseUploadedSpreadsheet(req.file.buffer, req.file.originalname);
  } catch (e) {
    return res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }

  const drivers = await dbAll(`
    SELECT d.id, d.name, fs.name AS fixed_store_name, fs.area AS fixed_store_area
    FROM drivers d LEFT JOIN stores fs ON fs.id = d.fixed_store_id
  `);
  const driverByName = new Map(drivers.map(d => [d.name.trim(), d]));

  const now = new Date().toISOString();
  let imported = 0;
  const errors = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const driverName = pickField(row, AVAILABILITY_ALIASES.driver_name);
    const desired_date = normalizeDateStr(pickField(row, AVAILABILITY_ALIASES.desired_date));
    if (!driverName || !desired_date) { errors.push(`${i + 2}行目: ドライバー名または希望日が読み取れませんでした`); continue; }

    const driver = driverByName.get(driverName);
    if (!driver) { errors.push(`${i + 2}行目: ドライバー「${driverName}」がドライバーマスタに見つかりません(先に登録してください)`); continue; }
    const driver_id = driver.id;

    // ファイルに希望店舗/エリアの記載がなければ、固定希望店舗を初期値として使う
    let desired_area = pickField(row, AVAILABILITY_ALIASES.desired_area);
    let desired_store = pickField(row, AVAILABILITY_ALIASES.desired_store);
    if (!desired_area && !desired_store && driver.fixed_store_name) {
      desired_store = driver.fixed_store_name;
      desired_area = driver.fixed_store_area || '';
    }
    const time_start = pickField(row, AVAILABILITY_ALIASES.time_start);
    const time_end = pickField(row, AVAILABILITY_ALIASES.time_end);
    const requests = pickField(row, AVAILABILITY_ALIASES.requests);

    // 同じドライバー×同じ希望日の行が既にあれば上書き更新する(同じファイルの再取込みで重複登録されないように)
    const existing = await dbGet(
      'SELECT id FROM driver_availability WHERE driver_id = ? AND desired_date = ? AND archived_month IS NULL',
      [driver_id, desired_date]
    );
    if (existing) {
      await dbRun(
        `UPDATE driver_availability SET desired_area=?, desired_store=?, time_start=?, time_end=?, requests=? WHERE id=?`,
        [desired_area, desired_store, time_start, time_end, requests, existing.id]
      );
    } else {
      await dbRun(
        `INSERT INTO driver_availability (driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        [driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests, now]
      );
    }
    imported++;
  }

  res.json({ success: true, imported, total: rows.length, errors });
});

// アップロードされたExcelファイルのシート名一覧を返す(CSVは単一シート扱いで空配列)。
// 1ファイルに複数シートがある場合に、取込み前にどのシートを使うか選べるようにするため
app.post('/api/sheet-names', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, message: 'ファイルが必要です' });
  const ext = (req.file.originalname || '').toLowerCase();
  if (ext.endsWith('.csv') || ext.endsWith('.txt')) return res.json({ success: true, sheets: [] });
  try {
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer', bookSheets: true });
    res.json({ success: true, sheets: workbook.SheetNames });
  } catch (e) {
    res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }
});

// filenameを「uploads/」フォルダ直下のファイル名に限定して安全な絶対パスにする(..などでの脱出を防ぐ)。
// 存在しない・フォルダ外を指す場合はnullを返す
function resolveIncomingFilePath(filename) {
  if (!filename) return null;
  const safeName = path.basename(String(filename));
  const fullPath = path.join(UPLOADS_DIR, safeName);
  if (path.dirname(fullPath) !== UPLOADS_DIR) return null;
  if (!fs.existsSync(fullPath) || !fs.statSync(fullPath).isFile()) return null;
  return fullPath;
}

// 所定フォルダ(uploads/)に置かれているファイルの一覧を返す(名前・サイズ・更新日時、更新が新しい順)
app.get('/api/incoming-files', (req, res) => {
  const files = fs.readdirSync(UPLOADS_DIR)
    .filter(name => INCOMING_FILE_EXTENSIONS.includes(path.extname(name).toLowerCase()))
    .map(name => {
      const stat = fs.statSync(path.join(UPLOADS_DIR, name));
      return { name, size: stat.size, mtime: stat.mtime.toISOString() };
    })
    .sort((a, b) => b.mtime.localeCompare(a.mtime));
  res.json({ success: true, files });
});

// 各種取込みエンドポイントで、multerでのアップロードと所定フォルダ(uploads/)内のファイル名指定の
// どちらからでも取込み元を得られるようにする共通ヘルパー。{buffer, filename}を返す。どちらも無ければnull
function getImportSource(req) {
  if (req.file) return { buffer: req.file.buffer, filename: req.file.originalname };
  const fullPath = resolveIncomingFilePath(req.body.source_filename);
  if (fullPath) return { buffer: fs.readFileSync(fullPath), filename: path.basename(fullPath) };
  return null;
}

// 所定フォルダ内の指定ファイルのシート名一覧を返す(/api/sheet-namesの、アップロードでなくファイル名指定版)
app.post('/api/incoming-files/sheet-names', (req, res) => {
  const fullPath = resolveIncomingFilePath(req.body.filename);
  if (!fullPath) return res.status(404).json({ success: false, message: 'ファイルが見つかりません' });
  const ext = path.extname(fullPath).toLowerCase();
  if (ext === '.csv' || ext === '.txt') return res.json({ success: true, sheets: [] });
  try {
    const workbook = XLSX.read(fs.readFileSync(fullPath), { type: 'buffer', bookSheets: true });
    res.json({ success: true, sheets: workbook.SheetNames });
  } catch (e) {
    res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }
});

// ドライバーの月間シフト表を「氏名×日付」のワイド形式(1行=1人、列=日付見出し「1(火)」等、
// セル=勤務時間「10:00〜22:00」or「休み」or 空欄)で取込む。年月はファイルに含まれないためフォームで指定してもらう。
// 空欄・「休み」のセルは希望シフトを作らない(その日は稼働しない扱い)
app.post('/api/driver-availability/import-wide', upload.single('file'), async (req, res) => {
  const source = getImportSource(req);
  if (!source) return res.status(400).json({ success: false, message: 'ファイルが必要です(アップロードするか、所定フォルダのファイルを指定してください)' });
  const year = parseInt(req.body.year, 10);
  const month = parseInt(req.body.month, 10);
  if (!year || !month || month < 1 || month > 12) {
    return res.status(400).json({ success: false, message: '対象年月を指定してください' });
  }

  let rows;
  try {
    rows = parseWideSpreadsheet(source.buffer, source.filename, req.body.sheet);
  } catch (e) {
    return res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }
  if (rows.length < 2) return res.status(400).json({ success: false, message: 'データ行が見つかりませんでした' });

  const header = rows[0];
  // 「1(火)」のような見出しから日付の数字部分だけを取り出す
  const dayColumns = [];
  for (let c = 1; c < header.length; c++) {
    const m = String(header[c] ?? '').match(/^(\d{1,2})/);
    if (m) dayColumns.push({ colIndex: c, day: parseInt(m[1], 10) });
  }
  if (dayColumns.length === 0) {
    return res.status(400).json({ success: false, message: '日付の列見出しが読み取れませんでした(例: 「1(火)」のような形式を想定しています)' });
  }

  const drivers = await dbAll(`
    SELECT d.id, d.name, fs.name AS fixed_store_name, fs.area AS fixed_store_area
    FROM drivers d LEFT JOIN stores fs ON fs.id = d.fixed_store_id
  `);
  const driverByName = new Map(drivers.map(d => [d.name.trim(), d]));

  let imported = 0;
  const errors = [];
  const now = new Date().toISOString();

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const driverName = String(row[0] ?? '').trim();
    if (!driverName) continue; // 空行はスキップ

    const driver = driverByName.get(driverName);
    if (!driver) { errors.push(`${r + 1}行目: ドライバー「${driverName}」がドライバーマスタに見つかりません(先に登録してください)`); continue; }
    const driver_id = driver.id;
    // 希望シフト自体には店舗/エリアの列がないため、固定希望店舗が設定済みならその値を初期値として入れておく
    // (マッチングロジックは元々この値を動的にフォールバック利用していたが、一覧画面でも実態が見えるようにする)
    const fallbackStore = driver.fixed_store_name || '';
    const fallbackArea = driver.fixed_store_area || '';

    for (const { colIndex, day } of dayColumns) {
      const cell = String(row[colIndex] ?? '').trim();
      if (!cell || cell === '休み') continue; // 空欄・休みは希望シフトを作らない

      const timeMatch = cell.match(/^(\d{1,2}:\d{2})\s*[-－‐–—〜～~ー―−]\s*(\d{1,2}:\d{2})$/);
      if (!timeMatch) { errors.push(`${r + 1}行目 ${driverName} ${day}日: 「${cell}」を勤務時間として読み取れませんでした`); continue; }

      const desired_date = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const time_start = timeMatch[1];
      const time_end = timeMatch[2];

      const existing = await dbGet(
        'SELECT id, desired_store, desired_area FROM driver_availability WHERE driver_id = ? AND desired_date = ? AND archived_month IS NULL',
        [driver_id, desired_date]
      );
      if (existing) {
        // 既に希望店舗/エリアが入っている(LINE取込みや手動編集による本人の実際の希望)場合は上書きしない
        const desired_store = existing.desired_store || fallbackStore;
        const desired_area = existing.desired_area || fallbackArea;
        await dbRun(
          'UPDATE driver_availability SET time_start=?, time_end=?, desired_store=?, desired_area=? WHERE id=?',
          [time_start, time_end, desired_store, desired_area, existing.id]
        );
      } else {
        await dbRun(
          `INSERT INTO driver_availability (driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
          [driver_id, desired_date, fallbackArea, fallbackStore, time_start, time_end, '', now]
        );
      }
      imported++;
    }
  }

  res.json({ success: true, imported, errors });
});

// LINEなどで届く自由文の希望シフトメッセージをAI(Claude)で解析する(試験的機能)。
// 既存の正規表現ベースの解析(parseLineShiftText、クライアント側)と同じ形の結果を返し、
// どちらで解析しても同じプレビュー・確認画面を経由してから登録される(AIの解析結果を無条件に信用しない)
app.post('/api/ai/parse-line-shift', async (req, res) => {
  if (!anthropic) return res.status(400).json({ success: false, message: 'AI解析は未設定です(.envにANTHROPIC_API_KEYを設定してサーバーを再起動してください)' });
  const { text, year, month } = req.body;
  if (!text || !year || !month) return res.status(400).json({ success: false, message: 'text, year, month は必須です' });

  try {
    const response = await anthropic.messages.create({
      model: AI_MODEL,
      max_tokens: 2000,
      tools: [{
        name: 'extract_shift_rows',
        description: '個人事業主がLINEで送ってきた希望シフトの自由文から、勤務可能な日付ごとの時間帯を抽出する',
        input_schema: {
          type: 'object',
          properties: {
            rows: {
              type: 'array',
              description: '休みの日・出勤しない日は含めない(登録対象は勤務可能な日のみ)',
              items: {
                type: 'object',
                properties: {
                  desired_date: { type: 'string', description: `YYYY-MM-DD形式。年は${year}年、月の記載が省略されている行は${month}月として補う` },
                  time_start: { type: 'string', description: 'HH:MM形式。読み取れなければ空文字' },
                  time_end: { type: 'string', description: 'HH:MM形式。読み取れなければ空文字' },
                  raw: { type: 'string', description: '元のメッセージのうち、この行に対応する部分の抜粋' },
                  unrecognized: { type: 'boolean', description: '日付や時間の解釈に自信が持てない場合はtrue' }
                },
                required: ['desired_date', 'time_start', 'time_end', 'raw', 'unrecognized']
              }
            }
          },
          required: ['rows']
        }
      }],
      tool_choice: { type: 'tool', name: 'extract_shift_rows' },
      messages: [{
        role: 'user',
        content: `以下はドライバー(個人事業主)がLINEで送ってきた希望シフトのメッセージです。勤務可能な日付と時間帯を抽出してください。\n\n---\n${text}\n---`
      }]
    });
    const toolUse = response.content.find(c => c.type === 'tool_use');
    res.json({ success: true, rows: (toolUse && toolUse.input.rows) || [] });
  } catch (e) {
    res.status(500).json({ success: false, message: 'AI解析に失敗しました: ' + e.message });
  }
});

// ===== 店舗からの人員要請 =====
app.get('/api/store-requests', async (req, res) => {
  const rows = await dbAll('SELECT * FROM store_requests WHERE archived_month IS NULL ORDER BY request_date DESC, id DESC');
  res.json({ success: true, requests: rows });
});

app.post('/api/store-requests', async (req, res) => {
  const { id, store_name, area, address, request_date, time_start, time_end, required_count, requests } = req.body;
  if (!store_name || !request_date) return res.status(400).json({ success: false, message: 'store_name, request_date は必須です' });

  let finalArea = area || '', finalAddress = address || '', lat = null, lng = null;
  if (address) {
    const existing = id ? await dbGet('SELECT address, lat, lng FROM store_requests WHERE id = ?', [id]) : null;
    if (existing && existing.address === address && existing.lat != null) {
      lat = existing.lat; lng = existing.lng;
    } else {
      const geo = await geocodeAddress(address);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }
  }
  // 店舗名は自由入力のまま、裏側で店舗マスタに名寄せしておく(相性・派遣履歴を店舗単位で扱うため)
  const store_id = await resolveStoreId({ store_name, area, address, lat, lng });

  // 住所が未入力の場合は、店舗マスタに登録済みの住所を自動で補完する(距離計算ができるように)
  if (!address) {
    const storeInfo = await getStoreMasterInfo(store_id);
    if (storeInfo && storeInfo.address) {
      finalArea = finalArea || storeInfo.area || '';
      finalAddress = storeInfo.address;
      lat = storeInfo.lat;
      lng = storeInfo.lng;
    }
  }

  const now = new Date().toISOString();
  if (id) {
    await dbRun(
      `UPDATE store_requests SET store_name=?, area=?, address=?, lat=?, lng=?, request_date=?, time_start=?, time_end=?, required_count=?, requests=?, store_id=? WHERE id=?`,
      [store_name, finalArea, finalAddress, lat, lng, request_date, time_start || '', time_end || '', required_count || 1, requests || '', store_id, id]
    );
    return res.json({ success: true, id });
  }
  const result = await dbRun(
    `INSERT INTO store_requests (store_name, area, address, lat, lng, request_date, time_start, time_end, required_count, requests, created_at, store_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [store_name, finalArea, finalAddress, lat, lng, request_date, time_start || '', time_end || '', required_count || 1, requests || '', now, store_id]
  );
  res.json({ success: true, id: result.lastID });
});

app.delete('/api/store-requests/:id', async (req, res) => {
  await dbRun('DELETE FROM store_requests WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// 店舗からのExcel依頼を取り込む想定。列名は実際の形式が分かり次第STORE_ALIASESに追記する
app.post('/api/store-requests/import', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, message: 'ファイルが必要です' });
  let rows;
  try {
    rows = parseUploadedSpreadsheet(req.file.buffer, req.file.originalname);
  } catch (e) {
    return res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }

  const now = new Date().toISOString();
  let imported = 0;
  const errors = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const store_name = pickField(row, STORE_ALIASES.store_name);
    const request_date = normalizeDateStr(pickField(row, STORE_ALIASES.request_date));
    if (!store_name || !request_date) { errors.push(`${i + 2}行目: 店舗名または依頼日が読み取れませんでした`); continue; }

    let address = pickField(row, STORE_ALIASES.address);
    let area = pickField(row, STORE_ALIASES.area);
    let lat = null, lng = null;
    if (address) {
      const geo = await geocodeAddress(address);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }
    const store_id = await resolveStoreId({ store_name, area, address, lat, lng });

    // 住所が未入力の場合は、店舗マスタに登録済みの住所を自動で補完する(距離計算ができるように)
    if (!address) {
      const storeInfo = await getStoreMasterInfo(store_id);
      if (storeInfo && storeInfo.address) {
        area = area || storeInfo.area || '';
        address = storeInfo.address;
        lat = storeInfo.lat;
        lng = storeInfo.lng;
      }
    }

    const time_start = pickField(row, STORE_ALIASES.time_start);
    const time_end = pickField(row, STORE_ALIASES.time_end);
    const required_count = parseInt(pickField(row, STORE_ALIASES.required_count), 10) || 1;
    const requests = pickField(row, STORE_ALIASES.requests);

    // 同じ店舗×同じ依頼日×同じ開始時刻の行が既にあれば上書き更新する(同じファイルの再取込みで重複登録されないように。
    // 1日に時間帯違いで複数依頼が来る店舗もあるため、開始時刻もキーに含めて別依頼として区別する)
    const existing = await dbGet(
      'SELECT id FROM store_requests WHERE store_id = ? AND request_date = ? AND time_start = ? AND archived_month IS NULL',
      [store_id, request_date, time_start || '']
    );
    if (existing) {
      await dbRun(
        `UPDATE store_requests SET store_name=?, area=?, address=?, lat=?, lng=?, time_end=?, required_count=?, requests=? WHERE id=?`,
        [store_name, area, address, lat, lng, time_end, required_count, requests, existing.id]
      );
    } else {
      await dbRun(
        `INSERT INTO store_requests (store_name, area, address, lat, lng, request_date, time_start, time_end, required_count, requests, created_at, store_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [store_name, area, address, lat, lng, request_date, time_start, time_end, required_count, requests, now, store_id]
      );
    }
    imported++;
  }

  res.json({ success: true, imported, total: rows.length, errors });
});

// 店舗の「曜日ごとの週間必要枠」表(1行=1人分の必要枠。同じ店舗が複数行あれば必要人数として合算する。
// セルに「10-22*2」のように*(全角＊も可)+数字が付いていれば、その曜日だけ人数分(例:2人)として数える)を取込む。
// 列は固定位置ではなく見出しのテキスト(店舗名/月/火/水/木/金/土/日/フラグ/備考)で探すため、
// 見出し行がどこにあっても(先頭にメモ行が入っていても)対応できる。
// 年月はファイルに含まれないためフォームで指定してもらい、対象月の該当曜日すべてに展開して店舗依頼を作る。
// 備考に社員番号(5〜7桁の数字)らしき記載があれば、その社員のドライバーマスタの「固定希望店舗」にこの店舗を自動設定する
app.post('/api/store-requests/import-weekly', upload.single('file'), async (req, res) => {
  const source = getImportSource(req);
  if (!source) return res.status(400).json({ success: false, message: 'ファイルが必要です(アップロードするか、所定フォルダのファイルを指定してください)' });
  const year = parseInt(req.body.year, 10);
  const month = parseInt(req.body.month, 10);
  if (!year || !month || month < 1 || month > 12) {
    return res.status(400).json({ success: false, message: '対象年月を指定してください' });
  }

  let rows;
  try {
    rows = parseWideSpreadsheet(source.buffer, source.filename, req.body.sheet);
  } catch (e) {
    return res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }

  // 見出し行を「店舗名」という見出しが含まれる行として自動検出する(先頭にメモ行が挿入されていることがあるため)
  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    if (rows[i].some(cell => String(cell ?? '').trim() === '店舗名')) { headerRowIndex = i; break; }
  }
  if (headerRowIndex === -1) {
    return res.status(400).json({ success: false, message: '見出し行(「店舗名」の列)が見つかりませんでした' });
  }
  const header = rows[headerRowIndex].map(h => String(h ?? '').trim());
  const colOf = (label) => header.indexOf(label);
  const colStoreName = colOf('店舗名');
  const colFlag = colOf('フラグ');
  const colNotes = colOf('備考');
  const WEEKDAY_LABELS = ['月', '火', '水', '木', '金', '土', '日'];
  const WEEKDAY_JSDAY = [1, 2, 3, 4, 5, 6, 0]; // Date.getDay()に合わせる(0=日,1=月,...,6=土)
  const weekdayCols = WEEKDAY_LABELS.map(colOf);

  if (weekdayCols.some(c => c === -1)) {
    return res.status(400).json({ success: false, message: '曜日(月〜日)の列が見つかりませんでした' });
  }

  const drivers = await dbAll('SELECT id, driver_code, fixed_store_id FROM drivers');
  // 社員番号の先頭0の有無(「540094」と「00540094」等)の表記ゆれを吸収するため、先頭0を除いた形をキーにする
  const normalizeEmployeeCode = (code) => String(code || '').trim().replace(/^0+(?=\d)/, '');
  const driverByCode = new Map(drivers.filter(d => d.driver_code).map(d => [normalizeEmployeeCode(d.driver_code), d.id]));
  const currentFixedStoreById = new Map(drivers.map(d => [d.id, d.fixed_store_id]));
  // 備考の社員番号から固定希望店舗を自動設定する際、その店舗が既にNG設定されていれば矛盾するため設定しない
  const ngPairs = await dbAll(`SELECT driver_id, store_id FROM driver_store_preferences WHERE preference = 'NG'`);
  const ngPairSet = new Set(ngPairs.map(p => `${p.driver_id}:${p.store_id}`));

  // 対象月の日付を曜日ごとにまとめておく(この曜日は月内のこの日付たち、という対応表)
  const daysInMonth = new Date(year, month, 0).getDate();
  const datesByJsDay = {};
  for (let d = 1; d <= daysInMonth; d++) {
    const jsDay = new Date(year, month - 1, d).getDay();
    (datesByJsDay[jsDay] = datesByJsDay[jsDay] || []).push(d);
  }

  const errors = [];
  let fixedDriverLinks = 0;

  // 1周目: 同じ店舗×同じ曜日×同じ時間帯の行を合算し、必要人数を数える(1行=1人分のため)
  const slotMap = new Map(); // "storeId|weekdayIndex|time_start|time_end" -> { count, storeName, notesSet }
  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r];
    const store_name = String(row[colStoreName] ?? '').trim();
    if (!store_name) continue;

    const flag = colFlag !== -1 ? String(row[colFlag] ?? '').trim() : '';
    const notesRaw = colNotes !== -1 ? String(row[colNotes] ?? '').trim() : '';
    const note = flag ? `(${flag}) ${notesRaw}`.trim() : notesRaw;

    const store_id = await resolveStoreId({ store_name });
    // 店舗マスタに住所が登録されていれば、それを店舗依頼側にも使う(距離計算ができるように)
    const storeInfo = await getStoreMasterInfo(store_id);

    // 備考に社員番号らしき数字があれば、その社員の固定希望店舗をこの店舗に設定する
    // (ただし、その店舗が店舗相性マスタで既にNGに設定されている場合は矛盾するため設定しない)
    const codeMatch = notesRaw.match(/(\d{5,8})/);
    const normalizedCode = codeMatch ? normalizeEmployeeCode(codeMatch[1]) : null;
    if (normalizedCode && driverByCode.has(normalizedCode)) {
      const matchedDriverId = driverByCode.get(normalizedCode);
      if (ngPairSet.has(`${matchedDriverId}:${store_id}`)) {
        errors.push(`${r + 1}行目「${store_name}」: 備考の社員番号(${codeMatch[1]})は固定希望店舗の対象ですが、この店舗は店舗相性マスタでNGに設定されているため、固定希望店舗には反映しませんでした(手動で確認してください)`);
      } else {
        const oldFixedStoreId = currentFixedStoreById.get(matchedDriverId);
        await dbRun('UPDATE drivers SET fixed_store_id = ? WHERE id = ?', [store_id, matchedDriverId]);
        if (oldFixedStoreId !== store_id) {
          await logFixedStoreChange(matchedDriverId, oldFixedStoreId, store_id, 'auto_remarks');
          currentFixedStoreById.set(matchedDriverId, store_id); // 同じ取込み内で複数回ヒットしても重複記録しないように
        }
        fixedDriverLinks++;
      }
    }

    let hasAnySchedule = false;
    for (let w = 0; w < WEEKDAY_LABELS.length; w++) {
      const cell = String(row[weekdayCols[w]] ?? '').trim();
      if (!cell) continue;
      const timeMatch = cell.match(/(\d{1,2})[:時]?(\d{2})?\s*[-－‐–—〜～~ー―−]\s*(\d{1,2})[:時]?(\d{2})?/);
      if (!timeMatch) continue;
      hasAnySchedule = true;
      const time_start = `${timeMatch[1].padStart(2, '0')}:${(timeMatch[2] || '00').padStart(2, '0')}`;
      const time_end = `${timeMatch[3].padStart(2, '0')}:${(timeMatch[4] || '00').padStart(2, '0')}`;
      // セルに「*2」「＊3」のような倍数指定があれば、その人数分を必要枠として数える(無ければ1人分)
      const multiplierMatch = cell.match(/[*＊]\s*(\d+)/);
      const multiplier = multiplierMatch ? parseInt(multiplierMatch[1], 10) : 1;

      const key = `${store_id}|${w}|${time_start}|${time_end}`;
      const entry = slotMap.get(key) || {
        count: 0, storeName: store_name, notesSet: new Set(), weekdayIndex: w, time_start, time_end, storeId: store_id,
        area: storeInfo ? (storeInfo.area || '') : '', address: storeInfo ? (storeInfo.address || '') : '',
        lat: storeInfo ? storeInfo.lat : null, lng: storeInfo ? storeInfo.lng : null
      };
      entry.count += multiplier;
      if (note) entry.notesSet.add(note);
      slotMap.set(key, entry);
    }

    if (!hasAnySchedule && notesRaw) {
      let message = `${r + 1}行目「${store_name}」: 曜日の時間帯が読み取れず、備考「${notesRaw}」があるため要確認です(手動で確認してください)`;
      const aiGuess = await interpretNoteWithAI(notesRaw, year, month);
      if (aiGuess) {
        message += ` ／ 🤖AI解釈(参考・自動登録はされていません): ${aiGuess.desired_date || '?'} ${aiGuess.time_start || '?'}〜${aiGuess.time_end || '?'} ${aiGuess.required_count}名`;
      }
      errors.push(message);
    }
  }

  // 2周目: 曜日×時間帯の枠を、対象月の実際の日付に展開して店舗依頼を作る(既存の重複防止と同じキーで上書き更新)
  const now = new Date().toISOString();
  let imported = 0;
  for (const entry of slotMap.values()) {
    const jsDay = WEEKDAY_JSDAY[entry.weekdayIndex];
    const requests = [...entry.notesSet].join(' / ');
    for (const day of (datesByJsDay[jsDay] || [])) {
      const request_date = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const existing = await dbGet(
        'SELECT id FROM store_requests WHERE store_id = ? AND request_date = ? AND time_start = ? AND archived_month IS NULL',
        [entry.storeId, request_date, entry.time_start]
      );
      if (existing) {
        await dbRun(
          `UPDATE store_requests SET store_name=?, area=?, address=?, lat=?, lng=?, time_end=?, required_count=?, requests=? WHERE id=?`,
          [entry.storeName, entry.area, entry.address, entry.lat, entry.lng, entry.time_end, entry.count, requests, existing.id]
        );
      } else {
        await dbRun(
          `INSERT INTO store_requests (store_name, area, address, lat, lng, request_date, time_start, time_end, required_count, requests, created_at, store_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          [entry.storeName, entry.area, entry.address, entry.lat, entry.lng, request_date, entry.time_start, entry.time_end, entry.count, requests, now, entry.storeId]
        );
      }
      imported++;
    }
  }

  res.json({ success: true, imported, fixedDriverLinks, errors });
});

// ===== 店舗マスタ =====
app.get('/api/stores', async (req, res) => {
  const rows = await dbAll('SELECT * FROM stores ORDER BY name ASC');
  res.json({ success: true, stores: rows });
});

app.post('/api/stores', async (req, res) => {
  const { id, name, area, address, manager_name, sv_name, store_code, notes } = req.body;
  if (!name) return res.status(400).json({ success: false, message: '店舗名は必須です' });

  let lat = null, lng = null;
  if (address) {
    const existing = id ? await dbGet('SELECT address, lat, lng FROM stores WHERE id = ?', [id]) : null;
    if (existing && existing.address === address && existing.lat != null) {
      lat = existing.lat; lng = existing.lng;
    } else {
      const geo = await geocodeAddress(address);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }
  }

  const now = new Date().toISOString();
  try {
    if (id) {
      await dbRun(
        `UPDATE stores SET name=?, area=?, address=?, lat=?, lng=?, manager_name=?, sv_name=?, store_code=?, notes=? WHERE id=?`,
        [name, area || '', address || '', lat, lng, manager_name || '', sv_name || '', store_code || null, notes || '', id]
      );
      return res.json({ success: true, id });
    }
    const result = await dbRun(
      `INSERT INTO stores (name, area, address, lat, lng, manager_name, sv_name, store_code, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [name, area || '', address || '', lat, lng, manager_name || '', sv_name || '', store_code || null, notes || '', now]
    );
    res.json({ success: true, id: result.lastID });
  } catch (e) {
    res.status(400).json({ success: false, message: '同じ店舗名が既に登録されています' });
  }
});

app.delete('/api/stores/:id', async (req, res) => {
  await dbRun('DELETE FROM stores WHERE id = ?', [req.params.id]);
  await dbRun('DELETE FROM driver_store_preferences WHERE store_id = ?', [req.params.id]);
  res.json({ success: true });
});

// 店舗マスタをCSV/Excelでまとめて取込む(初期データ投入用。店舗名が既存と一致する場合は住所等を更新する)
app.post('/api/stores/import', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, message: 'ファイルが必要です' });
  let rows;
  try {
    rows = parseUploadedSpreadsheet(req.file.buffer, req.file.originalname);
  } catch (e) {
    return res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }

  const now = new Date().toISOString();
  let imported = 0;
  const errors = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const name = pickField(row, STORE_MASTER_ALIASES.store_name);
    if (!name) { errors.push(`${i + 2}行目: 店舗名が読み取れませんでした`); continue; }

    const area = pickField(row, STORE_MASTER_ALIASES.area);
    const address = pickField(row, STORE_MASTER_ALIASES.address);
    const manager_name = pickField(row, STORE_MASTER_ALIASES.manager_name);
    const sv_name = pickField(row, STORE_MASTER_ALIASES.sv_name);
    const store_code = pickField(row, STORE_MASTER_ALIASES.store_code) || null;
    const notes = pickField(row, STORE_MASTER_ALIASES.notes);
    let lat = null, lng = null;
    if (address) {
      const geo = await geocodeAddress(address);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }

    const existing = await dbGet('SELECT id FROM stores WHERE name = ?', [name]);
    if (existing) {
      await dbRun('UPDATE stores SET area=?, address=?, lat=?, lng=?, manager_name=?, sv_name=?, store_code=?, notes=? WHERE id=?', [area, address, lat, lng, manager_name, sv_name, store_code, notes, existing.id]);
    } else {
      await dbRun(
        'INSERT INTO stores (name, area, address, lat, lng, manager_name, sv_name, store_code, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
        [name, area, address, lat, lng, manager_name, sv_name, store_code, notes, now]
      );
    }
    imported++;
  }

  res.json({ success: true, imported, total: rows.length, errors });
});

// ===== ドライバー×店舗の相性(好き/NG) =====
// 設定済みの相性を全件返す(ダッシュボードの統計表示用)
app.get('/api/driver-store-preferences', async (req, res) => {
  const rows = await dbAll(`
    SELECT p.*, d.name AS driver_name, s.name AS store_name, s.area
    FROM driver_store_preferences p
    JOIN drivers d ON d.id = p.driver_id
    JOIN stores s ON s.id = p.store_id
    ORDER BY p.id DESC
  `);
  res.json({ success: true, preferences: rows });
});

// 指定ドライバーについて、登録済み全店舗と現在の相性設定(未設定はnull)を一覧で返す
app.get('/api/drivers/:id/preferences', async (req, res) => {
  const rows = await dbAll(`
    SELECT s.id AS store_id, s.name AS store_name, s.area,
           p.id AS preference_id, p.preference, p.notes
    FROM stores s
    LEFT JOIN driver_store_preferences p ON p.store_id = s.id AND p.driver_id = ?
    ORDER BY s.name ASC
  `, [req.params.id]);
  res.json({ success: true, preferences: rows });
});

// 指定店舗について、登録済み全ドライバーと現在の相性設定(未設定はnull)を一覧で返す(上記の店舗版)
app.get('/api/stores/:id/preferences', async (req, res) => {
  const rows = await dbAll(`
    SELECT d.id AS driver_id, d.name AS driver_name,
           p.id AS preference_id, p.preference, p.notes
    FROM drivers d
    LEFT JOIN driver_store_preferences p ON p.driver_id = d.id AND p.store_id = ?
    ORDER BY d.name ASC
  `, [req.params.id]);
  res.json({ success: true, preferences: rows });
});

// 相性の登録・更新(driver_id+store_idの組でupsert)
app.post('/api/driver-store-preferences', async (req, res) => {
  const { driver_id, store_id, preference, notes } = req.body;
  if (!driver_id || !store_id || !['NG', '1', '2', '3', '4', '5', 'OK'].includes(preference)) {
    return res.status(400).json({ success: false, message: 'driver_id, store_id, preference(NG/1/2/3/4/5/OK) は必須です' });
  }
  const now = new Date().toISOString();
  const existing = await dbGet('SELECT id FROM driver_store_preferences WHERE driver_id = ? AND store_id = ?', [driver_id, store_id]);
  if (existing) {
    await dbRun('UPDATE driver_store_preferences SET preference=?, notes=? WHERE id=?', [preference, notes || '', existing.id]);
    return res.json({ success: true, id: existing.id });
  }
  const result = await dbRun(
    'INSERT INTO driver_store_preferences (driver_id, store_id, preference, notes, created_at) VALUES (?, ?, ?, ?, ?)',
    [driver_id, store_id, preference, notes || '', now]
  );
  res.json({ success: true, id: result.lastID });
});

// 相性設定の解除(「普通」に戻す)
app.delete('/api/driver-store-preferences/:id', async (req, res) => {
  await dbRun('DELETE FROM driver_store_preferences WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// ===== 派遣履歴(店舗の勝手・配達エリアの知見の蓄積) =====
app.get('/api/dispatch-history', async (req, res) => {
  const rows = await dbAll(`
    SELECT h.*, d.name AS driver_name, s.name AS store_name, s.area
    FROM dispatch_history h
    JOIN drivers d ON d.id = h.driver_id
    LEFT JOIN stores s ON s.id = h.store_id
    ORDER BY h.work_date DESC, h.id DESC
  `);
  res.json({ success: true, history: rows });
});

app.post('/api/dispatch-history', async (req, res) => {
  const { id, driver_id, store_name, work_date, notes } = req.body;
  if (!driver_id || !store_name || !work_date) {
    return res.status(400).json({ success: false, message: 'driver_id, store_name, work_date は必須です' });
  }
  const store_id = await resolveStoreId({ store_name });
  const now = new Date().toISOString();
  if (id) {
    await dbRun(
      'UPDATE dispatch_history SET driver_id=?, store_id=?, work_date=?, notes=? WHERE id=?',
      [driver_id, store_id, work_date, notes || '', id]
    );
    return res.json({ success: true, id });
  }
  const result = await dbRun(
    'INSERT INTO dispatch_history (driver_id, store_id, work_date, notes, created_at) VALUES (?, ?, ?, ?, ?)',
    [driver_id, store_id, work_date, notes || '', now]
  );
  res.json({ success: true, id: result.lastID });
});

app.delete('/api/dispatch-history/:id', async (req, res) => {
  await dbRun('DELETE FROM dispatch_history WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// 過去の派遣実績をCSV/Excelでまとめて取込む(アプリ運用前の実績・記憶をバックフィルする想定)
app.post('/api/dispatch-history/import', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, message: 'ファイルが必要です' });
  let rows;
  try {
    rows = parseUploadedSpreadsheet(req.file.buffer, req.file.originalname);
  } catch (e) {
    return res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }

  const drivers = await dbAll('SELECT id, name FROM drivers');
  const driverByName = new Map(drivers.map(d => [d.name.trim(), d.id]));

  const now = new Date().toISOString();
  let imported = 0;
  const errors = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const driverName = pickField(row, HISTORY_ALIASES.driver_name);
    const storeName = pickField(row, HISTORY_ALIASES.store_name);
    const work_date = normalizeDateStr(pickField(row, HISTORY_ALIASES.work_date));
    if (!driverName || !storeName || !work_date) { errors.push(`${i + 2}行目: ドライバー名・店舗名・実施日のいずれかが読み取れませんでした`); continue; }

    const driver_id = driverByName.get(driverName);
    if (!driver_id) { errors.push(`${i + 2}行目: ドライバー「${driverName}」がドライバーマスタに見つかりません(先に登録してください)`); continue; }

    const store_id = await resolveStoreId({ store_name: storeName });
    await dbRun(
      'INSERT INTO dispatch_history (driver_id, store_id, work_date, notes, created_at) VALUES (?, ?, ?, ?, ?)',
      [driver_id, store_id, work_date, pickField(row, HISTORY_ALIASES.notes), now]
    );
    imported++;
  }

  res.json({ success: true, imported, total: rows.length, errors });
});

// ===== マッチング =====
app.get('/api/matches', async (req, res) => {
  const rows = await dbAll(`
    SELECT m.*, d.name AS driver_name, d.phone AS driver_phone, d.home_address, d.driver_code,
           s.store_name, s.area, s.address AS store_address, s.time_start, s.time_end, s.requests AS store_requests
    FROM matches m
    JOIN drivers d ON d.id = m.driver_id
    JOIN store_requests s ON s.id = m.store_request_id
    WHERE m.archived_month IS NULL
    ORDER BY m.match_date DESC, m.id DESC
  `);
  res.json({ success: true, matches: rows });
});

// 自動マッチングを実行する(スコアリングつきの貪欲法):
// 各店舗依頼(日付+必要人数)ごとに、同じ日付で空いている希望を持つドライバーを探す。
// その店舗を「NG」に設定しているドライバーは候補から除外し(嫌な店舗は割り当てない)、
// 残った候補を「距離」「希望店舗/エリアの一致」「その店舗が好き」「その店舗/エリアでの過去派遣実績
// (店舗の勝手・配達エリアの知見)」を加味したスコアで並べ替えて、必要人数まで割り当てる。
// 距離がDIST_WARNING_KMを超える場合はマッチング自体は作るが「遠い」警告フラグを立てる(除外はしない)。
// 既に他のマッチングで同日確定しているドライバーは対象から外す(ダブルブッキング防止)。
app.post('/api/matches/run', async (req, res) => {
  // 一旦候補を作り直す(確定済みの運用に育ったら「候補のみ削除」に変更する想定)。アーカイブ済み(過去に月次クローズしたもの)は対象外にする
  await dbRun('DELETE FROM matches WHERE archived_month IS NULL');

  const storeRequests = await dbAll('SELECT * FROM store_requests WHERE archived_month IS NULL ORDER BY request_date ASC');
  const availabilityRows = await dbAll(`
    SELECT a.*, d.name AS driver_name, d.home_lat, d.home_lng, d.fixed_store_id, fs.name AS fixed_store_name
    FROM driver_availability a
    JOIN drivers d ON d.id = a.driver_id
    LEFT JOIN stores fs ON fs.id = d.fixed_store_id
    WHERE a.archived_month IS NULL
  `);
  // 希望シフトで店舗が未入力の場合は、ドライバーマスタの「固定希望店舗」を初期値として使う
  // desired_store_explicit: 固定希望店舗からの自動補完が入る前の、本人がその日に実際に入力した希望店舗
  // (「希望シフトで指定した店舗」の優先階層で、固定希望店舗と区別するために使う)
  const availability = availabilityRows.map(a => ({ ...a, desired_store_explicit: a.desired_store || '', desired_store: a.desired_store || a.fixed_store_name || '' }));
  // 実際に希望シフトを提出した(driver_id, date)の組を把握しておく(エリア固定の自動補完で、
  // 本人が別の希望を出している日を上書きしないようにするため)
  const explicitAvailabilitySet = new Set(availability.map(a => `${a.driver_id}|${a.desired_date}`));

  // エリア固定ドライバー(店舗を1つに固定するのではなく、曜日ごとの決まった時間帯+複数の候補店舗群の中から
  // 優先的に割り当てる人)を読み込んでおく。希望シフト未提出の日だけ、このパターンから仮の候補を作る
  const areaFixedRows = await dbAll(`
    SELECT id, name, home_lat, home_lng, area_fixed_pattern, area_fixed_store_ids
    FROM drivers WHERE area_fixed_enabled = 1
  `);
  const areaFixedDrivers = areaFixedRows.map(d => {
    let pattern = {}, storeIds = [];
    try { pattern = d.area_fixed_pattern ? JSON.parse(d.area_fixed_pattern) : {}; } catch (e) { /* ignore */ }
    try { storeIds = d.area_fixed_store_ids ? JSON.parse(d.area_fixed_store_ids) : []; } catch (e) { /* ignore */ }
    return { ...d, pattern, storeIds };
  }).filter(d => d.storeIds.length > 0);

  const { preferenceByDriverStore, storeVisitCount, areaVisitCount } = await loadScoringContext();
  const areaFixedStoreIdsByDriverId = new Map(areaFixedDrivers.map(d => [d.id, d.storeIds]));
  const tierCtx = { preferenceByDriverStore, storeVisitCount, areaVisitCount, areaFixedStoreIdsByDriverId };
  const enabledTierKeys = await getEnabledTierKeys();

  const now = new Date().toISOString();
  const assignedDriverIdsByDate = {}; // date -> Set(driver_id) 同日の重複割当を防ぐ
  for (const store of storeRequests) {
    assignedDriverIdsByDate[store.request_date] = assignedDriverIdsByDate[store.request_date] || new Set();
  }
  let createdCount = 0;
  let noCandidateCount = 0;
  const filledCountByRequestId = {}; // store_request.id -> 既に埋まった人数(優先階層で埋めた分)

  async function insertMatch(store, p) {
    const isFar = p.distance != null && p.distance > DIST_WARNING_KM;
    await dbRun(
      `INSERT INTO matches (store_request_id, driver_id, match_date, distance_km, is_far_warning, status, created_at, score, preference_flag, experience_count) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [store.id, p.driver_id, store.request_date, p.distance, isFar ? 1 : 0, '候補', now, p.score, p.preference, p.experienceCount]
    );
    assignedDriverIdsByDate[store.request_date].add(p.driver_id);
    createdCount++;
  }

  // 優先階層: 画面(マッチング設定)で有効化・並び替えされた順に、各階層に当てはまる候補を、店舗の処理順に
  // 関係なく全店舗を通して先に確保する。これをしないと、たまたま先に処理された別の店舗に本来優先される
  // べき人が取られてしまうことがある
  for (const tierKey of enabledTierKeys) {
    const tier = TIER_CATALOG[tierKey];
    if (!tier) continue;
    for (const store of storeRequests) {
      const assignedToday = assignedDriverIdsByDate[store.request_date];
      const alreadyFilled = filledCountByRequestId[store.id] || 0;
      const needed = (store.required_count || 1) - alreadyFilled;
      if (needed <= 0) continue;

      const tierCandidates = availability.filter(a =>
        a.desired_date === store.request_date &&
        !assignedToday.has(a.driver_id) &&
        tier.test(a, store, tierCtx) &&
        preferenceByDriverStore.get(`${a.driver_id}:${store.store_id}`) !== 'NG'
      ).map(a => ({ ...a, ...scoreCandidate(a, store, preferenceByDriverStore, storeVisitCount, areaVisitCount) }))
        .sort((x, y) => y.score - x.score);

      const picked = tierCandidates.slice(0, needed);
      for (const p of picked) await insertMatch(store, p);
      filledCountByRequestId[store.id] = alreadyFilled + picked.length;
    }
  }

  // 2階層目: 残りの枠を、これまで通りのスコアリング(距離・希望店舗/エリア一致・相性・経験)で埋める
  for (const store of storeRequests) {
    const assignedToday = assignedDriverIdsByDate[store.request_date];
    const alreadyFilled = filledCountByRequestId[store.id] || 0;
    const needed = (store.required_count || 1) - alreadyFilled;
    if (needed <= 0) continue;

    // 同じ日付の希望を持ち、この店舗をNGにしていないドライバーを候補にする
    const normalCandidates = availability.filter(a =>
      a.desired_date === store.request_date &&
      !assignedToday.has(a.driver_id) &&
      preferenceByDriverStore.get(`${a.driver_id}:${store.store_id}`) !== 'NG'
    );

    // エリア固定ドライバーのうち、この店舗が候補店舗に含まれ、その曜日のパターンがあり、
    // 本人が別途希望シフトを出していない人を、希望店舗=この店舗として仮の候補に加える
    // (strong-matchボーナスが働き、候補店舗の中で優先的に割り当てられるようになる)
    const weekdayKey = weekdayKeyOf(store.request_date);
    const areaFixedCandidates = [];
    for (const af of areaFixedDrivers) {
      if (!af.storeIds.includes(store.store_id)) continue;
      if (assignedToday.has(af.id)) continue;
      if (explicitAvailabilitySet.has(`${af.id}|${store.request_date}`)) continue;
      if (preferenceByDriverStore.get(`${af.id}:${store.store_id}`) === 'NG') continue;
      const timeRange = parseTimeRangeText(af.pattern[weekdayKey]);
      if (!timeRange) continue;
      areaFixedCandidates.push({
        driver_id: af.id, driver_name: af.name, home_lat: af.home_lat, home_lng: af.home_lng,
        desired_store: store.store_name, desired_area: store.area,
        time_start: timeRange.start, time_end: timeRange.end
      });
    }

    const candidates = [...normalCandidates, ...areaFixedCandidates];

    const scored = candidates
      .map(a => ({ ...a, ...scoreCandidate(a, store, preferenceByDriverStore, storeVisitCount, areaVisitCount) }))
      .sort((x, y) => y.score - x.score);

    const picked = scored.slice(0, needed);
    if (picked.length === 0 && alreadyFilled === 0) { noCandidateCount++; continue; }

    for (const p of picked) await insertMatch(store, p);
  }

  res.json({ success: true, created: createdCount, unmatched_requests: noCandidateCount });
});

// マッチングの優先順位階層の設定(カタログ全件+現在有効化・並び替え済みのキー配列)を返す
app.get('/api/settings/priority-tiers', async (req, res) => {
  const enabledKeys = await getEnabledTierKeys();
  const catalog = Object.entries(TIER_CATALOG).map(([key, t]) => ({ key, label: t.label, description: t.description }));
  res.json({ success: true, catalog, enabledKeys });
});

// 優先順位階層の設定を保存する(有効化したキーを、優先したい順に並べた配列で受け取る)
app.post('/api/settings/priority-tiers', async (req, res) => {
  const { keys } = req.body;
  if (!Array.isArray(keys)) return res.status(400).json({ success: false, message: 'keys(配列)は必須です' });
  const invalid = keys.filter(k => !TIER_CATALOG[k]);
  if (invalid.length > 0) return res.status(400).json({ success: false, message: `不明な階層キーです: ${invalid.join(', ')}` });
  const value = JSON.stringify(keys);
  const existing = await dbGet('SELECT key FROM app_settings WHERE key = ?', ['priority_tiers']);
  if (existing) await dbRun('UPDATE app_settings SET value = ? WHERE key = ?', [value, 'priority_tiers']);
  else await dbRun('INSERT INTO app_settings (key, value) VALUES (?, ?)', ['priority_tiers', value]);
  res.json({ success: true });
});

app.post('/api/matches/:id/confirm', async (req, res) => {
  await dbRun(`UPDATE matches SET status = '確定' WHERE id = ?`, [req.params.id]);
  const autoFixedStoreUpdated = await applyAutoFixedStoreFromMatches();
  res.json({ success: true, autoFixedStoreUpdated });
});

// 確定済みのマッチングを「完了」にし、派遣履歴に記録する(次回以降のマッチングで店舗/エリアの知見として使われる)
app.post('/api/matches/:id/complete', async (req, res) => {
  const match = await dbGet(`
    SELECT m.*, s.store_id FROM matches m JOIN store_requests s ON s.id = m.store_request_id WHERE m.id = ?
  `, [req.params.id]);
  if (!match) return res.status(404).json({ success: false, message: 'マッチングが見つかりません' });

  await dbRun(`UPDATE matches SET status = '完了' WHERE id = ?`, [req.params.id]);

  const already = await dbGet('SELECT id FROM dispatch_history WHERE match_id = ?', [req.params.id]);
  if (!already) {
    const now = new Date().toISOString();
    await dbRun(
      'INSERT INTO dispatch_history (driver_id, store_id, match_id, work_date, notes, created_at) VALUES (?, ?, ?, ?, ?, ?)',
      [match.driver_id, match.store_id, match.id, match.match_date, '', now]
    );
  }
  const autoFixedStoreUpdated = await applyAutoFixedStoreFromMatches();
  res.json({ success: true, autoFixedStoreUpdated });
});

// 「完了」を取り消して「確定」に戻す(押し間違い・後からの取り消し用)。あわせて自動作成された派遣履歴も削除する
app.post('/api/matches/:id/undo-complete', async (req, res) => {
  const match = await dbGet(`SELECT id, status FROM matches WHERE id = ?`, [req.params.id]);
  if (!match) return res.status(404).json({ success: false, message: 'マッチングが見つかりません' });
  if (match.status !== '完了') return res.status(400).json({ success: false, message: '「完了」のマッチングのみ取り消せます' });

  await dbRun(`UPDATE matches SET status = '確定' WHERE id = ?`, [req.params.id]);
  await dbRun('DELETE FROM dispatch_history WHERE match_id = ?', [req.params.id]);
  res.json({ success: true });
});

// 欠勤にする(削除はせず、理由付きで履歴として残す。実際には稼働していないので派遣履歴には記録しない)
app.post('/api/matches/:id/absence', async (req, res) => {
  const { reason } = req.body;
  const match = await dbGet('SELECT id FROM matches WHERE id = ?', [req.params.id]);
  if (!match) return res.status(404).json({ success: false, message: 'マッチングが見つかりません' });
  await dbRun(`UPDATE matches SET status = '欠勤', absence_reason = ? WHERE id = ?`, [reason || '', req.params.id]);
  res.json({ success: true });
});

// 欠勤以外の理由(店舗都合・ドライバー都合など)で担当を変更する。欠勤と同じく代替候補探しの対象になる
app.post('/api/matches/:id/change', async (req, res) => {
  const { reason } = req.body;
  const match = await dbGet('SELECT id FROM matches WHERE id = ?', [req.params.id]);
  if (!match) return res.status(404).json({ success: false, message: 'マッチングが見つかりません' });
  await dbRun(`UPDATE matches SET status = '変更', absence_reason = ? WHERE id = ?`, [reason || '', req.params.id]);
  res.json({ success: true });
});

// 欠勤したマッチングについて、代替候補を登録済み全ドライバーからランキングして返す
// (その日の希望シフトを提出していない人も対象にする=緊急時に幅広く探せるようにするため)
app.get('/api/matches/:id/substitutes', async (req, res) => {
  const match = await dbGet(`
    SELECT m.*, s.store_id, s.store_name, s.area, s.lat, s.lng, s.request_date
    FROM matches m JOIN store_requests s ON s.id = m.store_request_id
    WHERE m.id = ?
  `, [req.params.id]);
  if (!match) return res.status(404).json({ success: false, message: 'マッチングが見つかりません' });

  // 同日に他のマッチングで既に候補/確定/完了になっているドライバーは除外(ダブルブッキング防止。欠勤した本人も除外)
  const busyRows = await dbAll(
    `SELECT driver_id FROM matches WHERE match_date = ? AND status IN ('候補', '確定', '完了') AND archived_month IS NULL`,
    [match.match_date]
  );
  const excluded = new Set(busyRows.map(r => r.driver_id));
  excluded.add(match.driver_id);

  const drivers = await dbAll(`
    SELECT d.id, d.name, d.phone, d.home_lat, d.home_lng, fs.name AS fixed_store_name
    FROM drivers d
    LEFT JOIN stores fs ON fs.id = d.fixed_store_id
  `);
  const availabilityToday = await dbAll(
    'SELECT driver_id, desired_store, desired_area FROM driver_availability WHERE desired_date = ? AND archived_month IS NULL',
    [match.match_date]
  );
  const availabilityByDriver = new Map(availabilityToday.map(a => [a.driver_id, a]));
  const { preferenceByDriverStore, storeVisitCount, areaVisitCount } = await loadScoringContext();

  const candidates = drivers
    .filter(d => !excluded.has(d.id))
    .filter(d => preferenceByDriverStore.get(`${d.id}:${match.store_id}`) !== 'NG')
    .map(d => {
      const av = availabilityByDriver.get(d.id);
      // 希望シフトで店舗未入力(または当日の希望シフトが無い)場合は、固定希望店舗を初期値として使う
      const candidateLike = {
        driver_id: d.id, home_lat: d.home_lat, home_lng: d.home_lng,
        desired_store: (av && av.desired_store) || d.fixed_store_name || null,
        desired_area: av ? av.desired_area : null
      };
      const scored = scoreCandidate(candidateLike, match, preferenceByDriverStore, storeVisitCount, areaVisitCount);
      return {
        driver_id: d.id, name: d.name, phone: d.phone,
        has_availability_today: !!av,
        is_far_warning: scored.distance != null && scored.distance > DIST_WARNING_KM,
        ...scored
      };
    })
    .sort((x, y) => y.score - x.score);

  res.json({ success: true, candidates });
});

// 代替候補の中から1人を選んで新規マッチング(候補)を作成し、欠勤マッチングと紐付ける
app.post('/api/matches/:id/substitute', async (req, res) => {
  const { driver_id } = req.body;
  if (!driver_id) return res.status(400).json({ success: false, message: 'driver_id は必須です' });

  const match = await dbGet(`
    SELECT m.*, s.store_id, s.store_name, s.area, s.lat, s.lng
    FROM matches m JOIN store_requests s ON s.id = m.store_request_id
    WHERE m.id = ?
  `, [req.params.id]);
  if (!match) return res.status(404).json({ success: false, message: 'マッチングが見つかりません' });

  const driver = await dbGet(`
    SELECT d.id, d.home_lat, d.home_lng, fs.name AS fixed_store_name
    FROM drivers d LEFT JOIN stores fs ON fs.id = d.fixed_store_id
    WHERE d.id = ?
  `, [driver_id]);
  if (!driver) return res.status(404).json({ success: false, message: 'ドライバーが見つかりません' });

  const av = await dbGet(
    'SELECT desired_store, desired_area FROM driver_availability WHERE driver_id = ? AND desired_date = ? AND archived_month IS NULL',
    [driver_id, match.match_date]
  );
  const { preferenceByDriverStore, storeVisitCount, areaVisitCount } = await loadScoringContext();
  const candidateLike = {
    driver_id: driver.id, home_lat: driver.home_lat, home_lng: driver.home_lng,
    desired_store: (av && av.desired_store) || driver.fixed_store_name || null,
    desired_area: av ? av.desired_area : null
  };
  const scored = scoreCandidate(candidateLike, match, preferenceByDriverStore, storeVisitCount, areaVisitCount);
  const isFar = scored.distance != null && scored.distance > DIST_WARNING_KM;

  const now = new Date().toISOString();
  const result = await dbRun(
    `INSERT INTO matches (store_request_id, driver_id, match_date, distance_km, is_far_warning, status, created_at, score, preference_flag, experience_count) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [match.store_request_id, driver_id, match.match_date, scored.distance, isFar ? 1 : 0, '候補', now, scored.score, scored.preference, scored.experienceCount]
  );
  await dbRun('UPDATE matches SET replaced_by_match_id = ? WHERE id = ?', [result.lastID, req.params.id]);

  res.json({ success: true, id: result.lastID });
});

// 特定の日付(・任意で店舗)について、その日に空いているドライバーを単独で検索する。
// 既存のマッチングに紐付けずに「今日は誰が空いているか」をすぐ調べたい場合に使う(読み取り専用、何も変更しない)
app.get('/api/availability-search', async (req, res) => {
  const date = req.query.date;
  const storeId = req.query.store_id ? parseInt(req.query.store_id, 10) : null;
  if (!date) return res.status(400).json({ success: false, message: '日付を指定してください' });

  const busyRows = await dbAll(
    `SELECT driver_id FROM matches WHERE match_date = ? AND status IN ('候補', '確定', '完了') AND archived_month IS NULL`,
    [date]
  );
  const excluded = new Set(busyRows.map(r => r.driver_id));

  const drivers = await dbAll(`
    SELECT d.id, d.name, d.phone, d.status, d.home_lat, d.home_lng, fs.name AS fixed_store_name
    FROM drivers d
    LEFT JOIN stores fs ON fs.id = d.fixed_store_id
  `);
  const availabilityToday = await dbAll(
    'SELECT driver_id, desired_store, desired_area, time_start, time_end FROM driver_availability WHERE desired_date = ? AND archived_month IS NULL',
    [date]
  );
  const availabilityByDriver = new Map(availabilityToday.map(a => [a.driver_id, a]));

  let store = null;
  const { preferenceByDriverStore, storeVisitCount, areaVisitCount } = await loadScoringContext();
  if (storeId) {
    store = await dbGet('SELECT id AS store_id, name AS store_name, area, lat, lng FROM stores WHERE id = ?', [storeId]);
  }

  const candidates = drivers
    .filter(d => !excluded.has(d.id))
    .filter(d => d.status !== '契約終了')
    .filter(d => !store || preferenceByDriverStore.get(`${d.id}:${store.store_id}`) !== 'NG')
    .map(d => {
      const av = availabilityByDriver.get(d.id);
      let scored = { distance: null, score: null, preference: null, experienceCount: 0 };
      if (store) {
        const candidateLike = {
          driver_id: d.id, home_lat: d.home_lat, home_lng: d.home_lng,
          desired_store: (av && av.desired_store) || d.fixed_store_name || null,
          desired_area: av ? av.desired_area : null
        };
        scored = scoreCandidate(candidateLike, store, preferenceByDriverStore, storeVisitCount, areaVisitCount);
      }
      return {
        driver_id: d.id, name: d.name, phone: d.phone, driver_status: d.status,
        has_availability_today: !!av,
        desired_area: av ? av.desired_area : null,
        desired_store: av ? av.desired_store : null,
        time_start: av ? av.time_start : null,
        time_end: av ? av.time_end : null,
        is_far_warning: scored.distance != null && scored.distance > DIST_WARNING_KM,
        ...scored
      };
    })
    .sort((a, b) => {
      if (store) return (b.score ?? -Infinity) - (a.score ?? -Infinity);
      return a.name.localeCompare(b.name, 'ja');
    });

  res.json({ success: true, candidates, storeSelected: !!store });
});

app.delete('/api/matches/:id', async (req, res) => {
  await dbRun('DELETE FROM matches WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// ===== 月次クローズ =====
// 指定した年月の希望シフト・店舗依頼・マッチング結果(その月の作業データ)を「アーカイブ済み」にして作業画面から隠し、
// 次の月の準備をする。削除はしない(archived_monthに「YYYY-MM」を入れるだけ)ので、後から「過去の月を見る」で参照できる。
// 店舗マスタ・ドライバーマスタ・派遣履歴はそもそも対象外(実績として恒久的に残す)。
// 「確定」のまま「完了」になっていないマッチングが残っている場合は、force指定が無ければ一旦確認を促す
app.post('/api/month-close', async (req, res) => {
  const { year, month, force } = req.body;
  if (!year || !month || month < 1 || month > 12) {
    return res.status(400).json({ success: false, message: '対象年月を指定してください' });
  }
  const startDate = `${year}-${String(month).padStart(2, '0')}-01`;
  const daysInMonth = new Date(year, month, 0).getDate();
  const endDate = `${year}-${String(month).padStart(2, '0')}-${String(daysInMonth).padStart(2, '0')}`;
  const archivedMonth = `${year}-${String(month).padStart(2, '0')}`;

  const unfinished = await dbGet(
    `SELECT COUNT(*) AS cnt FROM matches WHERE match_date >= ? AND match_date <= ? AND status = '確定' AND archived_month IS NULL`,
    [startDate, endDate]
  );
  if (unfinished.cnt > 0 && !force) {
    return res.json({ success: true, needsConfirmation: true, unfinishedCount: unfinished.cnt });
  }

  const matchesResult = await dbRun(
    'UPDATE matches SET archived_month = ? WHERE match_date >= ? AND match_date <= ? AND archived_month IS NULL',
    [archivedMonth, startDate, endDate]
  );
  const requestsResult = await dbRun(
    'UPDATE store_requests SET archived_month = ? WHERE request_date >= ? AND request_date <= ? AND archived_month IS NULL',
    [archivedMonth, startDate, endDate]
  );
  const availabilityResult = await dbRun(
    'UPDATE driver_availability SET archived_month = ? WHERE desired_date >= ? AND desired_date <= ? AND archived_month IS NULL',
    [archivedMonth, startDate, endDate]
  );

  res.json({
    success: true,
    archivedMonth,
    archived: { matches: matchesResult.changes, storeRequests: requestsResult.changes, availability: availabilityResult.changes }
  });
});

// クローズ済み(アーカイブ済み)の月の一覧を返す
app.get('/api/archive/months', async (req, res) => {
  const rows = await dbAll(`
    SELECT archived_month FROM store_requests WHERE archived_month IS NOT NULL
    UNION SELECT archived_month FROM driver_availability WHERE archived_month IS NOT NULL
    UNION SELECT archived_month FROM matches WHERE archived_month IS NOT NULL
    ORDER BY archived_month DESC
  `);
  res.json({ success: true, months: rows.map(r => r.archived_month) });
});

// 指定した月にクローズされた店舗依頼・希望シフト・マッチング結果を参照専用で返す
app.get('/api/archive/:month', async (req, res) => {
  const month = req.params.month;
  // ドライバーが後から削除されていても、アーカイブされた行自体は消えずに(ドライバー名だけ「削除済み」として)
  // 見えるように、INNER JOINではなくLEFT JOIN+COALESCEにしている
  const [storeRequests, availability, matches] = await Promise.all([
    dbAll('SELECT * FROM store_requests WHERE archived_month = ? ORDER BY request_date ASC, id ASC', [month]),
    dbAll(`
      SELECT a.*, COALESCE(d.name, '(削除済みドライバー)') AS driver_name FROM driver_availability a
      LEFT JOIN drivers d ON d.id = a.driver_id
      WHERE a.archived_month = ? ORDER BY a.desired_date ASC, a.id ASC
    `, [month]),
    dbAll(`
      SELECT m.*, COALESCE(d.name, '(削除済みドライバー)') AS driver_name, s.store_name
      FROM matches m
      LEFT JOIN drivers d ON d.id = m.driver_id
      JOIN store_requests s ON s.id = m.store_request_id
      WHERE m.archived_month = ? ORDER BY m.match_date ASC, m.id ASC
    `, [month]),
  ]);
  res.json({ success: true, storeRequests, availability, matches });
});

// クローズ済みの月を作業画面に戻す(archived_monthを外す)。間違えてクローズしてしまった場合や、
// 過去の月に戻って作業を再開したい場合に使う
app.post('/api/archive/:month/restore', async (req, res) => {
  const month = req.params.month;
  const storeRequestsResult = await dbRun('UPDATE store_requests SET archived_month = NULL WHERE archived_month = ?', [month]);
  const availabilityResult = await dbRun('UPDATE driver_availability SET archived_month = NULL WHERE archived_month = ?', [month]);
  const matchesResult = await dbRun('UPDATE matches SET archived_month = NULL WHERE archived_month = ?', [month]);
  res.json({
    success: true,
    restored: { matches: matchesResult.changes, storeRequests: storeRequestsResult.changes, availability: availabilityResult.changes }
  });
});

// 「候補」を一括で「確定」にする(1件ずつ確定ボタンを押さなくても、シフト表送付の準備がすぐできるように)
app.post('/api/matches/confirm-all', async (req, res) => {
  const result = await dbRun(`UPDATE matches SET status = '確定' WHERE archived_month IS NULL AND status = '候補'`);
  const autoFixedStoreUpdated = await applyAutoFixedStoreFromMatches();
  res.json({ success: true, confirmed: result.changes, autoFixedStoreUpdated });
});

function weekdayLabelOf(dateStr) {
  const WEEKDAY_LABELS_JP = ['日', '月', '火', '水', '木', '金', '土'];
  const d = new Date(dateStr + 'T00:00:00');
  return WEEKDAY_LABELS_JP[d.getDay()];
}

// Excelのシート名制約(31文字以内、: \ / ? * [ ] 不可、重複不可)に収まるように名前を整形する
function sanitizeSheetName(name, usedNames) {
  let base = String(name || '').replace(/[:\\/?*[\]]/g, '').trim() || 'シート';
  base = base.slice(0, 28); // 重複時の番号サフィックス分の余裕を残す
  let candidate = base;
  let i = 2;
  while (usedNames.has(candidate)) {
    candidate = `${base}(${i})`;
    i++;
  }
  usedNames.add(candidate);
  return candidate;
}

// 個人事業主ごとの月間シフト表をExcelで一括出力する(1人1シート)。「確定」「完了」のみが対象(候補はまだ未確定のため含めない)
app.get('/api/export/shift-by-driver', async (req, res) => {
  const matches = await dbAll(`
    SELECT m.*, d.name AS driver_name, d.phone AS driver_phone,
           s.store_name, s.area, s.address AS store_address, s.time_start, s.time_end
    FROM matches m
    JOIN drivers d ON d.id = m.driver_id
    JOIN store_requests s ON s.id = m.store_request_id
    WHERE m.archived_month IS NULL AND m.status IN ('確定', '完了')
    ORDER BY d.name ASC, m.match_date ASC
  `);
  if (matches.length === 0) return res.status(400).json({ success: false, message: '「確定」または「完了」のマッチングがありません(候補のままの場合は先に確定してください)' });

  const byDriver = new Map();
  for (const m of matches) {
    if (!byDriver.has(m.driver_id)) byDriver.set(m.driver_id, { name: m.driver_name, phone: m.driver_phone, rows: [] });
    byDriver.get(m.driver_id).rows.push(m);
  }

  const wb = XLSX.utils.book_new();
  const usedNames = new Set();
  for (const info of byDriver.values()) {
    const aoa = [
      [`氏名: ${info.name}`, `電話: ${info.phone || '-'}`],
      [],
      ['日付', '曜日', '店舗名', '開始', '終了', 'ステータス'],
      ...info.rows.map(m => [m.match_date, weekdayLabelOf(m.match_date), m.store_name, m.time_start || '', m.time_end || '', m.status])
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 12 }, { wch: 6 }, { wch: 22 }, { wch: 8 }, { wch: 8 }, { wch: 8 }];
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(info.name, usedNames));
  }

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="driver_shifts_${new Date().toISOString().slice(0, 10)}.xlsx"`);
  res.send(buffer);
});

// 店舗ごとの月間シフト表(誰が何時から何時まで来るか)をExcelで一括出力する(1店舗1シート)
app.get('/api/export/shift-by-store', async (req, res) => {
  const matches = await dbAll(`
    SELECT m.*, d.name AS driver_name, d.phone AS driver_phone,
           s.store_name, s.area, s.address AS store_address, s.time_start, s.time_end
    FROM matches m
    JOIN drivers d ON d.id = m.driver_id
    JOIN store_requests s ON s.id = m.store_request_id
    WHERE m.archived_month IS NULL AND m.status IN ('確定', '完了')
    ORDER BY s.store_name ASC, m.match_date ASC
  `);
  if (matches.length === 0) return res.status(400).json({ success: false, message: '「確定」または「完了」のマッチングがありません(候補のままの場合は先に確定してください)' });

  const byStore = new Map();
  for (const m of matches) {
    const key = m.store_name; // store_requestsは店舗名の表記ゆれ吸収済み(resolveStoreId)だが、念のため店舗名単位でまとめる
    if (!byStore.has(key)) byStore.set(key, { name: m.store_name, rows: [] });
    byStore.get(key).rows.push(m);
  }

  const wb = XLSX.utils.book_new();
  const usedNames = new Set();
  for (const info of byStore.values()) {
    const aoa = [
      [`店舗名: ${info.name}`],
      [],
      ['日付', '曜日', '開始', '終了', '担当ドライバー', '電話番号', 'ステータス'],
      ...info.rows.map(m => [m.match_date, weekdayLabelOf(m.match_date), m.time_start || '', m.time_end || '', m.driver_name, m.driver_phone || '-', m.status])
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 12 }, { wch: 6 }, { wch: 8 }, { wch: 8 }, { wch: 14 }, { wch: 14 }, { wch: 8 }];
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(info.name, usedNames));
  }

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="store_shifts_${new Date().toISOString().slice(0, 10)}.xlsx"`);
  res.send(buffer);
});

// 「HH:MM」同士の勤務時間(時間数)を計算する。終了が開始より前なら日またぎとみなし24時間分足す
function hoursBetween(time_start, time_end) {
  if (!time_start || !time_end) return null;
  const toMinutes = t => { const [h, m] = t.split(':').map(Number); return h * 60 + m; };
  let diff = toMinutes(time_end) - toMinutes(time_start);
  if (diff <= 0) diff += 24 * 60;
  return Math.round((diff / 60) * 100) / 100;
}

// 実データで共有された「ＳＶ確認シート」と同じ列構成(日付/社員番号/名前/シフト/出勤場所/勤務時間)で
// マッチング結果を一覧出力する(JSON版はダッシュボードの表示用、Excel版はそのまま提出できる形式)
async function getSvSheetRows() {
  const matches = await dbAll(`
    SELECT m.match_date, d.driver_code, d.name AS driver_name, s.store_name, s.time_start, s.time_end, m.status
    FROM matches m
    JOIN drivers d ON d.id = m.driver_id
    JOIN store_requests s ON s.id = m.store_request_id
    WHERE m.archived_month IS NULL AND m.status IN ('確定', '完了')
    ORDER BY m.match_date ASC, d.name ASC
  `);
  return matches.map(m => ({
    date: m.match_date,
    driver_code: m.driver_code || '',
    driver_name: m.driver_name,
    shift: (m.time_start && m.time_end) ? `${m.time_start}〜${m.time_end}` : '',
    store_name: m.store_name,
    hours: hoursBetween(m.time_start, m.time_end)
  }));
}

app.get('/api/export/sv-sheet', async (req, res) => {
  const rows = await getSvSheetRows();
  res.json({ success: true, rows });
});

app.get('/api/export/sv-sheet.xlsx', async (req, res) => {
  const rows = await getSvSheetRows();
  if (rows.length === 0) return res.status(400).json({ success: false, message: '「確定」または「完了」のマッチングがありません(候補のままの場合は先に確定してください)' });
  const aoa = [
    ['日付', '社員番号', '名前(個人事業主)', 'シフト', '出勤場所', '勤務時間'],
    ...rows.map(r => [r.date, r.driver_code, r.driver_name, r.shift, r.store_name, r.hours])
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 8 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'ＳＶ確認シート');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="sv_sheet_${new Date().toISOString().slice(0, 10)}.xlsx"`);
  res.send(buffer);
});

// 店舗依頼(週間必要枠表)の、取込みにそのまま使えるひな形Excelを生成する。
// 「これ聞取り固定デポ」等の実ファイルで実績のある構造(店舗名+月〜日の曜日列+備考)に合わせてあり、
// 余計な分析用の列は含めない(あくまで取込みに必要な最小限の列のみ)
app.get('/api/templates/store-requests-weekly.xlsx', (req, res) => {
  const header = ['店舗名', '月', '火', '水', '木', '金', '土', '日', '備考'];
  const sampleRows = [
    ['（記入例）銀座SS', '10:00-22:00', '10:00-22:00', '10:00-22:00', '10:00-22:00', '10:00-22:00', '10:00-22:00', '10:00-22:00', ''],
    ['（記入例）築地店', '10:00-22:00', '', '10:00-22:00', '', '10:00-22:00', '10:00-22:00', '', '平日休みがある店舗の例(空欄の曜日は依頼なし)'],
    ['（記入例）東雲店', '10:00-22:00*2', '10:00-22:00*2', '10:00-22:00*2', '10:00-22:00*2', '10:00-22:00*2', '10:00-22:00*3', '10:00-22:00*3', '1日に複数人必要な場合は「*人数」を付ける例(土日は3人)'],
  ];
  const rulesSheet = [
    ['店舗依頼(週間必要枠表)ひな形の使い方'],
    [],
    ['① 1行=1つの店舗の、曜日ごとの必要枠パターンです。実際に取込む際は、記入例の行を削除して店舗名を入れ替えてください。'],
    ['② 月〜日の各列には、その曜日に必要な時間帯を「開始-終了」の形式で入れてください(例: 10:00-22:00)。'],
    ['③ その曜日に依頼が無い場合は、セルを空欄のままにしてください。'],
    ['④ 1日に複数人必要な場合は、時間帯の後ろに「*人数」を付けてください(例: 10:00-22:00*2 で2人分)。'],
    ['⑤ 同じ店舗・同じ曜日・同じ時間帯の行が複数あっても、必要人数として自動的に合算されます。'],
    ['⑥ 備考欄は自由記述です。社員番号(5〜8桁の数字)を書くと、該当ドライバーの固定希望店舗に自動反映されます(ただしその店舗がNG設定の場合は反映されません)。'],
    ['⑦ 取込み時に「対象年月」を指定すると、その月のうち該当する曜日すべてに展開されて店舗依頼が登録されます。'],
  ];

  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet([header, ...sampleRows]);
  ws1['!cols'] = [{ wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, ws1, '店舗依頼ひな形');
  const ws2 = XLSX.utils.aoa_to_sheet(rulesSheet);
  ws2['!cols'] = [{ wch: 80 }];
  XLSX.utils.book_append_sheet(wb, ws2, '記入ルール');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="store_requests_template.xlsx"');
  res.send(buffer);
});

// 希望シフト(長形式、1行=1人×1日)の、取込みにそのまま使えるひな形Excelを生成する。
// 楽シフ等の実際のCSV列名がまだ分かっていないため、現時点の想定列名(AVAILABILITY_ALIASESの主要なもの)で
// 用意している。実際の楽シフCSVのヘッダー行が分かり次第、列名を合わせて調整する想定
app.get('/api/templates/driver-availability-long.xlsx', (req, res) => {
  const header = ['ドライバー名', '希望日', '希望エリア', '希望店舗', '開始時刻', '終了時刻', '備考'];
  const sampleRows = [
    ['（記入例）山田太郎', '2026-11-01', '', '銀座SS', '10:00', '22:00', ''],
    ['（記入例）山田太郎', '2026-11-02', '', '', '10:00', '20:00', '希望店舗が空欄でも、固定希望店舗が設定されていればそれが使われます'],
    ['（記入例）山田太郎', '2026-11-04', '渋谷区', '', '12:00', '22:00', '店舗名が分からない場合は希望エリアだけでもOK'],
  ];
  const rulesSheet = [
    ['希望シフト(長形式)ひな形の使い方'],
    [],
    ['① 1行=1人のドライバーの、1日分の希望シフトです。実際に取込む際は、記入例の行を削除してドライバー名を入れ替えてください。'],
    ['② ドライバー名は「ドライバーマスタ」に登録済みの氏名と完全一致している必要があります(先に登録してください)。'],
    ['③ 休み(その日は稼働しない)の場合は、その日の行自体を入れないでください(空のセルではなく、行ごと無しにする形式です)。'],
    ['④ 希望エリア・希望店舗は分かる範囲で構いません。両方空欄でも、ドライバーマスタの「固定希望店舗」が設定されていれば自動的に使われます。'],
    ['⑤ 同じドライバー・同じ希望日の行が複数あると、後に読み込んだ方で上書きされます。'],
    [],
    ['※ このひな形は現時点の想定列名です。楽シフ等のエクスポート形式が分かれば、実際の列名に合わせて調整できます。'],
  ];

  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet([header, ...sampleRows]);
  ws1['!cols'] = [{ wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 10 }, { wch: 10 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, ws1, '希望シフトひな形');
  const ws2 = XLSX.utils.aoa_to_sheet(rulesSheet);
  ws2['!cols'] = [{ wch: 80 }];
  XLSX.utils.book_append_sheet(wb, ws2, '記入ルール');

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="driver_availability_template.xlsx"');
  res.send(buffer);
});

app.listen(PORT, () => console.log(`マッチングアプリ起動: http://localhost:${PORT}`));
