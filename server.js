require('dotenv').config();
const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawnSync } = require('child_process');
const multer = require('multer');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs'); // 罫線等のセル装飾が必要なひな形生成に使う(xlsxパッケージは装飾の書き出しに非対応のため)
const iconv = require('iconv-lite');
const jschardet = require('jschardet');
const Anthropic = require('@anthropic-ai/sdk');
const crypto = require('crypto');
const archiver = require('archiver'); // 個人事業主別シフトPDFをまとめてzipでダウンロードさせるために使う

const app = express();
// ローカルPCでは今まで通り4001番・プロジェクト直下のファイルを使う。クラウド環境(Render等)では
// PORT/DB_PATH/UPLOADS_DIRを環境変数で指定し、永続ディスクのパスを使うようにする
const PORT = process.env.PORT || 4001;
const DIST_WARNING_KM = 30; // これを超える距離のマッチングは「距離が遠い」として画面上で警告表示する
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

// 「所定フォルダ」: 担当者から共有されたExcel/CSVをここに置いておけば、ブラウザでファイルを都度選ばなくても
// 取込み画面から一覧表示→選択して取込みできる(uploads/フォルダをそのまま使う)
const UPLOADS_DIR = process.env.UPLOADS_DIR || path.join(__dirname, 'uploads');
fs.mkdirSync(UPLOADS_DIR, { recursive: true });
const INCOMING_FILE_EXTENSIONS = ['.xlsx', '.xls', '.csv', '.txt'];

// 備考・LINEメッセージのAI解析(任意機能)。.envにANTHROPIC_API_KEYが設定されていない場合はnullのままで、
// 関連エンドポイントは「AI未設定」を返す(既存の正規表現ベースの解析は影響を受けない)
const anthropic = process.env.ANTHROPIC_API_KEY ? new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY }) : null;
const AI_MODEL = 'claude-haiku-5-5'; // 解析用途のため、速く安価なモデルを使う

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
const SCORE_AI_PREFERRED_BONUS = 20;     // 店舗依頼の備考をAIが解析し、特定ドライバーを「できれば希望」と判断した場合のボーナス
const SCORE_AI_EXCLUDED_PENALTY = -99999; // 同、「NG」と判断した場合のペナルティ(候補の絞り込み側で既に除外している想定の、念のための二重の安全策)

// ベーシック認証(最低限の保護)。.envにAUTH_USER/AUTH_PASSが設定されている時だけ有効にする
// (このPC上の開発用インスタンスは未設定のままにして、今まで通り認証無しで使える)
const AUTH_USER = process.env.AUTH_USER;
const AUTH_PASS = process.env.AUTH_PASS;
if (AUTH_USER && AUTH_PASS) {
  app.use((req, res, next) => {
    const header = req.headers.authorization || '';
    const [scheme, encoded] = header.split(' ');
    if (scheme === 'Basic' && encoded) {
      const decoded = Buffer.from(encoded, 'base64').toString('utf8');
      const sep = decoded.indexOf(':');
      if (sep !== -1 && decoded.slice(0, sep) === AUTH_USER && decoded.slice(sep + 1) === AUTH_PASS) {
        return next();
      }
    }
    res.set('WWW-Authenticate', 'Basic realm="matching-app"');
    res.status(401).send('認証が必要です(IDとパスワードを入力してください)');
  });
  console.log('🔒 ベーシック認証が有効です');
}

// ===== 利用者ログイン(利用者マスタに登録したログインID・パスワードでの個別アカウント認証) =====
// 権限の区分けはせず「ログインできるかどうか」のみ。社内の少人数利用のため、ログインしたら
// 長期間(180日)はセッションを保持する。パスワードはcrypto.scryptでソルト付きハッシュ化して保存する
const SESSION_COOKIE_NAME = 'matching_session';
const SESSION_MAX_AGE_SEC = 60 * 60 * 24 * 180;

function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(password, salt, 64).toString('hex');
  return `${salt}:${hash}`;
}
function verifyPassword(password, stored) {
  if (!stored || !stored.includes(':')) return false;
  const [salt, hashHex] = stored.split(':');
  const hashBuffer = Buffer.from(hashHex, 'hex');
  const testHash = crypto.scryptSync(password, salt, 64);
  return hashBuffer.length === testHash.length && crypto.timingSafeEqual(hashBuffer, testHash);
}
function parseCookies(req) {
  const header = req.headers.cookie;
  const cookies = {};
  if (!header) return cookies;
  for (const part of header.split(';')) {
    const idx = part.indexOf('=');
    if (idx === -1) continue;
    cookies[part.slice(0, idx).trim()] = decodeURIComponent(part.slice(idx + 1).trim());
  }
  return cookies;
}
function setSessionCookie(res, token) {
  res.setHeader('Set-Cookie', `${SESSION_COOKIE_NAME}=${token}; HttpOnly; Path=/; Max-Age=${SESSION_MAX_AGE_SEC}; SameSite=Lax`);
}
function clearSessionCookie(res) {
  res.setHeader('Set-Cookie', `${SESSION_COOKIE_NAME}=; HttpOnly; Path=/; Max-Age=0; SameSite=Lax`);
}
async function getSessionUser(req) {
  const token = parseCookies(req)[SESSION_COOKIE_NAME];
  if (!token) return null;
  return await dbGet(
    'SELECT u.id, u.name, u.login_id FROM app_sessions s JOIN app_users u ON u.id = s.user_id WHERE s.token = ?',
    [token]
  );
}

// ログイン画面本体・ログイン/ログアウト/初回セットアップ関連のAPIだけは未ログインでもアクセスできるようにする
const AUTH_PUBLIC_PATHS = new Set(['/login.html', '/api/auth/login', '/api/auth/logout', '/api/auth/setup-needed', '/api/auth/setup']);
app.use(async (req, res, next) => {
  if (AUTH_PUBLIC_PATHS.has(req.path)) return next();
  try {
    const user = await getSessionUser(req);
    if (!user) {
      if (req.path.startsWith('/api/')) return res.status(401).json({ success: false, message: 'ログインが必要です' });
      return res.redirect('/login.html');
    }
    req.user = user;
    next();
  } catch (e) {
    console.error('セッション確認に失敗しました:', e);
    res.status(500).json({ success: false, message: 'サーバーエラーが発生しました' });
  }
});

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'matching.db');
fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });
const db = new sqlite3.Database(DB_PATH);

// 複数プロセスから同時にこのDBファイルが開かれても書き込みが失われにくいようにする。
// WALモードは読み取りと書き込みが競合しにくく、busy_timeoutは他プロセスが書き込み中でも
// すぐにエラーにせず一定時間リトライしてから諦めるようにする(誤って二重起動してしまった場合の保険)
db.run('PRAGMA journal_mode = WAL');
db.run('PRAGMA busy_timeout = 5000');

db.serialize(() => {
  // 利用者マスタ(このアプリにログインできる担当者の一覧。氏名のみを管理項目とし、
  // ログインID・パスワードはログイン用に別途持たせる。権限の区分けはせず、ログイン可否のみ)
  db.run(`
    CREATE TABLE IF NOT EXISTS app_users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      login_id TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      created_at TEXT
    )
  `);

  // ログインセッション(トークンをCookieに持たせ、ここと突き合わせて本人確認する)
  db.run(`
    CREATE TABLE IF NOT EXISTS app_sessions (
      token TEXT PRIMARY KEY,
      user_id INTEGER NOT NULL,
      created_at TEXT,
      FOREIGN KEY (user_id) REFERENCES app_users(id)
    )
  `);

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

  // 固定希望店舗の「候補」。店舗依頼の備考(社員番号・AI解析)や、確定/完了マッチングの実績パターンから
  // 固定希望店舗が示唆された時に、ドライバーマスタを自動で書き換える代わりにここに記録しておき、
  // コーディネーターがドライバーマスタ画面で見て手動で採用/却下できるようにする
  db.run(`
    CREATE TABLE IF NOT EXISTS fixed_store_suggestions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      driver_id INTEGER NOT NULL,
      suggested_store_id INTEGER NOT NULL,
      source TEXT NOT NULL,
      reason TEXT,
      status TEXT NOT NULL DEFAULT 'pending',
      created_at TEXT,
      resolved_at TEXT,
      FOREIGN KEY (driver_id) REFERENCES drivers(id),
      FOREIGN KEY (suggested_store_id) REFERENCES stores(id)
    )
  `);

  // 一括取込み(店舗依頼/希望シフト/エリア固定)を1回につき1件記録し、間違えた/テストで入れた取込みを
  // まとめて取り消せるようにする(行ごとにimport_batch_idの印をつけ、このバッチ単位で取り消す)
  db.run(`
    CREATE TABLE IF NOT EXISTS import_batches (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      type TEXT NOT NULL,
      filename TEXT,
      label TEXT,
      row_count INTEGER DEFAULT 0,
      created_at TEXT,
      undone_at TEXT
    )
  `);

  // ドライバーグループ(「2店舗を3人で回す」のように、複数のドライバーと複数の店舗をまとめて紐付ける仕組み)。
  // 旧「エリア固定」機能の置き換え。1人のドライバーが複数のグループに所属することもできる
  db.run(`
    CREATE TABLE IF NOT EXISTS driver_groups (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      created_at TEXT
    )
  `);
  db.run(`
    CREATE TABLE IF NOT EXISTS driver_group_stores (
      group_id INTEGER NOT NULL REFERENCES driver_groups(id),
      store_id INTEGER NOT NULL REFERENCES stores(id),
      PRIMARY KEY (group_id, store_id)
    )
  `);
  db.run(`
    CREATE TABLE IF NOT EXISTS driver_group_members (
      group_id INTEGER NOT NULL REFERENCES driver_groups(id),
      driver_id INTEGER NOT NULL REFERENCES drivers(id),
      PRIMARY KEY (group_id, driver_id)
    )
  `);

  // 既存テーブルへのカラム追加(運用中のDBを壊さないよう、無ければ追加する形で行う)
  ensureColumn('store_requests', 'store_id', 'INTEGER REFERENCES stores(id)');
  ensureColumn('matches', 'score', 'REAL');
  ensureColumn('matches', 'preference_flag', 'TEXT');
  ensureColumn('matches', 'experience_count', 'INTEGER DEFAULT 0');
  ensureColumn('matches', 'absence_reason', 'TEXT');
  ensureColumn('matches', 'replaced_by_match_id', 'INTEGER');
  // 'auto'(自動マッチングで作成)か'manual'(マッチング・結果画面の空き枠右クリックから手動で割り当て)かの区別。
  // 画面側で色分けして見分けられるようにするために使う
  ensureColumn('matches', 'created_via', "TEXT DEFAULT 'auto'");
  // このマッチングが実際にカバーする時間(担当者自身の希望時間。無ければ店舗依頼の時間をそのまま使う)。
  // 店舗依頼の時間を1人でカバーしきれない場合に、複数人のマッチングで時間を分担しているかどうかの
  // 判定(時間の「穴」検知)に使う
  ensureColumn('matches', 'time_start', 'TEXT');
  ensureColumn('matches', 'time_end', 'TEXT');
  ensureColumn('stores', 'manager_name', 'TEXT'); // 店舗担当者(店長)
  ensureColumn('stores', 'sv_name', 'TEXT');      // SV(スーパーバイザー)
  ensureColumn('stores', 'store_code', 'TEXT');   // 店番(店舗を一意に表す社内コード。現時点では未入力で運用)
  ensureColumn('drivers', 'driver_code', 'TEXT');          // 社員コード(ドライバーを一意に表す社内コード)
  ensureColumn('drivers', 'furigana', 'TEXT');             // ふりがな(個人事業主ごとの作業依頼表PDFに印字する)
  // ログインのブルートフォース対策(パスワード総当たり防止)。一定回数連続で失敗したら
  // 一時的にロックする(ログインIDごとに管理。locked_untilを過ぎれば自動的に再試行できる)
  ensureColumn('app_users', 'failed_login_count', 'INTEGER DEFAULT 0');
  ensureColumn('app_users', 'locked_until', 'TEXT');
  ensureColumn('drivers', 'first_contract_date', 'TEXT');  // 初回委託日
  ensureColumn('drivers', 'status', 'TEXT');               // ステータス(稼働中/休止/契約終了)
  ensureColumn('drivers', 'email', 'TEXT');                // メールアドレス
  ensureColumn('drivers', 'insurance_info', 'TEXT');        // 保険加入状況(貨物保険等のメモ)
  ensureColumn('drivers', 'company_name', 'TEXT');          // 会社名(法人として契約している個人事業主向け。個人の場合は空欄)
  ensureColumn('drivers', 'fixed_store_id', 'INTEGER REFERENCES stores(id)'); // 固定希望店舗(希望シフトで店舗未入力の日の初期値として使う)
  // 固定希望店舗が複数ドライバーで重複した場合の優先順位(1が最優先)。同じ固定希望店舗のドライバー間でのみ意味を持つ
  ensureColumn('drivers', 'fixed_store_priority', 'INTEGER');
  // エリア固定: 店舗を1つに固定するのではなく、曜日ごとの決まった時間帯+複数の候補店舗群(想定デポ)の中から
  // 優先的に割り当てる仕組み。area_fixed_patternは{mon:"10:00-22:00", ...}形式のJSON文字列、
  // area_fixed_store_idsは候補店舗idの配列のJSON文字列
  ensureColumn('drivers', 'area_fixed_enabled', 'INTEGER DEFAULT 0');
  ensureColumn('drivers', 'area_fixed_pattern', 'TEXT');
  ensureColumn('drivers', 'area_fixed_store_ids', 'TEXT');
  ensureColumn('store_requests', 'archived_month', 'TEXT');     // 月次クローズで「YYYY-MM」を入れ、作業画面から隠す(データは消さない)
  ensureColumn('driver_availability', 'archived_month', 'TEXT');
  ensureColumn('matches', 'archived_month', 'TEXT');
  // 取込みの取消し機能用: 行がどの取込みバッチで作成/更新されたかの印(無ければ手入力/旧データ)
  ensureColumn('store_requests', 'import_batch_id', 'INTEGER');
  // 備考(requests列)をAIで解析した結果のキャッシュ。ai_analyzed_requestsは解析時点のrequests原文を
  // 保存しておき、再度マッチングを実行する時に内容が変わっていなければ再解析(API呼び出し)をスキップする
  ensureColumn('store_requests', 'ai_preference_json', 'TEXT');
  ensureColumn('store_requests', 'ai_analyzed_requests', 'TEXT');
  // エリア固定の「SV行」(特定の個人ではなく、SVが担当するエリア全体の需要)から生成された店舗依頼を
  // グループ化するキー。同じpool_group_idを持つ行は「このうちどれか1件が指定人数分埋まればよい」
  // 候補店舗群として扱う(通常の店舗依頼はNULLのまま、1件=1店舗の個別需要として今まで通り扱う)
  ensureColumn('store_requests', 'pool_group_id', 'TEXT');
  // この依頼(特定の店舗・特定の日)の「時間の穴」警告を対応しないと決めた場合のフラグ。
  // 店舗単位・月単位ではなく、この依頼(1日分)だけに効く(マッチング・結果画面の右クリックから設定)
  ensureColumn('store_requests', 'gap_dismissed', 'INTEGER DEFAULT 0');
  ensureColumn('driver_availability', 'import_batch_id', 'INTEGER');
  ensureColumn('drivers', 'area_fixed_batch_id', 'INTEGER');
  migratePreferenceScale(); // 相性を好き/NGの2択からNG〜1〜5〜OKの7段階スケールに移行する(旧DB向け)
});

// 「時間の穴」機能を追加する前に作られたマッチングは、matches.time_start/time_endがまだ空のままなので、
// (computeTimeGapsのフォールバックにより)店舗依頼の時間を完全にカバーしていると誤認識されてしまう
// (例: 実際は9:00〜10:00が未カバーなのに、警告が出ない)。起動のたびに、担当者本人がその日に
// 出した希望シフトの時間を遡って設定し直す(無ければ店舗依頼の時間のまま=今まで通り)。
// 対象(time_start IS NULL)が無ければ何もしない
async function backfillMatchTimes() {
  try {
    const rows = await dbAll(`
      SELECT m.id, m.driver_id, m.match_date, s.time_start AS store_time_start, s.time_end AS store_time_end
      FROM matches m JOIN store_requests s ON s.id = m.store_request_id
      WHERE m.time_start IS NULL OR m.time_end IS NULL
    `);
    if (rows.length === 0) return;
    const availRows = await dbAll('SELECT driver_id, desired_date, time_start, time_end FROM driver_availability');
    const availByKey = new Map(availRows.map(a => [`${a.driver_id}|${a.desired_date}`, a]));
    for (const r of rows) {
      const av = availByKey.get(`${r.driver_id}|${r.match_date}`);
      const ts = (av && av.time_start) || r.store_time_start;
      const te = (av && av.time_end) || r.store_time_end;
      await dbRun('UPDATE matches SET time_start = ?, time_end = ? WHERE id = ?', [ts, te, r.id]);
    }
    console.log(`🕐 既存マッチング${rows.length}件の時間情報を補完しました(時間の穴の検知を正しく行うため)`);
  } catch (e) {
    console.error('既存マッチングの時間情報の補完に失敗しました:', e);
  }
}
backfillMatchTimes();

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

// 全角の英数字(Ａ-Ｚ、ａ-ｚ、０-９)を半角に変換する(「新宿１丁目店」と「新宿1丁目店」のような表記ゆれを吸収するため)
function foldWidth(str) {
  return String(str || '').replace(/[Ａ-Ｚａ-ｚ０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
}

// 店舗名の表記ゆれ(「class」接頭辞、「【Tax-Free】」等の装飾、全角/半角の違い)を取り除いて正規化する
function normalizeStoreName(store_name) {
  return foldWidth(
    (store_name || '')
      .replace(/^class/i, '')
      .replace(/【[^】]*】/g, '')
      .trim()
  );
}

// 社員番号の先頭0の有無(「540094」と「00540094」等)の表記ゆれを吸収するため、先頭0を除いた形をキーにする
function normalizeEmployeeCode(code) {
  return String(code || '').trim().replace(/^0+(?=\d)/, '');
}

// 氏名の照合用に、全角/半角スペース・連続スペースの違いを吸収する(「吉川  守人」等)
function normalizeNameForMatch(name) {
  return String(name || '').replace(/[\s　]+/g, '');
}

// 店舗名(自由入力)を店舗マスタに名寄せする。店番(拠点コード)が分かればまずそれで照合し
// (表記ゆれが起きやすい店舗名より確実なため)、無ければ店舗名で照合する。
// 既存店舗が見つかれば、住所/店番等の未設定項目を補完してそのidを返す。見つからなければ新規に登録する
async function resolveStoreId({ store_name, area, address, lat, lng, store_code }) {
  // 取込み元によって店舗名に「class」接頭辞や「【Tax-Free】」表記が付いたり付かなかったりするため、
  // 正規化してから既存の店舗マスタと照合する(そうしないと同じ店舗が表記違いで重複登録されてしまう)
  const name = normalizeStoreName(store_name);
  const code = store_code ? String(store_code).trim() : '';
  if (!name && !code) return null;
  const now = new Date().toISOString();

  let existing = code ? await dbGet('SELECT * FROM stores WHERE store_code = ?', [code]) : null;
  if (!existing && name) existing = await dbGet('SELECT * FROM stores WHERE name = ?', [name]);
  if (existing) {
    if ((!existing.address && address) || (!existing.store_code && code)) {
      await dbRun(
        'UPDATE stores SET area = COALESCE(NULLIF(area, \'\'), ?), address = COALESCE(NULLIF(address, \'\'), ?), lat = COALESCE(lat, ?), lng = COALESCE(lng, ?), store_code = COALESCE(NULLIF(store_code, \'\'), ?) WHERE id = ?',
        [area || '', address || '', lat, lng, code || null, existing.id]
      );
    }
    return existing.id;
  }
  if (!name) return null; // 店番だけでは店舗名が分からず新規登録できない
  const result = await dbRun(
    'INSERT INTO stores (name, area, address, lat, lng, notes, store_code, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
    [name, area || '', address || '', lat, lng, '', code || null, now]
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

// 固定希望店舗の「候補」を記録する(ドライバーマスタは書き換えない)。同じドライバー・同じ候補店舗の
// 保留中の候補が既にあれば、日時と理由だけ更新する(取込むたびに同じ候補が重複して積み上がらないように)
async function recordFixedStoreSuggestion(driverId, suggestedStoreId, source, reason) {
  const existing = await dbGet(
    `SELECT id FROM fixed_store_suggestions WHERE driver_id = ? AND suggested_store_id = ? AND status = 'pending'`,
    [driverId, suggestedStoreId]
  );
  const now = new Date().toISOString();
  if (existing) {
    await dbRun('UPDATE fixed_store_suggestions SET source = ?, reason = ?, created_at = ? WHERE id = ?', [source, reason, now, existing.id]);
  } else {
    await dbRun(
      `INSERT INTO fixed_store_suggestions (driver_id, suggested_store_id, source, reason, status, created_at) VALUES (?, ?, ?, ?, 'pending', ?)`,
      [driverId, suggestedStoreId, source, reason, now]
    );
  }
}

// 一括取込みの開始時に呼び、取込み履歴(import_batches)に1件記録してそのidを返す。
// 取込み処理中は、作成/更新した行にこのidを印として付けておき(import_batch_id列)、
// 後から「この取込みを取り消す」際にまとめて特定できるようにする
async function createImportBatch(type, filename, label) {
  const result = await dbRun(
    'INSERT INTO import_batches (type, filename, label, row_count, created_at) VALUES (?, ?, ?, 0, ?)',
    [type, filename || '', label || '', new Date().toISOString()]
  );
  return result.lastID;
}
// 取込み処理の最後に呼び、実際に取り込んだ件数を記録する(履歴画面での表示用)
async function finalizeImportBatch(batchId, rowCount) {
  await dbRun('UPDATE import_batches SET row_count = ? WHERE id = ?', [rowCount, batchId]);
}

// 「確定」「完了」のマッチングを見て、1ヶ月間ずっと同じ店舗に割り当てられているドライバーがいれば、
// 固定希望店舗の「候補」として記録する(2日以上、かつ全ての確定/完了マッチングが同じ店舗の場合のみ対象。
// 既に同じ店舗が設定済みなら何もしない)。確定・完了操作のたびに呼び出す想定。
// ※以前はここでドライバーマスタを自動で書き換えていたが、自動で勝手に書き換えないでほしいという
// 要望を受け、候補の記録のみに変更した(実際に設定するかはコーディネーターがドライバーマスタ画面で判断する)。
// グループに所属しているドライバーは対象外(たまたま数日同じ店舗が続いただけで固定希望店舗の候補に
// してしまうと、複数店舗を柔軟に回るグループの良さを分かりにくくしてしまうため)
async function applyAutoFixedStoreFromMatches() {
  const rows = await dbAll(`
    SELECT m.driver_id, s.store_id, st.name AS store_name, d.fixed_store_id
    FROM matches m
    JOIN store_requests s ON s.id = m.store_request_id
    JOIN drivers d ON d.id = m.driver_id
    LEFT JOIN stores st ON st.id = s.store_id
    WHERE m.archived_month IS NULL AND m.status IN ('確定', '完了')
      AND d.id NOT IN (SELECT driver_id FROM driver_group_members)
  `);
  const byDriver = new Map(); // driver_id -> { storeIds: Set, storeName, count, currentFixedStoreId }
  for (const r of rows) {
    const entry = byDriver.get(r.driver_id) || { storeIds: new Set(), storeName: r.store_name, count: 0, currentFixedStoreId: r.fixed_store_id };
    entry.storeIds.add(r.store_id);
    entry.count++;
    byDriver.set(r.driver_id, entry);
  }

  let updated = 0;
  for (const [driver_id, entry] of byDriver) {
    if (entry.count < 2 || entry.storeIds.size !== 1) continue; // 同じ店舗が2件以上続いている場合のみ対象
    const onlyStoreId = [...entry.storeIds][0];
    if (entry.currentFixedStoreId === onlyStoreId) continue; // 既に同じ設定なら何もしない
    await recordFixedStoreSuggestion(driver_id, onlyStoreId, 'match_pattern',
      `直近の確定/完了マッチングで、同じ店舗(${entry.storeName || ''})に${entry.count}件連続して割り当てられています`);
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

// 店舗依頼の備考欄(自由記述)をAI(Claude)で解析し、特定のドライバー(個人事業主)への
// 希望・必須指定・除外(NG)の意図を読み取る。正規表現(数字の社員番号)だけでは拾えない、
// 「できれば髙橋さん希望」のような自然文表現を読み取るために使う。
// 未設定時・解析失敗時・該当なしの場合はnullを返す(呼び出し側は従来通りの処理にフォールバックする)
async function analyzeRequestNoteForMatching(noteText) {
  if (!anthropic || !noteText || !noteText.trim()) return null;
  try {
    const response = await anthropic.messages.create({
      model: AI_MODEL,
      max_tokens: 500,
      tools: [{
        name: 'analyze_note',
        description: '配送委託(軽貨物)の店舗依頼の備考欄から、特定の個人事業主(ドライバー)への希望・指定・除外の意図を読み取る',
        input_schema: {
          type: 'object',
          properties: {
            driver_mentions: {
              type: 'array',
              description: '備考文に具体的な人物(社員番号または氏名)への言及があれば、その一覧。無ければ空配列',
              items: {
                type: 'object',
                properties: {
                  employee_code: { type: 'string', description: '文中の社員番号(数字のみ)。無ければ空文字' },
                  name_hint: { type: 'string', description: '文中の氏名・苗字の手がかり(例: 「髙橋」「髙橋さん」なら「髙橋」)。無ければ空文字' },
                  sentiment: {
                    type: 'string',
                    enum: ['required', 'preferred', 'excluded'],
                    description: 'required=必ずこの人にしてほしいという強い指定、preferred=できればこの人がいいという弱い希望、excluded=この人はNG/避けてほしい'
                  }
                },
                required: ['employee_code', 'name_hint', 'sentiment']
              }
            },
            summary: { type: 'string', description: '備考全体の内容を1文で要約(人物指定以外の内容も含めて)。特に内容が無ければ空文字' }
          },
          required: ['driver_mentions', 'summary']
        }
      }],
      tool_choice: { type: 'tool', name: 'analyze_note' },
      messages: [{ role: 'user', content: `配送委託の店舗依頼スプレッドシートの備考欄です。特定の個人事業主への希望・指定・除外が書かれていないか読み取ってください:\n「${noteText}」` }]
    });
    const toolUse = response.content.find(c => c.type === 'tool_use');
    if (!toolUse) return null;
    return toolUse.input;
  } catch (e) {
    return null; // AI解析に失敗しても処理は止めない(呼び出し側で従来通りのフォールバックを使う)
  }
}

// analyzeRequestNoteForMatchingが返したdriver_mentionsの1件を、実際のドライバーIDに解決する。
// 社員番号が一致すればそれを優先し、無ければ氏名の手がかりで部分一致を試みる(表記ゆれがあるため、
// 正規化した氏名同士が互いに含み合うかで判定する)。複数候補に一致する場合は誤爆を避けるため解決しない
function resolveDriverMention(mention, driverByCode, driversForNameSearch) {
  if (mention.employee_code) {
    const code = normalizeEmployeeCode(mention.employee_code);
    if (driverByCode.has(code)) return driverByCode.get(code);
  }
  if (mention.name_hint) {
    const hint = normalizeNameForMatch(mention.name_hint);
    if (hint) {
      const hits = driversForNameSearch.filter(d => d.normName.includes(hint) || hint.includes(d.normName));
      if (hits.length === 1) return hits[0].id;
    }
  }
  return null;
}

// 店舗依頼(store_requests)1件の備考をAIで解析した結果を返す。前回解析時から備考の内容が
// 変わっていなければ、DBに保存済みの結果をそのまま返す(呼び出すたびにAPIを叩いてコストが
// かさまないようにするため)。内容が変わっていれば(または未解析であれば)解析してDBに保存し直す
async function getOrAnalyzeStoreRequestNote(storeRequest) {
  const noteText = (storeRequest.requests || '').trim();
  if (!noteText) return null;
  if (storeRequest.ai_analyzed_requests === noteText && storeRequest.ai_preference_json) {
    try { return JSON.parse(storeRequest.ai_preference_json); } catch (e) { /* 壊れていたら再解析にフォールバック */ }
  }
  const result = await analyzeRequestNoteForMatching(noteText);
  if (result) {
    await dbRun('UPDATE store_requests SET ai_preference_json = ?, ai_analyzed_requests = ? WHERE id = ?',
      [JSON.stringify(result), noteText, storeRequest.id]);
  }
  return result;
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

// 指定した年月の日付を、曜日(Date.getDay()の0=日,1=月,...,6=土)ごとにまとめた対応表を返す
// (「この曜日は月内のこの日付たち」という対応。曜日パターンを実際の日付に展開する時に使う)
function datesByWeekdayInMonth(year, month) {
  const daysInMonth = new Date(year, month, 0).getDate();
  const datesByJsDay = {};
  for (let d = 1; d <= daysInMonth; d++) {
    const jsDay = new Date(year, month - 1, d).getDay();
    (datesByJsDay[jsDay] = datesByJsDay[jsDay] || []).push(d);
  }
  return datesByJsDay;
}

// 候補者(driver_id, home_lat, home_lng, desired_store, desired_area を持つオブジェクト)を
// 店舗依頼に対してスコアリングする(自動マッチングと欠勤時の代替候補探しの両方で使う共通ロジック)。
// aiSentiment: その店舗依頼の備考をAIが解析した結果、この候補者について判定された意図
// ('preferred'=できれば希望、'excluded'=NG、該当なしはnull/undefined)。省略可(後方互換)
function scoreCandidate(a, store, preferenceByDriverStore, storeVisitCount, areaVisitCount, aiSentiment) {
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
  if (aiSentiment === 'preferred') score += SCORE_AI_PREFERRED_BONUS;
  if (aiSentiment === 'excluded') score += SCORE_AI_EXCLUDED_PENALTY;

  return { distance, score, preference, experienceCount };
}

// マッチングの優先順位(階層)のカタログ。/api/matches/run は、画面(マッチング設定)で有効化・
// 並び替えされた階層の順に、店舗の処理順に関係なく全店舗を通して先に確保してから、
// 残りをスコア順で埋める(2段階処理)。test(a, store, ctx)がtrueを返す候補がその階層の対象。
// ここに新しい階層を追加すれば、画面側で有効化・並び替えできるようになる
const TIER_CATALOG = {
  ai_note_required: {
    label: '備考でAIが検出した「必須」指定',
    description: '店舗依頼の備考欄(自由記述)をAIが解析し、特定のドライバーを強く希望している(必須)と判断した場合に最優先する(.envにANTHROPIC_API_KEYが未設定の場合はこの階層は常に対象なし)',
    test: (a, store, ctx) => ctx.aiPreferenceByKey?.get(`${a.driver_id}:${store.id}`) === 'required'
  },
  fixed_store: {
    label: '固定希望店舗',
    description: '1ヶ月間ずっと同じ店舗に割り当てられた等でドライバーマスタの「固定希望店舗」に設定されている店舗を最優先する',
    test: (a, store) => a.fixed_store_id === store.store_id
  },
  group_fixed: {
    label: 'グループの紐付け店舗',
    description: '「2店舗を3人で回す」のように、ドライバーグループに紐付けた店舗を優先する(1人のドライバーが複数のグループに所属していてもよい)',
    test: (a, store, ctx) => (ctx.groupStoreIdsByDriverId.get(a.driver_id) || new Set()).has(store.store_id)
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
const DEFAULT_ENABLED_TIER_KEYS = ['fixed_store', 'group_fixed']; // 画面でまだ設定したことが無い場合の初期値

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

// 固定希望店舗が複数ドライバーで重複している組のうち、優先順位が未確定なもの(未設定のドライバーが
// いる、または同じ数字が重複している)を返す。自動マッチング実行前のチェックと、ドライバーマスタの
// 重複アラート表示の両方から使う
async function findUnresolvedFixedStoreDuplicates() {
  const drivers = await dbAll(`
    SELECT d.id, d.name, d.fixed_store_id, d.fixed_store_priority, s.name AS store_name
    FROM drivers d JOIN stores s ON s.id = d.fixed_store_id
  `);
  const byStore = new Map();
  for (const d of drivers) {
    if (!byStore.has(d.fixed_store_id)) byStore.set(d.fixed_store_id, []);
    byStore.get(d.fixed_store_id).push(d);
  }
  const unresolved = [];
  for (const [storeId, group] of byStore) {
    if (group.length <= 1) continue;
    const priorities = group.map(d => d.fixed_store_priority);
    const hasUnset = priorities.some(p => p == null);
    const setPriorities = priorities.filter(p => p != null);
    const hasDuplicate = new Set(setPriorities).size !== setPriorities.length;
    if (hasUnset || hasDuplicate) {
      unresolved.push({
        store_id: storeId,
        store_name: group[0].store_name,
        drivers: group.map(d => ({ id: d.id, name: d.name, priority: d.fixed_store_priority }))
      });
    }
  }
  return unresolved;
}

// 自動マッチングの対象を「固定希望店舗が設定されているドライバーのみ」に絞る設定(画面からON/OFFできる)。
// 担当者と相談の上、それ以外のドライバーは手作業で割り当てる運用に切り替えられるようにするため。
// 初期値はOFF(従来通り全員を対象にする)
async function getRestrictMatchingToFixedStore() {
  const row = await dbGet('SELECT value FROM app_settings WHERE key = ?', ['restrict_matching_to_fixed_store']);
  return row ? row.value === '1' : false;
}

// 「今日の日付」をアプリ内だけ任意の日付として扱えるようにする設定(PC本体の時刻は変更しない)。
// 月末の運用(来月分の取込み・マッチング確認)を本番の月末を待たずに練習・確認したい時に使う想定。
// 現時点ではマッチングや自動完了の判定自体は「今日」を使っていないため、主に画面の日付欄の初期値や
// 「今日は◯月◯日として扱っています」という表示に使われる
async function getMockToday() {
  const row = await dbGet('SELECT value FROM app_settings WHERE key = ?', ['mock_today']);
  return row && row.value ? row.value : null; // 'YYYY-MM-DD' または未設定ならnull(本当の今日を使う)
}

// 個人事業主向けPDF(作業依頼表)の「依頼者」「担当者」「電話番号」欄の表記。画面(ダッシュボード)
// から変更できるようにしてある(会社名・担当者が変わった時にコードを直さずに済むように)
const PDF_HEADER_DEFAULTS = {
  requester_name: '株式会社 ひとまいるロジスティクス',
  requester_address: '埼玉県和光市新倉7-7-25',
  contact_name: '上北　渉',
  contact_phone: '090-9146-9122'
};
async function getPdfHeaderSettings() {
  const row = await dbGet('SELECT value FROM app_settings WHERE key = ?', ['pdf_header_settings']);
  if (!row || !row.value) return { ...PDF_HEADER_DEFAULTS };
  try { return { ...PDF_HEADER_DEFAULTS, ...JSON.parse(row.value) }; } catch (e) { return { ...PDF_HEADER_DEFAULTS }; }
}

// 個人事業主向けPDF(作業依頼表)の一番下に記載する注記文言。テンプレート本来の文言を初期値にしてある
const PDF_FOOTER_DEFAULT = '本発注書に係る支払い条件そのほかの取引条件は、別途締結済みの「業務委託契約書」(令和8年03月02日付け)に基づくものとします。';
async function getPdfFooterNote() {
  const row = await dbGet('SELECT value FROM app_settings WHERE key = ?', ['pdf_footer_note']);
  return (row && row.value) ? row.value : PDF_FOOTER_DEFAULT;
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

// "HH:MM"形式の時刻を分に変換する(不正な形式はnullを返す)
function parseTimeToMinutes(t) {
  if (!t) return null;
  const m = String(t).trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}
function minutesToTimeStr(min) {
  return `${String(Math.floor(min / 60)).padStart(2, '0')}:${String(min % 60).padStart(2, '0')}`;
}

// 店舗依頼の希望時間帯(storeTimeStart〜storeTimeEnd)のうち、既存マッチング(それぞれ自身の
// time_start/time_endを持つ。無ければ店舗の時間でカバーしているとみなし、今まで通りの前提に合わせる)
// でまだカバーされていない「穴」の時間帯を返す([{start,end}, ...]、無ければ空配列)。
// 店舗の時間が未設定・不正(終了が開始以前等)の場合は判定できないため空配列を返す
function computeTimeGaps(storeTimeStart, storeTimeEnd, matches) {
  const storeStart = parseTimeToMinutes(storeTimeStart);
  const storeEnd = parseTimeToMinutes(storeTimeEnd);
  if (storeStart == null || storeEnd == null || storeEnd <= storeStart) return [];
  const intervals = matches
    .map(m => [parseTimeToMinutes(m.time_start) ?? storeStart, parseTimeToMinutes(m.time_end) ?? storeEnd])
    .filter(([s, e]) => e > s)
    .sort((a, b) => a[0] - b[0]);
  const gaps = [];
  let cursor = storeStart;
  for (const [s, e] of intervals) {
    if (s > cursor) gaps.push([cursor, Math.min(s, storeEnd)]);
    cursor = Math.max(cursor, e);
    if (cursor >= storeEnd) break;
  }
  if (cursor < storeEnd) gaps.push([cursor, storeEnd]);
  return gaps.filter(([s, e]) => e > s).map(([s, e]) => ({ start: minutesToTimeStr(s), end: minutesToTimeStr(e) }));
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

// Excel上で非表示設定になっているシート名の集合を返す(hidden=1, very hidden=2)。
// 非表示シートは古い下書き・参考データであることが多く、取込み先として自動選択・候補表示の
// どちらにも出したくない(実際に「これ聞取り固定デポ」という非表示の古いシートが自動選択されて
// しまい、間違ったデータが取り込まれる事故が起きたため導入した)
function getHiddenSheetNames(workbook) {
  const hidden = new Set();
  for (const s of (workbook.Workbook && workbook.Workbook.Sheets) || []) {
    if (s.Hidden) hidden.add(s.name);
  }
  return hidden;
}

// 「氏名×日付」のワイド形式(1行目=日付見出し、1列目=氏名、セル=勤務時間 or 休み)のシフト表を
// 見出し行+各行を配列のまま返す(通常のparseUploadedSpreadsheetは列名をキーにしたオブジェクトを返すため、
// 「1(火)」のような日付見出しを順序どおり扱いたいこちらの用途には配列のままの形が必要)
// sheetNameを指定すればそのシートを、未指定(またはファイル内に無い名前)なら先頭シートを読む
// ブックの中から、見出しにmarkerHeadersのいずれかを含む「表示状態の」シート名を探して返す(xlsx以外や、
// 一致するシートが無い場合はnull)。1つのExcelファイルに複数シートがあり(例: 1枚目が週間必要枠表、
// 2枚目がエリア固定の依頼、3枚目はその他のデータ)、取込み先ごとに読むべきシートが違う場合に、
// シート名を手で選ばなくても見出しの中身からそれらしいシートを自動で見つけるために使う
function findMarkerSheetName(buffer, filename, markerHeaders) {
  const ext = (filename || '').toLowerCase();
  if (ext.endsWith('.csv') || ext.endsWith('.txt') || !markerHeaders || !markerHeaders.length) return null;
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const hiddenNames = getHiddenSheetNames(workbook);
  return workbook.SheetNames.filter(name => !hiddenNames.has(name)).find(name => {
    const candidateRows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: false }).slice(0, 15);
    return candidateRows.some(row => row.some(cell => markerHeaders.includes(String(cell ?? '').trim())));
  }) || null;
}

function parseWideSpreadsheet(buffer, filename, sheetName, markerHeaders) {
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
  let targetName = (sheetName && workbook.SheetNames.includes(sheetName)) ? sheetName : null;
  if (!targetName) targetName = findMarkerSheetName(buffer, filename, markerHeaders);
  if (!targetName) targetName = workbook.SheetNames[0];
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
  furigana: ['ふりがな', 'フリガナ', 'よみがな', 'ヨミガナ', 'furigana'],
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

// ===== ログイン・利用者マスタ =====
// 利用者が1人もいない場合だけ、初回セットアップとして最初の1人を自分で登録できるようにする
// (このエンドポイント自体が、パスワードを誰かに教えてもらわずに最初の自分のアカウントを作る唯一の手段)
app.get('/api/auth/setup-needed', async (req, res) => {
  const count = await dbGet('SELECT COUNT(*) AS c FROM app_users');
  res.json({ success: true, needed: count.c === 0 });
});

app.post('/api/auth/setup', async (req, res) => {
  const count = await dbGet('SELECT COUNT(*) AS c FROM app_users');
  if (count.c > 0) return res.status(403).json({ success: false, message: '利用者が既に登録されているため、初回セットアップは使えません(利用者マスタ画面から追加してください)' });
  const { name, login_id, password } = req.body || {};
  if (!name || !login_id || !password) return res.status(400).json({ success: false, message: '氏名・ログインID・パスワードは必須です' });
  const result = await dbRun(
    'INSERT INTO app_users (name, login_id, password_hash, created_at) VALUES (?, ?, ?, ?)',
    [name, login_id, hashPassword(password), new Date().toISOString()]
  );
  const token = crypto.randomBytes(32).toString('hex');
  await dbRun('INSERT INTO app_sessions (token, user_id, created_at) VALUES (?, ?, ?)', [token, result.lastID, new Date().toISOString()]);
  setSessionCookie(res, token);
  res.json({ success: true, user: { id: result.lastID, name, login_id } });
});

// ブルートフォース対策: 連続して何回失敗したらロックするか、ロック時間はどれくらいか
const LOGIN_MAX_ATTEMPTS = 5;
const LOGIN_LOCKOUT_MINUTES = 15;

app.post('/api/auth/login', async (req, res) => {
  const { login_id, password } = req.body || {};
  if (!login_id || !password) return res.status(400).json({ success: false, message: 'ログインIDとパスワードを入力してください' });
  const user = await dbGet('SELECT * FROM app_users WHERE login_id = ?', [login_id]);

  // 既にロック中なら、パスワードを確認するまでもなく拒否する(ロック中はロック自体を延長しない)
  if (user && user.locked_until && new Date(user.locked_until) > new Date()) {
    const remainingMin = Math.ceil((new Date(user.locked_until) - new Date()) / 60000);
    return res.status(429).json({ success: false, message: `ログイン試行が多すぎるため、一時的にロックされています。約${remainingMin}分後に再度お試しください。` });
  }

  if (!user || !verifyPassword(password, user.password_hash)) {
    if (user) {
      const failedCount = (user.failed_login_count || 0) + 1;
      if (failedCount >= LOGIN_MAX_ATTEMPTS) {
        const lockedUntil = new Date(Date.now() + LOGIN_LOCKOUT_MINUTES * 60000).toISOString();
        await dbRun('UPDATE app_users SET failed_login_count = 0, locked_until = ? WHERE id = ?', [lockedUntil, user.id]);
        return res.status(429).json({ success: false, message: `ログインに${LOGIN_MAX_ATTEMPTS}回連続で失敗したため、${LOGIN_LOCKOUT_MINUTES}分間ロックされました。` });
      }
      await dbRun('UPDATE app_users SET failed_login_count = ? WHERE id = ?', [failedCount, user.id]);
    }
    return res.status(401).json({ success: false, message: 'ログインIDまたはパスワードが違います' });
  }

  // ログイン成功時は失敗回数・ロックをリセットする
  if (user.failed_login_count || user.locked_until) {
    await dbRun('UPDATE app_users SET failed_login_count = 0, locked_until = NULL WHERE id = ?', [user.id]);
  }
  const token = crypto.randomBytes(32).toString('hex');
  await dbRun('INSERT INTO app_sessions (token, user_id, created_at) VALUES (?, ?, ?)', [token, user.id, new Date().toISOString()]);
  setSessionCookie(res, token);
  res.json({ success: true, user: { id: user.id, name: user.name, login_id: user.login_id } });
});

app.post('/api/auth/logout', async (req, res) => {
  const token = parseCookies(req)[SESSION_COOKIE_NAME];
  if (token) await dbRun('DELETE FROM app_sessions WHERE token = ?', [token]);
  clearSessionCookie(res);
  res.json({ success: true });
});

app.get('/api/auth/me', (req, res) => {
  res.json({ success: true, user: req.user || null });
});

// 利用者マスタのCRUD(パスワードのハッシュは一覧・取得のレスポンスに含めない)
app.get('/api/users', async (req, res) => {
  const users = await dbAll('SELECT id, name, login_id, created_at FROM app_users ORDER BY id');
  res.json({ success: true, users });
});

app.post('/api/users', async (req, res) => {
  const { name, login_id, password } = req.body || {};
  if (!name || !login_id || !password) return res.status(400).json({ success: false, message: '氏名・ログインID・パスワードは必須です' });
  const existing = await dbGet('SELECT id FROM app_users WHERE login_id = ?', [login_id]);
  if (existing) return res.status(400).json({ success: false, message: 'このログインIDは既に使われています' });
  const result = await dbRun(
    'INSERT INTO app_users (name, login_id, password_hash, created_at) VALUES (?, ?, ?, ?)',
    [name, login_id, hashPassword(password), new Date().toISOString()]
  );
  res.json({ success: true, id: result.lastID });
});

app.put('/api/users/:id', async (req, res) => {
  const { name, login_id, password } = req.body || {};
  if (!name || !login_id) return res.status(400).json({ success: false, message: '氏名・ログインIDは必須です' });
  const dup = await dbGet('SELECT id FROM app_users WHERE login_id = ? AND id != ?', [login_id, req.params.id]);
  if (dup) return res.status(400).json({ success: false, message: 'このログインIDは既に使われています' });
  if (password) {
    await dbRun('UPDATE app_users SET name=?, login_id=?, password_hash=? WHERE id=?', [name, login_id, hashPassword(password), req.params.id]);
  } else {
    await dbRun('UPDATE app_users SET name=?, login_id=? WHERE id=?', [name, login_id, req.params.id]);
  }
  res.json({ success: true });
});

app.delete('/api/users/:id', async (req, res) => {
  const count = await dbGet('SELECT COUNT(*) AS c FROM app_users');
  if (count.c <= 1) return res.status(400).json({ success: false, message: '利用者が1人もいなくなるとログインできなくなるため、最後の1人は削除できません' });
  await dbRun('DELETE FROM app_users WHERE id = ?', [req.params.id]);
  await dbRun('DELETE FROM app_sessions WHERE user_id = ?', [req.params.id]);
  res.json({ success: true });
});

// ===== ドライバーマスタ =====
app.get('/api/drivers', async (req, res) => {
  const rows = await dbAll(`
    SELECT d.*, fs.name AS fixed_store_name
    FROM drivers d
    LEFT JOIN stores fs ON fs.id = d.fixed_store_id
    ORDER BY d.id DESC
  `);
  // 所属グループ(複数所属あり得る)を一覧表示用にまとめて付与する
  const memberships = await dbAll(`
    SELECT m.driver_id, g.id, g.name FROM driver_group_members m JOIN driver_groups g ON g.id = m.group_id
  `);
  const groupsByDriverId = new Map();
  for (const m of memberships) {
    if (!groupsByDriverId.has(m.driver_id)) groupsByDriverId.set(m.driver_id, []);
    groupsByDriverId.get(m.driver_id).push({ id: m.id, name: m.name });
  }
  for (const d of rows) d.groups = groupsByDriverId.get(d.id) || [];
  res.json({ success: true, drivers: rows });
});

app.post('/api/drivers', async (req, res) => {
  const { id, name, furigana, phone, vehicle_type, home_address, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes } = req.body;
  if (!name) return res.status(400).json({ success: false, message: '氏名は必須です' });

  const existing = id ? await dbGet('SELECT home_address, home_lat, home_lng, fixed_store_id FROM drivers WHERE id = ?', [id]) : null;

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

  const now = new Date().toISOString();
  if (id) {
    const newFixedStoreId = fixed_store_id || null;
    if (existing && existing.fixed_store_id !== newFixedStoreId) {
      await logFixedStoreChange(id, existing.fixed_store_id, newFixedStoreId, 'manual');
    }
    await dbRun(
      `UPDATE drivers SET name=?, furigana=?, phone=?, vehicle_type=?, home_address=?, home_lat=?, home_lng=?, driver_code=?, company_name=?, fixed_store_id=?, first_contract_date=?, status=?, email=?, insurance_info=?, notes=? WHERE id=?`,
      [name, furigana || '', phone || '', vehicle_type || '', home_address || '', lat, lng, driver_code || null, company_name || '', newFixedStoreId, first_contract_date || '', status || '', email || '', insurance_info || '', notes || '', id]
    );
    return res.json({ success: true, id });
  }
  const result = await dbRun(
    `INSERT INTO drivers (name, furigana, phone, vehicle_type, home_address, home_lat, home_lng, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [name, furigana || '', phone || '', vehicle_type || '', home_address || '', lat, lng, driver_code || null, company_name || '', fixed_store_id || null, first_contract_date || '', status || '', email || '', insurance_info || '', notes || '', now]
  );
  if (fixed_store_id) await logFixedStoreChange(result.lastID, null, fixed_store_id, 'manual');
  res.json({ success: true, id: result.lastID });
});

// 固定希望店舗の変更履歴を返す(手動設定か、どの自動反映によるものかを確認できるようにする)。
// driver_idを指定すると、そのドライバーの分だけに絞り込む(編集画面にその場で表示する用)
app.get('/api/logs/fixed-store-changes', async (req, res) => {
  const driverId = req.query.driver_id ? parseInt(req.query.driver_id, 10) : null;
  const rows = await dbAll(`
    SELECT l.*, d.name AS driver_name, os.name AS old_store_name, ns.name AS new_store_name
    FROM fixed_store_change_log l
    JOIN drivers d ON d.id = l.driver_id
    LEFT JOIN stores os ON os.id = l.old_store_id
    LEFT JOIN stores ns ON ns.id = l.new_store_id
    ${driverId ? 'WHERE l.driver_id = ?' : ''}
    ORDER BY l.id DESC
    LIMIT 200
  `, driverId ? [driverId] : []);
  res.json({ success: true, logs: rows });
});

// ドライバーマスタ一覧から、固定希望店舗だけをその場で直接編集できるようにする専用エンドポイント。
// (氏名など他の項目は一切触らない。/api/drivers(POST)は全項目を送る前提の更新なので、
// この用途にそのまま使うと他の項目を空で上書きしてしまう危険があるため、別エンドポイントにしている)
app.post('/api/drivers/:id/fixed-store', async (req, res) => {
  const id = parseInt(req.params.id, 10);
  const newFixedStoreId = req.body.fixed_store_id ? parseInt(req.body.fixed_store_id, 10) : null;
  const existing = await dbGet('SELECT fixed_store_id FROM drivers WHERE id = ?', [id]);
  if (!existing) return res.status(404).json({ success: false, message: 'ドライバーが見つかりません' });
  if (existing.fixed_store_id !== newFixedStoreId) {
    await dbRun('UPDATE drivers SET fixed_store_id = ? WHERE id = ?', [newFixedStoreId, id]);
    await logFixedStoreChange(id, existing.fixed_store_id, newFixedStoreId, 'manual');
  }
  res.json({ success: true });
});

// 固定希望店舗の「候補」一覧(保留中のみ、新しい順)。driver_idを指定すると、そのドライバーの分だけに絞る
// (ドライバーマスタ一覧に候補バッジを出す時に使う)
app.get('/api/fixed-store-suggestions', async (req, res) => {
  const driverId = req.query.driver_id ? parseInt(req.query.driver_id, 10) : null;
  const rows = await dbAll(`
    SELECT fs.*, d.name AS driver_name, s.name AS suggested_store_name
    FROM fixed_store_suggestions fs
    JOIN drivers d ON d.id = fs.driver_id
    JOIN stores s ON s.id = fs.suggested_store_id
    WHERE fs.status = 'pending' ${driverId ? 'AND fs.driver_id = ?' : ''}
    ORDER BY fs.id DESC
  `, driverId ? [driverId] : []);
  res.json({ success: true, suggestions: rows });
});

// 固定希望店舗の候補を採用する(=その内容でドライバーマスタのfixed_store_idを実際に設定する、唯一の経路)。
// コーディネーターがドライバーマスタ画面でボタンを押した時だけ呼ばれる、明示的な手動操作
app.post('/api/fixed-store-suggestions/:id/apply', async (req, res) => {
  const id = parseInt(req.params.id, 10);
  const suggestion = await dbGet(`SELECT * FROM fixed_store_suggestions WHERE id = ? AND status = 'pending'`, [id]);
  if (!suggestion) return res.status(404).json({ success: false, message: '候補が見つかりません(既に処理済みかもしれません)' });
  const driver = await dbGet('SELECT fixed_store_id FROM drivers WHERE id = ?', [suggestion.driver_id]);
  if (!driver) return res.status(404).json({ success: false, message: 'ドライバーが見つかりません' });
  await dbRun('UPDATE drivers SET fixed_store_id = ? WHERE id = ?', [suggestion.suggested_store_id, suggestion.driver_id]);
  if (driver.fixed_store_id !== suggestion.suggested_store_id) {
    await logFixedStoreChange(suggestion.driver_id, driver.fixed_store_id, suggestion.suggested_store_id, 'manual_from_suggestion');
  }
  await dbRun(`UPDATE fixed_store_suggestions SET status = 'applied', resolved_at = ? WHERE id = ?`, [new Date().toISOString(), id]);
  // 同じドライバーへの他の保留中候補も、採用した以上は役目を終えたとみなして消しておく(放置されたままにならないように)
  await dbRun(`UPDATE fixed_store_suggestions SET status = 'dismissed', resolved_at = ? WHERE driver_id = ? AND status = 'pending'`, [new Date().toISOString(), suggestion.driver_id]);
  res.json({ success: true });
});

// 固定希望店舗の候補を却下する(この内容では設定しない、という意思表示。ドライバーマスタは変更しない)
app.post('/api/fixed-store-suggestions/:id/dismiss', async (req, res) => {
  const id = parseInt(req.params.id, 10);
  const result = await dbRun(`UPDATE fixed_store_suggestions SET status = 'dismissed', resolved_at = ? WHERE id = ? AND status = 'pending'`, [new Date().toISOString(), id]);
  if (result.changes === 0) return res.status(404).json({ success: false, message: '候補が見つかりません(既に処理済みかもしれません)' });
  res.json({ success: true });
});

// 一括取込みの履歴一覧(新しい順、最大50件)。取り消し済みかどうか(undone_at)、現時点でまだ
// このバッチの印が付いたまま残っている行数(currentRowCount。手動削除や月次アーカイブで減ることがある)を返す
app.get('/api/import-batches', async (req, res) => {
  const batches = await dbAll(`SELECT * FROM import_batches ORDER BY id DESC LIMIT 50`);
  for (const b of batches) {
    if (b.type === 'store_requests_weekly' || b.type === 'store_requests') {
      b.currentRowCount = (await dbGet('SELECT COUNT(*) AS c FROM store_requests WHERE import_batch_id = ? AND archived_month IS NULL', [b.id])).c;
    } else if (b.type === 'driver_availability') {
      b.currentRowCount = (await dbGet('SELECT COUNT(*) AS c FROM driver_availability WHERE import_batch_id = ? AND archived_month IS NULL', [b.id])).c;
    } else if (b.type === 'area_fixed') {
      b.currentRowCount = (await dbGet('SELECT COUNT(*) AS c FROM drivers WHERE area_fixed_batch_id = ?', [b.id])).c;
    } else {
      b.currentRowCount = null;
    }
  }
  res.json({ success: true, batches });
});

// 1回分の取込みをまとめて取り消す。対象行は import_batch_id(店舗依頼/希望シフト)または
// area_fixed_batch_id(エリア固定)でこのバッチの印が付いているものに限る(後から別の取込みで
// 上書きされた行は印が最新の取込みに付け替わっているため、対象にならず誤って巻き込まない)
app.post('/api/import-batches/:id/undo', async (req, res) => {
  const batch = await dbGet('SELECT * FROM import_batches WHERE id = ?', [req.params.id]);
  if (!batch) return res.status(404).json({ success: false, message: '取込み履歴が見つかりません' });
  if (batch.undone_at) return res.status(400).json({ success: false, message: 'この取込みは既に取り消し済みです' });
  const now = new Date().toISOString();

  if (batch.type === 'store_requests_weekly' || batch.type === 'store_requests') {
    const targets = await dbAll('SELECT id FROM store_requests WHERE import_batch_id = ? AND archived_month IS NULL', [batch.id]);
    const targetIds = targets.map(t => t.id);
    if (targetIds.length === 0) {
      await dbRun('UPDATE import_batches SET undone_at = ? WHERE id = ?', [now, batch.id]);
      return res.json({ success: true, deletedStoreRequests: 0, deletedMatches: 0 });
    }
    const placeholders = targetIds.map(() => '?').join(',');
    const confirmedCount = await dbGet(
      `SELECT COUNT(*) AS c FROM matches WHERE store_request_id IN (${placeholders}) AND status IN ('確定','完了') AND archived_month IS NULL`,
      targetIds
    );
    if (confirmedCount.c > 0) {
      return res.status(400).json({
        success: false,
        message: `この取込みで作った店舗依頼に、既に「確定」または「完了」のマッチングが${confirmedCount.c}件あるため取り消せません。先にそのマッチングを取り消す(候補に戻す/削除する)などで確認してから、もう一度お試しください。`
      });
    }
    const deletedMatches = await dbRun(`DELETE FROM matches WHERE store_request_id IN (${placeholders}) AND status = '候補'`, targetIds);
    await dbRun(`DELETE FROM store_requests WHERE id IN (${placeholders})`, targetIds);
    await dbRun('UPDATE import_batches SET undone_at = ? WHERE id = ?', [now, batch.id]);
    return res.json({ success: true, deletedStoreRequests: targetIds.length, deletedMatches: deletedMatches.changes || 0 });
  }

  if (batch.type === 'driver_availability') {
    const result = await dbRun('DELETE FROM driver_availability WHERE import_batch_id = ? AND archived_month IS NULL', [batch.id]);
    await dbRun('UPDATE import_batches SET undone_at = ? WHERE id = ?', [now, batch.id]);
    return res.json({ success: true, deletedAvailability: result.changes || 0 });
  }

  if (batch.type === 'area_fixed') {
    const result = await dbRun(
      `UPDATE drivers SET area_fixed_enabled = 0, area_fixed_pattern = NULL, area_fixed_store_ids = NULL, area_fixed_batch_id = NULL WHERE area_fixed_batch_id = ?`,
      [batch.id]
    );
    await dbRun('UPDATE import_batches SET undone_at = ? WHERE id = ?', [now, batch.id]);
    return res.json({ success: true, clearedDrivers: result.changes || 0 });
  }

  return res.status(400).json({ success: false, message: '未対応の取込み種類です' });
});

// 練習用: 現在作業中(未アーカイブ)の店舗依頼・希望シフト・マッチング結果をまとめて削除し、
// 何も取り込んでいない状態に戻す(使い方タブのボタンから呼ぶ)。確定・完了済みのマッチングも
// 含めて消す点が取込み履歴の個別取消しと異なる(練習のやり直し専用なので安全確認はせず全部消す)。
// ドライバーマスタ・店舗マスタ・派遣履歴・各種設定・アーカイブ済みの過去月データは対象外
app.post('/api/practice-reset', async (req, res) => {
  const matchesResult = await dbRun(`DELETE FROM matches WHERE archived_month IS NULL`);
  const storeRequestsResult = await dbRun(`DELETE FROM store_requests WHERE archived_month IS NULL`);
  const availabilityResult = await dbRun(`DELETE FROM driver_availability WHERE archived_month IS NULL`);
  await dbRun(
    `UPDATE import_batches SET undone_at = ? WHERE undone_at IS NULL AND type IN ('store_requests_weekly', 'store_requests', 'driver_availability')`,
    [new Date().toISOString()]
  );
  res.json({
    success: true,
    deleted: { matches: matchesResult.changes || 0, storeRequests: storeRequestsResult.changes || 0, availability: availabilityResult.changes || 0 }
  });
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
    const furigana = pickField(row, DRIVER_MASTER_ALIASES.furigana);
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
        `UPDATE drivers SET name=?, furigana=?, phone=?, vehicle_type=?, home_address=?, home_lat=?, home_lng=?, driver_code=?, company_name=?, fixed_store_id=?, first_contract_date=?, status=?, email=?, insurance_info=?, notes=? WHERE id=?`,
        [name, furigana, phone, vehicle_type, home_address, lat, lng, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes, existing.id]
      );
    } else {
      await dbRun(
        `INSERT INTO drivers (name, furigana, phone, vehicle_type, home_address, home_lat, home_lng, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [name, furigana, phone, vehicle_type, home_address, lat, lng, driver_code, company_name, fixed_store_id, first_contract_date, status, email, insurance_info, notes, now]
      );
    }
    imported++;
  }

  res.json({ success: true, imported, total: rows.length, errors });
});

// ドライバーマスタをExcelでエクスポートする(他のPCへ渡して取込み直せるよう、一括取込みと同じ列構成・列名にしてある)
app.get('/api/export/drivers.xlsx', async (req, res) => {
  const drivers = await dbAll(`
    SELECT d.*, fs.name AS fixed_store_name
    FROM drivers d LEFT JOIN stores fs ON fs.id = d.fixed_store_id
    ORDER BY d.name
  `);
  const header = ['社員コード', '氏名', '会社名', 'お住まい住所', '初回委託日', 'ステータス', 'メールアドレス', '固定希望店舗', '保険加入状況', '電話番号', '車両種別', '備考'];
  const rows = drivers.map(d => [
    d.driver_code || '', d.name || '', d.company_name || '', d.home_address || '',
    d.first_contract_date || '', d.status || '', d.email || '', d.fixed_store_name || '',
    d.insurance_info || '', d.phone || '', d.vehicle_type || '', d.notes || ''
  ]);
  const wb = new ExcelJS.Workbook();
  addFilledSampleSheet(wb, 'ドライバーマスタ', header, rows, [10, 20, 14, 24, 14, 10, 18, 14, 14, 16, 16, 36]);
  const buffer = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="driver_master_${new Date().toISOString().slice(0, 10)}.xlsx"`);
  res.send(Buffer.from(buffer));
});

app.delete('/api/drivers/:id', async (req, res) => {
  await dbRun('DELETE FROM drivers WHERE id = ?', [req.params.id]);
  // 過去の月としてアーカイブ済みの希望シフトは、ドライバーがマスタから削除されても履歴として残す
  // (archived_month IS NULLの、現在進行中の希望シフトのみ削除する)
  await dbRun('DELETE FROM driver_availability WHERE driver_id = ? AND archived_month IS NULL', [req.params.id]);
  res.json({ success: true });
});

// 固定希望店舗が重複しているドライバー間の優先順位をまとめて保存する(ドライバーマスタの
// 重複アラートからの保存専用。updates: [{driver_id, priority(数値またはnull)}, ...])
app.post('/api/drivers/fixed-store-priorities', async (req, res) => {
  const { updates } = req.body || {};
  if (!Array.isArray(updates) || updates.length === 0) return res.status(400).json({ success: false, message: 'updatesは1件以上の配列で指定してください' });
  for (const u of updates) {
    await dbRun('UPDATE drivers SET fixed_store_priority = ? WHERE id = ?', [u.priority != null && u.priority !== '' ? parseInt(u.priority, 10) : null, u.driver_id]);
  }
  res.json({ success: true });
});

// 固定希望店舗なのに、その店舗を相性でNGに設定している矛盾した組み合わせを一覧で返す。
// 自動マッチングではNGが必ず優先される(固定希望店舗・優先順位があっても除外される)ため、
// 入力ミスや運用上の矛盾に気付けるよう、ドライバーマスタ画面に警告として表示する
app.get('/api/drivers/fixed-store-ng-conflicts', async (req, res) => {
  const rows = await dbAll(`
    SELECT d.id AS driver_id, d.name AS driver_name, d.fixed_store_id AS store_id, s.name AS store_name, p.id AS preference_id
    FROM drivers d
    JOIN driver_store_preferences p ON p.driver_id = d.id AND p.store_id = d.fixed_store_id
    JOIN stores s ON s.id = d.fixed_store_id
    WHERE p.preference = 'NG'
    ORDER BY d.name ASC
  `);
  res.json({ success: true, conflicts: rows });
});

// ===== ドライバーの希望シフト =====
app.get('/api/driver-availability', async (req, res) => {
  const rows = await dbAll(`
    SELECT a.*, d.name AS driver_name, d.phone AS driver_phone, d.driver_code AS driver_code
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
  const importBatchId = await createImportBatch('driver_availability', fixUploadedFilename(req.file.originalname), '長形式');

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
        `UPDATE driver_availability SET desired_area=?, desired_store=?, time_start=?, time_end=?, requests=?, import_batch_id=? WHERE id=?`,
        [desired_area, desired_store, time_start, time_end, requests, importBatchId, existing.id]
      );
    } else {
      await dbRun(
        `INSERT INTO driver_availability (driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests, created_at, import_batch_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests, now, importBatchId]
      );
    }
    imported++;
  }
  await finalizeImportBatch(importBatchId, imported);

  res.json({ success: true, imported, total: rows.length, errors, importBatchId });
});

// アップロードされたExcelファイルのシート名一覧を返す(CSVは単一シート扱いで空配列)。
// 1ファイルに複数シートがある場合に、取込み前にどのシートを使うか選べるようにするため。
// markerを指定すると、その見出し(例:「店舗名」)を含むシートだけをcandidatesとして別途返す
// (1つのファイルに同じ見出しを持つシートが複数あると、自動選択(findMarkerSheetNameと同じロジック)が
// 意図しない方のシートを選んでしまうことがあるため、複数ヒットした場合は画面側で選び直せるようにする)
app.post('/api/sheet-names', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, message: 'ファイルが必要です' });
  const ext = (req.file.originalname || '').toLowerCase();
  if (ext.endsWith('.csv') || ext.endsWith('.txt')) return res.json({ success: true, sheets: [], candidates: [] });
  try {
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    // Excel上で非表示にしているシートは、古い下書き・参考データであることが多く選択肢に出しても
    // 紛らわしいだけなので除外する(実際にこれが原因で、見えているシート数より選択肢が多く
    // 見えて混乱を招いた上、非表示の古いシートが誤って自動選択される事故も起きたため)
    const hiddenNames = getHiddenSheetNames(workbook);
    const visibleSheetNames = workbook.SheetNames.filter(name => !hiddenNames.has(name));
    const marker = (req.body.marker || '').trim();
    let candidates = [];
    if (marker) {
      candidates = visibleSheetNames.filter(name => {
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: false }).slice(0, 15);
        return rows.some(row => row.some(cell => String(cell ?? '').trim() === marker));
      });
    }
    res.json({ success: true, sheets: visibleSheetNames, candidates });
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
// multerが受け取るアップロードファイル名(multipart/form-dataのfilename)は、ブラウザがUTF-8の
// バイト列をそのまま送ってくる一方、multer(内部のbusboy)は仕様上それをlatin1として解釈してしまうため、
// 日本語ファイル名が文字化けする。Latin-1は1バイト=1文字なので情報の欠落なく元のバイト列に戻せる
// (取込み履歴のファイル名表示が文字化けしていた不具合の原因)
function fixUploadedFilename(name) {
  if (!name) return name;
  try { return Buffer.from(name, 'latin1').toString('utf8'); } catch (e) { return name; }
}

function getImportSource(req) {
  if (req.file) return { buffer: req.file.buffer, filename: fixUploadedFilename(req.file.originalname) };
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

// ===== ドライバーグループ(旧「エリア固定」の置き換え) =====
// 「2店舗を3人で回す」のように、複数のドライバーと複数の店舗をまとめて紐付けて管理する。
// マッチング時は、グループの誰か(希望シフトを出している人)を、グループに紐付いた店舗に
// 優先的に割り当てる(優先順位は画面の「マッチング設定」でどこに置くか決められる。
// 初期値は固定希望店舗の次)。1人のドライバーが複数のグループに所属することもできる
app.get('/api/driver-groups', async (req, res) => {
  const groups = await dbAll('SELECT * FROM driver_groups ORDER BY id DESC');
  const storeRows = await dbAll(`
    SELECT gs.group_id, s.id, s.name FROM driver_group_stores gs JOIN stores s ON s.id = gs.store_id
  `);
  const memberRows = await dbAll(`
    SELECT m.group_id, d.id, d.name FROM driver_group_members m JOIN drivers d ON d.id = m.driver_id
  `);
  const storesByGroup = new Map();
  for (const r of storeRows) { if (!storesByGroup.has(r.group_id)) storesByGroup.set(r.group_id, []); storesByGroup.get(r.group_id).push({ id: r.id, name: r.name }); }
  const membersByGroup = new Map();
  for (const r of memberRows) { if (!membersByGroup.has(r.group_id)) membersByGroup.set(r.group_id, []); membersByGroup.get(r.group_id).push({ id: r.id, name: r.name }); }
  for (const g of groups) {
    g.stores = storesByGroup.get(g.id) || [];
    g.members = membersByGroup.get(g.id) || [];
  }
  res.json({ success: true, groups });
});

app.post('/api/driver-groups', async (req, res) => {
  const { id, name, store_ids, driver_ids } = req.body;
  if (!name) return res.status(400).json({ success: false, message: 'グループ名は必須です' });
  const storeIds = Array.isArray(store_ids) ? store_ids : [];
  const driverIds = Array.isArray(driver_ids) ? driver_ids : [];
  if (storeIds.length === 0) return res.status(400).json({ success: false, message: '店舗を1件以上選んでください' });
  if (driverIds.length === 0) return res.status(400).json({ success: false, message: 'ドライバーを1人以上選んでください' });

  let groupId = id;
  if (groupId) {
    await dbRun('UPDATE driver_groups SET name = ? WHERE id = ?', [name, groupId]);
    await dbRun('DELETE FROM driver_group_stores WHERE group_id = ?', [groupId]);
    await dbRun('DELETE FROM driver_group_members WHERE group_id = ?', [groupId]);
  } else {
    const result = await dbRun('INSERT INTO driver_groups (name, created_at) VALUES (?, ?)', [name, new Date().toISOString()]);
    groupId = result.lastID;
  }
  for (const storeId of storeIds) await dbRun('INSERT INTO driver_group_stores (group_id, store_id) VALUES (?, ?)', [groupId, storeId]);
  for (const driverId of driverIds) await dbRun('INSERT INTO driver_group_members (group_id, driver_id) VALUES (?, ?)', [groupId, driverId]);
  res.json({ success: true, id: groupId });
});

app.delete('/api/driver-groups/:id', async (req, res) => {
  await dbRun('DELETE FROM driver_group_stores WHERE group_id = ?', [req.params.id]);
  await dbRun('DELETE FROM driver_group_members WHERE group_id = ?', [req.params.id]);
  await dbRun('DELETE FROM driver_groups WHERE id = ?', [req.params.id]);
  res.json({ success: true });
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
  const importBatchId = await createImportBatch('driver_availability', source.filename, `${year}年${month}月(ワイド形式)`);

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
  const unregisteredDriverNames = new Set(); // 画面側で「選んで登録」できるように、重複を除いた未登録の氏名を別途集める
  const now = new Date().toISOString();

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const driverName = String(row[0] ?? '').trim();
    if (!driverName) continue; // 空行はスキップ

    const driver = driverByName.get(driverName);
    if (!driver) {
      errors.push(`${r + 1}行目: ドライバー「${driverName}」がドライバーマスタに見つかりません(先に登録してください)`);
      unregisteredDriverNames.add(driverName);
      continue;
    }
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
          'UPDATE driver_availability SET time_start=?, time_end=?, desired_store=?, desired_area=?, import_batch_id=? WHERE id=?',
          [time_start, time_end, desired_store, desired_area, importBatchId, existing.id]
        );
      } else {
        await dbRun(
          `INSERT INTO driver_availability (driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests, created_at, import_batch_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          [driver_id, desired_date, fallbackArea, fallbackStore, time_start, time_end, '', now, importBatchId]
        );
      }
      imported++;
    }
  }
  await finalizeImportBatch(importBatchId, imported);

  res.json({ success: true, imported, errors, importBatchId, unregisteredDrivers: [...unregisteredDriverNames].sort((a, b) => a.localeCompare(b, 'ja')) });
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
  const rows = await dbAll(`
    SELECT sr.*, st.store_code
    FROM store_requests sr
    LEFT JOIN stores st ON st.id = sr.store_id
    WHERE sr.archived_month IS NULL
    ORDER BY sr.request_date DESC, sr.id DESC
  `);
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

// この依頼(特定の店舗・特定の日)の「時間の穴」警告を対応しないことにする/元に戻す
// (マッチング・結果画面で、穴に関するセルを右クリックして設定する。1件ずつの設定)
app.post('/api/store-requests/:id/gap-dismiss', async (req, res) => {
  const { dismissed } = req.body || {};
  await dbRun('UPDATE store_requests SET gap_dismissed = ? WHERE id = ?', [dismissed ? 1 : 0, req.params.id]);
  res.json({ success: true });
});

// 同じ行(店舗の同じ「N人目」枠)にまたがる複数の穴を、まとめて対応なし設定/解除する
app.post('/api/store-requests/gap-dismiss-batch', async (req, res) => {
  const { request_ids, dismissed } = req.body || {};
  if (!Array.isArray(request_ids) || request_ids.length === 0) return res.status(400).json({ success: false, message: 'request_idsは1件以上の配列で指定してください' });
  for (const id of request_ids) {
    await dbRun('UPDATE store_requests SET gap_dismissed = ? WHERE id = ?', [dismissed ? 1 : 0, id]);
  }
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
  const importBatchId = await createImportBatch('store_requests_weekly', source.filename, `${year}年${month}月`);

  let rows;
  try {
    rows = parseWideSpreadsheet(source.buffer, source.filename, req.body.sheet, ['店舗名']);
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
  // 店番(拠点コード)は任意項目。店舗名より表記ゆれが起きにくいため、あれば優先的に店舗の照合に使う
  const colStoreCode = ['店番', '店舗コード', '店舗CD', '拠点コード'].map(colOf).find(c => c !== -1) ?? -1;
  const colFlag = colOf('フラグ');
  const colNotes = colOf('備考');
  const WEEKDAY_LABELS = ['月', '火', '水', '木', '金', '土', '日'];
  const WEEKDAY_JSDAY = [1, 2, 3, 4, 5, 6, 0]; // Date.getDay()に合わせる(0=日,1=月,...,6=土)
  const weekdayCols = WEEKDAY_LABELS.map(colOf);

  if (weekdayCols.some(c => c === -1)) {
    return res.status(400).json({ success: false, message: '曜日(月〜日)の列が見つかりませんでした' });
  }

  const drivers = await dbAll('SELECT id, name, driver_code, fixed_store_id FROM drivers');
  const driverByCode = new Map(drivers.filter(d => d.driver_code).map(d => [normalizeEmployeeCode(d.driver_code), d.id]));
  const driversForNameSearch = drivers.map(d => ({ id: d.id, normName: normalizeNameForMatch(d.name) }));
  const driverNameById = new Map(drivers.map(d => [d.id, d.name]));
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
  let fixedStoreSuggestions = 0;

  // 備考の内容(社員番号またはAIが読み取った人物像)から、固定希望店舗の候補が見つかった場合の処理。
  // ※どんなに強い希望の記載があっても、ドライバーマスタのfixed_store_idは自動で書き換えない
  // (マスタの内容を勝手に変更しないでほしいという要望のため)。見つかったことは「要確認」として
  // 伝えるだけに留め、実際に設定するかどうかはコーディネーターがドライバーマスタの一覧から
  // 手動で判断・操作する(一覧の「固定希望店舗」列はその場で編集できる)
  async function suggestFixedStoreLink(matchedDriverId, store_id, rowNumber, store_name, mentionLabel, source) {
    const currentFixedStoreId = currentFixedStoreById.get(matchedDriverId);
    if (currentFixedStoreId === store_id) return false; // 既に同じ設定ならあらためて知らせる必要はない
    const driverName = driverNameById.get(matchedDriverId) || `id:${matchedDriverId}`;
    const ngNote = ngPairSet.has(`${matchedDriverId}:${store_id}`) ? '(この店舗は店舗相性マスタでNGに設定されています)' : '';
    errors.push(`${rowNumber}行目「${store_name}」: 備考の${mentionLabel}から、${driverName}さんの固定希望店舗として「${store_name}」が考えられます${ngNote}。必要であればドライバーマスタの一覧から手動で設定してください(自動では設定していません)`);
    await recordFixedStoreSuggestion(matchedDriverId, store_id, source, `店舗依頼の備考(${mentionLabel})より(${rowNumber}行目「${store_name}」)`);
    return true;
  }

  // 1周目: 行ごとにそのまま店舗依頼の枠として扱う(同じ店舗・同じ曜日・同じ時間帯の行が複数あっても
  // 合算はしない=「2名」のように1行にまとめず、元のシート通り複数行の店舗依頼として取り込む)
  const slots = []; // { storeId, storeName, weekdayIndex, time_start, time_end, count(セル内の*N倍率), note, area, address, lat, lng }
  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r];
    const store_name = String(row[colStoreName] ?? '').trim();
    if (!store_name) continue;

    const flag = colFlag !== -1 ? String(row[colFlag] ?? '').trim() : '';
    const notesRaw = colNotes !== -1 ? String(row[colNotes] ?? '').trim() : '';
    const note = flag ? `(${flag}) ${notesRaw}`.trim() : notesRaw;
    const store_code = colStoreCode !== -1 ? String(row[colStoreCode] ?? '').trim() : '';

    const store_id = await resolveStoreId({ store_name, store_code });
    // 店舗マスタに住所が登録されていれば、それを店舗依頼側にも使う(距離計算ができるように)
    const storeInfo = await getStoreMasterInfo(store_id);

    // 備考に社員番号らしき数字があれば、固定希望店舗の候補として要確認に表示する(確実なのでまずこちらを試す)
    const codeMatch = notesRaw.match(/(\d{5,8})/);
    const normalizedCode = codeMatch ? normalizeEmployeeCode(codeMatch[1]) : null;
    if (normalizedCode && driverByCode.has(normalizedCode)) {
      if (await suggestFixedStoreLink(driverByCode.get(normalizedCode), store_id, r + 1, store_name, `社員番号(${codeMatch[1]})`, 'note_code')) fixedStoreSuggestions++;
    } else if (anthropic && notesRaw) {
      // 社員番号の記載が無い(=数字で確実には分からない)場合のみ、AIで自然文から人物への言及を読み取る
      // (「できれば髙橋さん希望」のような表現を拾うため。数字で確実に分かる場合は無駄なAI呼び出しをしない)
      const aiResult = await analyzeRequestNoteForMatching(notesRaw);
      for (const mention of (aiResult?.driver_mentions || [])) {
        if (mention.sentiment !== 'required') continue; // 強い指定の時だけ固定希望店舗の候補として表示する(弱い希望はマッチング時のスコアのみで考慮する)
        const driverId = resolveDriverMention(mention, driverByCode, driversForNameSearch);
        if (!driverId) continue;
        if (await suggestFixedStoreLink(driverId, store_id, r + 1, store_name, `内容(AIが「${mention.name_hint || mention.employee_code}」への強い希望と判定)`, 'note_ai')) fixedStoreSuggestions++;
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
      // セルに「*2」「＊3」のような倍数指定があれば、その人数分を必要枠として数える(無ければ1人分。
      // これは同じ行の中での倍率なので、これとは別に行自体が複数ある場合は合算せず別の枠として扱う)
      const multiplierMatch = cell.match(/[*＊]\s*(\d+)/);
      const multiplier = multiplierMatch ? parseInt(multiplierMatch[1], 10) : 1;

      slots.push({
        storeId: store_id, storeName: store_name, weekdayIndex: w, time_start, time_end, count: multiplier, note,
        area: storeInfo ? (storeInfo.area || '') : '', address: storeInfo ? (storeInfo.address || '') : '',
        lat: storeInfo ? storeInfo.lat : null, lng: storeInfo ? storeInfo.lng : null
      });
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

  // 2周目: 曜日×時間帯の枠を、対象月の実際の日付に展開して店舗依頼を作る。
  // 同じ店舗・同じ曜日・同じ時間帯の枠が複数行ある場合は、合算して1行の必要人数にはせず、
  // 元のシートに出てきた順番のまま、既存の重複防止(店舗×日付×開始時刻)の行と1対1で対応づけて
  // 上書き更新する(再取込みのたびに重複登録されないように。既存の方が多い場合、余った古い行は
  // 既にマッチング済みの可能性があるため削除せずそのまま残す)
  const slotGroups = new Map(); // "storeId|weekdayIndex|time_start|time_end" -> [slot, slot, ...](シート登場順)
  for (const slot of slots) {
    const key = `${slot.storeId}|${slot.weekdayIndex}|${slot.time_start}|${slot.time_end}`;
    if (!slotGroups.has(key)) slotGroups.set(key, []);
    slotGroups.get(key).push(slot);
  }

  const now = new Date().toISOString();
  let imported = 0;
  for (const slotGroup of slotGroups.values()) {
    const first = slotGroup[0];
    const jsDay = WEEKDAY_JSDAY[first.weekdayIndex];
    for (const day of (datesByJsDay[jsDay] || [])) {
      const request_date = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      const existingRows = await dbAll(
        'SELECT id FROM store_requests WHERE store_id = ? AND request_date = ? AND time_start = ? AND archived_month IS NULL ORDER BY id',
        [first.storeId, request_date, first.time_start]
      );
      for (let i = 0; i < slotGroup.length; i++) {
        const slot = slotGroup[i];
        const requests = slot.note || '';
        if (existingRows[i]) {
          await dbRun(
            `UPDATE store_requests SET store_name=?, area=?, address=?, lat=?, lng=?, time_end=?, required_count=?, requests=?, import_batch_id=? WHERE id=?`,
            [slot.storeName, slot.area, slot.address, slot.lat, slot.lng, slot.time_end, slot.count, requests, importBatchId, existingRows[i].id]
          );
        } else {
          await dbRun(
            `INSERT INTO store_requests (store_name, area, address, lat, lng, request_date, time_start, time_end, required_count, requests, created_at, store_id, import_batch_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [slot.storeName, slot.area, slot.address, slot.lat, slot.lng, request_date, slot.time_start, slot.time_end, slot.count, requests, now, slot.storeId, importBatchId]
          );
        }
        imported++;
      }
    }
  }
  await finalizeImportBatch(importBatchId, imported);

  res.json({ success: true, imported, fixedStoreSuggestions, errors, importBatchId });
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

// 本部等から定期的に配布される「店舗対SV・店長一覧表」を取込み、店舗マスタのSV・店長を
// 最新の状態に更新する(人事異動のたびに再取込みする運用を想定)。「課・エリア／担当SV・責任者」の
// ような1つの列にまとまった自由文から、「ＳＶ」または「課長」の後ろの氏名を読み取る。
// 店番(確実)→店舗名の順で照合し、見つかった店舗のSV・店長は(前の値が何であれ)常に上書きする
app.post('/api/stores/import-sv-manager', upload.single('file'), async (req, res) => {
  const source = getImportSource(req);
  if (!source) return res.status(400).json({ success: false, message: 'ファイルが必要です(アップロードするか、所定フォルダのファイルを指定してください)' });
  const importBatchId = await createImportBatch('store_sv_manager', source.filename, '');

  let rows;
  try {
    rows = parseWideSpreadsheet(source.buffer, source.filename, req.body.sheet, ['店番']);
  } catch (e) {
    return res.status(400).json({ success: false, message: 'ファイルの読み込みに失敗しました: ' + e.message });
  }

  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const cells = rows[i].map(c => String(c ?? '').trim());
    if (cells.includes('店番') && cells.includes('店舗名')) { headerRowIndex = i; break; }
  }
  if (headerRowIndex === -1) {
    return res.status(400).json({ success: false, message: '見出し行(「店番」「店舗名」の列)が見つかりませんでした' });
  }
  const header = rows[headerRowIndex].map(h => String(h ?? '').trim());
  const colOf = (labels) => header.findIndex(h => labels.includes(h));
  const colCode = colOf(['店番']);
  const colName = colOf(['店舗名']);
  const colManager = colOf(['店長']);
  // SV名はこの列単独にあるとは限らず、「拠点運営七課　11店舗　ＳＶ　吉田　拓人」のように
  // 他の情報と1つの列にまとまっていることがあるため、その中から正規表現で拾う
  const colSvBlock = colOf(['課・エリア／担当SV・責任者', '課・エリア/担当SV・責任者', 'ＳＶ', 'SV']);

  const errors = [];
  let updated = 0;
  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r];
    const store_name = colName !== -1 ? String(row[colName] ?? '').trim() : '';
    const store_code = colCode !== -1 ? String(row[colCode] ?? '').trim() : '';
    if (!store_name && !store_code) continue;

    const manager_name = colManager !== -1 ? String(row[colManager] ?? '').trim() : '';
    let sv_name = '';
    if (colSvBlock !== -1) {
      const raw = String(row[colSvBlock] ?? '');
      const m = raw.match(/(?:ＳＶ|SV|課長)[\s：:　]*(.+)$/);
      if (m) {
        // 「川村　優太　070-1329-6078」のように電話番号が紛れ込むことがあるため、末尾の電話番号らしき
        // 部分は取り除く
        sv_name = m[1].replace(/[\s　]*[\d０-９]{2,4}[-－‐―]{1}[\d０-９]{2,4}[-－‐―]{1}[\d０-９]{3,4}\s*$/, '').trim();
      } else if (raw.trim()) {
        errors.push(`${r + 1}行目「${store_name || store_code}」: 「ＳＶ」「課長」の記載が見つからず、SV名を読み取れませんでした(内容: ${raw.trim()})`);
      }
    }

    const store_id = await resolveStoreId({ store_name, store_code });
    if (!store_id) {
      errors.push(`${r + 1}行目「${store_name || store_code}」: 店舗を特定できませんでした`);
      continue;
    }
    await dbRun('UPDATE stores SET sv_name = ?, manager_name = ?, store_code = COALESCE(NULLIF(store_code, \'\'), ?) WHERE id = ?', [sv_name, manager_name, store_code || null, store_id]);
    updated++;
  }
  await finalizeImportBatch(importBatchId, updated);

  res.json({ success: true, updated, errors, importBatchId });
});

// 店舗マスタをExcelでエクスポートする(他のPCへ渡して取込み直せるよう、一括取込みと同じ列構成・列名にしてある)
app.get('/api/export/stores.xlsx', async (req, res) => {
  const stores = await dbAll('SELECT * FROM stores ORDER BY name');
  const header = ['店番', '店舗名', 'エリア', '住所', '店長', 'SV', '備考'];
  const rows = stores.map(s => [
    s.store_code || '', s.name || '', s.area || '', s.address || '', s.manager_name || '', s.sv_name || '', s.notes || ''
  ]);
  const wb = new ExcelJS.Workbook();
  addFilledSampleSheet(wb, '店舗マスタ', header, rows, [10, 22, 12, 30, 14, 14, 30]);
  const buffer = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="store_master_${new Date().toISOString().slice(0, 10)}.xlsx"`);
  res.send(Buffer.from(buffer));
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
    SELECT m.*, m.time_start AS driver_time_start, m.time_end AS driver_time_end,
           d.name AS driver_name, d.phone AS driver_phone, d.home_address, d.driver_code,
           s.store_name, s.area, s.address AS store_address, s.time_start, s.time_end, s.requests AS store_requests,
           st.store_code AS store_code, st.sv_name AS store_sv_name
    FROM matches m
    JOIN drivers d ON d.id = m.driver_id
    JOIN store_requests s ON s.id = m.store_request_id
    LEFT JOIN stores st ON st.id = s.store_id
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
  // 固定希望店舗が重複しているドライバーの優先順位が決まっていない組が1つでもあれば、
  // (既存の「候補」を消す前に)ここで止める。誰を優先するか決めないまま自動マッチングすると、
  // スコアだけで決まってしまい意図しない人が割り当てられる恐れがあるため
  const unresolvedFixedStoreDuplicates = await findUnresolvedFixedStoreDuplicates();
  if (unresolvedFixedStoreDuplicates.length > 0) {
    return res.status(400).json({
      success: false,
      message: `固定希望店舗が重複しているドライバーの優先順位が決まっていないため、マッチングを実行しませんでした(${unresolvedFixedStoreDuplicates.length}件)。ドライバーマスタ画面の重複アラートから優先順位を設定してください。`,
      unresolvedFixedStoreDuplicates
    });
  }

  // 「候補」だけを作り直す。「確定」「完了」は既に実際の運用(シフト確定・派遣実績)で使われているため、
  // 再実行しても絶対に消さない(以前はここで全件消していたため、実行するたびに確定済みの予定が
  // 白紙に戻ってしまう重大な不具合があった。アーカイブ済み(過去に月次クローズしたもの)は対象外)
  await dbRun(`DELETE FROM matches WHERE archived_month IS NULL AND status = '候補'`);

  const storeRequests = await dbAll('SELECT * FROM store_requests WHERE archived_month IS NULL ORDER BY request_date ASC');
  const availabilityRows = await dbAll(`
    SELECT a.*, d.name AS driver_name, d.home_lat, d.home_lng, d.fixed_store_id, d.fixed_store_priority, fs.name AS fixed_store_name
    FROM driver_availability a
    JOIN drivers d ON d.id = a.driver_id
    LEFT JOIN stores fs ON fs.id = d.fixed_store_id
    WHERE a.archived_month IS NULL
  `);
  // 希望シフトで店舗が未入力の場合は、ドライバーマスタの「固定希望店舗」を初期値として使う
  // desired_store_explicit: 固定希望店舗からの自動補完が入る前の、本人がその日に実際に入力した希望店舗
  // (「希望シフトで指定した店舗」の優先階層で、固定希望店舗と区別するために使う)
  let availability = availabilityRows.map(a => ({ ...a, desired_store_explicit: a.desired_store || '', desired_store: a.desired_store || a.fixed_store_name || '' }));

  // 設定がONの場合、自動マッチングの対象を「固定希望店舗が設定されているドライバー」だけに絞る
  // (それ以外のドライバーは候補を作らず、手作業での割当に委ねる運用)
  const restrictToFixedStore = await getRestrictMatchingToFixedStore();
  if (restrictToFixedStore) {
    availability = availability.filter(a => a.fixed_store_id != null);
  }

  // ドライバーグループ(「2店舗を3人で回す」のように、複数のドライバー×複数の店舗を紐付けたもの)。
  // 所属ドライバーに、グループに紐付いた店舗群をstoreIdsとして持たせ、優先順位階層(group_fixed)で使う
  const groupRows = restrictToFixedStore ? [] : await dbAll(`
    SELECT m.driver_id, gs.store_id FROM driver_group_members m JOIN driver_group_stores gs ON gs.group_id = m.group_id
  `);
  const groupStoreIdsByDriverId = new Map();
  for (const r of groupRows) {
    if (!groupStoreIdsByDriverId.has(r.driver_id)) groupStoreIdsByDriverId.set(r.driver_id, new Set());
    groupStoreIdsByDriverId.get(r.driver_id).add(r.store_id);
  }

  const { preferenceByDriverStore, storeVisitCount, areaVisitCount } = await loadScoringContext();

  // 店舗依頼の備考(自由記述)をAIで解析し、特定ドライバーへの希望・必須指定・除外を読み取る
  // (.envにANTHROPIC_API_KEYが未設定の場合、getOrAnalyzeStoreRequestNoteは常にnullを返すため
  // aiPreferenceByKeyは空のままになり、全体としては従来通りの動作になる)
  const aiDriversForLookup = await dbAll('SELECT id, name, driver_code FROM drivers');
  const aiDriverByCode = new Map(aiDriversForLookup.filter(d => d.driver_code).map(d => [normalizeEmployeeCode(d.driver_code), d.id]));
  const aiDriversForNameSearch = aiDriversForLookup.map(d => ({ id: d.id, normName: normalizeNameForMatch(d.name) }));
  const aiPreferenceByKey = new Map(); // "driver_id:store_request_id" -> 'required'|'preferred'|'excluded'
  if (anthropic) {
    await Promise.all(storeRequests.filter(s => s.requests && s.requests.trim()).map(async (store) => {
      const analysis = await getOrAnalyzeStoreRequestNote(store);
      if (!analysis || !Array.isArray(analysis.driver_mentions)) return;
      for (const mention of analysis.driver_mentions) {
        const driverId = resolveDriverMention(mention, aiDriverByCode, aiDriversForNameSearch);
        if (driverId) aiPreferenceByKey.set(`${driverId}:${store.id}`, mention.sentiment);
      }
    }));
  }

  const tierCtx = { preferenceByDriverStore, storeVisitCount, areaVisitCount, groupStoreIdsByDriverId, aiPreferenceByKey };
  const enabledTierKeys = await getEnabledTierKeys();

  const now = new Date().toISOString();
  const assignedDriverIdsByDate = {}; // date -> Set(driver_id) 同日の重複割当を防ぐ
  for (const store of storeRequests) {
    assignedDriverIdsByDate[store.request_date] = assignedDriverIdsByDate[store.request_date] || new Set();
  }
  let createdCount = 0;
  let noCandidateCount = 0;
  const filledCountByRequestId = {}; // store_request.id -> 既に埋まった人数(優先階層で埋めた分)
  // エリア固定の「SV行」由来のプール需要(pool_group_id)は、グループ内のどの店舗で埋まっても
  // 合計でカウントする(「4店舗のうちどれか1つに合計N人」という意味のため、1店舗ごとの個別集計ではなく
  // グループ単位で必要人数を判定する)
  const filledCountByPoolGroup = {};
  function getFilledCount(store) {
    return store.pool_group_id ? (filledCountByPoolGroup[store.pool_group_id] || 0) : (filledCountByRequestId[store.id] || 0);
  }

  // 既に「確定」「完了」になっているマッチングは消さずそのまま活かすため、その分を
  // 「埋まった人数」「その日は既に割当済みのドライバー」として先に計上しておく
  // (そうしないと必要人数を超えて候補を追加したり、同じドライバーを同日に二重登録してしまう)
  const existingMatches = await dbAll(`
    SELECT m.store_request_id, m.driver_id, m.match_date, s.pool_group_id
    FROM matches m JOIN store_requests s ON s.id = m.store_request_id
    WHERE m.archived_month IS NULL AND m.status IN ('確定', '完了')
  `);
  for (const em of existingMatches) {
    (assignedDriverIdsByDate[em.match_date] = assignedDriverIdsByDate[em.match_date] || new Set()).add(em.driver_id);
    filledCountByRequestId[em.store_request_id] = (filledCountByRequestId[em.store_request_id] || 0) + 1;
    if (em.pool_group_id) filledCountByPoolGroup[em.pool_group_id] = (filledCountByPoolGroup[em.pool_group_id] || 0) + 1;
  }

  // createdVia: 画面側で色分け表示するための、このマッチングが決まった経緯('fixed_store_tier'=固定希望店舗の
  // 優先階層で決まった、未指定なら既定値'auto'=それ以外の自動マッチング)。手動割当は/api/matches/manualが
  // 別途'manual'を指定する(このinsertMatchは自動マッチング実行時専用)
  async function insertMatch(store, p, createdVia) {
    const isFar = p.distance != null && p.distance > DIST_WARNING_KM;
    await dbRun(
      `INSERT INTO matches (store_request_id, driver_id, match_date, distance_km, is_far_warning, status, created_at, score, preference_flag, experience_count, created_via, time_start, time_end) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [store.id, p.driver_id, store.request_date, p.distance, isFar ? 1 : 0, '候補', now, p.score, p.preference, p.experienceCount, createdVia || 'auto', p.time_start || store.time_start, p.time_end || store.time_end]
    );
    assignedDriverIdsByDate[store.request_date].add(p.driver_id);
    filledCountByRequestId[store.id] = (filledCountByRequestId[store.id] || 0) + 1;
    if (store.pool_group_id) filledCountByPoolGroup[store.pool_group_id] = (filledCountByPoolGroup[store.pool_group_id] || 0) + 1;
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
      const alreadyFilled = getFilledCount(store);
      const needed = (store.required_count || 1) - alreadyFilled;
      if (needed <= 0) continue;

      const tierCandidates = availability.filter(a =>
        a.desired_date === store.request_date &&
        !assignedToday.has(a.driver_id) &&
        tier.test(a, store, tierCtx) &&
        preferenceByDriverStore.get(`${a.driver_id}:${store.store_id}`) !== 'NG' &&
        aiPreferenceByKey.get(`${a.driver_id}:${store.id}`) !== 'excluded'
      ).map(a => ({ ...a, ...scoreCandidate(a, store, preferenceByDriverStore, storeVisitCount, areaVisitCount, aiPreferenceByKey.get(`${a.driver_id}:${store.id}`)) }))
        .sort((x, y) => {
          // 固定希望店舗の階層だけは、重複時にドライバーマスタで決めた優先順位(1が最優先)を
          // スコアより先に見る(/api/matches/run は優先順位が未確定だと実行自体をブロックするため、
          // ここに来る時点でこの階層の候補は全員fixed_store_priorityが決まっているはずだが、
          // 単独(重複なし)のドライバーはnullのままで構わないため、その場合はスコア順にフォールバックする)
          if (tierKey === 'fixed_store') {
            const pDiff = (x.fixed_store_priority ?? Infinity) - (y.fixed_store_priority ?? Infinity);
            if (pDiff !== 0) return pDiff;
          }
          return y.score - x.score;
        });

      const picked = tierCandidates.slice(0, needed);
      for (const p of picked) await insertMatch(store, p, tierKey === 'fixed_store' ? 'fixed_store_tier' : 'auto');
    }
  }

  // 2階層目: 残りの枠を、これまで通りのスコアリング(距離・希望店舗/エリア一致・相性・経験)で埋める
  for (const store of storeRequests) {
    const assignedToday = assignedDriverIdsByDate[store.request_date];
    const alreadyFilled = getFilledCount(store);
    const needed = (store.required_count || 1) - alreadyFilled;
    if (needed <= 0) continue;

    // 同じ日付の希望を持ち、この店舗をNGにしていない(かつAIが備考からNGと判定していない)ドライバーを候補にする。
    // グループに所属していても、本人が希望シフトを出していなければ候補にはならない(グループは優先順位階層
    // (group_fixed)でスコアに反映されるだけで、希望シフト未提出の日の仮の候補までは作らない)
    const candidates = availability.filter(a =>
      a.desired_date === store.request_date &&
      !assignedToday.has(a.driver_id) &&
      preferenceByDriverStore.get(`${a.driver_id}:${store.store_id}`) !== 'NG' &&
      aiPreferenceByKey.get(`${a.driver_id}:${store.id}`) !== 'excluded'
    );

    const scored = candidates
      .map(a => ({ ...a, ...scoreCandidate(a, store, preferenceByDriverStore, storeVisitCount, areaVisitCount, aiPreferenceByKey.get(`${a.driver_id}:${store.id}`)) }))
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

// 自動マッチングを「固定希望店舗が設定されているドライバーのみ」に絞る設定の取得/保存
app.get('/api/settings/restrict-matching-to-fixed-store', async (req, res) => {
  res.json({ success: true, enabled: await getRestrictMatchingToFixedStore() });
});
app.post('/api/settings/restrict-matching-to-fixed-store', async (req, res) => {
  const { enabled } = req.body;
  const value = enabled ? '1' : '0';
  const existing = await dbGet('SELECT key FROM app_settings WHERE key = ?', ['restrict_matching_to_fixed_store']);
  if (existing) await dbRun('UPDATE app_settings SET value = ? WHERE key = ?', [value, 'restrict_matching_to_fixed_store']);
  else await dbRun('INSERT INTO app_settings (key, value) VALUES (?, ?)', ['restrict_matching_to_fixed_store', value]);
  res.json({ success: true });
});


// アプリ内だけの「今日」設定の取得/保存(PC本体の時刻はそのまま)。date が空文字/未指定なら解除(本当の今日に戻す)
app.get('/api/settings/mock-today', async (req, res) => {
  const mockToday = await getMockToday();
  const now = new Date();
  const realToday = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  res.json({ success: true, mockToday, realToday, today: mockToday || realToday });
});
app.post('/api/settings/mock-today', async (req, res) => {
  const { date } = req.body;
  if (date && !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return res.status(400).json({ success: false, message: 'date は YYYY-MM-DD 形式で指定してください' });
  }
  const value = date || '';
  const existing = await dbGet('SELECT key FROM app_settings WHERE key = ?', ['mock_today']);
  if (existing) await dbRun('UPDATE app_settings SET value = ? WHERE key = ?', [value, 'mock_today']);
  else await dbRun('INSERT INTO app_settings (key, value) VALUES (?, ?)', ['mock_today', value]);
  res.json({ success: true });
});

// 個人事業主向けPDF(作業依頼表)の「依頼者」「担当者」「電話番号」欄の表記設定
app.get('/api/settings/pdf-header', async (req, res) => {
  res.json({ success: true, settings: await getPdfHeaderSettings() });
});
app.post('/api/settings/pdf-header', async (req, res) => {
  const current = await getPdfHeaderSettings();
  const next = {
    requester_name: (req.body.requester_name ?? current.requester_name) || '',
    requester_address: (req.body.requester_address ?? current.requester_address) || '',
    contact_name: (req.body.contact_name ?? current.contact_name) || '',
    contact_phone: (req.body.contact_phone ?? current.contact_phone) || ''
  };
  const value = JSON.stringify(next);
  const existing = await dbGet('SELECT key FROM app_settings WHERE key = ?', ['pdf_header_settings']);
  if (existing) await dbRun('UPDATE app_settings SET value = ? WHERE key = ?', [value, 'pdf_header_settings']);
  else await dbRun('INSERT INTO app_settings (key, value) VALUES (?, ?)', ['pdf_header_settings', value]);
  res.json({ success: true, settings: next });
});

// 個人事業主向けPDF(作業依頼表)の一番下の注記文言の設定
app.get('/api/settings/pdf-footer', async (req, res) => {
  res.json({ success: true, note: await getPdfFooterNote() });
});
app.post('/api/settings/pdf-footer', async (req, res) => {
  const note = (req.body.note ?? PDF_FOOTER_DEFAULT) || '';
  const existing = await dbGet('SELECT key FROM app_settings WHERE key = ?', ['pdf_footer_note']);
  if (existing) await dbRun('UPDATE app_settings SET value = ? WHERE key = ?', [note, 'pdf_footer_note']);
  else await dbRun('INSERT INTO app_settings (key, value) VALUES (?, ?)', ['pdf_footer_note', note]);
  res.json({ success: true, note });
});

// マッチング・結果の表で、空いている枠を右クリック→候補のドライバーをクリックした時に呼ばれる、
// その場で1件だけ手動でマッチングを作る専用エンドポイント(自動マッチングの候補生成ロジックは経由しない)。
// 同日の二重割当・必要人数を超えての割当(プール需要の合計も含む)を防ぐチェックのみ行う
// マッチング・結果の表で、空き枠を右クリックした時に出す候補一覧を返す。その日に希望シフトを
// 出していて、まだ何にも割り当てが無い(欠勤は除く)ドライバーを、その店舗への過去の派遣実績
// (dispatch_history)が多い順に並べて返す(実績が無い人は末尾、同数なら氏名順)。
// 「直近に手動で割り当てた人を上に出す」という並び替えは、画面側(JS)でこの結果にかぶせて行う
app.get('/api/matches/manual-candidates', async (req, res) => {
  const storeRequestId = parseInt(req.query.store_request_id, 10);
  if (!storeRequestId) return res.status(400).json({ success: false, message: 'store_request_id は必須です' });
  const storeReq = await dbGet('SELECT * FROM store_requests WHERE id = ?', [storeRequestId]);
  if (!storeReq) return res.status(404).json({ success: false, message: '店舗依頼が見つかりません' });

  const date = storeReq.request_date;
  const availability = await dbAll(`
    SELECT a.driver_id, a.time_start, a.time_end, d.name AS driver_name
    FROM driver_availability a
    JOIN drivers d ON d.id = a.driver_id
    WHERE a.desired_date = ? AND a.archived_month IS NULL
  `, [date]);

  // excludeMatchId: 右クリックした枠に既に入っている「候補」を差し替える場合、その枠自身は
  // 「既に埋まっている」扱いから除外する(今入っている人も選び直せるように候補に残す)
  const excludeMatchId = req.query.exclude_match_id ? parseInt(req.query.exclude_match_id, 10) : 0;
  // excludeDriverId: 確定済みの枠を「欠勤」にしてから交代を探す場合に指定する。欠勤にした本人は
  // 「既に埋まっている」扱いから外れてしまう(欠勤は稼働していない扱いのため)が、交代探しの
  // 一覧に欠勤にした当の本人が出てきてしまうと紛らわしいので、明示的に除外する
  const excludeDriverId = req.query.exclude_driver_id ? parseInt(req.query.exclude_driver_id, 10) : 0;
  const bookedRows = await dbAll(
    `SELECT DISTINCT driver_id FROM matches WHERE match_date = ? AND status != '欠勤' AND archived_month IS NULL AND id != ?`,
    [date, excludeMatchId]
  );
  const bookedSet = new Set(bookedRows.map(r => r.driver_id));
  if (excludeDriverId) bookedSet.add(excludeDriverId);
  const candidates = availability.filter(a => !bookedSet.has(a.driver_id));

  let visitCountByDriver = new Map();
  if (storeReq.store_id) {
    const historyRows = await dbAll(
      `SELECT driver_id, COUNT(*) AS c FROM dispatch_history WHERE store_id = ? GROUP BY driver_id`,
      [storeReq.store_id]
    );
    visitCountByDriver = new Map(historyRows.map(r => [r.driver_id, r.c]));
  }

  const result = candidates
    .map(a => ({
      driver_id: a.driver_id, driver_name: a.driver_name, time_start: a.time_start, time_end: a.time_end,
      visit_count: visitCountByDriver.get(a.driver_id) || 0
    }))
    .sort((x, y) => y.visit_count - x.visit_count || x.driver_name.localeCompare(y.driver_name, 'ja'));

  res.json({ success: true, candidates: result });
});

app.post('/api/matches/manual', async (req, res) => {
  const store_request_id = parseInt(req.body.store_request_id, 10);
  const driver_id = parseInt(req.body.driver_id, 10);
  // replace_match_id: 既に「候補」で入っている人を、右クリックから別の人に差し替える場合に指定する
  // (確定・完了したものは、確定するまでは変更可能という方針のため対象外にする)
  const replace_match_id = req.body.replace_match_id ? parseInt(req.body.replace_match_id, 10) : null;
  if (!store_request_id || !driver_id) return res.status(400).json({ success: false, message: 'store_request_id, driver_id は必須です' });

  const storeReq = await dbGet('SELECT * FROM store_requests WHERE id = ?', [store_request_id]);
  if (!storeReq) return res.status(404).json({ success: false, message: '店舗依頼が見つかりません' });
  const driver = await dbGet('SELECT * FROM drivers WHERE id = ?', [driver_id]);
  if (!driver) return res.status(404).json({ success: false, message: 'ドライバーが見つかりません' });

  let replaceTarget = null;
  if (replace_match_id) {
    replaceTarget = await dbGet(`SELECT * FROM matches WHERE id = ? AND archived_month IS NULL`, [replace_match_id]);
    if (!replaceTarget) return res.status(404).json({ success: false, message: '差し替え対象のマッチングが見つかりません' });
    if (replaceTarget.status !== '候補') {
      return res.status(400).json({ success: false, message: '確定・完了済みのマッチングは差し替えできません(候補のうちだけ変更できます)' });
    }
  }

  const already = await dbGet(
    `SELECT 1 FROM matches WHERE driver_id = ? AND match_date = ? AND status != '欠勤' AND archived_month IS NULL AND id != ?`,
    [driver_id, storeReq.request_date, replace_match_id || 0]
  );
  if (already) return res.status(400).json({ success: false, message: 'このドライバーは同じ日に既に別の割り当てがあります' });

  // プール需要(候補店舗のどれか1つで合計必要人数が埋まればよい需要)の場合は、グループ全体で
  // 既に必要人数に達していないか確認する(差し替えの場合は、差し替え対象自身の分は数えない)。
  // 通常の依頼は自分自身の件数だけで確認する
  const siblingIds = storeReq.pool_group_id
    ? (await dbAll('SELECT id FROM store_requests WHERE pool_group_id = ?', [storeReq.pool_group_id])).map(s => s.id)
    : [store_request_id];
  const filled = await dbGet(
    `SELECT COUNT(*) AS c FROM matches WHERE store_request_id IN (${siblingIds.map(() => '?').join(',')}) AND status IN ('候補', '確定', '完了') AND archived_month IS NULL AND id != ?`,
    [...siblingIds, replace_match_id || 0]
  );
  if (filled.c >= (storeReq.required_count || 1)) {
    // 必要人数自体は埋まっていても、店舗依頼の希望時間に「穴」(誰もカバーしていない時間帯)が
    // 残っている場合は、その穴を埋めるための追加の割当てとして許可する(プール需要は対象外。
    // 候補店舗ごとに時間が異なり得るため、時間の穴という考え方が馴染まないため)
    let gapAllows = false;
    if (!storeReq.pool_group_id) {
      const existingForGap = await dbAll(
        `SELECT time_start, time_end FROM matches WHERE store_request_id = ? AND status IN ('候補', '確定', '完了') AND archived_month IS NULL AND id != ?`,
        [store_request_id, replace_match_id || 0]
      );
      gapAllows = computeTimeGaps(storeReq.time_start, storeReq.time_end, existingForGap).length > 0;
    }
    if (!gapAllows) {
      return res.status(400).json({ success: false, message: 'この店舗依頼は既に必要人数が埋まっています' });
    }
  }

  if (replaceTarget) {
    await dbRun(`DELETE FROM matches WHERE id = ?`, [replace_match_id]);
  }

  // このマッチングが実際にカバーする時間は、担当者本人がその日に出した希望シフトの時間を使う
  // (無ければ店舗依頼の時間をそのままカバーするとみなす、今まで通りの前提にフォールバックする)
  const driverAvailability = await dbGet(
    'SELECT time_start, time_end FROM driver_availability WHERE driver_id = ? AND desired_date = ? AND archived_month IS NULL',
    [driver_id, storeReq.request_date]
  );
  const matchTimeStart = (driverAvailability && driverAvailability.time_start) || storeReq.time_start;
  const matchTimeEnd = (driverAvailability && driverAvailability.time_end) || storeReq.time_end;

  const hasCoords = driver.home_lat != null && driver.home_lng != null && storeReq.lat != null && storeReq.lng != null;
  const distance = hasCoords ? haversineKm(driver.home_lat, driver.home_lng, storeReq.lat, storeReq.lng) : null;
  const isFar = distance != null && distance > DIST_WARNING_KM;
  const now = new Date().toISOString();
  const result = await dbRun(
    `INSERT INTO matches (store_request_id, driver_id, match_date, distance_km, is_far_warning, status, created_at, created_via, time_start, time_end) VALUES (?, ?, ?, ?, ?, '候補', ?, 'manual', ?, ?)`,
    [store_request_id, driver_id, storeReq.request_date, distance, isFar ? 1 : 0, now, matchTimeStart, matchTimeEnd]
  );
  res.json({ success: true, id: result.lastID });
});

app.post('/api/matches/:id/confirm', async (req, res) => {
  await dbRun(`UPDATE matches SET status = '確定' WHERE id = ?`, [req.params.id]);
  const autoFixedStoreSuggested = await applyAutoFixedStoreFromMatches();
  res.json({ success: true, autoFixedStoreSuggested });
});

// 複数の「候補」をまとめて「確定」にする(画面の一括確定機能用)。
// status='候補'のものだけを対象にする(古い選択状態のまま送られてきても、既に確定/完了済みのものを
// 誤って巻き戻すことがないように)
app.post('/api/matches/bulk-confirm', async (req, res) => {
  const { ids } = req.body;
  if (!Array.isArray(ids) || ids.length === 0) return res.status(400).json({ success: false, message: 'ids(配列)は必須です' });
  const placeholders = ids.map(() => '?').join(',');
  const result = await dbRun(`UPDATE matches SET status = '確定' WHERE id IN (${placeholders}) AND status = '候補'`, ids);
  const autoFixedStoreSuggested = await applyAutoFixedStoreFromMatches();
  res.json({ success: true, confirmed: result.changes, autoFixedStoreSuggested });
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
  const autoFixedStoreSuggested = await applyAutoFixedStoreFromMatches();
  res.json({ success: true, autoFixedStoreSuggested });
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

// 複数の「完了」をまとめて「確定」に戻す(画面の一括取り消し機能用)。status='完了'のものだけを対象にする
app.post('/api/matches/bulk-undo-complete', async (req, res) => {
  const { ids } = req.body;
  if (!Array.isArray(ids) || ids.length === 0) return res.status(400).json({ success: false, message: 'ids(配列)は必須です' });
  const targets = await dbAll(
    `SELECT id FROM matches WHERE id IN (${ids.map(() => '?').join(',')}) AND status = '完了'`,
    ids
  );
  const targetIds = targets.map(t => t.id);
  if (targetIds.length === 0) return res.json({ success: true, undone: 0 });
  const placeholders = targetIds.map(() => '?').join(',');
  await dbRun(`UPDATE matches SET status = '確定' WHERE id IN (${placeholders})`, targetIds);
  await dbRun(`DELETE FROM dispatch_history WHERE match_id IN (${placeholders})`, targetIds);
  res.json({ success: true, undone: targetIds.length });
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
    'SELECT desired_store, desired_area, time_start, time_end FROM driver_availability WHERE driver_id = ? AND desired_date = ? AND archived_month IS NULL',
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

  // このマッチングが実際にカバーする時間は、担当者本人がその日に出した希望シフトの時間を使う
  // (無ければ元のマッチング=欠勤者が入っていた枠の時間をそのまま引き継ぐ。/api/matches/manual と同じ考え方)
  const matchTimeStart = (av && av.time_start) || match.time_start;
  const matchTimeEnd = (av && av.time_end) || match.time_end;

  const now = new Date().toISOString();
  const result = await dbRun(
    `INSERT INTO matches (store_request_id, driver_id, match_date, distance_km, is_far_warning, status, created_at, score, preference_flag, experience_count, time_start, time_end) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [match.store_request_id, driver_id, match.match_date, scored.distance, isFar ? 1 : 0, '候補', now, scored.score, scored.preference, scored.experienceCount, matchTimeStart, matchTimeEnd]
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
  // 今まさに作業中(未アーカイブ)のデータがある年月も調べておく。この年月をうっかりアーカイブから
  // 「戻す」と、今の作業データと混在してしまうため、画面側で選べないようにするのに使う
  const activeRows = await dbAll(`
    SELECT substr(request_date, 1, 7) AS ym FROM store_requests WHERE archived_month IS NULL
    UNION SELECT substr(desired_date, 1, 7) FROM driver_availability WHERE archived_month IS NULL
    UNION SELECT substr(match_date, 1, 7) FROM matches WHERE archived_month IS NULL
  `);
  res.json({ success: true, months: rows.map(r => r.archived_month), activeMonths: [...new Set(activeRows.map(r => r.ym))] });
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
  const autoFixedStoreSuggested = await applyAutoFixedStoreFromMatches();
  res.json({ success: true, confirmed: result.changes, autoFixedStoreSuggested });
});

function weekdayLabelOf(dateStr) {
  const WEEKDAY_LABELS_JP = ['日', '月', '火', '水', '木', '金', '土'];
  const d = new Date(dateStr + 'T00:00:00');
  return WEEKDAY_LABELS_JP[d.getDay()];
}

// 日本の祝日判定(固定日+ハッピーマンデー+春分/秋分の近似式+振替休日)。概算式のため、稀に公式発表と
// 1日ずれる可能性があるが、個人事業主向けPDFの土日祝の色分け用途であり実用上問題ない
// (フロント側のisJapaneseHoliday(index.html)と同じロジック)
function isJapaneseHoliday(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const y = d.getFullYear(), m = d.getMonth() + 1, day = d.getDate();
  function nthMonday(month, n) {
    const first = new Date(y, month - 1, 1);
    const firstMonday = 1 + ((8 - first.getDay()) % 7);
    return firstMonday + (n - 1) * 7;
  }
  const shunbun = Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
  const shubun = Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
  const fixed = new Set(['1-1', '2-11', '2-23', '4-29', '5-3', '5-4', '5-5', '8-11', '11-3', '11-23', `3-${shunbun}`, `9-${shubun}`]);
  const movable = new Set([`1-${nthMonday(1, 2)}`, `7-${nthMonday(7, 3)}`, `9-${nthMonday(9, 3)}`, `10-${nthMonday(10, 2)}`]);
  const key = `${m}-${day}`;
  if (fixed.has(key) || movable.has(key)) return true;
  if (d.getDay() === 1) {
    const prev = new Date(d); prev.setDate(d.getDate() - 1);
    if (prev.getDay() === 0) {
      const pKey = `${prev.getMonth() + 1}-${prev.getDate()}`;
      if (fixed.has(pKey) || movable.has(pKey)) return true;
    }
  }
  return false;
}

// ExcelJSはDateオブジェクトをUTC基準でシリアル値に変換するため、日本時間(UTC+9)のまま
// new Date(year, month-1, day)を渡すと、Excel上で前日の日付として表示されてしまう
// (例: 日本時間10/1 0:00 は UTC 9/30 15:00 であり、UTC基準では9/30扱いになる)。
// この「UTC上でもその年月日の0時になる」Dateを作ることで、ズレを防ぐ
function excelSafeDate(year, month, day) {
  return new Date(Date.UTC(year, month - 1, day));
}

// LibreOffice(soffice)の実行ファイルを探す(個人事業主向けPDF出力で、xlsx→PDF変換に使う)
function findSofficePath() {
  const candidates = [
    'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    '/usr/bin/soffice',
    '/usr/bin/libreoffice',
    '/opt/libreoffice/program/soffice',
  ];
  for (const c of candidates) {
    if (fs.existsSync(c)) return c;
  }
  // 上記の決め打ちパスに無い場合(Linux環境でパッケージ管理経由でインストールした場合など)は、
  // OSのコマンド検索(where/which)でPATHが通っているか確認する
  try {
    const finder = process.platform === 'win32' ? 'where' : 'which';
    const result = spawnSync(finder, ['soffice'], { encoding: 'utf8' });
    const found = (result.stdout || '').split(/\r?\n/).map(s => s.trim()).find(s => s && fs.existsSync(s));
    if (found) return found;
  } catch (e) { /* noop: 見つからなければ呼び出し元でnullとして扱われる */ }
  return null;
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

// あるシートの内容(値・スタイル・列幅・行高さ・結合セル)を、別のワークブックの新しいシートへ
// そのままコピーする(ExcelJSにはワークブックをまたいだシート複製の機能が無いため、自前で実装している)。
// 個人事業主ごとの作業依頼表PDFと同時に、全員分を1つのExcelファイル(氏名ごとにシート分け)でも
// 出力するために使う
function copyWorksheetContent(targetWs, sourceWs, maxRow, maxCol) {
  targetWs.pageSetup = { ...sourceWs.pageSetup };
  for (let c = 1; c <= maxCol; c++) {
    const srcCol = sourceWs.getColumn(c);
    if (srcCol && srcCol.width) targetWs.getColumn(c).width = srcCol.width;
  }
  for (let r = 1; r <= maxRow; r++) {
    const srcRow = sourceWs.getRow(r);
    const tgtRow = targetWs.getRow(r);
    if (srcRow.height) tgtRow.height = srcRow.height;
    for (let c = 1; c <= maxCol; c++) {
      const srcCell = srcRow.getCell(c);
      const tgtCell = tgtRow.getCell(c);
      tgtCell.value = srcCell.value;
      tgtCell.style = { ...srcCell.style };
    }
  }
  for (const range of (sourceWs.model.merges || [])) {
    try { targetWs.mergeCells(range); } catch (e) { /* 既にマージ済み等は無視してよい */ }
  }
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
// 個人事業主ごとの月間「作業依頼表」をPDFで一括出力する。実際に使われている実ファイル
// (★11月HML個人事業主作業依頼表_00540102鈴木　翔.xls)の一番新しい月のシートの書式を
// そのままひな形(templates/driver_monthly_shift_template.xlsx)として使い、氏名・社員コード・
// 年月・その月の確定/完了シフトだけを差し込んでPDF化する(LibreOfficeのsofficeコマンドで変換)。
// 日付・曜日はテンプレートの数式(DATE関数等)をそのまま使うと再計算されない恐れがあるため、
// 実際の値(Dateオブジェクト・曜日の漢字)をこちら側で計算して直接書き込む
app.post('/api/export/driver-shift-pdfs', async (req, res) => {
  const year = parseInt(req.body.year, 10);
  const month = parseInt(req.body.month, 10);
  if (!year || !month || month < 1 || month > 12) {
    return res.status(400).json({ success: false, message: '対象年月を指定してください' });
  }
  const statuses = Array.isArray(req.body.statuses) && req.body.statuses.length ? req.body.statuses : ['確定', '完了'];

  const sofficePath = findSofficePath();
  if (!sofficePath) {
    return res.status(400).json({ success: false, message: 'LibreOffice(soffice)が見つかりません。このPC(マッチングアプリを動かしているPC)にLibreOfficeをインストールしてから再度お試しください。' });
  }

  const templatePath = path.join(__dirname, 'templates', 'driver_monthly_shift_template.xlsx');
  if (!fs.existsSync(templatePath)) {
    return res.status(400).json({ success: false, message: 'PDFのひな形ファイルが見つかりません(templates/driver_monthly_shift_template.xlsx)' });
  }

  const monthStr = `${year}-${String(month).padStart(2, '0')}`;
  const placeholders = statuses.map(() => '?').join(',');
  const matches = await dbAll(`
    SELECT m.match_date, m.driver_id, d.name AS driver_name, d.driver_code, d.furigana,
           s.store_name, s.time_start, s.time_end
    FROM matches m
    JOIN drivers d ON d.id = m.driver_id
    JOIN store_requests s ON s.id = m.store_request_id
    WHERE m.archived_month IS NULL AND m.status IN (${placeholders}) AND substr(m.match_date, 1, 7) = ?
    ORDER BY d.name ASC, m.match_date ASC
  `, [...statuses, monthStr]);

  if (matches.length === 0) {
    return res.status(400).json({ success: false, message: `${year}年${month}月分で、指定したステータス(${statuses.join('/')})のマッチングが見つかりませんでした` });
  }

  const byDriver = new Map();
  for (const m of matches) {
    if (!byDriver.has(m.driver_id)) byDriver.set(m.driver_id, { name: m.driver_name, code: m.driver_code, furigana: m.furigana || '', rows: [] });
    byDriver.get(m.driver_id).rows.push(m);
  }

  const pdfHeader = await getPdfHeaderSettings();
  const pdfFooterNote = await getPdfFooterNote();
  const WEEKDAY_LABELS_JP = ['日', '月', '火', '水', '木', '金', '土'];
  const HEADER_ROW = 18, DAY_BLOCK_ROWS = 31;
  const TEMPLATE_MAX_ROW = 71, TEMPLATE_MAX_COL = 6; // ひな形の実際の範囲(B1:F71)。Excel全体コピー用
  const daysInMonth = new Date(year, month, 0).getDate();
  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'shift_pdf_'));

  // PDFとは別に、全員分を1つのExcelファイル(氏名ごとにシート分け)にもまとめる
  const combinedWb = new ExcelJS.Workbook();
  const usedSheetNames = new Set();

  const generatedFiles = [];
  const errors = [];
  for (const [, info] of byDriver) {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(templatePath);
      const ws = wb.worksheets[0];

      // ふりがな欄。ドライバーマスタの「ふりがな」に入力があれば、氏名と同じく半角/全角スペースで
      // 姓・名に分けてC/Dそれぞれに入れる(スペースが無ければ姓側のCにまとめて入れる)。
      // 未入力ならテンプレート本来の「すずき」「しょう」が他の人にも残ってしまわないよう空欄にする
      const furiganaParts = (info.furigana || '').trim().split(/[\s　]+/).filter(Boolean);
      ws.getRow(6).getCell(3).value = furiganaParts[0] || null;
      ws.getRow(6).getCell(4).value = furiganaParts[1] || null;
      ws.getRow(7).getCell(2).value = `氏名　（${info.name}）　殿`;
      ws.getRow(8).getCell(4).value = info.code || '';
      const today = new Date();
      ws.getRow(10).getCell(6).value = excelSafeDate(today.getFullYear(), today.getMonth() + 1, today.getDate()); // 依頼日(発行日)
      ws.getRow(11).getCell(6).value = `${pdfHeader.requester_address}\n${pdfHeader.requester_name}`;
      ws.getRow(12).getCell(6).value = pdfHeader.contact_name;
      ws.getRow(13).getCell(6).value = pdfHeader.contact_phone;
      ws.getRow(13).getCell(2).value = `${year}年`;
      ws.getRow(14).getCell(2).value = month;
      ws.getRow(15).getCell(3).value = year;
      ws.getRow(51).getCell(2).value = pdfFooterNote; // 一番下の注記文言(画面から変更可能)

      const byDate = new Map(info.rows.map(m => [m.match_date, m]));
      for (let i = 0; i < DAY_BLOCK_ROWS; i++) {
        const row = ws.getRow(HEADER_ROW + 1 + i);
        const dayNum = i + 1;
        const bCell = row.getCell(2), cCell = row.getCell(3), dCell = row.getCell(4), eCell = row.getCell(5), fCell = row.getCell(6);
        if (dayNum > daysInMonth) {
          // テンプレートは31行固定だが、月を跨いだ分(翌月の日付)は表示せず空欄にする
          bCell.value = null; cCell.value = null; dCell.value = null; eCell.value = null; fCell.value = null;
          continue;
        }
        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`;
        const dow = new Date(dateStr + 'T00:00:00').getDay();
        bCell.value = excelSafeDate(year, month, dayNum);
        cCell.value = WEEKDAY_LABELS_JP[dow];

        let fillArgb, fontArgb;
        if (dow === 6) { fillArgb = 'FF00B0F0'; fontArgb = 'FFFFFFFF'; }
        else if (dow === 0 || isJapaneseHoliday(dateStr)) { fillArgb = 'FFFF66CC'; fontArgb = 'FFFFFFFF'; }
        else { fillArgb = 'FFFFFFFF'; fontArgb = 'FF000000'; }
        // ※ cCell.fill / cCell.font を個別に代入すると、このテンプレート(既存の複雑なスタイルを
        // 持つセル)特有のExcelJSの不具合で、他の行と色が混ざってしまうことがあった(実際に発生を確認)。
        // border/alignment/numFmtを今の値のまま保持しつつ、スタイル全体を1回で置き換えることで回避する
        cCell.style = {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } },
          font: { ...cCell.font, color: { argb: fontArgb } },
          border: cCell.border,
          alignment: cCell.alignment,
          numFmt: cCell.numFmt
        };

        const match = byDate.get(dateStr);
        if (match) {
          dCell.value = match.time_start || '';
          eCell.value = match.time_end || '';
          fCell.value = match.store_name || '';
        }
      }

      const safeName = info.name.replace(/[\\/:*?"<>|]/g, '');
      const fileStem = `${month}月_${safeName}_シフト`;
      const xlsxPath = path.join(workDir, `${fileStem}.xlsx`);
      await wb.xlsx.writeFile(xlsxPath);
      generatedFiles.push(xlsxPath);

      // 同じ内容を、全員まとめた1つのExcelファイルのシートとしても追加しておく
      const combinedWs = combinedWb.addWorksheet(sanitizeSheetName(info.name, usedSheetNames));
      copyWorksheetContent(combinedWs, ws, TEMPLATE_MAX_ROW, TEMPLATE_MAX_COL);
    } catch (e) {
      errors.push(`${info.name}さんの作成に失敗しました: ${e.message}`);
    }
  }

  if (generatedFiles.length === 0) {
    fs.rmSync(workDir, { recursive: true, force: true });
    return res.status(400).json({ success: false, message: 'PDFを1件も作成できませんでした', errors });
  }

  // LibreOfficeで一括PDF変換(1人ずつsofficeを起動するより、まとめて1回で変換した方が大幅に速い)。
  // 変換後のPDFは、元のxlsxと同じworkDir内に同名(拡張子だけpdf)で出力される
  const result = spawnSync(sofficePath, ['--headless', '--norestore', '--convert-to', 'pdf', '--outdir', workDir, ...generatedFiles], { timeout: 300000 });
  if (result.error || result.status !== 0) {
    fs.rmSync(workDir, { recursive: true, force: true });
    return res.status(500).json({ success: false, message: 'PDF変換に失敗しました: ' + (result.error ? result.error.message : (result.stderr || '').toString()) });
  }

  // サーバーのPC(このPCでもクラウド上でも)のフォルダに直接保存するのではなく、PDF一式と
  // 全員分まとめたExcelをzipにまとめてブラウザからダウンロードしてもらう形にする
  // (クラウド環境ではサーバー側の「デスクトップ」を見に行くことができないため)
  const excelFileName = `個人事業主別シフト_${year}年${String(month).padStart(2, '0')}月.xlsx`;
  const zipFileName = `driver_shift_pdfs_${year}${String(month).padStart(2, '0')}.zip`;
  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename="${zipFileName}"`);
  res.setHeader('X-Export-Count', String(generatedFiles.length));
  if (errors.length > 0) res.setHeader('X-Export-Errors', encodeURIComponent(JSON.stringify(errors)));
  res.setHeader('Access-Control-Expose-Headers', 'X-Export-Count, X-Export-Errors');

  // zip作成中に何かエラーが起きても(archiverのバージョン不一致など)、必ずworkDirの
  // 一時ファイルを片付ける。レスポンスをまだ何も書き出していなければJSONエラーで返せるが、
  // 途中まで送信済みの場合はそのまま接続を切るしかない
  try {
    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.on('warning', (e) => console.warn('shift PDF zip warning:', e.message));
    archive.on('error', (e) => { console.error('shift PDF zip error:', e); res.destroy(e); });
    archive.on('end', () => fs.rmSync(workDir, { recursive: true, force: true }));
    archive.pipe(res);
    for (const xlsxPath of generatedFiles) {
      const pdfPath = xlsxPath.replace(/\.xlsx$/, '.pdf');
      if (fs.existsSync(pdfPath)) archive.file(pdfPath, { name: path.basename(pdfPath) });
    }
    archive.append(await combinedWb.xlsx.writeBuffer(), { name: excelFileName });
    await archive.finalize();
  } catch (e) {
    console.error('shift PDF zip creation failed:', e);
    fs.rmSync(workDir, { recursive: true, force: true });
    if (!res.headersSent) {
      return res.status(500).json({ success: false, message: 'zip作成に失敗しました: ' + e.message });
    }
    res.destroy(e);
  }
});

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

  // 「対応なし」に設定した(=手配を諦めた)時間の穴を、店舗ごとにまとめておき、シフト表に
  // 「この時間帯は手配できていません」と明記できるようにする。穴がその後埋まって解消済みなら出さない
  const dismissedRequests = await dbAll(`
    SELECT id, store_name, request_date, time_start, time_end
    FROM store_requests WHERE gap_dismissed = 1 AND archived_month IS NULL
  `);
  const gapNoteByStore = new Map(); // 店舗名 -> [{date, start, end}, ...]
  for (const r of dismissedRequests) {
    const rms = await dbAll(
      `SELECT time_start, time_end, status FROM matches WHERE store_request_id = ? AND archived_month IS NULL`,
      [r.id]
    );
    const gaps = computeTimeGaps(r.time_start, r.time_end, rms.filter(m => m.status !== '欠勤'));
    if (gaps.length === 0) continue;
    if (!gapNoteByStore.has(r.store_name)) gapNoteByStore.set(r.store_name, []);
    for (const g of gaps) gapNoteByStore.get(r.store_name).push({ date: r.request_date, start: g.start, end: g.end });
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
    const gapNotes = (gapNoteByStore.get(info.name) || []).sort((a, b) => a.date.localeCompare(b.date) || a.start.localeCompare(b.start));
    if (gapNotes.length > 0) {
      aoa.push([], ['⚠️ 手配できていない時間帯(ご了承ください)']);
      aoa.push(['日付', '曜日', '時間帯']);
      for (const g of gapNotes) aoa.push([g.date, weekdayLabelOf(g.date), `${g.start}〜${g.end}`]);
    }
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

// ひな形(記入用)シートを1枚追加する。見出し行は太字+背景色+罫線、記入例の行は斜体のグレー文字+罫線にして
// 「後で消す行」だと一目で分かるようにする(xlsxパッケージは罫線等の装飾書き出しに対応していないためexceljsを使う)
const THIN_BORDER = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
function addTemplateDataSheet(workbook, sheetName, header, sampleRows, colWidths) {
  const ws = workbook.addWorksheet(sheetName);
  ws.columns = colWidths.map(w => ({ width: w }));

  const headerRow = ws.addRow(header);
  headerRow.eachCell({ includeEmpty: true }, cell => {
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5E7EB' } };
    cell.border = THIN_BORDER;
    cell.alignment = { vertical: 'middle', wrapText: true };
  });

  for (const r of sampleRows) {
    const row = ws.addRow(r);
    row.eachCell({ includeEmpty: true }, cell => {
      cell.border = THIN_BORDER;
      cell.font = { italic: true, color: { argb: 'FF9CA3AF' } }; // 記入例は薄いグレーの斜体にして目立たせる
      cell.alignment = { vertical: 'middle', wrapText: true };
    });
  }
  ws.views = [{ state: 'frozen', ySplit: 1 }]; // 見出し行を固定して、スクロールしても見えるようにする
  return ws;
}

// 練習用サンプルデータのシートを追加する(addTemplateDataSheetと違い、データ行を「消してください」の
// 記入例ではなく、そのまま取込んで使う本物のサンプル行として扱うため、薄いグレーの斜体にはしない)
function addFilledSampleSheet(workbook, sheetName, header, dataRows, colWidths) {
  const ws = workbook.addWorksheet(sheetName);
  ws.columns = colWidths.map(w => ({ width: w }));
  const headerRow = ws.addRow(header);
  headerRow.eachCell({ includeEmpty: true }, cell => {
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5E7EB' } };
    cell.border = THIN_BORDER;
    cell.alignment = { vertical: 'middle', wrapText: true };
  });
  for (const r of dataRows) {
    const row = ws.addRow(r);
    row.eachCell({ includeEmpty: true }, cell => {
      cell.border = THIN_BORDER;
      cell.alignment = { vertical: 'middle', wrapText: true };
    });
  }
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  return ws;
}

// 記入ルール(文章だけの説明)シートを追加する。罫線は付けず、タイトル行だけ太字・少し大きめにする
function addRulesTextSheet(workbook, sheetName, lines) {
  const ws = workbook.addWorksheet(sheetName);
  ws.columns = [{ width: 95 }];
  lines.forEach((line, i) => {
    const row = ws.addRow(line);
    if (i === 0) row.getCell(1).font = { bold: true, size: 13 };
  });
  return ws;
}

// 店舗依頼(週間必要枠表)の、取込みにそのまま使えるひな形Excelを生成する。
// 「これ聞取り固定デポ」等の実ファイルで実績のある構造(店舗名+月〜日の曜日列+備考)に合わせてあり、
// 余計な分析用の列は含めない(あくまで取込みに必要な最小限の列のみ)
app.get('/api/templates/store-requests-weekly.xlsx', async (req, res) => {
  const header = ['店番', '店舗名', '月', '火', '水', '木', '金', '土', '日', '備考'];
  const sampleRows = [
    ['235', '（記入例）銀座SS', '10:00-22:00', '10:00-22:00', '10:00-22:00', '10:00-22:00', '10:00-22:00', '10:00-22:00', '10:00-22:00', ''],
    ['721', '（記入例）築地店', '10:00-22:00', '', '10:00-22:00', '', '10:00-22:00', '10:00-22:00', '', '火・木・日はお休み(人がいらない曜日)の例。セルは空っぽのままでOK'],
    ['', '（記入例）東雲店', '10:00-22:00*2', '10:00-22:00*2', '10:00-22:00*2', '10:00-22:00*2', '10:00-22:00*2', '10:00-22:00*3', '10:00-22:00*3', '店番が分からない時は空欄でOK。1日に2人以上ほしい時は「*3」のように書く'],
  ];
  const rulesSheetText = [
    ['【店舗依頼】の書き方(むずかしく考えなくて大丈夫です)'],
    [],
    ['1行で、1つの店舗の「いつ・何時から何時まで・何人」をあらわします。'],
    ['使うときは、2〜4行目の(記入例)はぜんぶ消してから、自分のお店の分を書いてください。'],
    [],
    ['① 「店番」の列に、お店の番号(拠点コード)を書く。わからなければ空欄でもOK'],
    ['② 「店舗名」の列に、お店の名前を書く'],
    ['③ 月〜日の列に、その曜日に人がほしい時間を「開始時刻-終了時刻」で書く'],
    ['　　　書き方の例 → 10:00-22:00 (10時から22時までの意味)'],
    ['④ その曜日は人がいらない(休み)なら、何も書かずに空っぽのままにする'],
    ['⑤ 1日に2人以上ほしいときは、時間の右側に「*(ほしい人数)」を付け足す'],
    ['　　　書き方の例 → 10:00-22:00*2 (2人ほしいという意味。何も付けなければ1人の意味になります)'],
    ['⑥ 同じ店舗で、時間帯が違う人がほしい時は、行をもう1行足して時間帯を変えて書く'],
    ['　　　書き方の例 → 1行目「10:00-18:00」、2行目(同じ店舗名)「18:00-22:00」'],
    ['⑦ 備考は自由に書いてOK。何も書かなくても構いません'],
    [],
    ['よくある質問'],
    ['Q. 店番ってなに？なぜあった方がいいの？'],
    ['A. お店ごとに割り振られている番号です。店舗名は表記ゆれ(スペースの有無など)で別のお店として扱われてしまうことがありますが、店番があれば間違いなく同じお店だと分かります。無くても取込みはできます。'],
    ['Q. 同じお店・同じ曜日・同じ時間帯の行が2行あってもいい？'],
    ['A. 大丈夫です。合算されて必要人数として扱われます(時間帯が違う場合は別々の依頼として扱われます)。'],
  ];

  const wb = new ExcelJS.Workbook();
  addTemplateDataSheet(wb, '店舗依頼ひな形', header, sampleRows, [10, 22, 15, 15, 15, 15, 15, 15, 15, 42]);
  addRulesTextSheet(wb, '記入ルール', rulesSheetText);

  const buffer = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="monthly_request_template.xlsx"');
  res.send(Buffer.from(buffer));
});

// 希望シフト(長形式、1行=1人×1日)の、取込みにそのまま使えるひな形Excelを生成する。
// 楽シフ等の実際のCSV列名がまだ分かっていないため、現時点の想定列名(AVAILABILITY_ALIASESの主要なもの)で
// 用意している。実際の楽シフCSVのヘッダー行が分かり次第、列名を合わせて調整する想定
app.get('/api/templates/driver-availability-long.xlsx', async (req, res) => {
  const header = ['ドライバー名', '希望日', '希望エリア', '希望店舗', '開始時刻', '終了時刻', '備考'];
  const sampleRows = [
    ['（記入例）山田太郎', '2026-11-01', '', '銀座SS', '10:00', '22:00', ''],
    ['（記入例）山田太郎', '2026-11-02', '', '', '10:00', '20:00', '希望店舗が空欄でも、固定希望店舗が設定されていればそれが使われます'],
    ['（記入例）山田太郎', '2026-11-04', '渋谷区', '', '12:00', '22:00', '店舗名が分からない場合は希望エリアだけでもOK'],
  ];
  const rulesSheetText = [
    ['希望シフト(長形式)ひな形の使い方'],
    [],
    ['① 1行=1人のドライバーの、1日分の希望シフトです。実際に取込む際は、記入例の行を削除してドライバー名を入れ替えてください。'],
    ['② ドライバー名は「ドライバーマスタ」に登録済みの氏名と完全一致している必要があります(先に登録してください)。'],
    ['③ 休み(その日は稼働しない)の場合は、その日の行自体を入れないでください(空のセルではなく、行ごと無しにする形式です)。'],
    ['④ 希望エリア・希望店舗は分かる範囲で構いません。両方空欄でも、ドライバーマスタの「固定希望店舗」が設定されていれば自動的に使われます。'],
    ['⑤ 同じドライバー・同じ希望日の行が複数あると、後に読み込んだ方で上書きされます。'],
    [],
    ['※ 楽シフのエクスポートは、この長形式ではなく「氏名×日付」のワイド形式です(そちらは別の取込み機能で対応済みです)。こちらは手入力やその他ツール向けの形式です。'],
  ];

  const wb = new ExcelJS.Workbook();
  addTemplateDataSheet(wb, '希望シフトひな形', header, sampleRows, [16, 12, 12, 16, 10, 10, 42]);
  addRulesTextSheet(wb, '記入ルール', rulesSheetText);

  const buffer = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="driver_availability_template.xlsx"');
  res.send(Buffer.from(buffer));
});

// ===== 使い方を練習するためのサンプルデータ(担当者への説明・操作練習用) =====
// 店舗名・氏名はすべて「（サンプル）」を付けて、本物のデータと見分けやすくしてある。
// 取込みの練習後は、ドライバーマスタ/店舗マスタから削除するか、店舗依頼・希望シフトは
// 「📥 取込み履歴」からまとめて取り消せば元の状態に戻せる

// ①サンプルのドライバー3名(ドライバーマスタの一括取込みの練習用)
app.get('/api/samples/drivers.xlsx', async (req, res) => {
  const header = ['社員コード', '氏名', '会社名', 'お住まい住所', '初回委託日', 'ステータス', 'メールアドレス', '固定希望店舗', '保険加入状況', '電話番号', '車両種別', '備考'];
  const rows = [
    ['99001', '（サンプル）山田太郎', '', '', '2026-04-01', '稼働中', '', '', '加入済み', '090-0000-0001', '軽貨物（バン）', '練習用のサンプルデータです。練習が終わったら削除してOKです'],
    ['99002', '（サンプル）佐藤花子', '', '', '2026-04-01', '稼働中', '', '', '加入済み', '090-0000-0002', '軽貨物（バン）', '練習用のサンプルデータです。練習が終わったら削除してOKです'],
    ['99003', '（サンプル）鈴木次郎', '', '', '2026-04-01', '稼働中', '', '', '加入済み', '090-0000-0003', '軽貨物（バン）', '練習用のサンプルデータです。練習が終わったら削除してOKです'],
  ];
  const rulesText = [
    ['（サンプル）ドライバーマスタ練習用データ'],
    [],
    ['このファイルを「ドライバーマスタ」タブの「Excel/CSVから一括取込み」にそのままアップロードしてみてください。'],
    ['氏名の先頭に「（サンプル）」を付けてあるので、一覧の中でもすぐ見分けられます。'],
    ['練習が終わったら、この3名はドライバーマスタの画面から削除してください(削除すると、この後取り込む希望シフトも一緒に消えます)。'],
  ];
  const wb = new ExcelJS.Workbook();
  addFilledSampleSheet(wb, 'サンプルドライバー', header, rows, [10, 20, 14, 10, 14, 10, 18, 14, 14, 16, 16, 36]);
  addRulesTextSheet(wb, '説明', rulesText);
  const buffer = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="sample_drivers.xlsx"');
  res.send(Buffer.from(buffer));
});

// ②サンプルの店舗依頼3店舗分(週間必要枠表の一括取込みの練習用。番号は実在の店番と被らないよう9900番台にしてある)
app.get('/api/samples/store-requests.xlsx', async (req, res) => {
  const header = ['店番', '店舗名', '月', '火', '水', '木', '金', '土', '日', '備考'];
  const rows = [
    ['9901', '（サンプル）すずらん通り店', '10:00-19:00', '10:00-19:00', '10:00-19:00', '10:00-19:00', '10:00-19:00', '', '', '練習用のサンプルデータです'],
    ['9902', '（サンプル）ひまわり公園前店', '9:00-18:00', '9:00-18:00', '', '9:00-18:00', '9:00-18:00', '9:00-18:00*2', '', '土曜日は2人ほしい例(*2)'],
    ['9903', '（サンプル）みどり橋店', '13:00-21:00', '', '13:00-21:00', '', '13:00-21:00', '', '13:00-21:00', ''],
  ];
  const rulesText = [
    ['（サンプル）店舗依頼(週間必要枠表)練習用データ'],
    [],
    ['このファイルを「店舗依頼」タブの「週間必要枠表(曜日パターン)から一括取込み」にそのままアップロードしてみてください。'],
    ['取込み時に年月を指定する欄がありますが、何月にしても練習には問題ありません(例:来月を指定してみてください)。'],
    ['店舗名の先頭に「（サンプル）」を付けてあるので、一覧の中でもすぐ見分けられます。'],
    ['練習が終わったら、ダッシュボードの「📥 取込み履歴」からこの取込みを選んで「取り消す」を押せば、作成された店舗依頼がまとめて削除されます。'],
    ['(店舗マスタに残る「（サンプル）◯◯店」自体を消したい場合は、店舗マスタの画面から個別に削除してください)'],
  ];
  const wb = new ExcelJS.Workbook();
  addFilledSampleSheet(wb, 'サンプル店舗依頼', header, rows, [10, 22, 15, 15, 15, 15, 15, 15, 15, 30]);
  addRulesTextSheet(wb, '説明', rulesText);
  const buffer = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="sample_store_requests.xlsx"');
  res.send(Buffer.from(buffer));
});

// ③サンプルの希望シフト(氏名×日付のワイド形式、楽シフのエクスポートと同じ形)。
// ①のサンプルドライバー3名と氏名が一致しているので、①→③の順に取込むとエラーなく練習できる
app.get('/api/samples/driver-availability-wide.xlsx', async (req, res) => {
  const header = ['氏名', '1(月)', '2(火)', '3(水)', '4(木)', '5(金)', '6(土)', '7(日)', '8(月)', '9(火)', '10(水)', '11(木)', '12(金)', '13(土)', '14(日)'];
  const rows = [
    ['（サンプル）山田太郎', '10:00-19:00', '10:00-19:00', '休み', '10:00-19:00', '10:00-19:00', '', '', '10:00-19:00', '10:00-19:00', '休み', '10:00-19:00', '10:00-19:00', '', ''],
    ['（サンプル）佐藤花子', '9:00-18:00', '休み', '9:00-18:00', '9:00-18:00', '9:00-18:00', '9:00-18:00', '', '9:00-18:00', '休み', '9:00-18:00', '9:00-18:00', '9:00-18:00', '9:00-18:00', ''],
    ['（サンプル）鈴木次郎', '', '13:00-21:00', '13:00-21:00', '', '13:00-21:00', '13:00-21:00', '13:00-21:00', '', '13:00-21:00', '13:00-21:00', '', '13:00-21:00', '13:00-21:00', '13:00-21:00'],
  ];
  const rulesText = [
    ['（サンプル）希望シフト(ワイド形式/楽シフと同じ形式)練習用データ'],
    [],
    ['先に「①サンプルドライバー」を取り込んでから、このファイルを「ドライバーマスタ」タブの「楽シフのエクスポート(氏名×日付のワイド形式)から一括取込み」にアップロードしてください。'],
    ['氏名が①のサンプルドライバーと完全に一致しているので、先に①を取り込んでいればそのまま登録できます。'],
    ['取込み時に年月を指定する欄がありますが、何月にしても練習には問題ありません(「1(月)」などの曜日表記は見た目だけのものです)。'],
    ['「休み」と書いてある日・空欄の日は、その日は稼働しない扱いになります。'],
    ['練習が終わったら、①のサンプルドライバーをドライバーマスタから削除すれば、この希望シフトも一緒に消えます。'],
  ];
  const wb = new ExcelJS.Workbook();
  addFilledSampleSheet(wb, 'サンプル希望シフト', header, rows, [18, ...Array(14).fill(11)]);
  addRulesTextSheet(wb, '説明', rulesText);
  const buffer = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="sample_driver_availability_wide.xlsx"');
  res.send(Buffer.from(buffer));
});

app.listen(PORT, () => console.log(`マッチングアプリ起動: http://localhost:${PORT}`));
