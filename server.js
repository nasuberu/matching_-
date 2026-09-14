const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');
const iconv = require('iconv-lite');
const jschardet = require('jschardet');

const app = express();
const PORT = 4001;
const DIST_WARNING_KM = 30; // これを超える距離のマッチングは「距離が遠い」として画面上で警告表示する
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

// 自動マッチングのスコアリング重み(調整しやすいようファイル冒頭に定数化しておく)
const SCORE_WEIGHT_DISTANCE_KM = -1;     // 距離1kmごとの減点
const SCORE_STRONG_MATCH_BONUS = 50;     // 希望店舗/希望エリアが一致する場合のボーナス
const SCORE_PREFERENCE_GOOD_BONUS = 30;  // その店舗を「好き」に設定している場合のボーナス
const SCORE_EXPERIENCE_PER_VISIT = 5;    // その店舗での過去派遣1回あたりのボーナス(店舗の勝手を知っている)
const SCORE_EXPERIENCE_MAX_VISITS = 5;   // 店舗経験ボーナスの上限回数
const SCORE_AREA_EXPERIENCE_PER_VISIT = 2; // 同エリアでの過去派遣1回あたりのボーナス(配達エリアの知見)
const SCORE_AREA_EXPERIENCE_MAX_VISITS = 5; // エリア経験ボーナスの上限回数

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const db = new sqlite3.Database(path.join(__dirname, 'matching.db'));

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

  // ドライバーごとの店舗との相性(好き/NG)。NGはマッチング候補から除外し、好きは優先度を上げる
  db.run(`
    CREATE TABLE IF NOT EXISTS driver_store_preferences (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      driver_id INTEGER NOT NULL,
      store_id INTEGER NOT NULL,
      preference TEXT NOT NULL CHECK (preference IN ('good', 'ng')),
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
});

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

// 店舗名(自由入力)を店舗マスタに名寄せする。既存店舗が見つかればそのidを返し、
// 住所/緯度経度が未設定であれば補完する。見つからなければ新規に登録する
async function resolveStoreId({ store_name, area, address, lat, lng }) {
  const name = (store_name || '').trim();
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
  if (preference === 'good') score += SCORE_PREFERENCE_GOOD_BONUS;
  score += experienceCount * SCORE_EXPERIENCE_PER_VISIT;
  score += areaExperienceCount * SCORE_AREA_EXPERIENCE_PER_VISIT;

  return { distance, score, preference, experienceCount };
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
  const rows = await dbAll('SELECT * FROM drivers ORDER BY id DESC');
  res.json({ success: true, drivers: rows });
});

app.post('/api/drivers', async (req, res) => {
  const { id, name, phone, vehicle_type, home_address, driver_code, company_name, first_contract_date, status, email, insurance_info, notes } = req.body;
  if (!name) return res.status(400).json({ success: false, message: '氏名は必須です' });

  // 住所が変わった場合のみ再ジオコーディングする(毎回叩くと無駄なため)
  let lat = null, lng = null;
  if (home_address) {
    const existing = id ? await dbGet('SELECT home_address, home_lat, home_lng FROM drivers WHERE id = ?', [id]) : null;
    if (existing && existing.home_address === home_address && existing.home_lat != null) {
      lat = existing.home_lat; lng = existing.home_lng;
    } else {
      const geo = await geocodeAddress(home_address);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }
  }

  const now = new Date().toISOString();
  if (id) {
    await dbRun(
      `UPDATE drivers SET name=?, phone=?, vehicle_type=?, home_address=?, home_lat=?, home_lng=?, driver_code=?, company_name=?, first_contract_date=?, status=?, email=?, insurance_info=?, notes=? WHERE id=?`,
      [name, phone || '', vehicle_type || '', home_address || '', lat, lng, driver_code || null, company_name || '', first_contract_date || '', status || '', email || '', insurance_info || '', notes || '', id]
    );
    return res.json({ success: true, id });
  }
  const result = await dbRun(
    `INSERT INTO drivers (name, phone, vehicle_type, home_address, home_lat, home_lng, driver_code, company_name, first_contract_date, status, email, insurance_info, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [name, phone || '', vehicle_type || '', home_address || '', lat, lng, driver_code || null, company_name || '', first_contract_date || '', status || '', email || '', insurance_info || '', notes || '', now]
  );
  res.json({ success: true, id: result.lastID });
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
        `UPDATE drivers SET name=?, phone=?, vehicle_type=?, home_address=?, home_lat=?, home_lng=?, driver_code=?, company_name=?, first_contract_date=?, status=?, email=?, insurance_info=?, notes=? WHERE id=?`,
        [name, phone, vehicle_type, home_address, lat, lng, driver_code, company_name, first_contract_date, status, email, insurance_info, notes, existing.id]
      );
    } else {
      await dbRun(
        `INSERT INTO drivers (name, phone, vehicle_type, home_address, home_lat, home_lng, driver_code, company_name, first_contract_date, status, email, insurance_info, notes, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [name, phone, vehicle_type, home_address, lat, lng, driver_code, company_name, first_contract_date, status, email, insurance_info, notes, now]
      );
    }
    imported++;
  }

  res.json({ success: true, imported, total: rows.length, errors });
});

app.delete('/api/drivers/:id', async (req, res) => {
  await dbRun('DELETE FROM drivers WHERE id = ?', [req.params.id]);
  await dbRun('DELETE FROM driver_availability WHERE driver_id = ?', [req.params.id]);
  res.json({ success: true });
});

// ===== ドライバーの希望シフト =====
app.get('/api/driver-availability', async (req, res) => {
  const rows = await dbAll(`
    SELECT a.*, d.name AS driver_name, d.phone AS driver_phone
    FROM driver_availability a JOIN drivers d ON d.id = a.driver_id
    ORDER BY a.desired_date DESC, a.id DESC
  `);
  res.json({ success: true, availability: rows });
});

app.post('/api/driver-availability', async (req, res) => {
  const { driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests } = req.body;
  if (!driver_id || !desired_date) return res.status(400).json({ success: false, message: 'driver_id, desired_date は必須です' });
  const now = new Date().toISOString();
  const result = await dbRun(
    `INSERT INTO driver_availability (driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
    [driver_id, desired_date, desired_area || '', desired_store || '', time_start || '', time_end || '', requests || '', now]
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

  const drivers = await dbAll('SELECT id, name FROM drivers');
  const driverByName = new Map(drivers.map(d => [d.name.trim(), d.id]));

  const now = new Date().toISOString();
  let imported = 0;
  const errors = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const driverName = pickField(row, AVAILABILITY_ALIASES.driver_name);
    const desired_date = normalizeDateStr(pickField(row, AVAILABILITY_ALIASES.desired_date));
    if (!driverName || !desired_date) { errors.push(`${i + 2}行目: ドライバー名または希望日が読み取れませんでした`); continue; }

    const driver_id = driverByName.get(driverName);
    if (!driver_id) { errors.push(`${i + 2}行目: ドライバー「${driverName}」がドライバーマスタに見つかりません(先に登録してください)`); continue; }

    await dbRun(
      `INSERT INTO driver_availability (driver_id, desired_date, desired_area, desired_store, time_start, time_end, requests, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        driver_id, desired_date,
        pickField(row, AVAILABILITY_ALIASES.desired_area),
        pickField(row, AVAILABILITY_ALIASES.desired_store),
        pickField(row, AVAILABILITY_ALIASES.time_start),
        pickField(row, AVAILABILITY_ALIASES.time_end),
        pickField(row, AVAILABILITY_ALIASES.requests),
        now
      ]
    );
    imported++;
  }

  res.json({ success: true, imported, total: rows.length, errors });
});

// ===== 店舗からの人員要請 =====
app.get('/api/store-requests', async (req, res) => {
  const rows = await dbAll('SELECT * FROM store_requests ORDER BY request_date DESC, id DESC');
  res.json({ success: true, requests: rows });
});

app.post('/api/store-requests', async (req, res) => {
  const { id, store_name, area, address, request_date, time_start, time_end, required_count, requests } = req.body;
  if (!store_name || !request_date) return res.status(400).json({ success: false, message: 'store_name, request_date は必須です' });

  let lat = null, lng = null;
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

  const now = new Date().toISOString();
  if (id) {
    await dbRun(
      `UPDATE store_requests SET store_name=?, area=?, address=?, lat=?, lng=?, request_date=?, time_start=?, time_end=?, required_count=?, requests=?, store_id=? WHERE id=?`,
      [store_name, area || '', address || '', lat, lng, request_date, time_start || '', time_end || '', required_count || 1, requests || '', store_id, id]
    );
    return res.json({ success: true, id });
  }
  const result = await dbRun(
    `INSERT INTO store_requests (store_name, area, address, lat, lng, request_date, time_start, time_end, required_count, requests, created_at, store_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [store_name, area || '', address || '', lat, lng, request_date, time_start || '', time_end || '', required_count || 1, requests || '', now, store_id]
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

    const address = pickField(row, STORE_ALIASES.address);
    const area = pickField(row, STORE_ALIASES.area);
    let lat = null, lng = null;
    if (address) {
      const geo = await geocodeAddress(address);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }
    const store_id = await resolveStoreId({ store_name, area, address, lat, lng });

    await dbRun(
      `INSERT INTO store_requests (store_name, area, address, lat, lng, request_date, time_start, time_end, required_count, requests, created_at, store_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        store_name,
        area,
        address, lat, lng, request_date,
        pickField(row, STORE_ALIASES.time_start),
        pickField(row, STORE_ALIASES.time_end),
        parseInt(pickField(row, STORE_ALIASES.required_count), 10) || 1,
        pickField(row, STORE_ALIASES.requests),
        now,
        store_id
      ]
    );
    imported++;
  }

  res.json({ success: true, imported, total: rows.length, errors });
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

// 相性の登録・更新(driver_id+store_idの組でupsert)
app.post('/api/driver-store-preferences', async (req, res) => {
  const { driver_id, store_id, preference, notes } = req.body;
  if (!driver_id || !store_id || !['good', 'ng'].includes(preference)) {
    return res.status(400).json({ success: false, message: 'driver_id, store_id, preference(good/ng) は必須です' });
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
  const { driver_id, store_name, work_date, notes } = req.body;
  if (!driver_id || !store_name || !work_date) {
    return res.status(400).json({ success: false, message: 'driver_id, store_name, work_date は必須です' });
  }
  const store_id = await resolveStoreId({ store_name });
  const now = new Date().toISOString();
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
    SELECT m.*, d.name AS driver_name, d.phone AS driver_phone, d.home_address,
           s.store_name, s.area, s.address AS store_address, s.time_start, s.time_end, s.requests AS store_requests
    FROM matches m
    JOIN drivers d ON d.id = m.driver_id
    JOIN store_requests s ON s.id = m.store_request_id
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
  await dbRun('DELETE FROM matches'); // 一旦候補を作り直す(確定済みの運用に育ったら「候補のみ削除」に変更する想定)

  const storeRequests = await dbAll('SELECT * FROM store_requests ORDER BY request_date ASC');
  const availability = await dbAll(`
    SELECT a.*, d.name AS driver_name, d.home_lat, d.home_lng
    FROM driver_availability a JOIN drivers d ON d.id = a.driver_id
  `);
  const { preferenceByDriverStore, storeVisitCount, areaVisitCount } = await loadScoringContext();

  const now = new Date().toISOString();
  const assignedDriverIdsByDate = {}; // date -> Set(driver_id) 同日の重複割当を防ぐ
  let createdCount = 0;
  let noCandidateCount = 0;

  for (const store of storeRequests) {
    assignedDriverIdsByDate[store.request_date] = assignedDriverIdsByDate[store.request_date] || new Set();
    const assignedToday = assignedDriverIdsByDate[store.request_date];

    // 同じ日付の希望を持ち、この店舗をNGにしていないドライバーを候補にする
    const candidates = availability.filter(a =>
      a.desired_date === store.request_date &&
      !assignedToday.has(a.driver_id) &&
      preferenceByDriverStore.get(`${a.driver_id}:${store.store_id}`) !== 'ng'
    );

    const scored = candidates
      .map(a => ({ ...a, ...scoreCandidate(a, store, preferenceByDriverStore, storeVisitCount, areaVisitCount) }))
      .sort((x, y) => y.score - x.score);

    const needed = store.required_count || 1;
    const picked = scored.slice(0, needed);
    if (picked.length === 0) { noCandidateCount++; continue; }

    for (const p of picked) {
      const isFar = p.distance != null && p.distance > DIST_WARNING_KM;
      await dbRun(
        `INSERT INTO matches (store_request_id, driver_id, match_date, distance_km, is_far_warning, status, created_at, score, preference_flag, experience_count) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [store.id, p.driver_id, store.request_date, p.distance, isFar ? 1 : 0, '候補', now, p.score, p.preference, p.experienceCount]
      );
      assignedToday.add(p.driver_id);
      createdCount++;
    }
  }

  res.json({ success: true, created: createdCount, unmatched_requests: noCandidateCount });
});

app.post('/api/matches/:id/confirm', async (req, res) => {
  await dbRun(`UPDATE matches SET status = '確定' WHERE id = ?`, [req.params.id]);
  res.json({ success: true });
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
    `SELECT driver_id FROM matches WHERE match_date = ? AND status IN ('候補', '確定', '完了')`,
    [match.match_date]
  );
  const excluded = new Set(busyRows.map(r => r.driver_id));
  excluded.add(match.driver_id);

  const drivers = await dbAll('SELECT id, name, phone, home_lat, home_lng FROM drivers');
  const availabilityToday = await dbAll(
    'SELECT driver_id, desired_store, desired_area FROM driver_availability WHERE desired_date = ?',
    [match.match_date]
  );
  const availabilityByDriver = new Map(availabilityToday.map(a => [a.driver_id, a]));
  const { preferenceByDriverStore, storeVisitCount, areaVisitCount } = await loadScoringContext();

  const candidates = drivers
    .filter(d => !excluded.has(d.id))
    .filter(d => preferenceByDriverStore.get(`${d.id}:${match.store_id}`) !== 'ng')
    .map(d => {
      const av = availabilityByDriver.get(d.id);
      const candidateLike = {
        driver_id: d.id, home_lat: d.home_lat, home_lng: d.home_lng,
        desired_store: av ? av.desired_store : null, desired_area: av ? av.desired_area : null
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

  const driver = await dbGet('SELECT id, home_lat, home_lng FROM drivers WHERE id = ?', [driver_id]);
  if (!driver) return res.status(404).json({ success: false, message: 'ドライバーが見つかりません' });

  const av = await dbGet(
    'SELECT desired_store, desired_area FROM driver_availability WHERE driver_id = ? AND desired_date = ?',
    [driver_id, match.match_date]
  );
  const { preferenceByDriverStore, storeVisitCount, areaVisitCount } = await loadScoringContext();
  const candidateLike = {
    driver_id: driver.id, home_lat: driver.home_lat, home_lng: driver.home_lng,
    desired_store: av ? av.desired_store : null, desired_area: av ? av.desired_area : null
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

app.delete('/api/matches/:id', async (req, res) => {
  await dbRun('DELETE FROM matches WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

app.listen(PORT, () => console.log(`マッチングアプリ起動: http://localhost:${PORT}`));
