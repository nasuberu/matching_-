const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();

// ミドルウェアの設定
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// 静的ファイルの提供（publicフォルダがある場合）
const publicDir = path.join(__dirname, 'public');
if (fs.existsSync(publicDir)) {
  app.use(express.static(publicDir));
}

// データベース設定（Vercelサーバーレス環境対応）
let db = null;
try {
  const sqlite3 = require('sqlite3').verbose();
  const dbPath = path.join(__dirname, 'matching.db');
  db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
      console.error('Database connection error:', err.message);
    } else {
      console.log('Connected to SQLite database.');
    }
  });
} catch (e) {
  console.warn('SQLite3 load warning (Vercel environment):', e.message);
}

// ルートパス（トップページ）のリクエスト処理
app.get('/', (req, res) => {
  const indexPath = path.join(__dirname, 'public', 'index.html');
  if (fs.existsSync(indexPath)) {
    res.sendFile(indexPath);
  } else {
    res.send('<h1>マッチングアプリサーバーが正常に稼働しています</h1>');
  }
});

// ローカル開発環境でのサーバー起動処理
if (process.env.NODE_ENV !== 'production' && !process.env.VERCEL) {
  const PORT = process.env.PORT || 4001;
  app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
  });
}

// Vercel用に app をエクスポート
module.exports = app;