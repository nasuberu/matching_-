// ローカル起動用（Vercel以外での実行時）
if (process.env.NODE_ENV !== 'production') {
  const PORT = process.env.PORT || 4001;
  app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
}

// Vercel用にappをエクスポート
module.exports = app;