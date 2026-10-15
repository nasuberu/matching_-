# クラウド(Render等)で動かす用。ローカルPCでの開発・起動(npm start)には影響しない。
# 個人事業主別シフトPDFの作成にLibreOffice(soffice)を使っているため、Node本体だけでなく
# LibreOfficeも一緒に入れたイメージにしている(素のNodeビルドパックだとOSパッケージを追加できないため)
FROM node:20-bookworm-slim

# LibreOffice(headless変換用)と、日本語フォント(PDF内の日本語が文字化け・豆腐化しないように)を入れる
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice \
      fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY . .

# DB・アップロードファイルは永続ディスク(/data)に保存する(render.yamlでボリュームをマウントする前提)
ENV DB_PATH=/data/matching.db
ENV UPLOADS_DIR=/data/uploads

EXPOSE 4001
CMD ["node", "server.js"]
