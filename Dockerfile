# --- 第一階段：建置環境 (Build Stage) ---
# 使用 Node.js 映像檔來進行安裝與打包
FROM node:18-alpine as builder

# 設定工作目錄
WORKDIR /app

# 複製 package.json 和 package-lock.json
COPY package*.json ./

# 安裝依賴套件
RUN npm install

# 複製所有程式碼到工作目錄
COPY . .

# 執行 Webpack 打包 (產出 dist 資料夾)
RUN npm run build

# --- 第二階段：執行環境 (Production Stage) ---
# 使用輕量的 Nginx 伺服器來提供靜態網頁服務
FROM nginx:alpine

# 將第一階段打包好的 dist 資料夾內容，複製到 Nginx 的預設目錄
COPY --from=builder /app/dist /usr/share/nginx/html

# 🔥🔥🔥 關鍵修改：修改 Nginx 設定，讓它監聽 8080 port (Cloud Run 的預設要求)
# 這行指令會把預設設定檔裡的 "listen 80;" 改成 "listen 8080;"
RUN sed -i 's/listen       80;/listen       8080;/' /etc/nginx/conf.d/default.conf

# 宣告監聽 8080
EXPOSE 8080

# 啟動 Nginx
CMD ["nginx", "-g", "daemon off;"]