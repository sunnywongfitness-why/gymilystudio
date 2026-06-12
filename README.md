# Gymily Studio 預約系統

一個 React + Vite 的場地預約系統，已設定成 PWA（可「加到主畫面」當 app 用），資料儲存喺裝置本機（localStorage）。

## 本機試行（可選）

如果你想喺電腦先試：

```bash
npm install
npm run dev
```

然後開瀏覽器去 http://localhost:5173

## 部署（跟 PWA 部署教學）

呢個 folder 入面已經有齊部署需要嘅檔案。跟住你份「PWA 部署教學」做：

1. 將呢個 `gymily-booking` folder **入面嘅所有檔案**上傳到 GitHub（唔係成個 folder）。
2. 用 Vercel 連接個 repo，撳 Deploy。
3. Vercel 會自動認到 Vite，build 完畀你一個網址。
4. 手機開個網址，「加到主畫面」。

> 注意：上傳時**唔好**包含 `node_modules` 同 `dist`（如果有），Vercel 會自己安裝同 build。

## 登入

- 管理員：帳號 `admin`，密碼 `admin123`
- 教練：用帳號名稱登入（預設 alex / betty / chris / diana / eric / fiona，密碼 1234）

管理員可以喺「教練」分頁新增／修改教練同帳號、改密碼。

## 資料儲存：兩種模式

這個 app 支援兩種模式，由 `src/config.js` 決定：

**A. 本機模式（預設，config 留空）**
- 資料存喺該裝置瀏覽器（localStorage）。
- 同一部機重新整頁／關開都仲喺度，但**唔同裝置之間唔會同步**。
- 適合單一裝置管理（你自己一部機落晒 booking）。

**B. 雲端同步模式（填咗 Supabase 設定）**
- 所有裝置共用**同一份即時資料**，一部改其他幾秒內更新。
- 適合多個教練各自登入、共用同一個場地表。
- 設定方法見 `SUPABASE設定.md`（免費，約 10 分鐘）。

App 右上角會顯示目前模式：「☁️ 已同步」＝雲端，「📱 本機」＝本機。

建議定期用「設定 → 匯出 Excel 備份」儲存資料。

## 更新 app

直接喺 GitHub 改檔案再 Commit，Vercel 會自動重新 deploy。

## 換名／換 icon

- App 名／主題色：改 `vite.config.js` 入面 `VitePWA` 的 `manifest`。
- Icon：換 `public/icon-192.png`、`public/icon-512.png`、`public/icon-512-maskable.png`。
