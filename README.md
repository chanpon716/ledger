# 記帳 · Instant Photo

Instant Photo Macau 的手機記帳 PWA。表單照 Airtable「Instant Photo 財務 2026」複刻，
資料直接寫進自己的 Google Sheet，可「加入主畫面」當 App 用。

- **線上版**：https://a15d9fee8248a8018.sg2.agentos-app.run
- **資料存放**：Google Sheets（`收入` / `支出` 兩張工作表）
- **後端**：Google Apps Script Web App（`apps-script/Code.gs`，v1.3.0）

## 檔案

| 檔案 | 用途 |
|---|---|
| `index.html` | 整個 App（HTML + CSS + JS 單檔，含內嵌的 Apps Script 原始碼） |
| `sw.js` | Service Worker，離線快取 |
| `manifest.webmanifest` | PWA 設定（名稱、圖示、全螢幕） |
| `icons/` | 主畫面圖示 |
| `apps-script/Code.gs` | Google Apps Script 後端 |
| `使用說明.md` | 完整操作說明（部署、匯入、排錯） |

## 部署到 GitHub Pages

1. 把這個 repo 的 `main` 分支設為 Pages 來源（Settings → Pages → Branch: `main` / root）
2. 網址會是 `https://<你的帳號>.github.io/<repo>/`
3. 手機開這個網址 → 分享 → 加入主畫面

不需要 build，純靜態檔案。

## 後端設定

1. 開 [script.google.com](https://script.google.com) → 新增專案
2. 貼上 `apps-script/Code.gs` 全部內容
3. 部署 → 新增部署作業 → 網頁應用程式
4. 執行身分「我」、誰可以存取「任何人」→ 部署並授權
5. 複製結尾 `/exec` 的網址，貼到 App 設定頁 → 儲存網址 → 測試連線

看到 `腳本 v1.3.0 · 支援編輯與刪除` 就代表版本正確。

> 改過程式後要走「部署 → 管理部署作業 → 鉛筆 → 版本選新版本 → 部署」，
> 網址不變。只按「新增部署作業」會多一條新網址，舊網址仍跑舊程式。

## 資料格式

每筆記錄都有 `id`，寫入時靠它對應 Sheet 裡的列：

- 新增：`POST` 不帶 `mode`，ID 已存在就跳過（避免重試寫出重複列）
- 編輯：`mode: "update"`，按 ID 找到原列覆蓋
- 刪除：`mode: "delete"`，按 ID 找到原列刪除
- 載入：`GET ?action=list`，回傳 `{ok, income[], expense[]}`

本機資料存在 `localStorage`（`ledger.records.v2` / `ledger.settings.v2`），
沒網路時照樣能記，回到線上自動補寫。

## 注意

- 編輯、刪除、載入歷史都需要後端 v1.3.0 以上；舊版腳本會靜默失敗
- App 只是輸入介面，**資料以 Google Sheet 為準**；換裝置按「載入歷史」即可恢復
- 建議定期在設定頁匯出 CSV 備份
