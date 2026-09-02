# Lab52 Inventory 庫存管理系統

Google Apps Script + Google Sheets 的庫存管理 Web App，支援 FBA 庫存同步。

## 架構
```
Code.gs       ← GAS 後端（主程式 + SP-API FBA 同步）
index.html    ← 前端 Web App 介面
appsscript.json ← GAS 部署設定
```

## Google Sheets 結構（4 個分頁）
- `庫存總覽` — 產品主檔（只讀），公式自動計算庫存
- `Movement` — 出入庫紀錄（Date / Name / SKU / EAN / ASIN / exp_date / Boxes / location / note）
- `Transits` — 在途貨物追蹤（由 `setupSheets()` 自動建立）
- `Billing` — 帳單紀錄（由 `setupSheets()` 自動建立）

## GAS 端點（action 參數）
- `getData` — 回傳全部庫存資料
- `syncFba` — 觸發 SP-API FBA 庫存同步，回傳最新 FBA 庫存

## 首次設定步驟
1. Google Sheets → 擴充功能 → Apps Script
2. 貼入 `Code.gs`，另建 `index.html` 貼入前端
3. 執行 `setupSheets()` 一次（建立 Transits / Billing 分頁）
4. 部署 → 網頁應用程式（執行身分：我 / 存取：知道連結的所有人）

## 注意
- FBA 同步需要 Amazon SP-API 憑證（設定在 GAS 的指令碼屬性內）
- `appsscript.json` 已申請 Gmail 傳送權限（用於異常通知）
