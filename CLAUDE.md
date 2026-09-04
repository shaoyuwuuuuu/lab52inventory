# Lab52 Inventory 庫存管理系統

Google Apps Script + Google Sheets 的庫存管理 Web App，支援 Amazon SP-API FBA 庫存同步。

## 架構
```
Code.gs         ← GAS 後端（主程式 + SP-API 同步 + 診斷）
index.html      ← 前端 Web App 介面（單檔，含 CSS/JS）
appsscript.json ← GAS 部署設定與 OAuth scope
deploy.bat      ← clasp push + deploy（注意結尾有 pause，自動化時請直接下 clasp 指令）
```

## Google Sheets 分頁
- `庫存總覽` — 產品主檔（只讀），公式自動計算庫存
- `Movement` — 海外倉出入庫紀錄（Date / Name / SKU / EAN / ASIN / exp_date / Boxes / location / note）
  - `location` 預設 `AMZLGS`，可能出現 `TW`；**計算海外倉庫存時必須排除 `TW` 的列**，否則會與台灣倉重複計算
- `TW_Movement` — 台灣倉庫存
- `TW_Backup` — 台灣倉匯入前的自動備份（隱藏分頁，供「還原上一動」使用）
- `Transits` — 在途貨物追蹤
- `Billing` — 帳單紀錄
- `FBA庫存` — SP-API 同步結果（16 欄，見下）
- `Shared` — 跨裝置共用的小資料（隱藏分頁，`key / value / updated_at`）
  - `manual_transit_<EAN>` 手動在途、`dash_note_fba_<EAN>` 與 `dash_note_prod_<EAN>` 儀表板備註
  - 這些原本只存在瀏覽器 localStorage，換裝置就看不到；而表格內的備註
    （`saveNoteInline`）卻是寫進分頁的 —— 外觀相同行為不同，已統一
  - 一個 key 一列，改不同產品不會互相覆蓋；但同一格是「最後寫入者勝出」，無衝突偵測
  - `setSharedBulk` 只補雲端沒有的 key、絕不覆蓋，供 localStorage 舊值一次性搬遷

## FBA 欄位口徑（已逐項對照賣家後台驗證，勿憑文件推論修改）

SP-API 的欄位分組**不等於**賣家後台的顯示口徑。以下三項都是實際對帳後才確定的：

| 分頁欄位 | 算法 | 對應後台 |
|---|---|---|
| 可出貨數量 | `fulfillableQuantity` | 現貨 → 可售 |
| 入庫中 | `inboundWorking + inboundShipped` | 入庫 |
| 預留數量 | `totalReserved − pendingTransshipment` | 預留 |

**兩個踩過的坑：**

1. **入庫中不能加 `inboundReceivingQuantity`。** API 把 working / shipped / receiving 都歸在
   `InboundQuantityBreakdown` 底下，但後台的「入庫」只算前兩段——receiving 是貨已到倉正在點收，
   不算在途。實例：B0DBTNJW4G working=0 shipped=0 receiving=34，後台入庫顯示 0。

2. **預留要扣掉 `pendingTransshipmentQuantity`。** 後台把「運營中心轉運」歸在**現貨**底下
   （現貨 765 = 可售 764 + 轉運 1），不算預留。實例：API total=60（49+1+10），後台預留顯示 59。

後 5 欄 `不可售 / 調查中 / 買家訂單 / FC處理中 / 轉運` 供 hover 明細面板使用。
`fbaTotalOf_()` 刻意不寫死巢狀欄位名（是數字就用、是物件就找 `total*`、再不然加總數值欄位）。

**對帳時務必同時重新整理兩邊。** Amazon 的數字變動很快，曾因看的時間點不同而誤判成算法錯誤。
`diagFbaSync()` 的 `[4 對照]` 已排除「分頁過期」這個變數，剩下唯一變數就是後台的時間點。

## 前端計算口徑
- **可用庫存 = 可出貨 + 入庫中**（不含預留，預留已被訂單佔用，算進來會低估補貨需求）
  前端 `fbaAvailQty()` 與後端 `fbaLevel_()` 共用同一套門檻：
  有日均銷量看週數（<4 週緊急 / <10 週注意），沒有則降級用件數 50
- **補貨目標 60 天**、5 箱倍數；來源優先序：海外倉 → 台灣倉
- 日均銷量 = Sales API `getOrderMetrics` 過去 7×24 小時的 `unitCount` ÷ 7
- 已知限制：日均銷量為 0 的滯銷品不計入「充足／注意／緊急」三格，加總會小於總計（刻意保留）

## 排程 Trigger
| 函式 | 頻率 | 備註 |
|---|---|---|
| `syncFbaInventory` | 每小時 | `setupFbaTrigger()` 建立 |
| `syncFbaSalesVelocity` | 每天 07:00 | `setupSalesTrigger()` 建立；順便跑看門狗 |
| `sendRestockAlert` | 每天 09:00 | `setupRestockAlertTrigger()` 建立 |

**這三個 trigger 都要各自執行對應的 setup 函式才會存在。**
`sendRestockAlert` 的 trigger 曾長期沒被建立，導致補貨警示從上線起一次都沒發過，
而當時的診斷只檢查兩支同步函式、驗不出來。`diagFbaSync()` 現在三支都會檢查。

## 監控與診斷

**同步失敗一定要會叫。** 2026-08-18 曾停擺 16 天無人察覺，原因是 `syncFbaInventory` 的 catch
把錯誤吞成 `Logger.log` + `return`——從 trigger 執行時回傳值沒人接、Logger 沒人看，
對 Google 而言每次都「執行成功」，既不會停用 trigger 也不會寄失敗摘要信。

現在的防線：
- `notifySyncFailure_()` — 失敗主動寄信到 `RESTOCK_ALERT_EMAIL`，6 小時節流
- `checkInventoryFreshness_()` — 每日看門狗，庫存超過 6 小時沒更新就通知
  （補上「trigger 被刪掉就完全沒聲音」這個死角）
- `testSyncAlert()` / `testSyncAlertTo()` — **驗證通知管道真的通，改過收件人或權限後請執行一次**
  - `testSyncAlertTo()` 寄到指令碼屬性 `TEST_ALERT_TO` 指定的信箱，可測任意位址
  - 兩支都**刻意不 catch**：寄不出去就要讓例外浮到執行紀錄。
    第一版走 `notifySyncFailure_`，結果它把自己的例外也吞掉、寄失敗仍回傳 `ok:true`，
    等於重蹈同步無聲失敗的覆轍
  - 通知管道已於 2026-09-04 實測寄達 `annicewu@toothfilm.com`
  - 不要為了診斷去加 `userinfo.email` 之類的 scope：新增 `oauthScopes` 會讓
    網頁應用程式在擁有者重新授權之前失效
- `diagFbaSync()` — 逐層檢查 憑證 → Trigger → LWA token → SP-API → 分頁，
  並列印 API 與分頁的值供對照。只印憑證長度，不印內容

## SP-API 憑證

存在 GAS「專案設定 → 指令碼屬性」：
`AMAZON_REFRESH_TOKEN` / `AMAZON_CLIENT_ID` / `AMAZON_CLIENT_SECRET` / `AMAZON_MARKETPLACE_ID`

**client secret 效期只有 180 天**，過期時 LWA 回 HTTP 401 `invalid_client`，兩個同步一起停擺。
2026-08-18 那次就是這個原因（憑證 2026-08-11 重發），**下次到期約在 2027-02**。

修法：**只需更新 `AMAZON_CLIENT_SECRET` 一個**。依 Amazon 官方文件，refresh token 綁定的是
client_id，輪替 secret 不需要重新產生 refresh token，也不需要請賣家重新授權。
（另注意 refresh token 效期 365 天、賣家授權需每年更新，那是不同的到期。）

- client ID 格式固定為 `amzn1.application-oa2-client.` 開頭，`diagFbaSync()` 會檢查有沒有跟 secret 貼反
- `getSpApiToken_()` 讀屬性時一律 `trim()`，避免從文件複製時夾帶的空白造成 `invalid_client`
- 憑證一律由人工填入 GAS 介面，**不要進到 git 或對話中**（`.gitignore` 已擋 `*.docx` / `LAB52 API*`）

## 部署

```bash
clasp push --force    # 更新編輯器程式碼；trigger 與編輯器執行的函式立即生效
clasp deploy --deploymentId <ID> --description "..."   # 網頁應用程式才需要
```

**動到 `index.html` 一定要 deploy**，因為網頁應用程式是綁版本的，只 push 不會生效。
純後端改動（同步邏輯、診斷、通知）只要 push 即可。

## 開發慣例
- 前端無建置流程，`node --check` 逐一檢查 `index.html` 的三個 inline `<script>` 區塊即可
- 欄位新增一律加在 `FBA_HDR_` 尾端（保留日均銷量的邏輯是按欄位**名稱**查找的）
- `FBA_COLS` 增減欄位會使 localStorage 的欄位順序失效（`validOrder` 比對長度），
  使用者自訂順序會重設一次，這是必要的，否則新欄位不會顯示
- 覆蓋型操作（如 `replaceTaiwanInventory`）務必先備份且前後端都要擋空清單
