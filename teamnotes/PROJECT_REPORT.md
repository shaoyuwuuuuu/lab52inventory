# Social Media Influencer Tracker
## 一個社群策略人 × AI，8 天打造一套真正在用的 KOL 管理工具

> **核心角色：** Shao Yu（社群行銷策略）× Claude Sonnet 4.6（AI 工程師）
> **開發週期：** 2026/05/19 – 2026/05/26（8 天）
> **部署成本：** $0 / 月

---

## 一、背景：組員的問題，我來解決

我的組員是一位社群企劃，每天要追蹤數十位 KOL 和 UGC 網紅的合作進度。每一個案子有自己的節奏：

```
洽談 → 寄產品 → 催腳本 → 等初稿 → 審核 → 確認上線 → 付尾款
```

**她試過的工具都不適合：**

- **Google Sheets**：欄位多、跨分頁找資料、算日期靠手動、完全沒有提醒
- **Notion**：介面好看，但沒有自動提醒、無法跟 Sheets 串接、進階功能要付費
- **市面 KOL 平台**：定價高、功能過剩、無法客製化我們的業務邏輯

**所以我自己下來做。** 業務邏輯我最懂，我規劃 UX，AI 負責工程。

---

## 二、從零開始前，要先能「跑起來」

### 雙點啟動：啟動追蹤表.bat

開發第一天，我就意識到一個問題：用 `file://` 直接開 HTML 檔案，瀏覽器會封鎖任何外部同步請求（CORS 限制）。

解法是架一個本機伺服器，但要求組員每次都要開終端機輸入指令太不現實。

所以我讓 AI 做了一個批次檔 **`啟動追蹤表.bat`**：

```bat
@echo off
cd /d "%~dp0"
start "" "http://localhost:8765/influencer_tracker.html"
node "%~dp0server.js"
```

**雙擊即啟動**：Node.js 伺服器自動開、瀏覽器自動跳到正確頁面。
`%~dp0` 讓路徑永遠相對於檔案所在位置，資料夾改名也不會壞掉。

這是整個工具的第一個「使用者體驗決策」：**不讓組員碰任何技術細節。**

---

## 三、最核心的架構演進：從單向拉取到雙向連動

這是整個專案技術上最有價值的演進，也是花最多時間設計的部分。

### Phase 1：只能讀（Sheets → Web）

最初架構很簡單：
```
Google Sheets ──(讀取)──→ Web App（顯示）
```
用 Google Visualization API 拉資料，但很快遇到 CORS 問題，整個同步失效。

### Phase 2：解決 CORS，建立穩定的讀取通道

GAS 不支援直接跨域 fetch，我們改用 **JSONP** 的方式繞過限制：

```
Web App 動態插入 <script src="GAS_URL?action=getData&callback=fn">
         ↓
GAS 後端返回：fn({ paid: [...], ugc: [...], posts: [...] })
         ↓
瀏覽器自動執行，資料進入 App
```

同時設計了**三道保險**，確保任何情況都能拿到資料：

| 模式 | 使用情境 |
|------|----------|
| **GAS 直連**（當 App 跑在 GAS 上時） | 部署版，`google.script.run` 直接讀 |
| **JSONP URL**（本機開發版） | 貼上 GAS URL，一鍵同步所有分頁 |
| **CSV 匯入**（備用） | 網路受限、或第一次設定前的過渡期 |

同步後資料會快取到瀏覽器 `localStorage`，下次開啟不需要等待即可看到資料。

### Phase 3：Web 可以寫回 Sheets（真正的雙向）

這是最關鍵的一步。光能「看」還不夠，在 Web 上改的東西要能**直接回寫 Google Sheets**，不需要再去開 Sheets 操作。

最終 Code.gs 實作了完整的寫回 API：

```
Web 編輯時間軸日期  →  updatePaidRow()  →  Sheets 對應儲存格更新
Web 新增一筆 KOL   →  addPaidRow()     →  Sheets 新增一列
Web 新增 UGC       →  addUGCRow()      →  Sheets 新增一列
Web 更新貼文狀態   →  saveAllPosts()   →  Posts 分頁整批寫回
Web 記錄外發 Log   →  saveAllOutreachLog() → 自動建立分頁（若不存在）再寫入
Web 儲存範本       →  saveTemplate()   →  Outreach 分頁更新
Web 刪除範本       →  deleteTemplate() →  Sheets 對應列刪除
```

**值得特別說的一個細節：** `saveAllOutreachLog()` 在寫入前會先偵測 Outreach_Log 這個分頁存不存在，若不存在就**自動建立，並加上正確的欄位 header**。這讓第一次使用的人不需要手動在 Sheets 建分頁。

### 現在的架構

```
Web App（編輯、查看）
      ↕ 雙向即時同步
Google Apps Script（後端 API）
      ↕
Google Sheets（5 個分頁：Paid_Collabs / UGC_Free / Posts / Outreach / Outreach_Log）
```

多人同時使用也沒問題——所有人讀寫同一份 Sheets，不需要額外的帳號系統。

---

## 四、我規劃的 UI/UX 細節

這個部分是整個專案我最想分享的。每一個細節都來自我對實際工作流程的理解，不是「因為技術上可以做」，而是「因為使用者真的需要」。

---

### 1. 自動待辦引擎（Today — Progress Check）

每天開啟，系統自動比對所有 KOL 資料，依照業務規則生成當天的待辦：

| 提醒類型 | 觸發條件 | 設計原因 |
|----------|----------|----------|
| 💰 今日需付首款 | 付款日 ≤ 今天 | 付款有時間壓力，不能漏 |
| 📝 3天內催腳本 | 狀態=Product Sent，腳本截止日 ≤ 3天 | 提早催才有時間來回 |
| 🎬 3天內催初稿 | 狀態=Shooting，提交日 ≤ 3天 | 同上 |
| 📢 今日上線檢查 | 狀態=Submitted，live_date ≤ 今天 | 確認貼文真的有上 |
| 💵 今日需付尾款 | 狀態=Published | 上線後才付 |
| 🎁 7天內預警催片 | UGC，ship_date + 30 天 ≤ 7天 | UGC 交片通常設 30 天，提前預警 |
| 📬 Follow-up 提醒 | Outreach Log 設了追蹤日期 | 外發後沒回音要跟進 |

逾期超過 14 天的項目**自動隱藏**，只顯示計數——避免舊案件淹沒今天真正要處理的事。

---

### 2. 今日已處理（當日 Push Back）

待辦清單每一列都有勾選框，勾了之後該項變灰淡，視覺上告訴你「這件事今天已經處理了」。

**關鍵設計：勾選狀態只存在瀏覽器本機，不寫回 Sheets。**

這是刻意的。「我今天已經發了催腳本的訊息」不需要記在資料庫，因為腳本狀態本身沒有變；隔天 KOL 還沒交稿，提醒會再出現。

每個人也有自己的「今日已處理」記錄，不會影響另一個使用者的視圖。

---

### 3. 待辦點擊展開 Sidebar：從看到提醒到發出聯絡，不離開頁面

點擊任何一條待辦，右側滑出 Panel 顯示：

```
▸ 觸發規則 + 天數狀態（紅/橘/綠）
▸ Email  →  直接點 mailto: 開信箱
▸ 💬 IG DM  →  一鍵連結 ig.me/m/username（自動從 handle 組成）
▸ 完整 Timeline（付款日 / 腳本日 / 初稿日 / 上線日…）
▸ 備注
▸ → View full record（跳至 Paid / UGC 分頁）
```

**我決定了這個要是唯讀的**——Sidebar 的用途是「快速拿到聯絡方式」，不是編輯資料。需要改資料的話，點「View full record」去對應分頁操作。這讓兩個功能的責任不混淆。

---

### 4. Paid Collabs Sidebar：直接在 Web 改 Timeline，即時同步到 Sheets

Paid 分頁的 Sidebar 是可編輯的，這裡設計了一個完整的 Timeline 面板：

- 每個日期節點（首款 / 腳本 / 初稿 / Revision / 尾款）可以直接在 Sidebar 點擊修改
- 日期點的顏色即時反映狀態：灰=未設定、橘=未到期、綠=已過
- 修改完按「Save」→ `updatePaidRow()` → 直接寫回 Google Sheets 對應儲存格
- 備注區也可以直接在 Sidebar 編輯並儲存

不需要打開 Sheets，不需要找到正確的列和欄。

---

### 5. 自製日曆 / 月份選擇器

原生的 `<input type="month">` 和 `<input type="date">` 在不同瀏覽器長得不一樣，而且通常很醜。

我要求 AI 做自製的浮動選擇器：

- **日期選擇器**：年份下拉 + 月份格 + 日期格，點即選取，跟 App 整體視覺一致
- **月份選擇器**（Overview 篩選用）：左右切換年份 + 12 格月份

這不是美觀問題，是**讓使用者覺得這是一個完整的產品，不是拼湊的工具**。

---

### 6. 狀態點擊循環切換（Posts Calendar）

在 Posts 日曆的每個貼文格，有一個小色點代表目前狀態。點擊這個色點，狀態就循環切換：

```
Idea → Drafting → Scheduled → Ready → Published → Idea…
```

不需要打開 Modal，不需要找下拉選單。**一個點，一次點擊，狀態就更新並回寫 Sheets。**

這個互動設計來自我自己在看日曆時的需求：「我只是想快速標記這篇已排程了，不想做其他事。」

---

### 7. Outreach Templates 變數高亮

DM 和 Email 外發範本裡有佔位符，如 `[Name]`、`[specific video]`。

在 App 裡預覽範本時，這些佔位符會**用橘色高亮顯示**，提醒發送者記得替換。

這個細節讓「複製貼上」這個動作更安全——不會漏掉沒改到的 `[Name]`。

---

### 8. Outreach Log → Paid 一鍵轉換

1. 跳到 Paid Collabs 分頁
2. 把這個 handle 預填到搜尋框

省去跨頁複製貼上的步驟。

---

### 9. 美國假日標注在 Posts 日曆

Posts 日曆上會自動標出美國的法定假日和重要節慶（情人節、萬聖節、感恩節…），以橘色標籤顯示。

這個功能是我在排貼文行程時直接提出的需求——美國假日影響受眾的活躍度，在排程時需要一眼看到。**AI 自己實作了所有浮動假日的算法**（第 n 個週幾這種），不是硬寫日期。

---

### 10. Brand Quick Reference

Outreach 分頁頂部有一排快速連結，可以存品牌的 Amazon 連結、過去合作的 IG 帖子等。

**設計巧思：這些連結不是存在瀏覽器書籤，而是存在 Google Sheets 的 Outreach 分頁裡（用 `type=BrandRef` 區分）。** 新增 / 刪除會即時同步到 Sheets，重新 Sync 或換瀏覽器開啟，連結還在。

---

### 11. Toast 通知取代原生 alert/confirm

所有的操作回饋（儲存成功、同步失敗、刪除確認）都用自製的 Toast 和 Confirm Dialog，不用瀏覽器原生的 `alert()`。

這讓整個 App 的互動**視覺一致**，也讓錯誤訊息更清楚（原生 alert 只顯示文字，無法控制樣式）。錯誤的 Toast 是紅底紅字，成功是綠底，2 秒自動消失。

---

### 12. Export CSV

Paid Collabs 和 UGC Free 都有 Export CSV 按鈕，可以把目前篩選後的結果匯出。

這讓 App 不是一個封閉的系統——資料可以隨時拿出來，對帳、給主管看、轉交給其他工具都方便。

---

### 13. Quick Lookup（跨分頁交叉查詢）

Outreach 分頁有一個搜尋框，輸入 handle，即時顯示這個 KOL 在 Paid Collabs 和 UGC Free 裡的所有記錄——合作時間、狀態、費用、備注。

這個功能解決了一個真實情境：**「這個人之前跟我們合作過嗎？費用是多少？」** 不需要跑去另一個分頁搜尋，在 Outreach 頁面就能查到。

---

## 五、雙向同步的完整架構（最終版）

```
┌─────────────────────────────────────────┐
│              Web App                    │
│  ┌──────────┐  ┌──────────┐            │
│  │ 讀取顯示 │  │ 編輯寫回 │            │
│  └────┬─────┘  └────┬─────┘            │
└───────┼─────────────┼───────────────────┘
        │             │
        ▼             ▼
┌─────────────────────────────────────────┐
│         Google Apps Script              │
│  getPaidData / getUGCData / ...  (讀)   │
│  updatePaidRow / addPaidRow / ...  (寫) │
│  saveAllPosts / saveAllOutreachLog      │
│  saveTemplate / deleteTemplate          │
└─────────────────┬───────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────┐
│           Google Sheets                 │
│  Paid_Collabs / UGC_Free / Posts /      │
│  Outreach / Outreach_Log                │
└─────────────────────────────────────────┘
```

**任何有連結的人都能開啟 App，讀寫同一份 Sheets。多人協作，零額外設定。**

---

## 六、數字

| 指標 | 數字 |
|------|------|
| 開發天數 | 8 天 |
| Git Commits | 9 次 |
| 最終程式碼行數 | ~4,500 行（HTML + Apps Script） |
| Apps Script 寫回函數數量 | 8 個 |
| 對比初始版本淨新增 | +1,866 行 |
| 月費 | $0 |
| 依賴的外部付費服務 | 0 個 |

---

## 七、這個協作模式最值得說的

**業務邏輯的決策全部在我這裡。**

- 14 天 sunset 規則是我決定的（幾天才合理？）
- 3 天催腳本是我決定的（幾天前催才來得及？）
- Sidebar 要唯讀而不是可編輯，是我決定的
- IG DM 連結要加，是我決定的
- 美國假日要標出來，是我決定的

AI 的角色是：接到這些需求，**判斷最好的實作方式，然後執行。**

我沒有要求加安全防護，但 AI 自己加了 XSS 過濾和 URL 驗證。
我沒有要求處理 CORS 問題，AI 找到了 JSONP 這個方案並解釋給我聽。
我說「日期選擇器太醜了」，AI 重新做了一個跟整體設計一致的版本。

**不是 AI 代替了設計思考，是 AI 讓一個沒有工程背景的人能夠執行自己的設計思考。**

---

*Built with Claude Sonnet 4.6 × Google Apps Script × Google Sheets*
*2026/05/19 – 2026/05/26 · Lab52 Social Media Team*
