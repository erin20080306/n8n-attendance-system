# 多倉出勤 / 請假查詢與報表系統（n8n Workflow）

## 專案概述

本專案提供一套可直接匯入 n8n 的完整 workflow JSON，用於建立「多倉出勤 / 請假查詢與報表系統」。  
使用者透過 Webhook 發送自然語言查詢指令，系統自動解析意圖、日期、倉別，讀取對應 Google Sheets 資料，產出查詢結果或統計報表。

---

## 支援倉別

| 倉別   | Spreadsheet ID                                  |
|--------|------------------------------------------------|
| TAO1   | `1_bhGQdx0YH7lsqPFEq5___6_Nwq_gbelJmIHv0bmaIE` |
| TAO3   | `1cffI2jIVZA1uSiAyaLLXXgPzDByhy87xznaN85O7wEE`  |
| TAO4   | `1tVxQbV0298fn2OXWAF0UqZa7FLbypsatciatxs4YVTU`   |
| TAO5   | `1jzVXC6gt36hJtlUHoxtTzZLMNj4EtTsd4k8eNB1bdiA`  |
| TAO6   | `1wwPLSLjl2abfM_OMdTNI9PoiPKo3waCV_y0wmx2DxAE`  |
| TAO7   | `16nGCqRO8DYDm0PbXFbdt-fiEFZCXxXjlOWjKU67p4LY`  |
| TAO10  | `1y0w49xdFlHvcVtgtG8fq6zdrF26y8j7HMFh5ujzUyR4`   |

---

## 節點用途說明

### 1. Webhook 接收查詢
- **類型**：Webhook（POST）
- **路徑**：`/attendance-query`
- **用途**：接收使用者 POST 請求，body 中需包含 `query` 欄位
- **回傳模式**：Response Node（由後方節點控制回傳）

### 2. 指令解析
- **類型**：Code Node
- **用途**：規則式解析自然語言 query
  - 解析日期（支援 `4/2`、`04/02`、`2026/4/2`、`2026-04-02`）
  - 解析意圖（請假狀況、病假統計、請假統計、出勤明細、出勤統計）
  - 解析倉別（TAO1~TAO10，未指定則查全部）
  - 解析假別（病假、事假、特休等）
- **輸出**：標準化 JSON（intent, date, warehouse, leaveType, needReport）

### 3. 解析成功？
- **類型**：IF Node
- **用途**：判斷解析是否成功
  - 成功 → 繼續讀取 Google Sheets
  - 失敗 → 回傳錯誤訊息

### 4. 組裝讀取參數
- **類型**：Code Node
- **用途**：根據解析結果，為每個倉別產生對應的讀取參數
  - 對應 spreadsheetId
  - 根據意圖決定讀取假況分頁或工時分頁
  - 多倉查詢時輸出多筆 item，讓後續節點逐一處理

### 5. 讀取 Google Sheets
- **類型**：Code Node（透過 Google Sheets API v4）
- **用途**：
  - 先取得該 Sheet 所有分頁名稱
  - 根據關鍵字找到目標分頁（班表/出勤記錄/假況 或 出勤時數/出勤時間/工時）
  - 若找不到匹配分頁，退回讀取第一個分頁並附上 warning
  - 讀取分頁全部資料
- **API 認證**：使用 `googleSheetsOAuth2Api` credential

### 6. 資料標準化
- **類型**：Code Node
- **用途**：
  - 將各倉原始資料映射為統一欄位格式
  - 工時分頁：根據表頭映射 date/department/shift/name/workHours 等
  - 假況分頁：H 欄以後為日期欄，展平為每人每日一列
  - 所有假況值進行標準化映射（病假/事假/特休/曠職/出勤等）
  - 所有 mapping 規則集中在 CONFIG 區塊

### 7. 合併所有倉別資料
- **類型**：Code Node
- **用途**：將各倉標準化後的資料合併為一份，收集 warnings 和 errors

### 8. 查詢與統計
- **類型**：Code Node
- **用途**：根據意圖進行處理
  - `leave_status`：某日請假狀況（列出請假人員）
  - `sick_leave_stats`：某日病假統計
  - `leave_stats`：某日所有請假類型統計
  - `attendance_detail`：出勤人員明細（班別/部門/姓名/工時）
  - `attendance_stats`：出勤統計（含總工時/加班時數）
- **輸出**：包含 summaryText、totals、rows 的標準 JSON
- **預留**：emailReady、emailSubject、emailBody 欄位

### 9. 回傳成功結果
- **類型**：Respond to Webhook
- **用途**：回傳 HTTP 200 + JSON 結果

### 10. 回傳錯誤結果
- **類型**：Respond to Webhook
- **用途**：回傳 HTTP 400 + JSON 錯誤訊息

---

## Google Sheets Credentials 設定說明（Service Account 方式）

### 步驟

1. **建立 Google Cloud Project**
   - 前往 [Google Cloud Console](https://console.cloud.google.com/)
   - 建立新專案或選擇已有專案

2. **啟用 Google Sheets API**
   - 在 API Library 中搜尋「Google Sheets API」並啟用

3. **建立 Service Account**
   - 進入「IAM & Admin → Service Accounts」
   - 點擊「Create Service Account」
   - 輸入名稱（例如 `n8n-sheets-reader`）→ Create
   - 角色可跳過（不需要專案層級角色）→ Done
   - 點擊剛建立的 Service Account → Keys 分頁
   - 「Add Key → Create new key → JSON」→ 下載 JSON 金鑰檔案
   - **記下 Service Account 的 Email**（格式如 `n8n-sheets-reader@your-project.iam.gserviceaccount.com`）

4. **共用 Google Sheets 給 Service Account**
   - 打開每個倉的 Google Sheet（共 7 個）
   - 點擊右上角「共用」
   - 將 Service Account Email 加入為「檢視者」
   - 7 個 Sheet 都要執行此步驟

5. **在 n8n 中設定 Credential**
   - 進入 n8n → Settings → Credentials → Add Credential
   - 搜尋並選擇「**Google API (Service Account)**」
   - **Service Account Email**：填入 Service Account 的 Email
   - **Private Key**：打開下載的 JSON 金鑰檔案，複製 `private_key` 欄位的完整值（含 `-----BEGIN PRIVATE KEY-----` 和 `-----END PRIVATE KEY-----`）
   - **Scopes**：填入 `https://www.googleapis.com/auth/spreadsheets.readonly`
   - 點擊 Save
   - Credential 名稱會自動對應為 `googleApi`（匯入後可在節點中重新綁定）

### 優點
- **不需要手動登入 Google 帳號**，適合伺服器自動化
- **Token 不會過期**，不需要定期重新授權
- **安全性高**，Service Account 只有被共用的 Sheet 才能存取

---

## Webhook 測試範例

Webhook URL 格式（以本機為例）：
```
POST http://localhost:5678/webhook/attendance-query
Content-Type: application/json
```

### 測試指令：

```bash
# 1. 4/2 請假狀況
curl -X POST http://localhost:5678/webhook/attendance-query \
  -H "Content-Type: application/json" \
  -d '{"query": "4/2請假狀況"}'

# 2. 4/2 病假統計
curl -X POST http://localhost:5678/webhook/attendance-query \
  -H "Content-Type: application/json" \
  -d '{"query": "4/2病假統計"}'

# 3. 4/2 請假統計
curl -X POST http://localhost:5678/webhook/attendance-query \
  -H "Content-Type: application/json" \
  -d '{"query": "4/2請假統計"}'

# 4. 4/2 出勤人員明細
curl -X POST http://localhost:5678/webhook/attendance-query \
  -H "Content-Type: application/json" \
  -d '{"query": "4/2有出勤人員班別/部門/姓名/工時"}'

# 5. 查詢 TAO1 4/2 病假
curl -X POST http://localhost:5678/webhook/attendance-query \
  -H "Content-Type: application/json" \
  -d '{"query": "查詢TAO1 4/2病假"}'

# 6. 查詢 TAO3 4/2 出勤報表
curl -X POST http://localhost:5678/webhook/attendance-query \
  -H "Content-Type: application/json" \
  -d '{"query": "查詢TAO3 4/2出勤報表"}'

# 7. 查詢全部倉別 4/2 請假統計
curl -X POST http://localhost:5678/webhook/attendance-query \
  -H "Content-Type: application/json" \
  -d '{"query": "查詢全部倉別 4/2 請假統計"}'

# 8. 查詢 TAO10 2026/4/2 出勤人員明細
curl -X POST http://localhost:5678/webhook/attendance-query \
  -H "Content-Type: application/json" \
  -d '{"query": "查詢TAO10 2026/4/2 出勤人員明細"}'
```

---

## 後續擴充說明

### 1. 如何擴充成 Email 報表

Workflow 已在「查詢與統計」節點預留以下欄位：
- `emailReady`: true
- `emailSubject`: 自動產生的郵件主旨
- `emailBody`: summaryText 內容

**擴充步驟：**
1. 在「回傳成功結果」節點之前，新增 IF 節點判斷 `needReport === true`
2. 新增 Send Email 節點（或 Gmail 節點）
3. 將 `emailSubject` 接到主旨，`emailBody` + `rows` 格式化後接到內文
4. 若需 Excel 附件，可新增 Spreadsheet File 節點將 rows 轉為 xlsx

### 2. 若不同倉分頁名稱不同，要修改哪個 config

修改位置集中在以下兩個 Code Node 的 `CONFIG` 區塊：

- **「組裝讀取參數」節點** → `CONFIG.sheetKeywords`
  ```javascript
  sheetKeywords: {
    leave: ['班表', '出勤記錄', '假況'],  // ← 新增或修改假況分頁關鍵字
    work:  ['出勤時數', '出勤時間', '工時'], // ← 新增或修改工時分頁關鍵字
  },
  ```

- **「讀取 Google Sheets」節點** → `CONFIG.sheetKeywords`（同上，兩處保持一致）

### 3. 若資料欄位再增加，要修改哪個 mapping

修改位置在 **「資料標準化」節點** 的 `CONFIG` 區塊：

- **新增欄位映射**：在 `CONFIG.workFieldMapping` 中新增 key-value
  ```javascript
  workFieldMapping: {
    // 新增範例：
    employeeId: ['員工編號', 'Employee ID', 'EmpID'],
    // ... 其他現有欄位
  },
  ```

- **新增假況值映射**：在 `CONFIG.leaveStatusMapping` 中新增
  ```javascript
  leaveStatusMapping: {
    // 新增範例：
    '育嬰假': '育嬰假',
    // ... 其他現有值
  },
  ```

- **調整假況分頁日期起始欄**：修改 `CONFIG.leaveDateStartCol`（目前為 7，即 H 欄）

### 4. 若要新增倉別

修改兩個位置：

- **「指令解析」節點** → `CONFIG.warehouses` 陣列新增倉別代號
- **「組裝讀取參數」節點** → `CONFIG.spreadsheetMap` 新增 spreadsheetId

---

## 匯入方式

1. 開啟 n8n
2. 點擊左側選單 → Workflows
3. 點擊右上角 ⋮ → Import from File
4. 選擇本專案的 `workflow.json`
5. 匯入後設定 Google API (Service Account) Credential，並將 7 個 Sheet 共用給 Service Account
6. 啟用 Workflow
7. 使用上方 curl 範例測試

---

## 技術架構

```
Webhook POST
    ↓
指令解析（Code Node - 規則式）
    ↓
解析成功？（IF Node）
    ├─ 是 → 組裝讀取參數
    │        ↓
    │     讀取 Google Sheets（API v4 + Service Account）
    │        ↓
    │     資料標準化（統一欄位格式）
    │        ↓
    │     合併所有倉別資料
    │        ↓
    │     查詢與統計
    │        ↓
    │     回傳成功結果（HTTP 200 JSON）
    │
    └─ 否 → 回傳錯誤結果（HTTP 400 JSON）
```
