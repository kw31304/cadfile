# Outlook 附件自動分類歸檔工具 — 系統規格書 (`SPEC.md`)

> 本文件整理自原始需求，已將設定規則檔分離至 [`rules.md`](./rules.md)，並補齊原本中斷的「附件歸檔命名、防重複下載、日誌與 Windows 自動排程」規格。

---

## 0. 檔案結構規劃

程式、規則檔、狀態庫與日誌集中存放於 `d:\圖檔\CAD\outlook_archiver\`，附件則依分類直接輸出至根目錄 `d:\圖檔\CAD\`：

```text
d:\圖檔\CAD\
├── outlook_archiver\               # 程式與設定專用資料夾
│   ├── SPEC.md                     # 本規格文件
│   ├── rules.md                    # Markdown 分類規則設定檔（使用者可直接編輯）
│   ├── outlook_archiver.py         # 主程式（讀取規則、連線 Outlook、分類下載附件）
│   ├── outlook_archiver_gui.py     # PySide6 視覺操作介面
│   ├── run_archiver.bat            # Windows 工作排程器呼叫用啟動腳本（設定 UTF-8 與工作目錄）
│   ├── run_gui.bat                 # 啟動 PySide6 視覺介面腳本
│   ├── state.db                    # 防重複下載紀錄檔（SQLite，自動產生）
│   └── logs\
│       └── archiver.log            # 執行日誌（自動輪替）
├── 合宏展\                         # 分類附件輸出資料夾（依規則動態建立）
│   └── 2026-11-26_AM0938\          # 依收件日期時間建立的子資料夾
├── 單井\                           # 分類附件輸出資料夾（依規則動態建立）
└── 其他未分類\                     # 未命中規則之附件輸出資料夾
```

---

## 1. Outlook 連線與自動啟動防呆

- **執行環境**：Python 3.10+ 與 `pywin32`（`win32com.client`），連接本機桌面版 **Classic Outlook**（`OUTLOOK.EXE`）。
  - 注意：不支援「新版 Outlook for Windows（`olk.exe`）」，因新版不提供 COM 自動化介面。
- **自動啟動防呆**：
  - 執行時先檢查本機是否已存在 `OUTLOOK.EXE` 行程，或嘗試連接執行中的 `Outlook.Application`。
  - 若尚未開啟，自動啟動 `OUTLOOK.EXE`，並採用**輪詢等待機制**（例如每 2 秒重試一次 `GetActiveObject("Outlook.Application")` 與 `GetNamespace("MAPI")`，最長等待 `OUTLOOK_STARTUP_TIMEOUT = 45` 秒；至少等待初始化完成後再繼續），避免載入大型信箱時固定等待 10 秒仍逾時失敗。
- **目標信箱與資料夾（頂部可調參數）**：
  - `TARGET_ACCOUNT_EMAIL`：預設為 `None`（使用預設帳號）；若指定特定 Email 或 Store 名稱（例如未來的 Yahoo 信箱），則搜尋 `namespace.Stores` 對應項目。
  - `TARGET_FOLDER_NAME`：預設為收件匣，實作上優先使用 MAPI 常數 `olFolderInbox = 6`（`store.GetDefaultFolder(6)`）以避免中英文介面（`收件匣` / `Inbox`）差異；若需指定自訂子資料夾再依名稱查找。
  - `DAYS_BACK`：掃描天數上限（例如預設 `14` 天，設為 `None` 則掃描全部），透過 `Items.Restrict` 過濾收件時間以確保大量信件下的執行效能。
- **郵件類型過濾**：
  - 僅處理一般郵件物件（`item.Class == 43`，即 `olMail`），自動跳過行事曆邀請、退信回條等非郵件項目。
- **Exchange 內部寄件人解析**：
  - 當 `mail.SenderEmailType == "EX"` 時，透過 `mail.Sender.GetExchangeUser().PrimarySmtpAddress` 嘗試解析真實 SMTP Email；若回傳 `None` 或發生例外，改以 MAPI 屬性 `PR_SMTP_ADDRESS`（`http://schemas.microsoft.com/mapi/proptag/0x39FE001E`）作為備援，避免比對不到網域關鍵字。
- **附件過濾**：
  - 僅處理實際檔案型附件（`attachment.Type == 1`，即 `olByValue`），忽略嵌入 OLE 物件。
  - `IGNORE_INLINE_IMAGES = True`：自動排除信件簽名檔常見內嵌小圖示，判斷條件包含：
    1. 檔名以 `image00` 開頭（不分大小寫）且大小 `< 15KB`；或
    2. 附件帶有內嵌 `PR_ATTACH_CONTENT_ID`（`0x3712001F`）或隱藏旗標 `PR_ATTACHMENT_HIDDEN`（`0x7FFE000B`）且檔案大小 `< 30KB`（可依需求開關/微調）。
  - （選用擴充）可預留 `ALLOWED_EXTENSIONS` / `IGNORED_EXTENSIONS` 頂部參數，方便日後限定 CAD/文件副檔名。

---

## 2. Markdown 分類規則引擎 (`rules.md`)

主程式每次執行時動態解析 [`rules.md`](./rules.md)：

- **解析邏輯**：
  - 每個 `## <資料夾名稱>` 二級標題代表一個分類目標資料夾（去除前後空白並過濾 Windows 非法檔名字元 `\ / : * ? " < > |`）。
  - 標題底下的清單項目以正規表達式解析，支援半形 `:` 或全形 `：` 分隔鍵值，支援五個欄位：
    - `- **主旨包含**:`
    - `- **寄件人包含**:`
    - `- **Email包含**:`
    - `- **附件包含**:`（亦可寫作 `附件檔名包含`，比對該信件內所有有效附件之檔名）
    - `- **內文包含**:`（亦可寫作 `內文關鍵字`、`正文包含`，僅於前四項未命中時進行第二階段比對）
  - 關鍵字以半形逗號 `,` 或全形頓號 `、` 分隔多個項目，切割後需去除空白並**過濾空字串**；留空代表該欄位不參與比對。
- **比對規則（兩階段比對）**：
  - **第一階段**：依 `rules.md` 由上往下比對所有規則的 `主旨包含`、`寄件人包含`、`Email包含`、`附件包含`（OR 邏輯、忽略英文大小寫、整封信連動），命中第一條即歸類停止。
  - **第二階段（備援）**：僅當全部規則的前四項均未命中時，才讀取信件內文（`mail.Body`）由上往下比對 `內文包含`。
  - 未命中任何規則的信件，歸類到預設資料夾（頂部可調參數 `DEFAULT_FOLDER = "其他未分類"`）。

---

## 3. 附件儲存與命名規則

- **輸出根目錄**：頂部可調參數 `OUTPUT_ROOT_DIR = r"d:\圖檔\CAD"`（程式與 `rules.md` 位於 `d:\圖檔\CAD\outlook_archiver\`，附件分類資料夾直接建於 `d:\圖檔\CAD\` 之下）。
- **收件日期時間子資料夾**：預設開啟 `USE_DATETIME_SUBFOLDER = True`、`DATETIME_SUBFOLDER_FORMAT = "%Y-%m-%d_%p%I%M"`，會在分類資料夾底下依信件收件日期時間建立子資料夾（例如 `d:\圖檔\CAD\合宏展\2026-11-26_AM0938\`，精確至分鐘），方便區分不同時間寄來的圖檔。
- **延遲建夾**：僅在信件有符合條件的有效附件需要儲存時，才建立目標分類資料夾與日期時間子資料夾（避免產生空資料夾）。
- **檔名清理與同名衝突處理**：
  - 清除附件原始檔名中的非法字元，避免存檔失敗。
  - 若目標路徑已存在同名檔案，**不可直接覆蓋**，保留原檔名並在檔名與副檔名之間自動加上流水號 `_1`、`_2`（例如 `圖面.dwg` 重複時存為 `圖面_1.dwg`、`圖面_2.dwg`；設為可調參數 `DUPLICATE_FILENAME_MODE = "suffix_counter"`）。

---

## 4. 防重複下載機制

- **紀錄儲存**：使用本機 SQLite 資料庫 `state.db`（或 JSON 紀錄檔），確保跨次執行與排程無人值守時不會重複下載。
- **識別 Key 設計**：
  - 以郵件唯一識別碼（優先讀取 `PR_INTERNET_MESSAGE_ID` `0x1035001E`，無值時退回 `mail.EntryID`）搭配「附件序號 / 原始檔名」作為下載紀錄 Key，支援單封信多個附件中斷後續傳。
  - 對於**無有效附件**或**僅含被排除內嵌小圖**的信件，亦可記錄「已檢查過」狀態，避免每次排程重複掃描。
- **Dry-Run 模式**：支援 CLI 參數 `--dry-run`，僅輸出測試結果與預期分類路徑而不實際寫入檔案與狀態庫，方便規則調整後驗證。

---

## 5. 日誌與錯誤處理

- 使用 Python 標準 `logging` 搭配 `RotatingFileHandler`，輸出至 `logs/archiver.log`（UTF-8 編碼，保留歷史輪替檔案）並同步輸出至主控台。
- 每筆處理記錄包含：執行時間、掃描信件數、信件主旨/寄件人、命中規則名稱、下載附件路徑或跳過原因。
- 單一信件或單一附件發生 COM / I/O 例外時記錄完整堆疊並繼續處理下一封，不中斷整體流程。

---

## 6. Windows 自動排程執行

- **啟動腳本 (`run_archiver.bat`)**：
  - 明確切換工作目錄至程式資料夾（`cd /d "d:\圖檔\CAD\outlook_archiver"`）。
  - 設定 `set PYTHONUTF8=1`，避免背景排程執行時因 Windows 系統預設代碼頁造成中文路徑或檔名編碼錯誤。
  - 呼叫指定 Python 環境執行 `outlook_archiver.py`。
- **Windows 工作排程器（Task Scheduler）設定注意事項**：
  - **安全性選項（關鍵）**：必須勾選「**只有在使用者登入時才執行**」，不可選擇「不管使用者是否登入都要執行」（否則程式會在背景 Session 0 執行，無法連接桌面登入中的 Outlook COM）。
  - **權限一致性**：排程執行的權限層級須與 `OUTLOOK.EXE` 一致（皆以一般使用者權限執行，勿混用系統管理員權限）。
  - **防止重疊執行**：排程設定勾選「如果工作已在執行，則不要啟動新的執行個體」。
