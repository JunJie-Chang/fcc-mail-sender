# Mail Sender

以 Python 透過 Gmail SMTP 批次發送 HTML 活動邀請信的工具。支援斷點續寄，每次執行自動跳過已發送的收件人。

## 功能

- 從 Excel（`.xlsx`）讀取收件人名單
- HTML 信件範本，支援 `{{name}}` 個人化替換
- 每次最多發 90 封（低於 Gmail 免費帳號每日上限 500 封）
- 每封發完即時將狀態寫回 xlsx（`sent` / `error`），中途中斷不遺失進度
- 執行結束後在 xlsx 末尾寫入發送摘要

## 專案結構

```
mail-sender/
├── send_emails.py    # 主程式
├── contacts.xlsx     # 收件人名單
├── template.html     # 信件 HTML 內容
├── .env              # 帳密設定（不進 git）
├── .env.example      # 帳密範本
└── requirements.txt
```

## Excel 欄位格式

| 欄位 | 內容 |
|------|------|
| A | 報到號碼 |
| B | 公司 |
| C | 職稱 |
| D | Email |
| E | 姓名 |
| F | status（程式自動新增與更新） |

`status` 欄位值：空白 = 未發送、`sent` = 已發送、`error` = 發送失敗

## 快速開始

### 1. 建立虛擬環境並安裝依賴

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. 設定帳密

複製範本並填入 Gmail 帳號與 App Password：

```bash
cp .env.example .env
```

`.env` 內容：

```
GMAIL_USER=你的帳號@gmail.com
GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx
```

> **取得 App Password**：Google 帳號 → 安全性 → 兩步驟驗證（需先開啟）→ 應用程式密碼 → 建立（名稱隨意）→ 複製 16 位密碼

### 3. 準備收件人名單

將收件人資料填入 `contacts.xlsx`，確保欄位順序符合上方格式。

### 4. 編輯信件內容

修改 `template.html`，使用 `{{name}}` 作為收件人姓名的佔位符，程式會自動替換。

### 5. 執行

```bash
python send_emails.py
```

250 人以每次 90 封計算，需執行 **3 天**。每次執行會自動從上次中斷處繼續。

## 設定參數

`send_emails.py` 頂部可調整以下設定：

| 參數 | 預設值 | 說明 |
|------|--------|------|
| `MAX_SEND` | `90` | 每次執行最多發送封數 |
| `SLEEP_EVERY` | `10` | 每幾封暫停一次 |
| `SLEEP_SECONDS` | `1` | 暫停秒數 |
| `SENDER_NAME` | `藍濤亞洲 FCC Partners` | 寄件人顯示名稱 |
| `SUBJECT` | （活動邀請主旨） | 信件主旨 |

## 注意事項

- `.env` 已列入 `.gitignore`，帳密不會上傳至 GitHub
- `contacts.xlsx` 含個人資料，上傳前請自行評估是否需要排除
