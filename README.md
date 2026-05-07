# WebChecker｜網站檢核小幫手

以 **Streamlit** 建置的網頁應用，協助對公開網站進行**站內連結爬梳與多項檢核指標**的自動化輔助檢查。核心爬蟲與檢測邏輯集中於 `webchecker_core.py`，介面與流程由 `app.py` 負責，適合機關／單位網站品質稽核或改版前的初步掃描。

---

## 功能概要

| 面向 | 說明 |
|------|------|
| **全站掃描** | 自起始網址起，於設定的**網域與路徑範圍**內以佇列方式逐一載入頁面（Playwright Chromium），擷取站內連結並擴充待掃清單；支援暫停、清除重置。 |
| **檢核指標（輔助參考）** | 涵蓋「基本要件、導覽、語系編碼、語言版本、有效連結、資料即時、動畫／文件格式、意見信箱、搜尋、跨裝置 viewport、載入速度、流量統計」等面向；部分項目為啟發式字串／DOM 檢查，**正式上線或稽核仍建議人工複核**。 |
| **PageSpeed Insights** | 可選是否在掃描前對起始 URL 呼叫 Google PageSpeed Insights API，取得手機／桌面效能分數並換算報表顯示（約需數十秒至數分鐘）。 |
| **Excel 匯出** | 掃描完成後可下載 `.xlsx`，內含站內掃描網址與外站連線驗證清單及「有效連結」欄位標註。 |
| **常用網站** | 於側欄「常用網站設定」維護名稱與網址清單（存於 `favorites.json`），掃描時可快速選取。 |
| **排除規則** | 依網域設定關鍵字排除規則（存於 `config.json`），略過不符合業務需求的連結行為請參考程式內 `webchecker_core` 與實際需求調整。 |
| **檢核指標說明** | 內建「檢核指標說明」頁面，對照每一項的檢測意義與程式判斷邏輯。 |
| **介面** | 側欄選單切換模式；支援淺色／深色主題（自訂 CSS）。 |

---

## 技術棧

- **Python 3**（建議 3.9+）
- **Streamlit**：Web UI、`session_state` 管理掃描狀態與報表
- **Playwright（異步 Chromium）**：頁面載入、DOM／連結擷取、部分分頁與檔案連結處理
- **BeautifulSoup**：HTML 解析與連結／指標輔助判斷
- **requests / urllib3 / aiohttp**：HTTP 請求與外站連線探測
- **pandas + xlsxwriter**：Excel 匯出
- **其他**：`pandas`、`tenacity`、`fake-useragent` 等（見 `requirements.txt`）

執行時側欄會顯示**檢核核心版本**（與 `webchecker_core.SCAN_ENGINE_BUILD` 一致），便於對照部署版本。

---

## 系統需求

- Windows／macOS／Linux 皆可；專案內附 **Windows**「一鍵啟動」批次檔。
- 需安裝 **Python**，並建議將 `python` 或 Windows 啟動器 `py -3` 加入 PATH。
- **Playwright** 需額外安裝 **Chromium** 瀏覽器二進位檔（見下方安裝步驟）。
- 掃描會對目標網站發出大量請求，請確認已取得適當授權並遵守該站 `robots.txt` 與使用條款。

---

## 安裝步驟

### 1. 取得程式碼

```bash
git clone <您的儲存庫 URL>
cd WebChecker
```

### 2. 建立虛擬環境（建議）

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

### 3. 安裝相依套件

```bash
pip install -r requirements.txt
pip install playwright
playwright install chromium
```

> **說明：** `requirements.txt` 列出 Streamlit 與解析／試算表相關套件；**Playwright** 為爬蟲核心所用，須另外安裝並執行 `playwright install chromium`。

### 4. 啟動應用

```bash
streamlit run app.py
```

瀏覽器會開啟本機預設連接埠（通常為 `http://localhost:8501`）。

### Windows：使用「一鍵啟動」

雙擊專案根目錄下的 **`一鍵啟動.bat`**，會於新路徑開啟命令視窗並執行 `python -m streamlit run app.py`（若無 `python` 則嘗試 `py -3`）。關閉該視窗即停止服務。

### 免安裝 Python／Playwright（Windows 可攜式套件）

若希望**使用端不必安裝 Python、也不必執行 `playwright install`**，可在**一台已安裝 Python 的 Windows 電腦**上製作可攜式資料夾，再壓成 ZIP 或再用 Inno Setup 等做成「安裝檔」分發：

1. 於專案根目錄執行 **`scripts\製作可攜式套件.bat`**（或 `powershell -ExecutionPolicy Bypass -File scripts\build_portable_windows.ps1`）。
2. 完成後會產生 **`dist\WebChecker-Portable\`**，內含 `.venv`、Chromium（`ms-playwright`）與程式檔。
3. 將 **`WebChecker-Portable`** 整包 ZIP 給使用者；使用者解壓後雙擊 **`啟動WebChecker.bat`** 即可（會設定 `PLAYWRIGHT_BROWSERS_PATH` 指向同目錄瀏覽器）。

**說明：** 製作機仍需 Python 與網路以下載套件；套件體積約數百 MB；Chromium 與 Python 皆綁定建置時的 Windows 環境，建議在與使用者相近的系統（例如 64 位元 Windows 10/11）上製作。

---

## 使用方法

### 側欄選單

1. **執行全站掃描**：輸入完整起始網址（或使用「常用網站」；自訂網址優先）。
2. **常用網站設定**：新增／編輯／刪除常用項目。
3. **排除規則設定**：依網域輸入關鍵字（每行一則）並儲存。
4. **檢核指標說明**：閱讀各指標定義與技術邏輯。
5. **外觀**：切換淺色／深色主題。

### 掃描頁面操作

1. 在「自訂網站」貼上 **完整 URL**（例如 `https://example.gov.tw/`），或僅從「常用網站」選擇（自訂為空時才生效）。
2. 視需要勾選 **「同時執行載入速度檢測 (PageSpeed API)」**。
3. 按 **「開始掃描」**。程式會先進行 PageSpeed（若勾選），再以 Playwright 分批處理佇列中的網址。
4. 可使用 **暫停掃描**、**清除重置** 控制流程。
5. 完成後檢視 **完整檢核指標報告**，並可 **下載掃描網址清單（Excel）**。

### 掃描範圍提示（與介面說明一致）

- 若入口為某層路徑下的單一頁面（例如 `…/Sale/Login`），程式會試著以合理的路徑前綴界定範圍，避免漏掃同層目錄。
- 若要**明確限定某一子目錄**，且該層名稱無副檔名，請在網址**末端加上 `/`**（例如 `…/Sale/`）。

---

## 設定檔與資料檔

| 檔案 | 用途 |
|------|------|
| `config.json` | 依網域儲存排除規則文字（程式啟動後於 UI 中編輯並儲存會寫入）。首次使用前可能不存在，儲存後自動建立。 |
| `favorites.json` | 常用網站清單（JSON 陣列）。可於 UI 維護或備份還原。 |

請將上述檔案視為**環境／使用者資料**，若使用版本控制，可依需求列入 `.gitignore`。

---

## 部署到 Streamlit Cloud：常用網站／排除規則持久化

Streamlit Community Cloud 容器之檔案系統屬**暫時性**，每次重啟會回到 git clone 狀態，
直接寫 `favorites.json` / `config.json` 會在重啟後遺失。本專案內建以 **GitHub Gist** 作為
雲端持久層，所有使用者共享同一份清單，重啟亦不會遺失。

### 一次性設定（約 5 分鐘）

1. **建立 Secret Gist**（資料儲存位）
   - 到 [gist.github.com](https://gist.github.com/) 用部署用 GitHub 帳號登入。
   - 建立一個 **Secret gist**，內含兩個檔案：
     - `favorites.json` 內容填 `[]`
     - `config.json` 內容填 `{}`
   - 建立後從網址列複製 **gist id**（網址尾段 32 位十六進位字串）。

2. **建立 Fine-grained Personal Access Token**
   - [github.com/settings/tokens?type=beta](https://github.com/settings/tokens?type=beta) → **Generate new token**
   - **Repository access：** 選 *No access*（不需 repo 權限）。
   - **Account permissions：** 將 **Gists** 設為 **Read and write**。
   - 設定到期日後產生，立刻複製 token 字串（離開頁面後就看不到）。

3. **填入 Streamlit Cloud Secrets**
   - 到 [share.streamlit.io](https://share.streamlit.io/) → 該 app → 右側 ⋯ → **Settings → Secrets**。
   - 貼上以下內容並把 token / gist id 換成步驟 1、2 取得的值，按 **Save**：

     ```toml
     [github]
     gist_token = "ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
     gist_id    = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
     ```

   - 回到 app 介面點 **Reboot app**。

4. （可選）**本機測試同步行為**
   - 將 `.streamlit/secrets.toml.example` 複製成 `.streamlit/secrets.toml`（**已被 `.gitignore` 排除**），填入相同憑證後本機執行 Streamlit 即可同步測試。

### 行為說明

- 容器啟動時：自動把 Gist 上之 `favorites.json` / `config.json` 拉下來覆寫本機檔。
- UI 新增／編輯／刪除：寫本機檔之外**最佳努力**（best-effort）以 `PATCH` 同步至 Gist；同步失敗只會在 Manage app log 留訊息（`[cloud_persistence] save '...' failed: ...`），不會影響本機操作。
- 未設定 Secrets：Cloud Sync 自動關閉、回到「僅本機檔」行為，與本機執行一致。

---

## 檔案目錄說明

```
WebChecker/
├── app.py                 # Streamlit 主程式：側欄、四種模式頁面、掃描狀態與報表 UI
├── webchecker_core.py     # 檢核引擎：URL／範圍、Playwright 掃描批次、指標判斷、PageSpeed、Excel 輸出等
├── cloud_persistence.py   # GitHub Gist 持久層（Streamlit Cloud 重啟後仍保留設定）
├── requirements.txt       # pip 相依套件（已包含 playwright；本機仍需 playwright install chromium）
├── packages.txt           # Streamlit Cloud apt 系統相依（Chromium 必要 lib）
├── README.md              # 本說明檔
├── 一鍵啟動.bat           # Windows 快速啟動 Streamlit（可選）
├── scripts/               # 可攜式套件建置腳本（Windows）
│   ├── build_portable_windows.ps1
│   └── 製作可攜式套件.bat
├── .streamlit/
│   ├── config.toml             # Streamlit 客戶端設定（例如極簡工具列）
│   └── secrets.toml.example    # Cloud 持久化之 Gist 憑證樣本（請複製為 secrets.toml 後填入）
├── config.json            # （選用）排除規則，執行後可能由程式建立
├── favorites.json         # （選用）常用網站清單，執行後可能由程式建立
└── .gitignore             # Git 忽略規則
```

---

## 限制與注意事項

- 檢核結果為**輔助性質**：字串比對、單一瀏覽器自動化等無法取代完整無障礙、資安或內容正確性審查。
- **載入速度**依賴 Google PageSpeed Insights 網路 API，可能受配額、網路或目標站無法被公開分析等因素影響。
- 外站連結數量多時，單輪掃描時間可能明顯增加。
- 本機 Streamlit 預設僅監聽本機；若需對外提供服務，請自行評估反向代理、驗證與資安設定。

---

## 授權與貢獻

若本專案將公開於儲存庫，請於此補充授權條款（例如 MIT、Apache-2.0）與貢獻指南；目前程式結構可依上述檔案分工擴充或撰寫單元測試。
