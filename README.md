# Shop YoY 分析系統

讀取 **Hugo_by Shop_YoY** 匯出的 Excel，配合內置 **stores-template**（店鋪主檔），
按 **Site／年份／月份** 篩選計算 **YoY%**，並可匯出 Excel 報告。以標準 Docker 封裝，可直接部署到 Railway 等雲端平台。

## 功能

- 上載兩款檔案（自動偵測）：
  - `Hugo_by Shop_YoY_*.xlsx`（店鋪級每月匯總）
  - `Hugo_all_YoY_*.xlsx`（**SKU 明細**：含 Brand／MCH1／MCH2／MC／Nature／Postal 等維度）
- **多維度分析**（明細檔限定）：「分組」可**多選（最多 3 層）**，維度包括：品牌 Brand／大分類 MCH1／中分類 MCH2／品類 MC／客群 Nature／客源 Postal／買手 Buyer／貨品狀態 Status／客戶類型 Customer／推廣 PromoID1、PromoID2；預設仍然係「**店鋪 Site**」
- **多層嵌套檢視**：例如「品牌 › 品類」＝每個品牌總數＋其下品類明細（縮排顯示、每層小計）；累計、按月矩陣、全部指標、匯出 Excel 全部支援
- 自動對照 `stores-template.csv`，為每間店補上 Regional／Class 1／Class 2／Size／OM／Type／需特別關注 等屬性
- 篩選：**Site**（可搜尋、多選）、**年份**、**月份**（單月 = 單月 YoY；多月 = 累計 YoY，例如 1–9 月 YTD）
- 檢視模式：**累計**（合併所選月份）或 **按月矩陣**（每店一行、每月一欄，尾加「合計」欄；指標可揀：淨銷售／銷量／GP$／GP%／數值／**全部指標**／**全部指標＋值**——後兩者四組（或八組）指標並排、可橫向捲動）
- 進階篩選：Regional／OM／Class 1／Class 2／Size／Type、只顯示需特別關注店鋪、**只顯示精簡SKU店舖**（18 間：HA19、HA21、HA33、HA37、HB30、HC44、HB86、HB77、HB91、HB94、HC26、HC60、HC62、HC68、HC13、HC69、HB87、HB80）、只顯示模板內店鋪
- 分組檢視：店鋪／Regional／OM／Class 1／Class 2／Size／Type（連分組匯總 YoY）
- 指標：淨銷售、銷量、毛利 GP$ 的 YoY%，以及毛利率變化（± pp）
- 一鍵下載 Excel 報告（格式與畫面一致，附自動篩選、凍結窗格）
- 上載大檔案時顯示**實時解析進度（%）**（背景解析、可邊睇邊等）
- **上載抗斷線**：進度輪詢有 30 秒逾時及自動重試（畫面會顯示「重試中（n/6）」），短暫斷線或服務重新部署都唔會再卡死
- **私隱保護**：資料只存伺服器記憶體、逾時自動刪除（預設 8 小時）、可隨時一鍵「清除資料」、可選密碼保護
- 自動提示不完整月份（例如月中匯出的當月資料）

## 檔案結構

```
shop-yoy-app/
├── app.py                 # Flask 後端（解析 Excel、計算 YoY、產生匯出檔）
├── templates/index.html   # 前端介面（單頁，免編譯）
├── stores-template.csv    # 內置店鋪主檔（定期更新）
├── requirements.txt       # 套件版本已固定，確保每次重建成一樣
├── Dockerfile             # Docker 部署用（Railway 等）
├── .dockerignore          # 建置時排除非必要檔案（加快建置）
└── .gitignore             # 防止資料檔／快取被 push 上 GitHub
```

## 本機執行（測試用）

```bash
cd shop-yoy-app
pip install -r requirements.txt
python app.py
# 開啟 http://127.0.0.1:5000
```

## 部署到 Railway

### 1. 推送到 GitHub

```bash
cd shop-yoy-app
git add .
git commit -m "更新"
git push
```

（如未設定 remote：`git remote add origin https://github.com/<你的帳號>/<repo>.git`，再 `git push -u origin main`。）

### 2. Railway 建立服務

1. 以 GitHub 帳號登入 [Railway](https://railway.com) → **New Project → Deploy from GitHub repo** → 選擇你的 repo（首次需授權 Railway 讀取 GitHub）
2. Railway 偵測到 `Dockerfile` 會自動以 Docker 建置（Metal builder，設有層快取：之後只改程式碼唔會重新安裝套件，重新部署快好多），毋須任何設定
3. 如 repo 根目錄唔係 app 本身（例如 repo 內再包一層 `shop-yoy-app/` 資料夾）：**Settings → Source → Root Directory** 填 `shop-yoy-app`
4. **Settings → Regions** 揀 **Southeast Asia (Singapore)**（最近香港）
5. **Settings → Deploy → Healthcheck Path** 填 `/healthz`（app 已內置此端點；部署時 Railway 會確認新版本健康先切換流量，做到零停機更新）
6. 服務啟動後，於 **Settings → Networking → Public Networking → Generate Domain** 產生公開網址（`*.up.railway.app`）
7. ⚠️ **Replicas 必須保持 1**（Free 計劃固定係 1；Hobby 亦唔好加）——已上載資料只存喺**單一實例**嘅記憶體，開多個實例會出現「一時有資料、一時話過期」嘅情況
8. **Settings → Deploy → Replica Limits**：CPU **1 vCPU**、Memory **1–2 GB** 已好夠——上限只係「安全網」，按**實際用量**計費，設高啲唔會預先加錢；但設得太低（例如 0.5 GB），大檔解析或大矩陣匯出時會爆記憶體遭強制重啟（資料清空）

> 服務會讀取 Railway 自動注入嘅 `PORT` 環境變數（Dockerfile 預設 8080），毋須手動設定。

### 3. 收費

- 新帳戶有 **30 日 US$5 試用額度**（免信用卡；試用期每服務上限 2 vCPU / 1 GB）
- 之後可選：
  - **Free**：US$0/月，每月 US$1 用量額度；每服務上限 **1 vCPU / 0.5 GB RAM、1 replica**，日誌保留 3 日——明細大檔（15 萬行級）或一次過保留多份資料時，記憶體可能唔夠
  - **Hobby**：US$5/月（含 US$5 用量）；每服務上限 **48 vCPU / 48 GB、6 replicas**——明細大檔建議用呢個（但 **replicas 保持 1**）
- 用量**按秒計**：記憶體 **US$10/GB/月**、CPU **US$20/vCPU/月**、對外流量 **US$0.05/GB**；瞓覺期間唔計費
- **預算參考**（已開 Serverless）：間中使用嘅內部工具，實際用量估計約 **US$0.1–1/月**（成本主要係清醒期間嘅記憶體；CPU 同流量近乎唔計）——Hobby 用量唔過 US$5 就實收 **US$5/月**；Free 嘅 US$1 額度處理細檔一般都夠
- 詳情以 [railway.com/pricing](https://railway.com/pricing) 為準

### 4. 暫停服務（偶爾使用慳錢）

- **自動瞓覺（推薦）**：服務 → **Settings → Deploy → Serverless** → 開啟。最後一次對外通訊後 5–10 分鐘自動停止計費，有人開網址自動喚醒（冷啟動幾秒至十幾秒；瞓覺期間嘅請求會**自動排隊**，等容器起身後照常處理，一般唔使手動整理）。設定於**下次部署**先生效。注意：瞓覺後容器記憶體會清空——**已上載資料要重新上載**
- **完全停止**：服務 → **Deployments** → 最新 deployment 嘅 **⋮ → Remove**，即時停止計費；之後喺 dashboard 按 **Deploy** 或 push 任何 commit 即可重新啟動（環境變數及設定保留）
- **防超支保險**：Workspace → Usage 設定 **Hard limit**（最低 US$10），達到上限會自動停低服務

### 5. 更新

- **更新店鋪主檔**：修改 `stores-template.csv` → push 到 GitHub → Railway 自動重新部署
- **更新「精簡SKU店舖」名單**：修改 `app.py` 最上方的 `LEAN_SKU_SITES` 清單 → push
- **更新程式**：push 到 GitHub 即自動重新部署（層快取關係，唔改 `requirements.txt` 就唔會重新安裝套件）
- **部署設定**：本 app **唔使用** `railway.toml`／`railway.json`——Railway 將於 2026-12-01 停用 Config-as-code，所有設定（Region／Healthcheck／Serverless／Variables／Replicas）一律喺 dashboard 做

## 使用流程

1. 開啟網址 → 上載最新匯出的 `Hugo_by Shop_YoY_*.xlsx`
2. 系統自動列出可選的年份／月份／店鋪，預設：最新年份、全部月份、全部店鋪
3. 選擇篩選條件 → 即時看到 YoY 結果
4. 想逐月睇就將「檢視」切換成「按月矩陣」，可用「指標」下拉選擇顯示咩數字
5. 按「⬇ 下載 Excel」匯出目前篩選結果（累計或按月矩陣都會照畫面匯出）

## 資料注意事項

- Excel 內的 `Total` 行及底部備註行會自動排除；只讀取 `Month` 為 `YYYY-MM` 的資料列
- 店鋪以此格式對照主檔：`HA02 - 駱克` → Site `HA02`
- 不在 stores-template 內的店鋪仍會顯示（標示「模板外」），可在進階篩選中排除
- **上載的資料只存於服務「記憶體」（RAM），不會寫入硬碟或資料庫；服務重啟即消失**
- **其他人開啟同一網址只會見到上載畫面，無法看到你的數據**（數據以隨機 token 存取，只有上載嗰個瀏覽器持有）
- 未手動清除的資料會於上載後 **8 小時自動刪除**（可用環境變數 `DATA_TTL_HOURS` 調整；`0` = 不自動過期）
- 可隨時按畫面右上角「**清除資料**」即時刪除所有已上載資料
- 金額及銷量維持來源數據；畫面整數顯示，匯出檔保留完整數值

## 私隱與安全（環境變數）

| 環境變數 | 預設 | 說明 |
| --- | --- | --- |
| `APP_PASSWORD` | （未設定） | 設定後，開啟網站需先輸入密碼（瀏覽器彈出登入框：用戶名稱隨意，密碼填此值）。公開部署建議設定 |
| `DATA_TTL_HOURS` | `8` | 上載資料於記憶體保留時數，逾時自動刪除；`0` = 不自動過期 |

在 Railway 設定：服務頁面 → **Variables** → 新增，儲存後會自動重新部署。

> 如果完全唔想有公開網址：唔好按 **Generate Domain**（服務就只存在於 Railway 內部網絡）；或改為只在本機執行 `python app.py`（只有自己電腦用得到，不會曝露上網）。

## 效能

- Excel 解析引擎：**python-calamine**（Rust 實作，比 openpyxl 快約 10 倍）
- 實測速度：店鋪級檔（約 2 千行）< 0.1 秒；明細檔 39k 行約 **0.5 秒**；155k 行（4 個月）約 **2 秒**
- 如環境未安裝 python-calamine，會自動退回 openpyxl（功能相同、較慢）
- 萬一 calamine 讀唔到某檔案，亦會自動嘗試 openpyxl，唔會失敗
- **Railway 相關**：Docker 層快取（唔改 `requirements.txt` 就唔會重裝套件）；容器設 `MALLOC_ARENA_MAX=2` 減少記憶體碎片；大數據輪詢／載入自動 gzip（省對外流量 US$0.05/GB）
- 記憶體提示：已上載資料保留喺 RAM（最多 5 份、預設 8 小時後自動刪除）。如處理大檔或想即時釋放記憶體，按右上角「清除資料」

## 常見問題

| 情況 | 處理 |
| --- | --- |
| 提示「找不到資料表」 | 確認上載的是 Hugo_by Shop_YoY 匯出檔（需有 Month、Shop、Net Sales 欄） |
| 提示「資料已過期」 | 服務曾重啟、瞓覺被喚醒（記憶體清空）、逾時（預設 8 小時）或已按「清除資料」，重新上載 Excel 即可 |
| 想即時刪除數據 | 按右上角「清除資料」；或等自動逾期刪除 |
| 想加密碼先入得 | 在 Railway 設定環境變數 `APP_PASSWORD`（服務 → Variables） |
| 去年同月顯示「—」 | 所選月份的去年同期沒有資料（例如 2026 年只有 1–10 月） |
| 上載時進度條好似停住 | 會自動重試（顯示「連線不穩或服務正在更新，重試中（n/6）」）；如重試後仍失敗，多數係服務正在重新部署或有數據未上載完，重新上載即可 |
| 開網址第一次好慢（或逾時） | 服務正由「瞓覺」喚醒中——瞓覺期間嘅請求會自動排隊，通常等幾秒就會照常處理；如逾時或見到錯誤，重新整理一次即可 |
| Railway 建置失敗（pip 連線中斷） | Dockerfile 已內置自動重試（最多 3 次）；若仍失敗，多數係 PyPI 短暫問題——過幾分鐘於 dashboard 按 **Deploy** 重試即可 |
| 月末數字偏低 | 當月為進行中月份，系統會顯示黃色提示 |
