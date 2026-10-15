# Shop YoY 分析系統

讀取 **Hugo_by Shop_YoY** 匯出的 Excel，配合內置 **stores-template**（店鋪主檔），
按 **Site／年份／月份** 篩選計算 **YoY%**，並可匯出 Excel 報告。專為 Zeabur 部署而設。

## 功能

- 上載 `Hugo_by Shop_YoY_*.xlsx`（拖放或點擊選擇）
- 自動對照 `stores-template.csv`，為每間店補上 Regional／Class 1／Class 2／Size／OM／Type／需特別關注 等屬性
- 篩選：**Site**（可搜尋、多選）、**年份**、**月份**（單月 = 單月 YoY；多月 = 累計 YoY，例如 1–9 月 YTD）
- 檢視模式：**累計**（合併所選月份）或 **按月矩陣**（每店一行、每月一欄顯示該月 YoY%，可切換指標：淨銷售／銷量／GP$／GP%／數值，尾加「合計」欄）
- 進階篩選：Regional／OM／Class 1／Class 2／Size／Type、只顯示需特別關注店鋪、只顯示模板內店鋪
- 分組檢視：店鋪／Regional／OM／Class 1／Class 2／Size／Type（連分組匯總 YoY）
- 指標：淨銷售、銷量、毛利 GP$ 的 YoY%，以及毛利率變化（± pp）
- 一鍵下載 Excel 報告（格式與畫面一致，附自動篩選、凍結窗格）
- 自動提示不完整月份（例如月中匯出的當月資料）

## 檔案結構

```
shop-yoy-app/
├── app.py                 # Flask 後端（解析 Excel、計算 YoY、產生匯出檔）
├── templates/index.html   # 前端介面（單頁，免編譯）
├── stores-template.csv    # 內置店鋪主檔（定期更新）
├── requirements.txt
└── Dockerfile             # Zeabur 部署用
```

## 本機執行（測試用）

```bash
cd shop-yoy-app
pip install -r requirements.txt
python app.py
# 開啟 http://127.0.0.1:5000
```

## 部署到 Zeabur

### 1. 推送到 GitHub

```bash
cd shop-yoy-app
git init
git add .
git commit -m "Shop YoY 分析系統"
git branch -M main
git remote add origin https://github.com/<你的帳號>/<repo>.git
git push -u origin main
```

### 2. Zeabur 建立服務

1. 登入 [Zeabur](https://zeabur.com) → **Create Project**（選一個地區，建議 Hong Kong 或鄰近地區）
2. **Add Service → Git**（或 GitHub）→ 選擇你的 repo
3. Zeabur 偵測到 `Dockerfile` 會自動以 Docker 建置，毋須其他設定
4. 服務啟動後，於 **Networking → Generate Domain** 產生公開網址，即可使用

> 服務會讀取 `PORT` 環境變數（Zeabur 自動提供，預設 8080），毋須手動設定。

### 3. 更新

- **更新店鋪主檔**：修改 `stores-template.csv` → push 到 GitHub → Zeabur 自動重新部署
- **更新程式**：push 到 GitHub 即自動重新部署

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
- 上載的資料只存於服務記憶體（保留最近 5 份），服務重啟後需重新上載
- 金額及銷量維持來源數據；畫面整數顯示，匯出檔保留完整數值

## 常見問題

| 情況 | 處理 |
| --- | --- |
| 提示「找不到資料表」 | 確認上載的是 Hugo_by Shop_YoY 匯出檔（需有 Month、Shop、Net Sales 欄） |
| 提示「資料已過期」 | 服務曾重啟，重新上載 Excel 即可 |
| 去年同月顯示「—」 | 所選月份的去年同期沒有資料（例如 2026 年只有 1–10 月） |
| 月末數字偏低 | 當月為進行中月份，系統會顯示黃色提示 |
