# tiltdot-data-transfer

從 WMS 網站批次下載 tiltDot 傾斜計原始資料，依時間範圍過濾後換算成角秒（arcsec），
產出含圖表的 Excel 檔，並統整各感測器的溫度補償係數。

使用者操作說明請看 [docs/使用教學.md](docs/使用教學.md)。

## 開發環境

```bash
python -m venv .venv
.venv\Scripts\activate          # macOS / Linux: source .venv/bin/activate
python -m pip install -r requirements.txt
python -m playwright install chromium
python tiltdotdatatransfer.py
```

## 離線自我驗證

不需要網路，也不會開瀏覽器；用假造的原始資料跑完整條處理鏈：

```bash
python tools/selftest.py
```

涵蓋：時間範圍過濾（含未補零日期 `2026/8/12`、`7:45` 這類輸入）、分隔符號偵測、
異常值過濾與基準列一致性、TRANSFER / TEST 兩種模式的 Excel 欄位與公式、
校正參數總表、報告輸出、欄位設定錯誤的提示。

## 打包成安裝檔（Windows）

```bat
build\build.bat
```

會依序安裝相依套件、下載 Chromium（`PLAYWRIGHT_BROWSERS_PATH=0`，裝進 playwright
套件目錄好讓 PyInstaller 收得到）、跑自我驗證、再用 `build\tiltdot.spec` 打包成
`dist\tiltdot\`。接著用 Inno Setup 編譯 `build\installer.iss` 產生安裝檔。

幾個不能改的前提：

- **必須是 console 模式**：程式全靠 `input()` 互動，`--windowed` 會直接崩潰。
- **用 onedir 不用 onefile**：onefile 每次啟動都要解壓數百 MB 的 Chromium。
- **安裝到 `%LOCALAPPDATA%\Programs`**：程式預設把輸出建在 exe 旁邊，
  裝進 Program Files 會被權限擋住（程式本身有 fallback，但會多一次警告）。

## 出貨前必須確認

- [ ] 在**沒裝過 Python 與 Playwright** 的機器上跑過一次完整下載
- [ ] 網站目前是否需要登入才能查詢？程式碼**沒有任何登入流程**，
      若網站已改為需登入，所有序號都會回報「無資料或逾時」
- [ ] 原始資料的日期格式是 `DD/MM/YYYY` 還是 `MM/DD/YYYY`
      （程式會自動判定，日、月皆 ≤ 12 時預設 DMY，可在設定檔用 `date_order` 指定）
