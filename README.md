# Excel Checklist 產生工具

這個工具會讀取來源清單 Excel 的名稱資料，把每一筆名稱套進 Checklist 範本，最後下載一個 zip。zip 裡面的輸出資料夾會自動加上日期時間，避免覆蓋之前的產出。

## 使用方式

開啟 `dist\ExcelTemplateFiller.exe` 後，程式會自動開啟瀏覽器頁面。

1. 在「來源清單資料」選擇來源 Excel，設定來源工作表、名稱欄位、起始列。
2. 在「範本檔案」選擇 Checklist 範本。
3. 在「替換設定」確認替換文字，以及要替換哪些範本工作表與儲存格。
4. 依需求勾選「替換範本內容」與「替換檔名」。
5. 按「預覽筆數」確認資料，再按「產生並下載 zip」。

下載檔名會像這樣：

```text
output_checklists_20260531_181530.zip
```

zip 裡面的資料夾也會使用相同名稱。

## 常用設定

- 模板一：來源工作表 `CR FSD (模板一)`，名稱欄位 `E`，範本工作表 `FSD Checklist`，儲存格 `A1`。
- 模板二：來源工作表 `CR UTFT (模板二)`，名稱欄位 `E`，範本工作表 `檢核表`，儲存格 `A1`。

## 彈性替換

- 可設定多組「範本工作表 + 儲存格」。
- 可只替換範本內容。
- 可只替換檔名。
- 若不替換檔名，工具會自動在輸出檔名後加序號，避免重複覆蓋。

## 提供給同事

一般只需要提供這個檔案：

```text
dist\ExcelTemplateFiller.exe
```

同事不需要安裝 Python，也不需要拿到 `src`、`.venv`、`build`、`examples` 或其他專案檔。來源清單與範本 Excel 由同事在工具畫面自行選擇即可。

## 技術實作筆記

這個專案的目標是讓沒有 Python 環境的同事也能直接使用，因此採用「Python 核心邏輯 + 本機瀏覽器介面 + PyInstaller 打包」的方式。

### 主要技術

- Python 3：主程式語言。
- `openpyxl`：讀取來源 Excel、讀取範本 Excel、替換儲存格內容、另存新 Excel。
- Python 內建 `http.server`：啟動只綁定 `127.0.0.1` 的本機 HTTP server。
- HTML / CSS / JavaScript：提供使用者介面，讓同事用瀏覽器選檔、設定欄位、預覽筆數、下載 zip。
- Python 內建 `zipfile`：把批次產出的 Excel 打包成 zip 下載。
- PyInstaller：把 Python 程式與必要相依包成 Windows `.exe`。

### 為什麼不用 Tkinter

一開始曾使用 Tkinter 做桌面 GUI，但在 Miniconda + PyInstaller + Windows 的組合下，容易遇到 Tcl/Tk DLL 或 `init.tcl` 版本衝突。後來改成：

1. exe 啟動本機 HTTP server。
2. 程式自動用 Chrome 或 Edge 開啟 `http://127.0.0.1:<port>/`。
3. 使用者在瀏覽器頁面操作。

這樣可以避開 Tkinter/Tcl 相依問題，介面也比較容易用 HTML/CSS 調整。

### 程式結構

- `app_web.py`：本機 HTTP server 與瀏覽器版 UI。負責收檔、回傳工作表清單、預覽、產生 zip。
- `src/excel_template_filler.py`：Excel 產生核心邏輯。負責讀名稱、驗證設定、替換範本內容、替換檔名、建立 timestamp 輸出資料夾。
- `build_exe.ps1`：建立 `.venv`、安裝相依套件、打包 exe。
- `requirements.txt`：列出打包與執行需要的 Python 套件。
- `examples/`：開發與自我測試用的範例 Excel。

### 打包注意事項

`build_exe.ps1` 會：

- 建立或重用 `.venv`。
- 安裝 `openpyxl` 與 `pyinstaller`。
- 停止正在執行的 `ExcelTemplateFiller.exe`，避免打包時檔案被鎖住。
- 建立乾淨的 `build_package_examples`，排除 Excel 開啟時產生的 `~$...xlsx` 暫存檔。
- 額外加入 `expat.dll` 與 `libexpat.dll`，避免打包後讀取 `.xlsx` 時發生 XML parser DLL 錯誤。
- 排除不需要的模組，例如 `pkg_resources`、`setuptools`、`numpy`、`pythonnet`、`clr`，降低 exe 體積與啟動風險。

### 類似工具可複用的設計

如果之後要做類似「同事不用安裝 Python，也能處理 Office 檔案」的小工具，可以沿用這個模式：

1. 把真正的處理邏輯獨立成純 Python module。
2. 用本機 HTTP server + HTML 表單做介面。
3. 用瀏覽器的檔案選擇器上傳檔案到本機 server。
4. 後端處理完成後直接回傳 zip 或結果檔。
5. 用 PyInstaller 打成單一 exe。

這個模式比 Tkinter 更容易做出清楚的表單，也比較不容易被 Windows 桌面 GUI 相依套件卡住。

## 打包 exe

在本資料夾執行：

```powershell
.\build_exe.ps1
```

完成後 exe 會在：

```text
dist\ExcelTemplateFiller.exe
```

## 可公開範例檔

此 repo 只上傳假資料範例，放在 `examples`：

- `sample_source_list.xlsx`
- `sample_fsd_template.xlsx`
- `sample_utft_template.xlsx`

實際專案用的清單與範本檔案不應上傳到 GitHub，已透過 `.gitignore` 排除。
