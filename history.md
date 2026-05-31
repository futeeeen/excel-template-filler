# 專案上傳變更紀錄

## 2026.06.01_00:26:58
* 將真實專案 Excel 範例檔從 git 追蹤中移除，並在 `.gitignore` 排除同類型實際清單與範本檔案。
* 新增可公開上傳的假資料範例：`sample_source_list.xlsx`、`sample_fsd_template.xlsx`、`sample_utft_template.xlsx`。
* 調整 self-test 與打包腳本，只使用 `sample_*.xlsx` 作為可公開範例，避免本機實際範例檔被包入 exe 或推上 GitHub。
* 更新 README，說明 repo 只保留假資料範例，實際專案清單與範本不得上傳。
* Validation: ran `python -m compileall app_web.py src` and `python app_web.py --self-test`.

## 2026.06.01_00:19:05
* 建立 Excel Checklist 產生工具，將舊版手動修改 Python 設定的流程改成可由同事自行操作的瀏覽器介面。
* 使用 `openpyxl` 實作來源清單讀取、範本內容替換、檔名替換、timestamp 輸出資料夾與 zip 下載。
* 採用本機 HTTP server + HTML/CSS/JavaScript 介面，避開 Tkinter/Tcl 在 Windows 打包後的 DLL 與版本衝突。
* 新增彈性替換設定，可設定多組「範本工作表 + 儲存格」，並可勾選是否替換範本內容或檔名。
* 新增 `build_exe.ps1`，負責建立 `.venv`、安裝相依套件、排除 Excel 暫存檔、補入 XML parser DLL，並用 PyInstaller 打包 exe。
* 補充 README 的使用方式、提供同事檔案範圍、打包方式與技術實作筆記，方便日後複用類似架構。
* Validation: ran `python -m compileall app_web.py src` and `python app_web.py --self-test`.
