$ErrorActionPreference = "Stop"

$venvPython = ".\.venv\Scripts\python.exe"

Get-Process -Name "ExcelTemplateFiller" -ErrorAction SilentlyContinue | Stop-Process -Force

$stagedExamples = ".\build_package_examples"
if (Test-Path $stagedExamples) {
  Remove-Item -LiteralPath $stagedExamples -Recurse -Force
}
New-Item -ItemType Directory -Path $stagedExamples | Out-Null
Get-ChildItem -LiteralPath ".\examples" -File -Force |
  Where-Object { $_.Name -like "sample_*.xlsx" } |
  Copy-Item -Destination $stagedExamples -Force

if (-not (Test-Path $venvPython)) {
  python -m venv .venv
}

& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r requirements.txt

$condaRoot = (& $venvPython -c "import sys; print(sys.base_prefix)")
$expatDll = Join-Path $condaRoot "Library\bin\expat.dll"
$libexpatDll = Join-Path $condaRoot "Library\bin\libexpat.dll"

& $venvPython -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --name "ExcelTemplateFiller" `
  --add-data "$stagedExamples;examples" `
  --add-binary "$expatDll;." `
  --add-binary "$libexpatDll;." `
  --hidden-import pyexpat `
  --hidden-import xml.parsers.expat `
  --exclude-module pkg_resources `
  --exclude-module setuptools `
  --exclude-module numpy `
  --exclude-module pythonnet `
  --exclude-module clr `
  app_web.py

Write-Host ""
Write-Host "Build complete: dist\ExcelTemplateFiller.exe"
