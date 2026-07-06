$ErrorActionPreference = "Stop"
$python = "python"

Set-Location $PSScriptRoot
& $python -m pip install -r requirements.txt

if (Test-Path ".\build") {
  Remove-Item -LiteralPath ".\build" -Recurse -Force
}
$appName = "LedgerOCRDesktop"

if (Test-Path ".\dist\$appName") {
  Remove-Item -LiteralPath ".\dist\$appName" -Recurse -Force
}

& $python -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --name $appName `
  --collect-all webview `
  ".\desktop_app.py"

Write-Host "Built: $PSScriptRoot\dist\$appName\$appName.exe"
