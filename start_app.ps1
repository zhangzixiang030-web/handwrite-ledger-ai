$ErrorActionPreference = "Stop"
$python = "python"
Set-Location $PSScriptRoot
& $python -m pip install -r requirements.txt
& $python app.py
