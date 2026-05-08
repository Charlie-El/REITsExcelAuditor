$ErrorActionPreference = "Stop"
$ProjectRoot = $PSScriptRoot
$IconPath = Join-Path $ProjectRoot "reit_excel_auditor\assets\app_icon.ico"
$TemplateDir = Join-Path $ProjectRoot "standard_templates"
$ConfigDir = Join-Path $ProjectRoot "config"

if (-not (Test-Path -LiteralPath $IconPath)) {
    throw "Application icon was not found: $IconPath"
}
if (-not (Test-Path -LiteralPath $TemplateDir)) {
    throw "Standard template folder was not found: $TemplateDir"
}
if (-not (Test-Path -LiteralPath $ConfigDir)) {
    throw "Config folder was not found: $ConfigDir"
}

Write-Host "Checking PyInstaller..."
python -m PyInstaller --version *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller not found. Installing PyInstaller..."
    python -m pip install pyinstaller
}

Write-Host "Building REITs Excel auditor executable..."
python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --icon "$IconPath" `
    --add-data "$IconPath;reit_excel_auditor\assets" `
    --add-data "$TemplateDir;standard_templates" `
    --add-data "$ConfigDir;config" `
    --exclude-module numpy `
    --exclude-module pandas `
    --exclude-module matplotlib `
    --exclude-module IPython `
    --exclude-module pytest `
    --exclude-module PIL `
    --exclude-module pygame `
    --exclude-module scipy `
    --exclude-module zmq `
    --exclude-module cryptography `
    --exclude-module gevent `
    --exclude-module lxml `
    --name "REITsExcelAuditor" `
    --distpath "dist" `
    --workpath "build\pyinstaller" `
    --specpath "build" `
    "reit_excel_auditor\app.py"

Write-Host ""
Write-Host "Build finished: dist\REITsExcelAuditor.exe"
