param(
    [string]$PresentationPath = "C:\! BASTA\pptx\TEST.pptx"
)

$ErrorActionPreference = "Stop"

$pythonScriptPath = Join-Path $PSScriptRoot "morphos_ui_harness.py"
python $pythonScriptPath --presentation $PresentationPath --mode autoscan

if ($LASTEXITCODE -ne 0) {
    throw "PowerPoint auto-scan verification failed."
}

Write-Host "PowerPoint auto-scan user flow passed." -ForegroundColor Green
