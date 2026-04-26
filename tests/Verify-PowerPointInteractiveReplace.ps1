param(
    [string]$PresentationPath = "C:\! BASTA\pptx\TEST.pptx"
)

$ErrorActionPreference = "Stop"

$pythonScriptPath = Join-Path $PSScriptRoot "morphos_ui_harness.py"

foreach ($mode in @("autoscan", "open-font-dialog", "open-color-dialog")) {
    python $pythonScriptPath --presentation $PresentationPath --mode $mode
    if ($LASTEXITCODE -ne 0) {
        throw "PowerPoint interactive pane verification failed during mode '$mode'."
    }

    Stop-Process -Name powerpnt -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

Write-Host "PowerPoint interactive pane and dialog flows passed." -ForegroundColor Green
