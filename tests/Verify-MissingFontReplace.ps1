param(
    [string]$PresentationPath = "C:\! BASTA\pptx\TEST.pptx"
)

$ErrorActionPreference = "Stop"

$fixtureScriptPath = Join-Path $PSScriptRoot "create_missing_font_fixture.py"
$uiHarnessPath = Join-Path $PSScriptRoot "morphos_ui_harness.py"
$fixturePath = Join-Path ([System.IO.Path]::GetTempPath()) ("morphos-missing-font-" + [guid]::NewGuid().ToString("N") + ".pptx")

try {
    python $fixtureScriptPath --source $PresentationPath --output $fixturePath
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create the missing-font verification fixture."
    }

    python $uiHarnessPath --presentation $fixturePath --mode open-font-dialog
    if ($LASTEXITCODE -ne 0) {
        throw "Missing-font replace dialog verification failed."
    }

    Write-Host "Missing-font replace dialog opened successfully." -ForegroundColor Green
}
finally {
    Stop-Process -Name powerpnt -Force -ErrorAction SilentlyContinue
    if (Test-Path $fixturePath) {
        Remove-Item -LiteralPath $fixturePath -Force
    }
}
