param(
    [string]$Mode = "HKCU"
)

$projectName = 'MorphosPowerPointAddIn'
$registryPath = if ($Mode -eq "HKLM") { "HKLM:\Software\Microsoft\Office\PowerPoint\Addins\$projectName" } else { "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$projectName" }

Write-Host "Verifying Morphos Installation (Mode: $Mode)..." -ForegroundColor Cyan

if (-not (Test-Path $registryPath)) {
    Write-Host "FAIL: Registry key not found at $registryPath" -ForegroundColor Red
    exit 1
}

$reg = Get-ItemProperty -Path $registryPath
$manifest = $reg.Manifest
$loadBehavior = $reg.LoadBehavior

Write-Host "Manifest: $manifest"
Write-Host "LoadBehavior: $loadBehavior"

if ($manifest -notlike "*.vsto*") {
    Write-Host "FAIL: Invalid manifest path." -ForegroundColor Red
    exit 1
}

if ($loadBehavior -ne 3) {
    Write-Host "WARNING: LoadBehavior is not 3 (Load at startup). Current value: $loadBehavior" -ForegroundColor Yellow
}

# Check files
$manifestPath = $manifest.Replace("file:///", "").Replace("/", "\")
if ($manifestPath.Contains("|")) {
    $manifestPath = $manifestPath.Split("|")[0]
}

if (Test-Path $manifestPath) {
    Write-Host "SUCCESS: Manifest file found at $manifestPath" -ForegroundColor Green
} else {
    Write-Host "FAIL: Manifest file not found at $manifestPath" -ForegroundColor Red
    exit 1
}

Write-Host "Installation Verification PASSED." -ForegroundColor Green
exit 0
