[CmdletBinding()]
param(
    [switch]$Sign
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSCommandPath
if ($projectRoot -notlike "*packaging") {
    $projectRoot = Join-Path $projectRoot "packaging"
}
$repoRoot = Split-Path -Parent $projectRoot
$projectName = 'MorphosPowerPointAddIn'

# 1. Resolve MSBuild
Write-Host "Resolving MSBuild..." -ForegroundColor Cyan
$vswherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path $vswherePath)) {
    $vswherePath = "vswhere.exe"
}
$msbuildPath = & $vswherePath -latest -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1

if (-not $msbuildPath) {
    throw "MSBuild not found."
}
Write-Host "Using MSBuild: $msbuildPath"

# 2. Build Release
Write-Host "Building Release configuration..." -ForegroundColor Cyan
$csprojPath = Join-Path $repoRoot "$projectName.csproj"
& $msbuildPath $csprojPath /restore /t:Build /p:Configuration=Release /p:Platform=x64 /v:m /nologo

if ($LASTEXITCODE -ne 0) {
    throw "Build failed."
}

$releaseDir = Join-Path $repoRoot "bin\x64\Release"
if (-not (Test-Path $releaseDir)) {
    throw "Release directory not found at $releaseDir"
}

# 3. Prepare Staging Area
$stagingDir = Join-Path $projectRoot "staging"
if (Test-Path $stagingDir) { Remove-Item $stagingDir -Recurse -Force }
New-Item -Path $stagingDir -ItemType Directory -Force | Out-Null

$artifacts = @(
    "$projectName.dll",
    "$projectName.dll.manifest",
    "$projectName.vsto",
    "DocumentFormat.OpenXml.dll",
    "DocumentFormat.OpenXml.Framework.dll"
)

foreach ($file in $artifacts) {
    Copy-Item (Join-Path $releaseDir $file) $stagingDir -Force
}
Copy-Item (Join-Path $projectRoot "install.ps1") $stagingDir -Force

# 4. Generate IExpress SED File
Write-Host "Generating IExpress configuration..." -ForegroundColor Cyan
$setupExe = Join-Path $projectRoot "MorphosSetup.exe"
$sedPath = Join-Path $projectRoot "setup.sed"

$sedContent = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=0
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=1
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=
TargetName=$setupExe
FriendlyName=$projectName Installer
AppLaunched=powershell.exe -ExecutionPolicy Bypass -File install.ps1
PostInstallCmd=<None>
AdminQuietInstCmd=
UserQuietInstCmd=
SourceFiles=SourceFiles
[SourceFiles]
SourceFiles0=$stagingDir
[SourceFiles0]
%1=
"@

$filesSection = ""
$i = 0
foreach ($file in (Get-ChildItem $stagingDir)) {
    $filesSection += "FILE$i=$($file.Name)`r`n"
    $i++
}

$sedContent = $sedContent.Replace("%1=", $filesSection)
$sedContent | Out-File $sedPath -Encoding ASCII

# 5. Run IExpress
Write-Host "Creating self-extracting EXE..." -ForegroundColor Cyan
& "C:\Windows\System32\iexpress.exe" /n /q /m $sedPath

if (-not (Test-Path $setupExe)) {
    throw "Failed to create MorphosSetup.exe"
}

Write-Host "SUCCESS: Installer created at $setupExe" -ForegroundColor Green

# 6. Signing (Optional)
if ($Sign) {
    Write-Host "Signing is requested. Searching for signtool.exe..." -ForegroundColor Yellow
    # Note: User must provide PFX and Password. This is a template.
    Write-Host "Please use signtool.exe to sign the generated installer."
}
