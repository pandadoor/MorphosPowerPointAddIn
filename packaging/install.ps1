param(
    [switch]$Silent,
    [switch]$AllUsers,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
$projectName = 'MorphosPowerPointAddIn'
$displayName = 'Morphos'
$addinDescription = 'Morphos PowerPoint cleanup workspace'

$installDir = if ($AllUsers) {
    Join-Path $env:ProgramFiles $projectName
} else {
    Join-Path $env:LOCALAPPDATA $projectName
}

$registryBase = if ($AllUsers) { "HKLM:\Software\Microsoft\Office\PowerPoint\Addins\$projectName" } else { "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$projectName" }

function Write-Info {
    param([string]$Message)
    if (-not $Silent) {
        Write-Host $Message -ForegroundColor Cyan
    }
}

function Show-Error {
    param([string]$Message)
    if ($Silent) {
        Write-Error $Message
    } else {
        [System.Windows.MessageBox]::Show($Message, "Morphos Setup Error", "OK", "Error")
    }
    exit 1
}

if ($Uninstall) {
    Write-Info "Uninstalling $displayName..."
    if (Test-Path $registryBase) {
        Remove-Item -Path $registryBase -Recurse -Force
    }
    if (Test-Path $installDir) {
        Remove-Item -Path $installDir -Recurse -Force
    }
    Write-Info "Uninstallation complete."
    exit 0
}

Write-Info "Starting $displayName installation..."

# 1. Dependency Checks
Write-Info "Verifying dependencies..."

# .NET 4.8 Check
$net48Path = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
if (Test-Path $net48Path) {
    $release = Get-ItemProperty -Path $net48Path -Name Release -ErrorAction SilentlyContinue
    if ($release.Release -lt 528040) {
        Show-Error "This add-in requires .NET Framework 4.8 or higher. Please install it before continuing."
    }
} else {
    Show-Error ".NET Framework 4.8 was not detected. Please install it before continuing."
}

# VSTO Runtime Check
$vstoPath = "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
if (-not (Test-Path $vstoPath)) {
    Show-Error "Visual Studio Tools for Office Runtime was not detected. Please install it before continuing."
}

# 2. Copy Files
Write-Info "Copying files to $installDir..."
if (-not (Test-Path $installDir)) {
    New-Item -Path $installDir -ItemType Directory -Force | Out-Null
}

# Copy files from current directory (where installer extracted them)
$filesToCopy = Get-ChildItem -Path $PSScriptRoot -Exclude "install.ps1", "uninstall.ps1", "*.sed", "*.exe"
foreach ($file in $filesToCopy) {
    Copy-Item -Path $file.FullName -Destination $installDir -Force -Recurse
}

# 3. Registry Registration
Write-Info "Registering add-in..."
$vstoManifestPath = Join-Path $installDir "$projectName.vsto"
$manifestUri = ([System.Uri]$vstoManifestPath).AbsoluteUri

if (-not (Test-Path $registryBase)) {
    New-Item -Path $registryBase -Force | Out-Null
}

New-ItemProperty -Path $registryBase -Name FriendlyName -Value $displayName -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registryBase -Name Description -Value $addinDescription -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registryBase -Name Manifest -Value $manifestUri -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registryBase -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null

# 4. Post-Install Verification
Write-Info "Verifying installation..."
$reg = Get-ItemProperty -Path $registryBase -ErrorAction SilentlyContinue
if ($reg.Manifest -eq $manifestUri) {
    Write-Info "Installation successful and verified."
} else {
    Show-Error "Installation failed: Registry state is inconsistent."
}

if (-not $Silent) {
    Write-Host ""
    Write-Host "Morphos has been successfully installed." -ForegroundColor Green
}
