[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',

    [ValidateSet('x64')]
    [string]$Platform = 'x64',

    [switch]$NoStart,

    [switch]$SkipLoadVerification
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectName = 'MorphosPowerPointAddIn'
$legacyProjectName = 'SlidewisePowerPointAddIn'
$displayName = 'Morphos'
$addinDescription = 'Morphos PowerPoint cleanup workspace'
$projectRoot = Split-Path -Parent $PSCommandPath
$csprojPath = Join-Path $projectRoot ($projectName + '.csproj')
$outputDirectory = Join-Path $projectRoot ('bin\{0}\{1}' -f $Platform, $Configuration)
$assemblyPath = Join-Path $outputDirectory ($projectName + '.dll')
$applicationManifestPath = Join-Path $outputDirectory ($projectName + '.dll.manifest')
$deploymentManifestPath = Join-Path $outputDirectory ($projectName + '.vsto')
$registryPath = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\' + $projectName
$legacyRegistryPath = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\' + $legacyProjectName
$doNotDisableAddinListPath = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\DoNotDisableAddinList'
$buildStartedUtc = [DateTime]::UtcNow

function Write-Step {
    param([string]$Message)

    Write-Host ''
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Detail {
    param([string]$Message)

    Write-Host ('   ' + $Message) -ForegroundColor DarkGray
}

function Resolve-MSBuildPath {
    $vswherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path $vswherePath)) {
        throw 'vswhere.exe was not found. Install Visual Studio Build Tools or Visual Studio with the MSBuild workload.'
    }

    $resolvedPath = & $vswherePath `
        -latest `
        -products * `
        -requires Microsoft.Component.MSBuild `
        -find 'MSBuild\**\Bin\MSBuild.exe' |
        Select-Object -First 1

    if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
        throw 'MSBuild.exe was not found. Install the MSBuild component in Visual Studio.'
    }

    return $resolvedPath
}

function Stop-PowerPointProcesses {
    $runningProcesses = Get-Process -Name powerpnt -ErrorAction SilentlyContinue
    if (-not $runningProcesses) {
        Write-Detail 'PowerPoint is not currently running.'
        return
    }

    $processList = @($runningProcesses)
    Write-Detail ('Closing {0} running PowerPoint process(es).' -f $processList.Count)
    $processList | Stop-Process -Force
    Start-Sleep -Seconds 2
}

function Reset-BuildOutputs {
    $pathsToReset = @(
        $outputDirectory,
        (Join-Path $projectRoot 'obj')
    )

    foreach ($path in $pathsToReset) {
        if (-not (Test-Path $path)) {
            continue
        }

        Write-Detail ('Removing stale build output: ' + $path)
        Remove-PathWithRetry -Path $path
    }
}

function Remove-PathWithRetry {
    param(
        [string]$Path,
        [int]$MaxAttempts = 8
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force
            return
        }
        catch {
            if ($attempt -ge $MaxAttempts) {
                throw
            }

            Write-Detail ('Retrying cleanup for locked path (' + $attempt + '/' + $MaxAttempts + '): ' + $Path)
            Start-Sleep -Milliseconds (250 * $attempt)
        }
    }
}

function Assert-BuildArtifact {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        throw ('Expected build artifact was not created: ' + $Path)
    }

    $item = Get-Item $Path
    if ($item.LastWriteTimeUtc -lt $buildStartedUtc.AddSeconds(-5)) {
        throw ('Build artifact was not refreshed by this run: ' + $Path)
    }

    return $item
}

function Get-ManifestUri {
    param([string]$Path)

    $resolvedPath = (Resolve-Path $Path).Path
    return ([System.Uri]$resolvedPath).AbsoluteUri
}

function Get-DeploymentManifestInfo {
    param([string]$Path)

    [xml]$manifestXml = Get-Content -Raw -Path $Path
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($manifestXml.NameTable)
    $namespaceManager.AddNamespace('asmv2', 'urn:schemas-microsoft-com:asm.v2')
    $namespaceManager.AddNamespace('asmv1', 'urn:schemas-microsoft-com:asm.v1')
    $namespaceManager.AddNamespace('dsig', 'http://www.w3.org/2000/09/xmldsig#')

    $codebase = $manifestXml.SelectSingleNode('//asmv2:dependentAssembly/@codebase', $namespaceManager)
    $digest = $manifestXml.SelectSingleNode('//dsig:DigestValue', $namespaceManager)

    return [pscustomobject]@{
        DeploymentManifestUri = Get-ManifestUri -Path $Path
        ApplicationManifestCodebase = if ($codebase) { $codebase.Value } else { '' }
        DeploymentDigest = if ($digest) { $digest.InnerText } else { '' }
    }
}

function Ensure-AddinRegistryState {
    param([string]$ManifestUri)

    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    New-ItemProperty -Path $registryPath -Name FriendlyName -Value $displayName -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name Description -Value $addinDescription -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name Manifest -Value $ManifestUri -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null

    if (-not (Test-Path $doNotDisableAddinListPath)) {
        New-Item -Path $doNotDisableAddinListPath -Force | Out-Null
    }

    New-ItemProperty -Path $doNotDisableAddinListPath -Name $projectName -Value 1 -PropertyType DWord -Force | Out-Null
}

function Remove-LegacyAddinRegistration {
    if (Test-Path $legacyRegistryPath) {
        Write-Detail ('Removing legacy add-in registration: ' + $legacyRegistryPath)
        Remove-Item -LiteralPath $legacyRegistryPath -Recurse -Force
    }

    if (Test-Path $doNotDisableAddinListPath) {
        Remove-ItemProperty -Path $doNotDisableAddinListPath -Name $legacyProjectName -ErrorAction SilentlyContinue
    }
}

function Get-AddinRegistryState {
    if (-not (Test-Path $registryPath)) {
        return $null
    }

    return Get-ItemProperty -Path $registryPath
}

function Wait-ForPowerPointStartup {
    param([int]$TimeoutSeconds = 30)

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        try {
            return [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        }
        catch {
            Start-Sleep -Milliseconds 750
        }
    }
    while ((Get-Date) -lt $deadline)

    throw 'PowerPoint did not expose an active COM application within the timeout window.'
}

function Verify-ConnectedAddin {
    param(
        [string]$ProgId,
        [int]$TimeoutSeconds = 30
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    do {
        $application = $null
        $addin = $null

        try {
            $application = Wait-ForPowerPointStartup -TimeoutSeconds 5
            for ($index = 1; $index -le $application.COMAddIns.Count; $index++) {
                $candidate = $application.COMAddIns.Item($index)
                if ($candidate -and $candidate.ProgId -eq $ProgId) {
                    $addin = $candidate
                    break
                }
            }

            if ($addin -ne $null) {
                if (-not $addin.Connect) {
                    $addin.Connect = $true
                    Start-Sleep -Milliseconds 500
                }

                if ($addin.Connect) {
                    return [pscustomobject]@{
                        ProgId = $addin.ProgId
                        Description = $addin.Description
                        Connect = [bool]$addin.Connect
                    }
                }
            }
        }
        catch {
        }
        finally {
            if ($addin -ne $null) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($addin)
            }

            if ($application -ne $null) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($application)
            }

            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }

        Start-Sleep -Seconds 1
    }
    while ((Get-Date) -lt $deadline)

    throw ('PowerPoint did not load the "' + $ProgId + '" COM add-in within the timeout window.')
}

function Start-PowerPointAndVerifyAddin {
    param(
        [string]$ProgId,
        [string]$ManifestUri
    )

    $lastError = $null
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        Ensure-AddinRegistryState -ManifestUri $ManifestUri
        Write-Detail ('Launch attempt ' + $attempt + ' with local manifest registration.')
        Start-Process -FilePath 'powerpnt.exe' | Out-Null

        try {
            $addinState = Verify-ConnectedAddin -ProgId $ProgId
            Ensure-AddinRegistryState -ManifestUri $ManifestUri
            return $addinState
        }
        catch {
            $lastError = $_
            Write-Warning ('PowerPoint did not finish loading Morphos on attempt ' + $attempt + '. Retrying with a clean PowerPoint restart.')
            Stop-PowerPointProcesses
        }
    }

    if ($lastError) {
        throw $lastError
    }

    throw ('PowerPoint did not load the "' + $ProgId + '" COM add-in.')
}

if (-not (Test-Path $csprojPath)) {
    throw ('Project file not found: ' + $csprojPath)
}

Write-Step '1. Closing PowerPoint'
Stop-PowerPointProcesses

Write-Step '2. Cleaning stale build outputs'
Reset-BuildOutputs
Remove-LegacyAddinRegistration

Write-Step '3. Resolving toolchain'
$msbuildPath = Resolve-MSBuildPath
Write-Detail ('MSBuild: ' + $msbuildPath)

Write-Step ('4. Rebuilding ' + $projectName)
& $msbuildPath `
    $csprojPath `
    /restore `
    /t:Rebuild `
    ('/p:Configuration=' + $Configuration) `
    ('/p:Platform=' + $Platform) `
    /nologo `
    /verbosity:minimal

if ($LASTEXITCODE -ne 0) {
    throw ('MSBuild failed with exit code ' + $LASTEXITCODE + '.')
}

$assemblyItem = Assert-BuildArtifact -Path $assemblyPath
$applicationManifestItem = Assert-BuildArtifact -Path $applicationManifestPath
$deploymentManifestItem = Assert-BuildArtifact -Path $deploymentManifestPath
$deploymentManifestInfo = Get-DeploymentManifestInfo -Path $deploymentManifestPath
$assemblyHash = Get-FileHash -Path $assemblyPath -Algorithm SHA256

Write-Detail ('Assembly: ' + $assemblyItem.FullName)
Write-Detail ('Assembly SHA256: ' + $assemblyHash.Hash)
Write-Detail ('Application manifest: ' + $applicationManifestItem.FullName)
Write-Detail ('Deployment manifest: ' + $deploymentManifestItem.FullName)
Write-Detail ('Deployment manifest codebase: ' + $deploymentManifestInfo.ApplicationManifestCodebase)

Write-Step '5. Reapplying local add-in registration'
$localManifestUri = $deploymentManifestInfo.DeploymentManifestUri + '|vstolocal'
Ensure-AddinRegistryState -ManifestUri $localManifestUri

Write-Step '6. Normalizing PowerPoint registry state'
$registryState = Get-AddinRegistryState
Write-Detail ('Registry manifest: ' + $registryState.Manifest)
Write-Detail ('Registry load behavior: ' + $registryState.LoadBehavior)

if (-not $NoStart) {
    Write-Step '7. Starting PowerPoint'

    if (-not $SkipLoadVerification) {
        Write-Step '8. Verifying COM add-in load state'
        $addinState = Start-PowerPointAndVerifyAddin -ProgId $projectName -ManifestUri $localManifestUri
        Write-Detail ('COM add-in: ' + $addinState.ProgId)
        Write-Detail ('Connected: ' + $addinState.Connect)
    }
    else {
        Start-Process -FilePath 'powerpnt.exe' | Out-Null
    }
}
elseif (-not $SkipLoadVerification) {
    Write-Step '7. Skipping live load verification'
    Write-Detail 'PowerPoint was not started because -NoStart was supplied. Registry and manifest verification completed.'
}

Write-Host ''
Write-Host 'Morphos is rebuilt, reinstalled, and pointed at the latest manifest.' -ForegroundColor Green
