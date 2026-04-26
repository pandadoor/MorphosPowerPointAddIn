param(
    [string]$PresentationPath = "C:\! BASTA\pptx\TEST.pptx"
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$binPath = Join-Path $projectRoot "bin\x64\Debug"
$assemblyPath = Join-Path $binPath "MorphosPowerPointAddIn.dll"
$openXmlPath = Join-Path $binPath "DocumentFormat.OpenXml.dll"
$openXmlFrameworkPath = Join-Path $binPath "DocumentFormat.OpenXml.Framework.dll"
$powerPointInteropPath = "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Microsoft.Office.Interop.PowerPoint.dll"
$officeInteropPath = "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Office.dll"
$fixturePath = Join-Path ([System.IO.Path]::GetTempPath()) ("morphos-service-font-" + [guid]::NewGuid().ToString("N") + ".pptx")

Add-Type -Path $openXmlPath
Add-Type -Path $openXmlFrameworkPath
Add-Type -Path $assemblyPath

$code = @"
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using MorphosPowerPointAddIn.Services;

public static class MorphosServiceFontReplaceHarness
{
    [STAThread]
    public static string[] Run(string path)
    {
        var errors = new System.Collections.Generic.List<string>();
        Application app = null;
        Presentation presentation = null;
        try
        {
            app = new Application();
            app.Visible = MsoTriState.msoTrue;
            presentation = app.Presentations.Open(path, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoTrue);
            if (presentation.Windows.Count > 0)
            {
                presentation.Windows[1].Activate();
            }

            var service = new PowerPointPresentationService(app);
            var scanBefore = service.AnalyzeActivePresentationAsync(null, null, CancellationToken.None).GetAwaiter().GetResult();
            var sourceFont = scanBefore.FontItems
                .Where(x => x != null && !string.IsNullOrWhiteSpace(x.FontName))
                .Select(x => x.FontName)
                .FirstOrDefault(name => string.Equals(name, "Berlin Sans FB", StringComparison.OrdinalIgnoreCase))
                ?? scanBefore.FontItems
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.FontName))
                    .Select(x => x.FontName)
                    .FirstOrDefault(name => !string.Equals(name, "Arial", StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(sourceFont))
            {
                errors.Add("Could not find a non-Arial source font in the test presentation.");
                return errors.ToArray();
            }

            var result = service.ReplaceFonts(new[] { sourceFont }, "Arial");
            var scanAfter = service.AnalyzeActivePresentationAsync(null, null, CancellationToken.None).GetAwaiter().GetResult();
            var stillPresent = scanAfter.FontItems.Any(
                x => x != null
                    && string.Equals(x.FontName, sourceFont, StringComparison.OrdinalIgnoreCase));

            if (stillPresent)
            {
                errors.Add("Source font still appeared after replacement: " + sourceFont);
            }

            if (result != null && result.HasWarnings)
            {
                errors.Add("Replacement finished with warnings: " + result.WarningMessage);
            }
        }
        catch (Exception ex)
        {
            errors.Add("Service font replace verification failed: " + ex.GetType().FullName + ": " + ex.Message);
        }
        finally
        {
            if (presentation != null)
            {
                try
                {
                    presentation.Close();
                }
                catch
                {
                }

                try
                {
                    Marshal.ReleaseComObject(presentation);
                }
                catch
                {
                }
            }

            if (app != null)
            {
                try
                {
                    app.Quit();
                }
                catch
                {
                }

                try
                {
                    Marshal.ReleaseComObject(app);
                }
                catch
                {
                }
            }
        }

        return errors.ToArray();
    }
}
"@

Add-Type -TypeDefinition $code -ReferencedAssemblies @(
    $officeInteropPath,
    $powerPointInteropPath,
    $assemblyPath
) -Language CSharp

try {
    Copy-Item -LiteralPath $PresentationPath -Destination $fixturePath
    $errors = [MorphosServiceFontReplaceHarness]::Run($fixturePath)
    if ($errors.Length -gt 0) {
        throw ($errors -join '; ')
    }

    Write-Host "Service font replacement passed on a live PowerPoint presentation." -ForegroundColor Green
}
finally {
    Stop-Process -Name powerpnt -Force -ErrorAction SilentlyContinue
    if (Test-Path $fixturePath) {
        Start-Sleep -Seconds 2
        Remove-Item -LiteralPath $fixturePath -Force -ErrorAction SilentlyContinue
    }
}
