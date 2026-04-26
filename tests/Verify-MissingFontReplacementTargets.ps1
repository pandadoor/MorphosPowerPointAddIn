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
$fixtureScriptPath = Join-Path $PSScriptRoot "create_missing_font_fixture.py"
$fixturePath = Join-Path ([System.IO.Path]::GetTempPath()) ("morphos-missing-targets-" + [guid]::NewGuid().ToString("N") + ".pptx")

Add-Type -Path $openXmlPath
Add-Type -Path $openXmlFrameworkPath
Add-Type -Path $assemblyPath

$code = @"
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using MorphosPowerPointAddIn.Services;

public static class MorphosMissingFontReplacementTargetHarness
{
    private static readonly string[] MissingFonts = new[]
    {
        "A Morphos Missing Verification Font",
        "B Morphos Missing Verification Font"
    };

    [STAThread]
    public static string[] Run(string path)
    {
        var errors = new List<string>();
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
            service.AnalyzeActivePresentationAsync(null, null, CancellationToken.None).GetAwaiter().GetResult();
            var targetNames = new HashSet<string>(
                service.GetReplacementTargets()
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.NormalizedName))
                    .Select(x => x.NormalizedName),
                StringComparer.OrdinalIgnoreCase);

            var surfacedMissingFonts = MissingFonts
                .Where(targetNames.Contains)
                .ToArray();

            if (surfacedMissingFonts.Length > 0)
            {
                errors.Add("Replacement targets still surfaced missing fonts: " + string.Join(", ", surfacedMissingFonts));
            }

            if (!targetNames.Contains("+mj-lt") || !targetNames.Contains("+mn-lt"))
            {
                errors.Add("Replacement targets should still include Morphos theme font tokens.");
            }
        }
        catch (Exception ex)
        {
            errors.Add("Missing-font replacement target verification failed: " + ex.GetType().FullName + ": " + ex.Message);
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
    python $fixtureScriptPath --source $PresentationPath --output $fixturePath
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create the missing-font replacement-target fixture."
    }

    $errors = [MorphosMissingFontReplacementTargetHarness]::Run($fixturePath)
    if ($errors.Length -gt 0) {
        throw ($errors -join '; ')
    }

    Write-Host "Missing-font replacement targets exclude non-installed trap fonts." -ForegroundColor Green
}
finally {
    if (Test-Path $fixturePath) {
        Remove-Item -LiteralPath $fixturePath -Force
    }
}
