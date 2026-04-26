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

public static class MorphosFontReplacementVerificationHarness
{
    private static readonly string[] RequiredThemeTargets = new[] { "+mj-lt", "+mn-lt" };

    [STAThread]
    public static string[] Run(string path)
    {
        var errors = new List<string>();
        errors.AddRange(VerifySavedPresentationTargets(path));
        return errors.ToArray();
    }

    private static string[] VerifySavedPresentationTargets(string path)
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
            var scan = service.AnalyzeActivePresentationAsync(null, null, CancellationToken.None).GetAwaiter().GetResult();
            var targetNames = new HashSet<string>(
                service.GetReplacementTargets()
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.NormalizedName))
                    .Select(x => x.NormalizedName),
                StringComparer.OrdinalIgnoreCase);

            var missingScannedFonts = scan.FontItems
                .Where(x => x != null
                    && !string.IsNullOrWhiteSpace(x.FontName)
                    && (x.IsInstalled || x.IsThemeFont))
                .Select(x => x.FontName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(font => !targetNames.Contains(font))
                .OrderBy(font => font, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var surfacedMissingFonts = scan.FontItems
                .Where(x => x != null
                    && !string.IsNullOrWhiteSpace(x.FontName)
                    && !x.IsInstalled
                    && !x.IsThemeFont)
                .Select(x => x.FontName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(font => targetNames.Contains(font))
                .OrderBy(font => font, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var missingThemeTargets = RequiredThemeTargets
                .Where(required => !targetNames.Contains(required))
                .ToArray();

            if (missingScannedFonts.Length > 0)
            {
                errors.Add("Replacement targets are missing installed scanned fonts: " + string.Join(", ", missingScannedFonts));
            }

            if (surfacedMissingFonts.Length > 0)
            {
                errors.Add("Replacement targets still surfaced missing scanned fonts: " + string.Join(", ", surfacedMissingFonts));
            }

            if (missingThemeTargets.Length > 0)
            {
                errors.Add("Replacement targets are missing theme fonts: " + string.Join(", ", missingThemeTargets));
            }
        }
        catch (Exception ex)
        {
            errors.Add("Saved presentation verification failed: " + ex.GetType().FullName + ": " + ex.Message);
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

$errors = [MorphosFontReplacementVerificationHarness]::Run($PresentationPath)

if ($errors.Length -gt 0) {
    throw ($errors -join '; ')
}

Write-Host "Replacement targets include installed scanned fonts and theme fonts for saved presentations." -ForegroundColor Green
