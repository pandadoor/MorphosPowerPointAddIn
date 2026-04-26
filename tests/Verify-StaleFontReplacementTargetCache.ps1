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
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using MorphosPowerPointAddIn.Models;
using MorphosPowerPointAddIn.Services;

public static class MorphosStaleReplacementTargetHarness
{
    private const string InvalidTargetName = "A Morphos Missing Verification Font";

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

            var serviceType = typeof(PowerPointPresentationService);
            var cacheField = serviceType.GetField("_fontScanSessionCache", BindingFlags.Instance | BindingFlags.NonPublic);
            if (cacheField == null)
            {
                errors.Add("Could not access Morphos font scan cache.");
                return errors.ToArray();
            }

            var cache = cacheField.GetValue(service);
            if (cache == null)
            {
                errors.Add("Morphos font scan cache was null.");
                return errors.ToArray();
            }

            var getSnapshotMethod = cache.GetType().GetMethod("GetOrCreateSnapshot", BindingFlags.Instance | BindingFlags.Public);
            var snapshot = getSnapshotMethod == null ? null : getSnapshotMethod.Invoke(cache, new object[] { presentation });
            if (snapshot == null)
            {
                errors.Add("Could not access Morphos scan snapshot.");
                return errors.ToArray();
            }

            var replacementTargetsProperty = snapshot.GetType().GetProperty("ReplacementTargets", BindingFlags.Instance | BindingFlags.Public);
            if (replacementTargetsProperty == null)
            {
                errors.Add("Could not access snapshot replacement targets.");
                return errors.ToArray();
            }

            replacementTargetsProperty.SetValue(
                snapshot,
                new[]
                {
                    new FontReplacementTarget
                    {
                        DisplayName = InvalidTargetName,
                        NormalizedName = InvalidTargetName,
                        IsInstalled = false,
                        IsPresentationFont = true,
                        IsThemeFont = false,
                        SortKey = 0
                    },
                    new FontReplacementTarget
                    {
                        DisplayName = "Theme Heading (+mj-lt)",
                        NormalizedName = "+mj-lt",
                        IsInstalled = true,
                        IsPresentationFont = false,
                        IsThemeFont = true,
                        SortKey = 1
                    }
                });

            var targets = service.GetReplacementTargets();
            var targetNames = new HashSet<string>(
                targets
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.NormalizedName))
                    .Select(x => x.NormalizedName),
                StringComparer.OrdinalIgnoreCase);

            if (targetNames.Contains(InvalidTargetName))
            {
                errors.Add("Stale replacement target cache still surfaced a non-installed font.");
            }

            if (!targetNames.Contains("+mj-lt"))
            {
                errors.Add("Theme heading token should survive cache sanitization.");
            }
        }
        catch (Exception ex)
        {
            errors.Add("Stale replacement-target verification failed: " + ex.GetType().FullName + ": " + ex.Message);
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

$errors = [MorphosStaleReplacementTargetHarness]::Run($PresentationPath)
if ($errors.Length -gt 0) {
    throw ($errors -join '; ')
}

Write-Host "Stale replacement-target cache is sanitized before Morphos shows the dialog." -ForegroundColor Green
