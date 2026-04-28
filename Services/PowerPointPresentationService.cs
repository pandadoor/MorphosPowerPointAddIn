using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MorphosPowerPointAddIn.Models;
using MorphosPowerPointAddIn.Utilities;
using WinForms = System.Windows.Forms;

namespace MorphosPowerPointAddIn.Services
{
    public sealed class PowerPointPresentationService
    {
        private const int RetryDelayMilliseconds = 150;
        private const int ReplacementTargetCacheVersion = 2;
        private readonly PowerPoint.Application _application;
        private readonly OpenXmlColorReplacer _openXmlColorReplacer;
        private readonly OpenXmlFontReplacer _openXmlFontReplacer;
        private readonly OpenXmlScanService _openXmlScanService;
        private readonly FontScanSessionCache _fontScanSessionCache;
        private bool _lastMutationReloadedPresentation;
        private HashSet<string> _cachedInstalledFonts;
        private DateTime _cachedInstalledFontsTimestamp;

        public PowerPointPresentationService(PowerPoint.Application application)
        {
            _application = application;
            _openXmlColorReplacer = new OpenXmlColorReplacer();
            _openXmlFontReplacer = new OpenXmlFontReplacer();
            _openXmlScanService = new OpenXmlScanService();
            _fontScanSessionCache = new FontScanSessionCache(
                () => GetInstalledFontsSet(),
                CaptureScanSnapshot,
                RefreshPackageMetadata);
        }

        private HashSet<string> GetInstalledFontsSet()
        {
            var now = DateTime.UtcNow;
            if (_cachedInstalledFonts != null && (now - _cachedInstalledFontsTimestamp).TotalSeconds < 30)
            {
                return _cachedInstalledFonts;
            }

            _cachedInstalledFonts = new HashSet<string>(SystemFontRegistry.GetInstalledFontNames(), StringComparer.OrdinalIgnoreCase);
            _cachedInstalledFontsTimestamp = now;
            return _cachedInstalledFonts;
        }

        internal bool LastMutationReloadedPresentation => _lastMutationReloadedPresentation;

        public IReadOnlyList<FontReplacementTarget> GetReplacementTargets(IEnumerable<string> sourceFontNames = null)
        {
            var presentation = GetActivePresentation();
            if (presentation != null)
            {
                var snapshot = _fontScanSessionCache.GetOrCreateSnapshot(presentation);
                EnsureReplacementTargets(snapshot);
                return FilterReplacementTargets(snapshot.ReplacementTargets, sourceFontNames);
            }

            return Array.Empty<FontReplacementTarget>();
        }

        public async Task<PresentationScanResult> AnalyzeActivePresentationAsync(
            IProgress<ScanProgressInfo> fontProgress,
            IProgress<ScanProgressInfo> colorProgress,
            CancellationToken cancellationToken)
        {
            var presentation = GetActivePresentation();
            if (presentation == null)
            {
                fontProgress?.Report(new ScanProgressInfo { Message = "Open a presentation to scan fonts." });
                colorProgress?.Report(new ScanProgressInfo { Message = "Open a presentation to scan colors." });
                return new PresentationScanResult();
            }

            var installedFonts = GetInstalledFontsSet();
            var snapshot = _fontScanSessionCache.GetOrCreateSnapshot(presentation);
            var cachedAnalysisResult = TryGetCachedPresentationScanResult(snapshot);
            if (cachedAnalysisResult != null)
            {
                fontProgress?.Report(new ScanProgressInfo { CompletedItems = 1, TotalItems = 1, Message = "Using cached font inventory." });
                colorProgress?.Report(new ScanProgressInfo { CompletedItems = 1, TotalItems = 1, Message = "Using cached color inventory." });
                return cachedAnalysisResult;
            }

            IDictionary<string, FontAccumulator> accumulators = null;
            var usedPackageScan = false;
            var colorScanResult = new ColorScanResult();

            using (var packageScanSource = AcquirePackageScanSource(presentation, snapshot))
            {
                if (packageScanSource != null)
                {
                    var packageScanResult = await Task.Run(
                        () => _openXmlScanService.ScanPackage(packageScanSource.FilePath, fontProgress, colorProgress, cancellationToken),
                        cancellationToken).ConfigureAwait(true);

                    if (packageScanResult != null)
                    {
                        RefreshThemeFontNames(snapshot, packageScanSource.FilePath);
                        colorScanResult = packageScanResult.ColorScanResult ?? new ColorScanResult();
                        accumulators = BuildAccumulatorsFromPackage(packageScanResult.FontUsages);
                        usedPackageScan = true;
                    }
                }
                else
                {
                    colorProgress?.Report(new ScanProgressInfo { Message = "Morphos could not create a color scan copy of the presentation." });
                }

                if (accumulators == null)
                {
                    fontProgress?.Report(new ScanProgressInfo
                    {
                        Message = string.IsNullOrWhiteSpace(snapshot.FilePath)
                            ? "Presentation is not saved yet. Running live PowerPoint scan."
                            : !snapshot.IsSaved
                                ? "Presentation has unsaved changes. Running live PowerPoint scan."
                                : "Package indexing was unavailable. Running live PowerPoint scan."
                    });

                    accumulators = new Dictionary<string, FontAccumulator>(StringComparer.OrdinalIgnoreCase);
                    var scopes = BuildScanScopes(presentation).ToList();
                    for (var i = 0; i < scopes.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        var scope = scopes[i];
                        fontProgress?.Report(new ScanProgressInfo
                        {
                            CompletedItems = i,
                            TotalItems = scopes.Count,
                            Message = "Scanning " + scope.Label
                        });

                        ScanShapes(scope.Shapes, scope.Scope, scope.Index, scope.Label, accumulators);

                        fontProgress?.Report(new ScanProgressInfo
                        {
                            CompletedItems = i + 1,
                            TotalItems = scopes.Count,
                            Message = "Scanned " + scope.Label
                        });

                        await Task.Yield();
                    }
                }
            }

            ApplyPresentationFontMetadata(snapshot, accumulators, !usedPackageScan);

            if (usedPackageScan && ShouldCaptureDisplayedFontMap(snapshot, accumulators))
            {
                ApplySubstitutionStateFromPowerPoint(accumulators, CaptureDisplayedFontMap(presentation));
            }

            var saveValidation = ShouldPerformEmbeddedSaveValidation(snapshot, accumulators)
                ? CreateEmbeddedSaveValidation(presentation)
                : null;
            ApplyValidatedEmbeddingState(snapshot, accumulators, saveValidation);
            ApplyWarningState(accumulators, saveValidation);

            var fontItems = accumulators.Values
                .OrderByDescending(x => x.UsesCount)
                .ThenBy(x => x.FontName, StringComparer.OrdinalIgnoreCase)
                .Select(x => x.ToInventoryItem(installedFonts))
                .ToList();

            CacheReplacementTargetsFromScan(snapshot, fontItems);
            var result = new PresentationScanResult
            {
                FontItems = fontItems,
                ColorScanResult = colorScanResult ?? new ColorScanResult()
            };

            CachePresentationScanResult(snapshot, result);
            return ClonePresentationScanResult(result);
        }

        public async Task<IReadOnlyList<FontInventoryItem>> ScanActivePresentationAsync(IProgress<ScanProgressInfo> progress, CancellationToken cancellationToken)
        {
            var presentation = GetActivePresentation();
            if (presentation == null)
            {
                progress?.Report(new ScanProgressInfo { Message = "Open a presentation to scan fonts." });
                return Array.Empty<FontInventoryItem>();
            }

            var installedFonts = GetInstalledFontsSet();
            var snapshot = _fontScanSessionCache.GetOrCreateSnapshot(presentation);
            var cachedFontItems = TryGetCachedFontItems(snapshot);
            if (cachedFontItems != null)
            {
                progress?.Report(new ScanProgressInfo { CompletedItems = 1, TotalItems = 1, Message = "Using cached font inventory." });
                return cachedFontItems;
            }

            IDictionary<string, FontAccumulator> accumulators = null;
            var usedPackageScan = false;

            using (var packageScanSource = AcquirePackageScanSource(presentation, snapshot))
            {
                if (packageScanSource != null)
                {
                    progress?.Report(new ScanProgressInfo { Message = "Indexing presentation package..." });
                    var packageUsages = await Task.Run(
                        () => PresentationPackageInspector.ReadFontUsages(packageScanSource.FilePath, progress, cancellationToken),
                        cancellationToken).ConfigureAwait(true);

                    if (packageUsages != null)
                    {
                        RefreshThemeFontNames(snapshot, packageScanSource.FilePath);
                        accumulators = BuildAccumulatorsFromPackage(packageUsages);
                        usedPackageScan = true;
                    }
                }

                if (accumulators == null)
                {
                    progress?.Report(new ScanProgressInfo
                    {
                        Message = string.IsNullOrWhiteSpace(snapshot.FilePath)
                            ? "Presentation is not saved yet. Running live PowerPoint scan."
                            : !snapshot.IsSaved
                                ? "Presentation has unsaved changes. Running live PowerPoint scan."
                                : "Package indexing was unavailable. Running live PowerPoint scan."
                    });

                    accumulators = new Dictionary<string, FontAccumulator>(StringComparer.OrdinalIgnoreCase);
                    var scopes = BuildScanScopes(presentation).ToList();
                    for (var i = 0; i < scopes.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        var scope = scopes[i];
                        progress?.Report(new ScanProgressInfo
                        {
                            CompletedItems = i,
                            TotalItems = scopes.Count,
                            Message = "Scanning " + scope.Label
                        });

                        ScanShapes(scope.Shapes, scope.Scope, scope.Index, scope.Label, accumulators);

                        progress?.Report(new ScanProgressInfo
                        {
                            CompletedItems = i + 1,
                            TotalItems = scopes.Count,
                            Message = "Scanned " + scope.Label
                        });

                        await Task.Yield();
                    }
                }
            }

            ApplyPresentationFontMetadata(snapshot, accumulators, !usedPackageScan);

            if (usedPackageScan && ShouldCaptureDisplayedFontMap(snapshot, accumulators))
            {
                var currentDisplayMap = CaptureDisplayedFontMap(presentation);
                ApplySubstitutionStateFromPowerPoint(accumulators, currentDisplayMap);
            }

            var saveValidation = ShouldPerformEmbeddedSaveValidation(snapshot, accumulators)
                ? CreateEmbeddedSaveValidation(presentation)
                : null;
            ApplyValidatedEmbeddingState(snapshot, accumulators, saveValidation);
            ApplyWarningState(accumulators, saveValidation);

            var fontItems = accumulators.Values
                .OrderByDescending(x => x.UsesCount)
                .ThenBy(x => x.FontName, StringComparer.OrdinalIgnoreCase)
                .Select(x => x.ToInventoryItem(installedFonts))
                .ToList();

            CacheReplacementTargetsFromScan(snapshot, fontItems);
            CacheFontItems(snapshot, fontItems);
            return CloneFontItems(fontItems);
        }

        public async Task<ColorScanResult> ScanActivePresentationColorsAsync(IProgress<ScanProgressInfo> progress, CancellationToken cancellationToken)
        {
            var presentation = GetActivePresentation();
            if (presentation == null)
            {
                progress?.Report(new ScanProgressInfo { Message = "Open a presentation to scan colors." });
                return new ColorScanResult();
            }

            var snapshot = _fontScanSessionCache.GetOrCreateSnapshot(presentation);
            var cachedColorScanResult = TryGetCachedColorScanResult(snapshot);
            if (cachedColorScanResult != null)
            {
                progress?.Report(new ScanProgressInfo { CompletedItems = 1, TotalItems = 1, Message = "Using cached color inventory." });
                return cachedColorScanResult;
            }

            using (var packageScanSource = AcquireColorScanSource(presentation))
            {
                if (packageScanSource == null)
                {
                    progress?.Report(new ScanProgressInfo { Message = "Morphos could not create a color scan copy of the presentation." });
                    return new ColorScanResult();
                }

                var scanResult = await Task.Run(
                    () => PresentationColorInspector.ReadDirectColorUsages(packageScanSource.FilePath, progress, cancellationToken),
                    cancellationToken).ConfigureAwait(true);
                CacheColorScanResult(snapshot, scanResult);
                return CloneColorScanResult(scanResult);
            }
        }

        public void ShowUsage(FontUsageLocation location)
        {
            if (location == null)
            {
                return;
            }

            var presentation = GetActivePresentation();
            if (presentation == null)
            {
                return;
            }

            try
            {
                if (location.Scope != PresentationScope.Slide || !location.SlideIndex.HasValue)
                {
                    return;
                }

                var activeWindow = _application.ActiveWindow;
                if (activeWindow == null)
                {
                    return;
                }

                var slide = presentation.Slides[location.SlideIndex.Value];
                var navigated = false;

                try
                {
                    activeWindow.ViewType = PowerPoint.PpViewType.ppViewNormal;
                }
                catch
                {
                }

                try
                {
                    activeWindow.View.GotoSlide(location.SlideIndex.Value);
                    navigated = true;
                }
                catch
                {
                }

                if (!navigated)
                {
                    try
                    {
                        slide.Select();
                        navigated = true;
                    }
                    catch
                    {
                    }
                }

                try
                {
                    if (activeWindow.View.Zoom < 70)
                    {
                        activeWindow.View.Zoom = 70;
                    }
                }
                catch
                {
                }

                var shape = FindShape(slide.Shapes, location.ShapeId, location.ShapeName);
                if (shape != null)
                {
                    try
                    {
                        shape.Select(MsoTriState.msoTrue);
                        return;
                    }
                    catch
                    {
                    }

                    try
                    {
                        slide.Shapes.Range(new[] { shape.Name }).Select();
                        return;
                    }
                    catch
                    {
                    }
                }

                if (!navigated)
                {
                    try
                    {
                        slide.Select();
                        navigated = true;
                    }
                    catch
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                _ = ex;
            }
        }

        public FontReplacementResult ReplaceFont(string oldFontName, string newFontName)
        {
            return ReplaceFonts(new[] { oldFontName }, newFontName);
        }

        public FontReplacementResult ReplaceFonts(IEnumerable<string> oldFontNames, string newFontName)
        {
            _lastMutationReloadedPresentation = false;
            var result = new FontReplacementResult();
            if (oldFontNames == null || string.IsNullOrWhiteSpace(newFontName))
            {
                return result;
            }

            var presentation = GetActivePresentation();
            if (presentation == null)
            {
                return result;
            }

            var sourceFonts = new HashSet<string>(
                oldFontNames
                    .Select(FontNameNormalizer.Normalize)
                    .Where(x => !string.IsNullOrWhiteSpace(x)),
                StringComparer.OrdinalIgnoreCase);

            if (sourceFonts.Count == 0)
            {
                return result;
            }

            var replacement = FontNameNormalizer.NormalizeReplacementFont(newFontName);
            if (!IsThemeFontName(replacement) && !SystemFontRegistry.SystemFontExists(replacement))
            {
                throw new InvalidOperationException("Install the replacement font on Windows before replacing it in PowerPoint.");
            }

            var snapshot = _fontScanSessionCache.GetOrCreateSnapshot(presentation);
            var packageReplacementCount = !ShouldPreferLiveFontReplacement(presentation, sourceFonts)
                && CanUsePackageMutation(snapshot, presentation)
                ? TryApplyOpenXmlFontReplacementsSafely(presentation, sourceFonts, replacement)
                : null;
            if (packageReplacementCount.HasValue)
            {
                _fontScanSessionCache.Invalidate();
                return ValidateEmbeddedSaveState(GetActivePresentation() ?? presentation);
            }

            foreach (var oldFontName in sourceFonts)
            {
                try
                {
                    dynamic fonts = presentation.Fonts;
                    fonts.Replace(oldFontName, replacement);
                }
                catch
                {
                }
            }

            ApplyToAllPresentationShapes(presentation, shape => ReplaceFontInShape(shape, sourceFonts, replacement));
            ApplyToPresentationTextStyles(presentation, sourceFonts, replacement);
            TryMarkPresentationDirty(presentation);
            _fontScanSessionCache.Invalidate();
            return ValidateEmbeddedSaveState(presentation);
        }

        public ColorReplacementSummary PreviewColorReplacements(IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            var presentation = GetActivePresentation();
            if (presentation == null)
            {
                return new ColorReplacementSummary();
            }

            var sanitizedInstructions = SanitizeColorInstructions(instructions);
            if (sanitizedInstructions.Count == 0)
            {
                return new ColorReplacementSummary();
            }

            using (var packageScanSource = CreateTransientColorCopy(presentation))
            {
                if (packageScanSource == null)
                {
                    return new ColorReplacementSummary();
                }

                var replacementCount = PresentationColorInspector.ApplyColorReplacements(packageScanSource.FilePath, sanitizedInstructions);
                var scanResult = PresentationColorInspector.ReadDirectColorUsages(packageScanSource.FilePath, null, CancellationToken.None);
                return BuildColorReplacementSummary(scanResult, replacementCount, false);
            }
        }

        public ColorReplacementSummary ReplaceColors(IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            _lastMutationReloadedPresentation = false;
            var presentation = GetActivePresentation();
            if (presentation == null)
            {
                return new ColorReplacementSummary();
            }

            var sanitizedInstructions = SanitizeColorInstructions(instructions);
            if (sanitizedInstructions.Count == 0)
            {
                return new ColorReplacementSummary();
            }

            var packageReplacementCount = ShouldPreferLiveColorReplacement(sanitizedInstructions)
                || !CanUsePackageMutation(_fontScanSessionCache.GetOrCreateSnapshot(presentation), presentation)
                ? null
                : TryApplyOpenXmlColorReplacementsSafely(presentation, sanitizedInstructions);
            var replacementCount = packageReplacementCount ?? ApplyColorReplacementsToPresentation(presentation, sanitizedInstructions);
            if (replacementCount > 0 && !packageReplacementCount.HasValue)
            {
                TryMarkPresentationDirty(presentation);
            }

            if (replacementCount > 0)
            {
                _fontScanSessionCache.Invalidate();
            }

            return new ColorReplacementSummary
            {
                ReplacementCount = replacementCount,
                Applied = replacementCount > 0
            };
        }

        public void UpdateEmbedding(FontEmbeddingStatus status)
        {
            _lastMutationReloadedPresentation = false;
            var presentation = GetActivePresentation();
            if (presentation == null)
            {
                return;
            }

            if (!EnsurePresentationSavedToDisk(presentation, out var filePath))
            {
                throw new InvalidOperationException("Save the presentation before changing embedding settings.");
            }

            var embedFonts = status != FontEmbeddingStatus.No;
            var saveSubsetFonts = status == FontEmbeddingStatus.Subset;

            try
            {
                dynamic dynamicPresentation = presentation;
                dynamicPresentation.SaveAs(filePath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, embedFonts ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                PresentationPackageInspector.TrySetFontEmbeddingFlags(filePath, embedFonts, saveSubsetFonts);
                _fontScanSessionCache.Invalidate();
            }
            catch (Exception ex)
            {
                ErrorReporter.Show("Morphos could not update PowerPoint font embedding.", ex);
            }
        }

        private PowerPoint.Presentation GetActivePresentation()
        {
            try
            {
                return _application.ActivePresentation;
            }
            catch
            {
                return null;
            }
        }

        private FontScanSnapshot CaptureScanSnapshot(PowerPoint.Presentation presentation, ISet<string> installedFonts)
        {
            var snapshot = new FontScanSnapshot();
            snapshot.FilePath = TryGetPresentationPath(presentation);
            snapshot.IsSaved = TryIsPresentationSaved(presentation);
            snapshot.CanUsePackageScan = snapshot.IsSaved && !string.IsNullOrWhiteSpace(snapshot.FilePath);

            try
            {
                dynamic fonts = presentation.Fonts;
                var count = fonts.Count;

                for (var i = 1; i <= count; i++)
                {
                    dynamic font;
                    try
                    {
                        font = fonts[i];
                    }
                    catch
                    {
                        continue;
                    }

                    string fontName;
                    if (!TryReadStringProperty(font, "Name", out fontName))
                    {
                        continue;
                    }

                    fontName = FontNameNormalizer.Normalize(fontName);
                    if (string.IsNullOrWhiteSpace(fontName))
                    {
                        continue;
                    }

                    var isThemeFont = IsThemeFontName(fontName);
                    var isEmbeddable = false;
                    var hasEmbeddableMetadata = isThemeFont || TryReadBooleanProperty(font, "Embeddable", out isEmbeddable);

                    var isEmbedded = false;
                    var hasEmbeddedMetadata = isThemeFont || TryReadBooleanProperty(font, "Embedded", out isEmbedded);

                    snapshot.Fonts.Add(new PresentationFontMetadata
                    {
                        FontName = fontName,
                        IsInstalled = isThemeFont || (installedFonts != null && installedFonts.Contains(fontName)),
                        IsEmbeddable = isThemeFont || isEmbeddable,
                        HasEmbeddableMetadata = hasEmbeddableMetadata,
                        IsEmbedded = isThemeFont
                            ? snapshot.RequestsEmbeddedFonts && snapshot.EmbeddedFontNames.Contains(fontName)
                            : isEmbedded,
                        HasEmbeddedMetadata = hasEmbeddedMetadata
                    });
                }
            }
            catch
            {
            }

            return snapshot;
        }

        private static void RefreshPackageMetadata(FontScanSnapshot snapshot, string filePath)
        {
            if (snapshot == null)
            {
                return;
            }

            snapshot.HasPackageEmbeddingSettings = false;
            snapshot.RequestsEmbeddedFonts = false;
            snapshot.SaveSubsetFonts = false;
            snapshot.HasEmbeddedFontDataKnown = false;
            snapshot.HasEmbeddedFontData = false;
            snapshot.EmbeddedFontNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            snapshot.ThemeFontNames = Array.Empty<string>();
            snapshot.ReplacementTargets = Array.Empty<FontReplacementTarget>();

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            bool embedFonts;
            bool saveSubsetFonts;
            snapshot.HasPackageEmbeddingSettings = PresentationPackageInspector.TryGetFontEmbeddingSettings(filePath, out embedFonts, out saveSubsetFonts);
            snapshot.RequestsEmbeddedFonts = snapshot.HasPackageEmbeddingSettings && embedFonts;
            snapshot.SaveSubsetFonts = snapshot.RequestsEmbeddedFonts && saveSubsetFonts;

            ISet<string> embeddedFontNames;
            snapshot.HasEmbeddedFontDataKnown = PresentationPackageInspector.TryGetEmbeddedFontNames(filePath, out embeddedFontNames);
            snapshot.EmbeddedFontNames = embeddedFontNames ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            snapshot.HasEmbeddedFontData = snapshot.EmbeddedFontNames.Count > 0;

            bool hasEmbeddedFontPackageData;
            if (PresentationPackageInspector.TryHasEmbeddedFontData(filePath, out hasEmbeddedFontPackageData))
            {
                snapshot.HasEmbeddedFontData = hasEmbeddedFontPackageData;
            }

            IReadOnlyList<string> themeFontNames;
            if (PresentationPackageInspector.TryGetThemeFontNames(filePath, out themeFontNames))
            {
                snapshot.ThemeFontNames = themeFontNames ?? Array.Empty<string>();
            }
        }

        private static void EnsureReplacementTargets(FontScanSnapshot snapshot)
        {
            if (snapshot == null)
            {
                return;
            }

            var installedFontNames = SystemFontRegistry.GetInstalledFontNames();
            if (HasCurrentReplacementTargets(snapshot, installedFontNames))
            {
                return;
            }

            var powerPointFontNames = (snapshot.CachedScannedFontNames != null && snapshot.CachedScannedFontNames.Count > 0
                    ? snapshot.CachedScannedFontNames
                    : snapshot.Fonts
                        .Where(font => font != null && !string.IsNullOrWhiteSpace(font.FontName))
                        .Select(font => font.FontName)
                        .ToList())
                .ToList();

            snapshot.ReplacementTargets = FontReplacementTargetBuilder.Build(
                powerPointFontNames,
                installedFontNames,
                Array.Empty<string>(),
                snapshot.ThemeFontNames ?? Array.Empty<string>());
            snapshot.ReplacementTargetsVersion = ReplacementTargetCacheVersion;
        }

        private static void RefreshThemeFontNames(FontScanSnapshot snapshot, string filePath)
        {
            if (snapshot == null || string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            IReadOnlyList<string> themeFontNames;
            if (PresentationPackageInspector.TryGetThemeFontNames(filePath, out themeFontNames))
            {
                snapshot.ThemeFontNames = themeFontNames ?? Array.Empty<string>();
            }
        }

        private static PresentationScanResult TryGetCachedPresentationScanResult(FontScanSnapshot snapshot)
        {
            return CanReuseCachedScanResult(snapshot) && snapshot.CachedPresentationScanResult != null
                ? ClonePresentationScanResult(snapshot.CachedPresentationScanResult)
                : null;
        }

        private static IReadOnlyList<FontInventoryItem> TryGetCachedFontItems(FontScanSnapshot snapshot)
        {
            return CanReuseCachedScanResult(snapshot) && snapshot.CachedFontItems != null
                ? CloneFontItems(snapshot.CachedFontItems)
                : null;
        }

        private static ColorScanResult TryGetCachedColorScanResult(FontScanSnapshot snapshot)
        {
            return CanReuseCachedScanResult(snapshot) && snapshot.CachedColorScanResult != null
                ? CloneColorScanResult(snapshot.CachedColorScanResult)
                : null;
        }

        private static void CachePresentationScanResult(FontScanSnapshot snapshot, PresentationScanResult result)
        {
            if (!CanReuseCachedScanResult(snapshot) || result == null)
            {
                return;
            }

            snapshot.CachedPresentationScanResult = ClonePresentationScanResult(result);
            snapshot.CachedFontItems = CloneFontItems(result.FontItems);
            snapshot.CachedColorScanResult = CloneColorScanResult(result.ColorScanResult);
        }

        private static void CacheFontItems(FontScanSnapshot snapshot, IReadOnlyList<FontInventoryItem> fontItems)
        {
            if (!CanReuseCachedScanResult(snapshot))
            {
                return;
            }

            snapshot.CachedFontItems = CloneFontItems(fontItems);
            if (snapshot.CachedPresentationScanResult != null)
            {
                snapshot.CachedPresentationScanResult = new PresentationScanResult
                {
                    FontItems = CloneFontItems(fontItems),
                    ColorScanResult = CloneColorScanResult(snapshot.CachedPresentationScanResult.ColorScanResult)
                };
            }
        }

        private static void CacheColorScanResult(FontScanSnapshot snapshot, ColorScanResult scanResult)
        {
            if (!CanReuseCachedScanResult(snapshot))
            {
                return;
            }

            snapshot.CachedColorScanResult = CloneColorScanResult(scanResult);
            if (snapshot.CachedPresentationScanResult != null)
            {
                snapshot.CachedPresentationScanResult = new PresentationScanResult
                {
                    FontItems = CloneFontItems(snapshot.CachedPresentationScanResult.FontItems),
                    ColorScanResult = CloneColorScanResult(scanResult)
                };
            }
        }

        private static bool CanReuseCachedScanResult(FontScanSnapshot snapshot)
        {
            return snapshot != null
                && snapshot.IsSaved
                && !string.IsNullOrWhiteSpace(snapshot.FilePath)
                && File.Exists(snapshot.FilePath);
        }

        private static void CacheReplacementTargetsFromScan(FontScanSnapshot snapshot, IReadOnlyList<FontInventoryItem> fontItems)
        {
            if (snapshot == null)
            {
                return;
            }

            var scannedFontNames = (fontItems ?? Array.Empty<FontInventoryItem>())
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.FontName))
                .Select(item => item.FontName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            snapshot.CachedScannedFontNames = scannedFontNames;
            snapshot.ReplacementTargets = FontReplacementTargetBuilder.Build(
                scannedFontNames,
                SystemFontRegistry.GetInstalledFontNames(),
                Array.Empty<string>(),
                snapshot.ThemeFontNames ?? Array.Empty<string>());
            snapshot.ReplacementTargetsVersion = ReplacementTargetCacheVersion;
        }

        private static bool HasCurrentReplacementTargets(
            FontScanSnapshot snapshot,
            IReadOnlyCollection<string> installedFontNames)
        {
            if (snapshot == null
                || snapshot.ReplacementTargetsVersion != ReplacementTargetCacheVersion
                || snapshot.ReplacementTargets == null
                || snapshot.ReplacementTargets.Count == 0)
            {
                return false;
            }

            var installedLookup = new HashSet<string>(
                (installedFontNames ?? Array.Empty<string>())
                    .Select(FontNameNormalizer.NormalizeReplacementFont)
                    .Where(name => !string.IsNullOrWhiteSpace(name)),
                StringComparer.OrdinalIgnoreCase);

            var themeLookup = new HashSet<string>(
                (snapshot.ThemeFontNames ?? Array.Empty<string>())
                    .Select(FontNameNormalizer.NormalizeReplacementFont)
                    .Where(name => !string.IsNullOrWhiteSpace(name)),
                StringComparer.OrdinalIgnoreCase);

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var target in snapshot.ReplacementTargets)
            {
                var normalizedName = FontNameNormalizer.NormalizeReplacementFont(
                    target == null ? string.Empty : target.NormalizedName ?? target.DisplayName);
                if (string.IsNullOrWhiteSpace(normalizedName) || !seen.Add(normalizedName))
                {
                    return false;
                }

                if (themeLookup.Contains(normalizedName))
                {
                    continue;
                }

                if (!installedLookup.Contains(normalizedName))
                {
                    return false;
                }
            }

            return true;
        }

        private static IReadOnlyList<FontReplacementTarget> FilterReplacementTargets(
            IReadOnlyList<FontReplacementTarget> targets,
            IEnumerable<string> sourceFontNames)
        {
            if (targets == null || targets.Count == 0)
            {
                return Array.Empty<FontReplacementTarget>();
            }

            var sourceLookup = new HashSet<string>(
                (sourceFontNames ?? Array.Empty<string>())
                    .Select(FontNameNormalizer.NormalizeReplacementFont)
                    .Where(name => !string.IsNullOrWhiteSpace(name)),
                StringComparer.OrdinalIgnoreCase);

            if (sourceLookup.Count == 0)
            {
                return targets;
            }

            return targets
                .Where(target => target != null && !sourceLookup.Contains(target.NormalizedName))
                .ToList();
        }

        private static IDictionary<string, FontAccumulator> BuildAccumulatorsFromPackage(IReadOnlyList<PackageFontUsageRecord> usages)
        {
            var accumulators = new Dictionary<string, FontAccumulator>(StringComparer.OrdinalIgnoreCase);
            if (usages == null)
            {
                return accumulators;
            }

            foreach (var usage in usages)
            {
                if (usage == null || usage.Location == null)
                {
                    continue;
                }

                AddUsage(accumulators, usage.FontName, usage.Location);
            }

            return accumulators;
        }

        private void ScanShapes(
            PowerPoint.Shapes shapes,
            PresentationScope scope,
            int? slideIndex,
            string scopeLabel,
            IDictionary<string, FontAccumulator> accumulators)
        {
            if (shapes == null)
            {
                return;
            }

            for (var i = 1; i <= shapes.Count; i++)
            {
                ScanShape(shapes[i], scope, slideIndex, scopeLabel, accumulators);
            }
        }

        private void ScanShape(
            PowerPoint.Shape shape,
            PresentationScope scope,
            int? slideIndex,
            string scopeLabel,
            IDictionary<string, FontAccumulator> accumulators)
        {
            if (shape == null)
            {
                return;
            }

            HarvestText(shape, scope, slideIndex, scopeLabel, accumulators);

            try
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    for (var i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        ScanShape(shape.GroupItems[i], scope, slideIndex, scopeLabel, accumulators);
                    }
                }
            }
            catch
            {
            }

            try
            {
                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    for (var row = 1; row <= shape.Table.Rows.Count; row++)
                    {
                        for (var column = 1; column <= shape.Table.Columns.Count; column++)
                        {
                            HarvestText(shape.Table.Cell(row, column).Shape, scope, slideIndex, scopeLabel, accumulators);
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void HarvestText(
            PowerPoint.Shape shape,
            PresentationScope scope,
            int? slideIndex,
            string scopeLabel,
            IDictionary<string, FontAccumulator> accumulators)
        {
            if (shape == null)
            {
                return;
            }

            var fontsInShape = CollectFontsFromShape(shape);

            foreach (var fontName in fontsInShape)
            {
                AddUsage(
                    accumulators,
                    fontName,
                    new FontUsageLocation
                    {
                        Scope = scope,
                        SlideIndex = slideIndex,
                        ShapeId = SafeGetShapeId(shape),
                        ScopeName = scopeLabel,
                        ShapeName = SafeGetShapeName(shape),
                        Label = BuildLocationLabel(scope, slideIndex, shape, scopeLabel),
                        IsSelectable = scope == PresentationScope.Slide && slideIndex.HasValue
                    });
            }
        }

        private HashSet<string> CollectFontsFromShape(PowerPoint.Shape shape)
        {
            var fontsInShape = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (shape == null)
            {
                return fontsInShape;
            }

            try
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    HarvestTextRangeFonts(shape.TextFrame.TextRange, fontsInShape);
                }
            }
            catch
            {
            }

            try
            {
                dynamic textFrame2 = shape.TextFrame2;
                if (textFrame2 != null && SafeMsoBoolean(textFrame2.HasText))
                {
                    HarvestTextRangeFonts(textFrame2.TextRange, fontsInShape);
                }
            }
            catch
            {
            }

            return fontsInShape;
        }

        private void HarvestTextRangeFonts(dynamic textRange, ISet<string> fontsInShape)
        {
            if (textRange == null || fontsInShape == null)
            {
                return;
            }

            foreach (var run in EnumerateRuns(textRange))
            {
                try
                {
                    CollectFontNames(run.Font, fontsInShape);
                }
                catch
                {
                }

                try
                {
                    CollectFontNames(run.ParagraphFormat.Bullet.Font, fontsInShape);
                }
                catch
                {
                }
            }

            try
            {
                CollectFontNames(textRange.Font, fontsInShape);
            }
            catch
            {
            }
        }

        private static void CollectFontNames(dynamic font, ISet<string> fontsInShape)
        {
            if (font == null || fontsInShape == null)
            {
                return;
            }

            TryAddFont(fontsInShape, TryGetFontProperty(font, "Name"));
            TryAddFont(fontsInShape, TryGetFontProperty(font, "NameAscii"));
            TryAddFont(fontsInShape, TryGetFontProperty(font, "NameFarEast"));
            TryAddFont(fontsInShape, TryGetFontProperty(font, "NameComplexScript"));
            TryAddFont(fontsInShape, TryGetFontProperty(font, "NameOther"));
        }

        private static string TryGetFontProperty(dynamic font, string propertyName)
        {
            try
            {
                return Convert.ToString(font.GetType().InvokeMember(propertyName, System.Reflection.BindingFlags.GetProperty, null, font, null));
            }
            catch
            {
                try
                {
                    return Convert.ToString(((dynamic)font)[propertyName]);
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        private static void TryAddFont(ISet<string> fonts, string fontName)
        {
            var normalized = NormalizeFontName(fontName);
            if (!string.IsNullOrWhiteSpace(normalized))
            {
                fonts.Add(normalized);
            }
        }

        private static void AddUsage(IDictionary<string, FontAccumulator> accumulators, string fontName, FontUsageLocation location)
        {
            var normalized = NormalizeFontName(fontName);
            if (string.IsNullOrWhiteSpace(normalized) || location == null)
            {
                return;
            }

            if (ShouldSuppressThemeFontUsage(normalized, location))
            {
                return;
            }

            FontAccumulator accumulator;
            if (!accumulators.TryGetValue(normalized, out accumulator))
            {
                accumulator = new FontAccumulator(normalized);
                accumulators[normalized] = accumulator;
            }

            accumulator.UsesCount++;
            accumulator.Locations.Add(location);
        }

        private IEnumerable<dynamic> EnumerateRuns(dynamic textRange)
        {
            dynamic runs = null;
            try
            {
                runs = textRange.Runs();
            }
            catch
            {
                runs = null;
            }

            if (runs == null)
            {
                yield return textRange;
                yield break;
            }

            int count;
            try
            {
                count = runs.Count;
            }
            catch
            {
                count = 0;
            }

            if (count <= 0)
            {
                yield return textRange;
                yield break;
            }

            for (var i = 1; i <= count; i++)
            {
                yield return runs[i];
            }
        }

        private void ApplyPresentationFontMetadata(FontScanSnapshot snapshot, IDictionary<string, FontAccumulator> accumulators, bool addMissingFonts)
        {
            if (snapshot == null || accumulators == null)
            {
                return;
            }

            foreach (var font in snapshot.Fonts)
            {
                if (font == null || string.IsNullOrWhiteSpace(font.FontName))
                {
                    continue;
                }

                FontAccumulator accumulator;
                if (!accumulators.TryGetValue(font.FontName, out accumulator))
                {
                    if (!addMissingFonts && !ShouldCreateAccumulatorFromPresentationFont(font))
                    {
                        continue;
                    }

                    accumulator = new FontAccumulator(font.FontName);
                    accumulators[font.FontName] = accumulator;
                }

                accumulator.HasPresentationMetadata = true;
                accumulator.IsInstalled = font.IsInstalled;
                accumulator.HasEmbeddableMetadata = font.HasEmbeddableMetadata;
                if (font.HasEmbeddableMetadata)
                {
                    accumulator.IsEmbeddable = font.IsEmbeddable;
                }

                var embeddingStatus = ResolveEmbeddingStatus(snapshot, font);
                if (embeddingStatus != FontEmbeddingStatus.Unknown)
                {
                    accumulator.EmbeddingStatus = embeddingStatus;
                }
            }

            foreach (var accumulator in accumulators.Values)
            {
                if (accumulator == null || accumulator.EmbeddingStatus != FontEmbeddingStatus.Unknown)
                {
                    continue;
                }

                if (snapshot.HasEmbeddedFontDataKnown)
                {
                    accumulator.EmbeddingStatus = snapshot.EmbeddedFontNames.Contains(accumulator.FontName)
                        ? (snapshot.SaveSubsetFonts ? FontEmbeddingStatus.Subset : FontEmbeddingStatus.Yes)
                        : FontEmbeddingStatus.No;
                    continue;
                }

                if (snapshot.HasPackageEmbeddingSettings && !snapshot.RequestsEmbeddedFonts)
                {
                    accumulator.EmbeddingStatus = FontEmbeddingStatus.No;
                    continue;
                }

                if (accumulator.IsThemeFont && snapshot.HasPackageEmbeddingSettings && snapshot.RequestsEmbeddedFonts && snapshot.HasEmbeddedFontData)
                {
                    accumulator.EmbeddingStatus = snapshot.SaveSubsetFonts
                        ? FontEmbeddingStatus.Subset
                        : FontEmbeddingStatus.Yes;
                }
            }
        }

        private DisplayedFontMap CaptureDisplayedFontMap(PowerPoint.Presentation presentation)
        {
            var displayedFontMap = new DisplayedFontMap();
            if (presentation == null)
            {
                return displayedFontMap;
            }

            foreach (var scope in BuildScanScopes(presentation))
            {
                CaptureDisplayedFontMap(scope.Shapes, scope.Scope, scope.Index, scope.Label, displayedFontMap);
            }

            return displayedFontMap;
        }

        private void CaptureDisplayedFontMap(
            PowerPoint.Shapes shapes,
            PresentationScope scope,
            int? slideIndex,
            string scopeLabel,
            DisplayedFontMap displayedFontMap)
        {
            if (shapes == null || displayedFontMap == null)
            {
                return;
            }

            for (var i = 1; i <= shapes.Count; i++)
            {
                CaptureDisplayedFontMap(shapes[i], scope, slideIndex, scopeLabel, displayedFontMap);
            }
        }

        private void CaptureDisplayedFontMap(
            PowerPoint.Shape shape,
            PresentationScope scope,
            int? slideIndex,
            string scopeLabel,
            DisplayedFontMap displayedFontMap)
        {
            if (shape == null || displayedFontMap == null)
            {
                return;
            }

            var fontsInShape = CollectFontsFromShape(shape);
            if (fontsInShape.Count > 0)
            {
                displayedFontMap.Add(
                    BuildLocationIdKey(scope, slideIndex, SafeGetShapeId(shape), SafeGetShapeName(shape), scopeLabel),
                    fontsInShape,
                    true);
                displayedFontMap.Add(
                    BuildLocationLabelKey(BuildLocationLabel(scope, slideIndex, shape, scopeLabel)),
                    fontsInShape,
                    false);
            }

            try
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    for (var i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        CaptureDisplayedFontMap(shape.GroupItems[i], scope, slideIndex, scopeLabel, displayedFontMap);
                    }
                }
            }
            catch
            {
            }

            try
            {
                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    for (var row = 1; row <= shape.Table.Rows.Count; row++)
                    {
                        for (var column = 1; column <= shape.Table.Columns.Count; column++)
                        {
                            CaptureDisplayedFontMap(shape.Table.Cell(row, column).Shape, scope, slideIndex, scopeLabel, displayedFontMap);
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void ApplySubstitutionStateFromPowerPoint(IDictionary<string, FontAccumulator> accumulators, DisplayedFontMap displayedFontMap)
        {
            if (accumulators == null || displayedFontMap == null)
            {
                return;
            }

            foreach (var accumulator in accumulators.Values)
            {
                if (accumulator == null || accumulator.IsThemeFont)
                {
                    continue;
                }

                foreach (var location in accumulator.Locations)
                {
                    ISet<string> displayedFonts;
                    if (!TryGetDisplayedFontsForLocation(displayedFontMap, location, out displayedFonts))
                    {
                        continue;
                    }

                    if (!displayedFonts.Contains(accumulator.FontName))
                    {
                        accumulator.IsSubstituted = true;
                        break;
                    }
                }
            }
        }

        private void ApplyWarningState(IDictionary<string, FontAccumulator> accumulators, EmbeddedSaveValidation saveValidation)
        {
            if (accumulators == null)
            {
                return;
            }

            foreach (var accumulator in accumulators.Values)
            {
                if (accumulator == null || accumulator.IsThemeFont)
                {
                    continue;
                }

                if (accumulator.EmbeddingStatus == FontEmbeddingStatus.Yes
                    || accumulator.EmbeddingStatus == FontEmbeddingStatus.Subset)
                {
                    accumulator.HasSaveWarning = false;
                    continue;
                }

                if (saveValidation != null
                    && saveValidation.HasEmbeddedFontNames
                    && saveValidation.EmbeddedFontNames.Contains(accumulator.FontName))
                {
                    accumulator.HasSaveWarning = false;
                    continue;
                }

                var missingOnSystem = !accumulator.IsInstalled;
                var explicitlyNonEmbeddable = accumulator.HasEmbeddableMetadata && !accumulator.IsEmbeddable;

                if (saveValidation != null)
                {
                    accumulator.HasSaveWarning = !saveValidation.CopySucceeded
                        && saveValidation.LooksLikeFontAvailabilityFailure
                        && (missingOnSystem || explicitlyNonEmbeddable);
                    continue;
                }

                accumulator.HasSaveWarning = false;
            }
        }

        private static void ApplyValidatedEmbeddingState(
            FontScanSnapshot snapshot,
            IDictionary<string, FontAccumulator> accumulators,
            EmbeddedSaveValidation saveValidation)
        {
            if (snapshot == null
                || accumulators == null
                || saveValidation == null
                || !snapshot.RequestsEmbeddedFonts
                || !saveValidation.CopySucceeded
                || !saveValidation.HasEmbeddedFontNames
                || saveValidation.EmbeddedFontNames == null)
            {
                return;
            }

            foreach (var accumulator in accumulators.Values)
            {
                if (accumulator == null || accumulator.IsThemeFont)
                {
                    continue;
                }

                if (saveValidation.EmbeddedFontNames.Contains(accumulator.FontName))
                {
                    accumulator.EmbeddingStatus = snapshot.SaveSubsetFonts
                        ? FontEmbeddingStatus.Subset
                        : FontEmbeddingStatus.Yes;
                }
            }
        }

        private static bool NeedsEmbeddedSaveValidation(IDictionary<string, FontAccumulator> accumulators)
        {
            return accumulators != null && accumulators.Values.Any(NeedsEmbeddedSaveValidation);
        }

        private static bool NeedsEmbeddedSaveValidation(FontAccumulator accumulator)
        {
            if (accumulator == null || accumulator.IsThemeFont || accumulator.IsSubstituted)
            {
                return false;
            }

            if (accumulator.EmbeddingStatus == FontEmbeddingStatus.Yes
                || accumulator.EmbeddingStatus == FontEmbeddingStatus.Subset)
            {
                return false;
            }

            return !accumulator.IsInstalled
                || (accumulator.HasEmbeddableMetadata && !accumulator.IsEmbeddable);
        }

        private static bool ShouldPerformEmbeddedSaveValidation(
            FontScanSnapshot snapshot,
            IDictionary<string, FontAccumulator> accumulators)
        {
            return snapshot != null
                && snapshot.RequestsEmbeddedFonts
                && NeedsEmbeddedSaveValidation(accumulators);
        }

        private static bool ShouldCaptureDisplayedFontMap(
            FontScanSnapshot snapshot,
            IDictionary<string, FontAccumulator> accumulators)
        {
            if (snapshot == null || accumulators == null || accumulators.Count == 0)
            {
                return false;
            }

            return accumulators.Values.Any(x =>
                x != null
                && !x.IsThemeFont
                && (!x.IsInstalled || !x.HasPresentationMetadata));
        }

        private EmbeddedSaveValidation CreateEmbeddedSaveValidation(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return null;
            }

            var validation = new EmbeddedSaveValidation();
            string tempFilePath = null;

            try
            {
                tempFilePath = Path.Combine(Path.GetTempPath(), "morphos-font-check-" + Guid.NewGuid().ToString("N") + ".pptx");
                dynamic dynamicPresentation = presentation;
                dynamicPresentation.SaveCopyAs(tempFilePath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                validation.CopySucceeded = true;

                ISet<string> embeddedFontNames;
                validation.HasEmbeddedFontNames = PresentationPackageInspector.TryGetEmbeddedFontNames(tempFilePath, out embeddedFontNames);
                validation.EmbeddedFontNames = embeddedFontNames ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                validation.DetectedEmbeddedFontData = validation.EmbeddedFontNames.Count > 0;
            }
            catch (Exception ex)
            {
                validation.CopySucceeded = false;
                validation.HasEmbeddedFontNames = false;
                validation.EmbeddedFontNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                validation.DetectedEmbeddedFontData = false;
                validation.LooksLikeFontAvailabilityFailure = LooksLikeFontAvailabilityFailure(ex);
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(tempFilePath) && File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                    }
                    catch
                    {
                    }
                }
            }

            return validation;
        }

        private static void ApplySaveValidationResult(EmbeddedSaveValidation saveValidation, FontReplacementResult result)
        {
            if (result == null || saveValidation == null)
            {
                return;
            }

            result.SaveValidationCopySucceeded = saveValidation.CopySucceeded;
            result.SaveValidationDetectedEmbeddedFontData = saveValidation.DetectedEmbeddedFontData;
        }

        private static bool TryGetDisplayedFontsForLocation(DisplayedFontMap displayedFontMap, FontUsageLocation location, out ISet<string> displayedFonts)
        {
            displayedFonts = null;
            if (displayedFontMap == null || location == null)
            {
                return false;
            }

            if (displayedFontMap.ById.TryGetValue(
                BuildLocationIdKey(location.Scope, location.SlideIndex, location.ShapeId, location.ShapeName, location.ScopeName),
                out displayedFonts))
            {
                return true;
            }

            return displayedFontMap.ByLabel.TryGetValue(BuildLocationLabelKey(location.Label), out displayedFonts);
        }

        private void ApplyToAllPresentationShapes(PowerPoint.Presentation presentation, Action<PowerPoint.Shape> action)
        {
            if (presentation == null || action == null)
            {
                return;
            }

            TryApplyToShapeCollection(presentation.SlideMaster?.Shapes, action);

            foreach (PowerPoint.CustomLayout layout in presentation.SlideMaster.CustomLayouts)
            {
                TryApplyToShapeCollection(layout.Shapes, action);
                YieldToOffice();
            }

            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                TryApplyToShapeCollection(slide.Shapes, action);
                YieldToOffice();
            }

            TryApplyToShapeCollection(presentation.NotesMaster?.Shapes, action);
        }

        private void ApplyToPresentationTextStyles(PowerPoint.Presentation presentation, ISet<string> sourceFonts, string replacementFont)
        {
            if (presentation == null || sourceFonts == null || sourceFonts.Count == 0 || string.IsNullOrWhiteSpace(replacementFont))
            {
                return;
            }

            TryApplyTextStyles(presentation.SlideMaster, sourceFonts, replacementFont);
            TryApplyTextStyles(presentation.NotesMaster, sourceFonts, replacementFont);
        }

        private void TryApplyTextStyles(object master, ISet<string> sourceFonts, string replacementFont)
        {
            if (master == null)
            {
                return;
            }

            dynamic textStyles = null;
            try
            {
                textStyles = ((dynamic)master).TextStyles;
            }
            catch
            {
                textStyles = null;
            }

            if (textStyles == null)
            {
                return;
            }

            foreach (PowerPoint.PpTextStyleType styleType in new[]
            {
                PowerPoint.PpTextStyleType.ppDefaultStyle,
                PowerPoint.PpTextStyleType.ppTitleStyle,
                PowerPoint.PpTextStyleType.ppBodyStyle
            })
            {
                dynamic textStyle = null;

                try
                {
                    textStyle = textStyles[(int)styleType];
                }
                catch
                {
                    try
                    {
                        textStyle = textStyles(styleType);
                    }
                    catch
                    {
                        textStyle = null;
                    }
                }

                TryApplyTextStyleLevels(textStyle, sourceFonts, replacementFont);
            }
        }

        private static void TryApplyTextStyleLevels(dynamic textStyle, ISet<string> sourceFonts, string replacementFont)
        {
            if (textStyle == null)
            {
                return;
            }

            for (var levelIndex = 1; levelIndex <= 5; levelIndex++)
            {
                dynamic level = null;
                try
                {
                    level = textStyle.Levels(levelIndex);
                }
                catch
                {
                    level = null;
                }

                if (level == null)
                {
                    continue;
                }

                try
                {
                    var font = level.Font;
                    ReplaceFontProperty(font, "Name", sourceFonts, replacementFont);
                    ReplaceFontProperty(font, "NameAscii", sourceFonts, replacementFont);
                    ReplaceFontProperty(font, "NameFarEast", sourceFonts, replacementFont);
                    ReplaceFontProperty(font, "NameComplexScript", sourceFonts, replacementFont);
                    ReplaceFontProperty(font, "NameOther", sourceFonts, replacementFont);
                }
                catch
                {
                }

                try
                {
                    var bulletFont = level.ParagraphFormat.Bullet.Font;
                    ReplaceFontProperty(bulletFont, "Name", sourceFonts, replacementFont);
                }
                catch
                {
                }
            }
        }

        private void TryApplyToShapeCollection(PowerPoint.Shapes shapes, Action<PowerPoint.Shape> action)
        {
            if (shapes == null)
            {
                return;
            }

            for (var i = 1; i <= shapes.Count; i++)
            {
                ApplyToShapeDeep(shapes[i], action);
                if (i % 12 == 0)
                {
                    YieldToOffice();
                }
            }
        }

        private void ApplyToShapeDeep(PowerPoint.Shape shape, Action<PowerPoint.Shape> action)
        {
            if (shape == null)
            {
                return;
            }

            action(shape);

            try
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    for (var i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        ApplyToShapeDeep(shape.GroupItems[i], action);
                    }
                }
            }
            catch
            {
            }

            try
            {
                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    for (var row = 1; row <= shape.Table.Rows.Count; row++)
                    {
                        for (var column = 1; column <= shape.Table.Columns.Count; column++)
                        {
                            action(shape.Table.Cell(row, column).Shape);
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void ReplaceFontInShape(PowerPoint.Shape shape, ISet<string> sourceFonts, string replacementFont)
        {
            try
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    dynamic textRange = shape.TextFrame.TextRange;
                    foreach (var run in EnumerateRuns(textRange))
                    {
                        ReplaceFontOnTextRange(run, sourceFonts, replacementFont);
                    }

                    try
                    {
                        ReplaceFontOnTextRange(textRange, sourceFonts, replacementFont);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            try
            {
                dynamic textFrame2 = shape.TextFrame2;
                if (textFrame2 != null && SafeMsoBoolean(textFrame2.HasText))
                {
                    dynamic textRange2 = textFrame2.TextRange;
                    foreach (var run in EnumerateRuns(textRange2))
                    {
                        ReplaceFontOnTextRange(run, sourceFonts, replacementFont);
                    }

                    try
                    {
                        ReplaceFontOnTextRange(textRange2, sourceFonts, replacementFont);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        private static void ReplaceFontOnTextRange(dynamic textRange, ISet<string> sourceFonts, string replacementFont)
        {
            if (textRange == null)
            {
                return;
            }

            try
            {
                var font = textRange.Font;
                ReplaceFontProperty(font, "Name", sourceFonts, replacementFont);
                ReplaceFontProperty(font, "NameAscii", sourceFonts, replacementFont);
                ReplaceFontProperty(font, "NameFarEast", sourceFonts, replacementFont);
                ReplaceFontProperty(font, "NameComplexScript", sourceFonts, replacementFont);
                ReplaceFontProperty(font, "NameOther", sourceFonts, replacementFont);
            }
            catch
            {
            }

            try
            {
                var bulletFont = textRange.ParagraphFormat.Bullet.Font;
                ReplaceFontProperty(bulletFont, "Name", sourceFonts, replacementFont);
            }
            catch
            {
            }
        }

        private static void ReplaceFontProperty(object font, string propertyName, ISet<string> sourceFonts, string replacementFont)
        {
            if (font == null || sourceFonts == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return;
            }

            try
            {
                string currentValue;
                if (!ComFontAccessorCache.TryGetString(font, propertyName, out currentValue))
                {
                    return;
                }

                currentValue = FontNameNormalizer.Normalize(currentValue);

                if (!sourceFonts.Contains(currentValue))
                {
                    return;
                }

                ComFontAccessorCache.TrySetString(font, propertyName, replacementFont);
            }
            catch
            {
            }
        }

        private static void TryMarkPresentationDirty(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return;
            }

            try
            {
                presentation.Saved = MsoTriState.msoFalse;
            }
            catch
            {
            }
        }

        private static void TryMarkPresentationSaved(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return;
            }

            try
            {
                presentation.Saved = MsoTriState.msoTrue;
            }
            catch
            {
            }
        }

        private void ApplyColorToShape(PowerPoint.Shape shape, int oleColor)
        {
            try
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    dynamic textRange = shape.TextFrame.TextRange;
                    foreach (var run in EnumerateRuns(textRange))
                    {
                        run.Font.Color.RGB = oleColor;
                    }
                }
            }
            catch
            {
            }

            try
            {
                if (shape.Fill != null && shape.Fill.Visible == MsoTriState.msoTrue)
                {
                    shape.Fill.ForeColor.RGB = oleColor;
                }
            }
            catch
            {
            }

            try
            {
                if (shape.Line != null && shape.Line.Visible == MsoTriState.msoTrue)
                {
                    shape.Line.ForeColor.RGB = oleColor;
                }
            }
            catch
            {
            }
        }

        private int ApplyColorReplacementsToPresentation(
            PowerPoint.Presentation presentation,
            IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            if (presentation == null || instructions == null || instructions.Count == 0)
            {
                return 0;
            }

            var lookup = instructions
                .GroupBy(x => BuildColorInstructionKey(x.UsageKind, x.SourceHexValue), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.Last(), StringComparer.OrdinalIgnoreCase);

            var hexFallbackLookup = instructions
                .GroupBy(x => NormalizeColorHex(x.SourceHexValue), StringComparer.OrdinalIgnoreCase)
                .Where(x => x.Count() == 1 || x.Select(BuildColorReplacementFingerprint).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1)
                .ToDictionary(x => x.Key, x => x.Last(), StringComparer.OrdinalIgnoreCase);

            if (lookup.Count == 0 && hexFallbackLookup.Count == 0)
            {
                return 0;
            }

            var targetedReplacements = ApplyColorReplacementsToTargetedShapes(presentation, instructions);
            if (targetedReplacements.HasValue)
            {
                return targetedReplacements.Value;
            }

            var replacements = 0;
            ApplyToAllPresentationShapes(
                presentation,
                shape => replacements += ApplyColorReplacementsToShape(shape, lookup, hexFallbackLookup));

            return replacements;
        }

        private int? ApplyColorReplacementsToTargetedShapes(
            PowerPoint.Presentation presentation,
            IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            if (presentation == null || instructions == null || instructions.Count == 0)
            {
                return null;
            }

            var targets = BuildColorReplacementTargets(instructions);
            if (targets == null || targets.Count == 0)
            {
                return null;
            }

            var replacements = 0;
            for (var i = 0; i < targets.Count; i++)
            {
                var target = targets[i];
                if (target == null || !target.SlideIndex.HasValue)
                {
                    continue;
                }

                try
                {
                    var slide = presentation.Slides[target.SlideIndex.Value];
                    var shape = FindShape(slide.Shapes, target.ShapeId, target.ShapeName);
                    if (shape == null)
                    {
                        continue;
                    }

                    IDictionary<string, ColorReplacementInstruction> lookup;
                    IDictionary<string, ColorReplacementInstruction> hexFallbackLookup;
                    BuildColorInstructionLookups(target.Instructions, out lookup, out hexFallbackLookup);
                    ApplyToShapeDeep(shape, currentShape => replacements += ApplyColorReplacementsToShape(currentShape, lookup, hexFallbackLookup));
                }
                catch
                {
                }

                if ((i + 1) % 12 == 0)
                {
                    YieldToOffice();
                }
            }

            return replacements;
        }

        private static IReadOnlyList<ColorReplacementTarget> BuildColorReplacementTargets(IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            var targets = new Dictionary<string, ColorReplacementTarget>(StringComparer.OrdinalIgnoreCase);
            foreach (var instruction in instructions)
            {
                var locations = instruction == null || instruction.TargetLocations == null
                    ? null
                    : instruction.TargetLocations;
                if (locations == null || locations.Count == 0)
                {
                    return null;
                }

                foreach (var location in locations)
                {
                    if (location == null
                        || location.Scope != PresentationScope.Slide
                        || !location.SlideIndex.HasValue)
                    {
                        continue;
                    }

                    var key = BuildColorReplacementTargetKey(location);
                    ColorReplacementTarget target;
                    if (!targets.TryGetValue(key, out target))
                    {
                        target = new ColorReplacementTarget(location);
                        targets[key] = target;
                    }

                    target.Instructions.Add(instruction);
                }
            }

            return targets.Values.ToList();
        }

        private static void BuildColorInstructionLookups(
            IReadOnlyList<ColorReplacementInstruction> instructions,
            out IDictionary<string, ColorReplacementInstruction> lookup,
            out IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            var safeInstructions = instructions == null
                ? (IReadOnlyList<ColorReplacementInstruction>)Array.Empty<ColorReplacementInstruction>()
                : instructions.Where(x => x != null).ToList();

            lookup = safeInstructions
                .GroupBy(x => BuildColorInstructionKey(x.UsageKind, x.SourceHexValue), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.Last(), StringComparer.OrdinalIgnoreCase);

            hexFallbackLookup = safeInstructions
                .GroupBy(x => NormalizeColorHex(x.SourceHexValue), StringComparer.OrdinalIgnoreCase)
                .Where(x => x.Count() == 1 || x.Select(BuildColorReplacementFingerprint).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1)
                .ToDictionary(x => x.Key, x => x.Last(), StringComparer.OrdinalIgnoreCase);
        }

        private static string BuildColorReplacementTargetKey(FontUsageLocation location)
        {
            if (location == null)
            {
                return string.Empty;
            }

            return (location.SlideIndex.HasValue ? location.SlideIndex.Value.ToString() : string.Empty) + "|"
                + (location.ShapeId.HasValue ? location.ShapeId.Value.ToString() : string.Empty) + "|"
                + (location.ShapeName ?? string.Empty);
        }

        private int ApplyColorReplacementsToShape(
            PowerPoint.Shape shape,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (shape == null)
            {
                return 0;
            }

            var replacements = 0;
            replacements += TryApplyColorToFillFormat(shape.Fill, ColorUsageKind.ShapeFill, lookup, hexFallbackLookup);
            replacements += TryApplyColorToLineFormat(shape.Line, ColorUsageKind.Line, lookup, hexFallbackLookup);
            replacements += TryApplyShapeEffectColors(shape, lookup, hexFallbackLookup);
            replacements += TryApplyTextColors(shape, lookup, hexFallbackLookup);
            return replacements;
        }

        private int TryApplyShapeEffectColors(
            PowerPoint.Shape shape,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (shape == null)
            {
                return 0;
            }

            var replacements = 0;
            replacements += TryApplyColorToPropertyPath(shape, ColorUsageKind.Effect, lookup, hexFallbackLookup, "Shadow", "ForeColor");
            replacements += TryApplyColorToPropertyPath(shape, ColorUsageKind.Effect, lookup, hexFallbackLookup, "Glow", "Color");
            return replacements;
        }

        private int TryApplyTextColors(
            PowerPoint.Shape shape,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (shape == null)
            {
                return 0;
            }

            var replacements = 0;

            try
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    dynamic textRange = shape.TextFrame.TextRange;
                    var usedRuns = false;
                    try
                    {
                        foreach (var run in EnumerateRuns(textRange))
                        {
                            replacements += TryApplyLegacyTextRangeColors(run, lookup, hexFallbackLookup);
                            usedRuns = true;
                        }
                    }
                    catch
                    {
                        if (!usedRuns)
                        {
                            replacements += TryApplyLegacyTextRangeColors(textRange, lookup, hexFallbackLookup);
                        }
                    }

                    replacements += TryApplyColorToPropertyPath(textRange, ColorUsageKind.TextFill, lookup, hexFallbackLookup, "ParagraphFormat", "Bullet", "Font", "Color");
                }
            }
            catch
            {
            }

            try
            {
                dynamic textFrame2 = shape.TextFrame2;
                if (textFrame2 != null && SafeMsoBoolean(textFrame2.HasText))
                {
                    dynamic textRange2 = textFrame2.TextRange;
                    var usedRuns = false;
                    try
                    {
                        foreach (var run in EnumerateRuns(textRange2))
                        {
                            replacements += TryApplyModernTextRangeColors(run, lookup, hexFallbackLookup);
                            usedRuns = true;
                        }
                    }
                    catch
                    {
                        if (!usedRuns)
                        {
                            replacements += TryApplyModernTextRangeColors(textRange2, lookup, hexFallbackLookup);
                        }
                    }

                    replacements += TryApplyColorToPropertyPath(textRange2, ColorUsageKind.TextFill, lookup, hexFallbackLookup, "ParagraphFormat", "Bullet", "Font", "Fill", "ForeColor");
                }
            }
            catch
            {
            }

            return replacements;
        }

        private int TryApplyLegacyTextRangeColors(
            dynamic textRange,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (textRange == null)
            {
                return 0;
            }

            var replacements = 0;
            replacements += TryApplyColorToPropertyPath(textRange, ColorUsageKind.TextFill, lookup, hexFallbackLookup, "Font", "Color");
            replacements += TryApplyColorToPropertyPath(textRange, ColorUsageKind.TextFill, lookup, hexFallbackLookup, "ParagraphFormat", "Bullet", "Font", "Color");
            return replacements;
        }

        private int TryApplyModernTextRangeColors(
            dynamic textRange,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (textRange == null)
            {
                return 0;
            }

            var replacements = 0;
            replacements += TryApplyColorToPropertyPath(textRange, ColorUsageKind.TextFill, lookup, hexFallbackLookup, "Font", "Color");

            object font;
            if (TryReadProperty(textRange, "Font", out font) && font != null)
            {
                object fill;
                if (TryReadProperty(font, "Fill", out fill))
                {
                    replacements += TryApplyColorToFillFormat(fill, ColorUsageKind.TextFill, lookup, hexFallbackLookup);
                }

                object line;
                if (TryReadProperty(font, "Line", out line))
                {
                    replacements += TryApplyColorToLineFormat(line, ColorUsageKind.Line, lookup, hexFallbackLookup);
                }

                replacements += TryApplyColorToPropertyPath(font, ColorUsageKind.Effect, lookup, hexFallbackLookup, "Glow", "Color");
            }

            replacements += TryApplyColorToPropertyPath(textRange, ColorUsageKind.TextFill, lookup, hexFallbackLookup, "ParagraphFormat", "Bullet", "Font", "Fill", "ForeColor");
            replacements += TryApplyColorToPropertyPath(textRange, ColorUsageKind.TextFill, lookup, hexFallbackLookup, "ParagraphFormat", "Bullet", "Font", "Color");
            return replacements;
        }

        private int TryApplyColorToFillFormat(
            object fillFormat,
            ColorUsageKind usageKind,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (fillFormat == null)
            {
                return 0;
            }

            var replacements = 0;
            replacements += TryApplyColorToPropertyPath(fillFormat, usageKind, lookup, hexFallbackLookup, "ForeColor");
            replacements += TryApplyColorToPropertyPath(fillFormat, usageKind, lookup, hexFallbackLookup, "BackColor");
            replacements += TryApplyGradientStopColors(fillFormat, usageKind, lookup, hexFallbackLookup);
            return replacements;
        }

        private int TryApplyColorToLineFormat(
            object lineFormat,
            ColorUsageKind usageKind,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (lineFormat == null)
            {
                return 0;
            }

            var replacements = 0;
            replacements += TryApplyColorToPropertyPath(lineFormat, usageKind, lookup, hexFallbackLookup, "ForeColor");
            replacements += TryApplyColorToPropertyPath(lineFormat, usageKind, lookup, hexFallbackLookup, "BackColor");
            return replacements;
        }

        private int TryApplyGradientStopColors(
            object formatObject,
            ColorUsageKind usageKind,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (formatObject == null)
            {
                return 0;
            }

            dynamic gradientStops = null;
            try
            {
                gradientStops = ((dynamic)formatObject).GradientStops;
            }
            catch
            {
                gradientStops = null;
            }

            if (gradientStops == null)
            {
                return 0;
            }

            int count;
            try
            {
                count = gradientStops.Count;
            }
            catch
            {
                return 0;
            }

            var replacements = 0;
            for (var i = 1; i <= count; i++)
            {
                try
                {
                    replacements += TryApplyColorToPropertyPath(gradientStops[i], usageKind, lookup, hexFallbackLookup, "Color");
                }
                catch
                {
                }
            }

            return replacements;
        }

        private int TryApplyColorToPropertyPath(
            object root,
            ColorUsageKind usageKind,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup,
            params string[] propertyPath)
        {
            if (root == null || propertyPath == null || propertyPath.Length == 0)
            {
                return 0;
            }

            object current = root;
            for (var i = 0; i < propertyPath.Length; i++)
            {
                if (!TryReadProperty(current, propertyPath[i], out current) || current == null)
                {
                    return 0;
                }
            }

            return TryApplyColorToColorFormat(current, usageKind, lookup, hexFallbackLookup);
        }

        private int TryApplyColorToColorFormat(
            object colorFormat,
            ColorUsageKind usageKind,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup)
        {
            if (colorFormat == null)
            {
                return 0;
            }

            string currentHexValue;
            if (!TryGetDirectRgbHexValue(colorFormat, out currentHexValue))
            {
                return 0;
            }

            ColorReplacementInstruction instruction;
            if (!TryResolveColorInstruction(usageKind, currentHexValue, lookup, hexFallbackLookup, out instruction))
            {
                return 0;
            }

            return ApplyColorInstruction(colorFormat, instruction) ? 1 : 0;
        }

        private static bool TryGetDirectRgbHexValue(object colorFormat, out string hexValue)
        {
            hexValue = string.Empty;
            if (colorFormat == null)
            {
                return false;
            }

            object rawType;
            if (TryReadProperty(colorFormat, "Type", out rawType))
            {
                try
                {
                    if (Convert.ToInt32(rawType) != (int)MsoColorType.msoColorTypeRGB)
                    {
                        return false;
                    }
                }
                catch
                {
                    return false;
                }
            }

            object rawRgb;
            if (!TryReadProperty(colorFormat, "RGB", out rawRgb))
            {
                return false;
            }

            hexValue = ConvertOleColorToHex(rawRgb);
            return !string.IsNullOrWhiteSpace(hexValue);
        }

        private static bool TryResolveColorInstruction(
            ColorUsageKind usageKind,
            string hexValue,
            IDictionary<string, ColorReplacementInstruction> lookup,
            IDictionary<string, ColorReplacementInstruction> hexFallbackLookup,
            out ColorReplacementInstruction instruction)
        {
            instruction = null;
            if (string.IsNullOrWhiteSpace(hexValue))
            {
                return false;
            }

            if (lookup != null
                && lookup.TryGetValue(BuildColorInstructionKey(usageKind, hexValue), out instruction))
            {
                return true;
            }

            return hexFallbackLookup != null
                && hexFallbackLookup.TryGetValue(NormalizeColorHex(hexValue), out instruction);
        }

        private static bool ApplyColorInstruction(object colorFormat, ColorReplacementInstruction instruction)
        {
            if (colorFormat == null || instruction == null)
            {
                return false;
            }

            if (instruction.UseThemeColor)
            {
                MsoThemeColorIndex themeColorIndex;
                if (TryMapThemeSchemeNameToThemeColorIndex(instruction.ThemeSchemeName, out themeColorIndex)
                    && TryWriteProperty(colorFormat, "ObjectThemeColor", themeColorIndex))
                {
                    return true;
                }

                int schemeColorIndex;
                return TryMapThemeSchemeNameToSchemeColorIndex(instruction.ThemeSchemeName, out schemeColorIndex)
                    && TryWriteProperty(colorFormat, "SchemeColor", schemeColorIndex);
            }

            var oleColor = ConvertHexToOleColor(instruction.ReplacementHexValue);
            return oleColor.HasValue && TryWriteProperty(colorFormat, "RGB", oleColor.Value);
        }

        private static bool TryMapThemeSchemeNameToThemeColorIndex(string schemeName, out MsoThemeColorIndex themeColorIndex)
        {
            themeColorIndex = MsoThemeColorIndex.msoNotThemeColor;
            switch ((schemeName ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "dk1":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorDark1;
                    return true;
                case "lt1":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorLight1;
                    return true;
                case "dk2":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorDark2;
                    return true;
                case "lt2":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorLight2;
                    return true;
                case "accent1":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorAccent1;
                    return true;
                case "accent2":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorAccent2;
                    return true;
                case "accent3":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorAccent3;
                    return true;
                case "accent4":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorAccent4;
                    return true;
                case "accent5":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorAccent5;
                    return true;
                case "accent6":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorAccent6;
                    return true;
                case "hlink":
                case "hyperlink":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorHyperlink;
                    return true;
                case "folhlink":
                case "followedhyperlink":
                    themeColorIndex = MsoThemeColorIndex.msoThemeColorFollowedHyperlink;
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryMapThemeSchemeNameToSchemeColorIndex(string schemeName, out int schemeColorIndex)
        {
            schemeColorIndex = 0;
            switch ((schemeName ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "dk1":
                    schemeColorIndex = 1;
                    return true;
                case "lt1":
                    schemeColorIndex = 2;
                    return true;
                case "dk2":
                    schemeColorIndex = 3;
                    return true;
                case "lt2":
                    schemeColorIndex = 4;
                    return true;
                case "accent1":
                    schemeColorIndex = 5;
                    return true;
                case "accent2":
                    schemeColorIndex = 6;
                    return true;
                case "accent3":
                    schemeColorIndex = 7;
                    return true;
                case "accent4":
                    schemeColorIndex = 8;
                    return true;
                case "accent5":
                    schemeColorIndex = 9;
                    return true;
                case "accent6":
                    schemeColorIndex = 10;
                    return true;
                case "hlink":
                case "hyperlink":
                    schemeColorIndex = 11;
                    return true;
                case "folhlink":
                case "followedhyperlink":
                    schemeColorIndex = 12;
                    return true;
                default:
                    return false;
            }
        }

        private static string ConvertOleColorToHex(object value)
        {
            try
            {
                var color = ColorTranslator.FromOle(Convert.ToInt32(value));
                return color.R.ToString("X2")
                    + color.G.ToString("X2")
                    + color.B.ToString("X2");
            }
            catch
            {
                return string.Empty;
            }
        }

        private static int? ConvertHexToOleColor(string hexValue)
        {
            var normalized = NormalizeColorHex(hexValue);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return null;
            }

            try
            {
                return ColorTranslator.ToOle(ColorTranslator.FromHtml("#" + normalized));
            }
            catch
            {
                return null;
            }
        }

        private static string BuildColorInstructionKey(ColorUsageKind usageKind, string hexValue)
        {
            return ((int)usageKind).ToString() + "|" + NormalizeColorHex(hexValue);
        }

        private static string BuildColorReplacementFingerprint(ColorReplacementInstruction instruction)
        {
            if (instruction == null)
            {
                return string.Empty;
            }

            if (instruction.UseThemeColor)
            {
                return "theme|" + (instruction.ThemeSchemeName ?? string.Empty).Trim();
            }

            return "rgb|" + NormalizeColorHex(instruction.ReplacementHexValue);
        }

        private static string NormalizeColorHex(string hexValue)
        {
            if (string.IsNullOrWhiteSpace(hexValue))
            {
                return string.Empty;
            }

            var normalized = hexValue.Trim().TrimStart('#').ToUpperInvariant();
            return normalized.Length == 6 ? normalized : string.Empty;
        }

        private void TryApplyMasterColorScheme(object master, int oleColor)
        {
            if (master == null)
            {
                return;
            }

            try
            {
                dynamic colorScheme = ((dynamic)master).ColorScheme;
                for (var i = 1; i <= 8; i++)
                {
                    try
                    {
                        colorScheme.Colors(i).RGB = oleColor;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        private IEnumerable<ScanScope> BuildScanScopes(PowerPoint.Presentation presentation)
        {
            if (presentation.SlideMaster != null)
            {
                yield return new ScanScope(PresentationScope.SlideMaster, null, "Slide master", presentation.SlideMaster.Shapes);

                foreach (PowerPoint.CustomLayout layout in presentation.SlideMaster.CustomLayouts)
                {
                    yield return new ScanScope(PresentationScope.CustomLayout, layout.Index, "Custom layout " + layout.Index, layout.Shapes);
                }
            }

            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                yield return new ScanScope(PresentationScope.Slide, slide.SlideIndex, "Slide " + slide.SlideIndex, slide.Shapes);
            }

            if (presentation.NotesMaster != null)
            {
                yield return new ScanScope(PresentationScope.NotesMaster, null, "Notes master", presentation.NotesMaster.Shapes);
            }
        }

        private PackageScanSource AcquirePackageScanSource(PowerPoint.Presentation presentation, FontScanSnapshot snapshot)
        {
            if (presentation == null || snapshot == null)
            {
                return null;
            }

            if (snapshot.CanUsePackageScan && File.Exists(snapshot.FilePath))
            {
                return new PackageScanSource(snapshot.FilePath, false);
            }

            string tempFilePath;
            return TryCreatePackageScanCopy(presentation, out tempFilePath)
                ? new PackageScanSource(tempFilePath, true)
                : null;
        }

        private PackageScanSource AcquireColorScanSource(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return null;
            }

            var filePath = TryGetPresentationPath(presentation);
            if (TryIsPresentationSaved(presentation)
                && !string.IsNullOrWhiteSpace(filePath)
                && File.Exists(filePath))
            {
                return new PackageScanSource(filePath, false);
            }

            return CreateTransientColorCopy(presentation);
        }

        private PackageScanSource CreateTransientColorCopy(PowerPoint.Presentation presentation)
        {
            string tempFilePath;
            return TryCreatePackageScanCopy(presentation, out tempFilePath)
                ? new PackageScanSource(tempFilePath, true)
                : null;
        }

        private int? TryApplyOpenXmlColorReplacements(
            PowerPoint.Presentation presentation,
            IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            return TryApplyOpenXmlPackageMutation(
                presentation,
                "morphos-color-replace",
                filePath => _openXmlColorReplacer.ApplyColorReplacements(filePath, instructions),
                null);
        }

        private int? TryApplyOpenXmlFontReplacements(
            PowerPoint.Presentation presentation,
            IReadOnlyCollection<string> sourceFonts,
            string replacementFont)
        {
            var windowSnapshot = CapturePresentationWindowSnapshot(presentation);
            return TryApplyOpenXmlPackageMutation(
                presentation,
                "morphos-font-replace",
                filePath => _openXmlFontReplacer.ApplyFontReplacements(filePath, sourceFonts, replacementFont),
                reopenedPresentation => ActivateReopenedFontPresentationWindow(reopenedPresentation, windowSnapshot));
        }

        private int? TryApplyOpenXmlColorReplacementsSafely(
            PowerPoint.Presentation presentation,
            IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            try
            {
                return TryApplyOpenXmlColorReplacements(presentation, instructions);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is SecurityException)
            {
                return null;
            }
        }

        private int? TryApplyOpenXmlFontReplacementsSafely(
            PowerPoint.Presentation presentation,
            IReadOnlyCollection<string> sourceFonts,
            string replacementFont)
        {
            try
            {
                return TryApplyOpenXmlFontReplacements(presentation, sourceFonts, replacementFont);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is SecurityException)
            {
                return null;
            }
        }

        private static bool ShouldPreferLiveColorReplacement(IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            if (instructions == null || instructions.Count == 0)
            {
                return true;
            }

            var targetedCount = instructions
                .Where(x => x != null && x.TargetLocations != null)
                .Sum(x => x.TargetLocations.Count);

            return targetedCount == 0 || targetedCount <= 320;
        }

        private static bool ShouldPreferLiveFontReplacement(
            PowerPoint.Presentation presentation,
            ISet<string> sourceFonts)
        {
            if (sourceFonts == null || sourceFonts.Count == 0)
            {
                return true;
            }

            if (sourceFonts.Count <= 2)
            {
                return true;
            }

            try
            {
                return presentation != null && presentation.Slides.Count <= 80;
            }
            catch
            {
                return false;
            }
        }

        private static bool CanUsePackageMutation(FontScanSnapshot snapshot, PowerPoint.Presentation presentation)
        {
            if (snapshot == null
                || !snapshot.IsSaved
                || string.IsNullOrWhiteSpace(snapshot.FilePath)
                || !File.Exists(snapshot.FilePath))
            {
                return false;
            }

            try
            {
                return presentation != null && presentation.Windows.Count <= 0;
            }
            catch
            {
                return false;
            }
        }

        private int? TryApplyOpenXmlPackageMutation(
            PowerPoint.Presentation presentation,
            string workingPrefix,
            Func<string, int> packageMutation,
            Action<PowerPoint.Presentation> afterReopen)
        {
            if (presentation == null || packageMutation == null)
            {
                return null;
            }

            string filePath;
            if (!TryGetExistingPresentationPath(presentation, out filePath))
            {
                return null;
            }

            string workingCopyPath;
            if (!TryCreateColorReplacementWorkingCopy(presentation, filePath, out workingCopyPath))
            {
                return null;
            }

            string backupPath = null;
            var hadUnsavedChanges = !TryIsPresentationSaved(presentation);
            var presentationClosed = false;
            var destinationWasReadOnly = IsFileMarkedReadOnly(filePath);

            try
            {
                EnsureWritableFile(filePath);
                EnsureWritableFile(workingCopyPath);
                WaitForFileAvailability(workingCopyPath, FileAccess.ReadWrite, FileShare.None);

                var replacementCount = packageMutation(workingCopyPath);
                if (replacementCount <= 0)
                {
                    return 0;
                }

                backupPath = BuildSiblingTemporaryPath(filePath, workingPrefix + "-backup");
                TryMarkPresentationSaved(presentation);
                presentation.Close();
                presentationClosed = true;

                WaitForPresentationToClose(filePath);
                ReplaceFileWithRetry(workingCopyPath, filePath, backupPath);

                var reopenedPresentation = OpenPresentationWithRetry(filePath);
                WaitForPresentationReady(reopenedPresentation);
                afterReopen?.Invoke(reopenedPresentation);
                if (hadUnsavedChanges)
                {
                    TryMarkPresentationDirty(reopenedPresentation);
                }

                _lastMutationReloadedPresentation = true;
                RestoreReadOnlyAttribute(filePath, destinationWasReadOnly);

                return replacementCount;
            }
            catch
            {
                if (!presentationClosed)
                {
                    TryMarkPresentationDirty(presentation);
                }
                else
                {
                    TryRestorePresentationFile(filePath, backupPath);
                    RestoreReadOnlyAttribute(filePath, destinationWasReadOnly);
                    TryOpenPresentation(filePath);
                }

                throw;
            }
            finally
            {
                RestoreReadOnlyAttribute(workingCopyPath, false);
                DeleteTemporaryFile(workingCopyPath);
                DeleteTemporaryFile(backupPath);
            }
        }

        private static bool TryCreatePackageScanCopy(PowerPoint.Presentation presentation, out string tempFilePath)
        {
            tempFilePath = null;
            if (presentation == null)
            {
                return false;
            }

            try
            {
                tempFilePath = Path.Combine(Path.GetTempPath(), "morphos-font-scan-" + Guid.NewGuid().ToString("N") + ".pptx");
                dynamic dynamicPresentation = presentation;
                dynamicPresentation.SaveCopyAs(tempFilePath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoFalse);
                return File.Exists(tempFilePath);
            }
            catch
            {
                if (!string.IsNullOrWhiteSpace(tempFilePath) && File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                    }
                    catch
                    {
                    }
                }

                tempFilePath = null;
                return false;
            }
        }

        private static bool EnsurePresentationSavedToDisk(PowerPoint.Presentation presentation, out string filePath)
        {
            filePath = TryGetPresentationPath(presentation);
            if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
            {
                try
                {
                    presentation.Save();
                }
                catch
                {
                }

                return File.Exists(filePath);
            }

            try
            {
                presentation.Save();
            }
            catch
            {
            }

            filePath = TryGetPresentationPath(presentation);
            return !string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath);
        }

        private static bool TryGetExistingPresentationPath(PowerPoint.Presentation presentation, out string filePath)
        {
            filePath = TryGetPresentationPath(presentation);
            return !string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath);
        }

        private static bool TryCreateColorReplacementWorkingCopy(
            PowerPoint.Presentation presentation,
            string filePath,
            out string tempFilePath)
        {
            tempFilePath = null;
            if (presentation == null)
            {
                return false;
            }

            try
            {
                tempFilePath = BuildSiblingTemporaryPath(filePath, "morphos-color-replace");
                dynamic dynamicPresentation = presentation;
                dynamicPresentation.SaveCopyAs(tempFilePath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoFalse);
                EnsureWritableFile(tempFilePath);
                WaitForFileAvailability(tempFilePath, FileAccess.ReadWrite, FileShare.None);
                return File.Exists(tempFilePath);
            }
            catch
            {
                DeleteTemporaryFile(tempFilePath);
                tempFilePath = null;

                if (!TryIsPresentationSaved(presentation)
                    || string.IsNullOrWhiteSpace(filePath)
                    || !File.Exists(filePath))
                {
                    return false;
                }

                try
                {
                    tempFilePath = BuildSiblingTemporaryPath(filePath, "morphos-color-replace");
                    File.Copy(filePath, tempFilePath, true);
                    EnsureWritableFile(tempFilePath);
                    WaitForFileAvailability(tempFilePath, FileAccess.ReadWrite, FileShare.None);
                    return File.Exists(tempFilePath);
                }
                catch
                {
                    DeleteTemporaryFile(tempFilePath);
                    tempFilePath = null;
                    return false;
                }
            }
        }

        private static string BuildSiblingTemporaryPath(string filePath, string prefix)
        {
            var extension = string.IsNullOrWhiteSpace(Path.GetExtension(filePath)) ? ".pptx" : Path.GetExtension(filePath);
            var directory = string.IsNullOrWhiteSpace(filePath) ? string.Empty : Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                return Path.Combine(
                    directory,
                    "~" + prefix + "-" + Guid.NewGuid().ToString("N") + extension);
            }

            return Path.Combine(
                Path.GetTempPath(),
                prefix + "-" + Guid.NewGuid().ToString("N") + extension);
        }

        private static IReadOnlyList<ColorReplacementInstruction> SanitizeColorInstructions(IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            if (instructions == null || instructions.Count == 0)
            {
                return Array.Empty<ColorReplacementInstruction>();
            }

            return instructions
                .Where(IsValidColorInstruction)
                .GroupBy(
                    x => x.UsageKind + "|" + (x.SourceHexValue ?? string.Empty),
                    StringComparer.OrdinalIgnoreCase)
                .Select(x => x.Last())
                .ToList();
        }

        private static bool IsValidColorInstruction(ColorReplacementInstruction instruction)
        {
            if (instruction == null || string.IsNullOrWhiteSpace(instruction.SourceHexValue))
            {
                return false;
            }

            if (instruction.UseThemeColor)
            {
                return !string.IsNullOrWhiteSpace(instruction.ThemeSchemeName);
            }

            return !string.IsNullOrWhiteSpace(instruction.ReplacementHexValue);
        }

        private static ColorReplacementSummary BuildColorReplacementSummary(ColorScanResult scanResult, int replacementCount, bool applied)
        {
            var items = scanResult == null || scanResult.Items == null
                ? Enumerable.Empty<ColorInventoryItem>()
                : scanResult.Items.Where(x => x.UsageKind != ColorUsageKind.ChartOverride);

            return new ColorReplacementSummary
            {
                ReplacementCount = replacementCount,
                RemainingDirectColors = items.Count(),
                RemainingUses = items.Sum(x => x.UsesCount),
                PreviewAvailable = scanResult != null,
                Applied = applied
            };
        }

        private static PresentationScanResult ClonePresentationScanResult(PresentationScanResult result)
        {
            return result == null
                ? null
                : new PresentationScanResult
                {
                    FontItems = CloneFontItems(result.FontItems),
                    ColorScanResult = CloneColorScanResult(result.ColorScanResult)
                };
        }

        private static IReadOnlyList<FontInventoryItem> CloneFontItems(IReadOnlyList<FontInventoryItem> items)
        {
            return (items ?? Array.Empty<FontInventoryItem>())
                .Where(item => item != null)
                .Select(CloneFontItem)
                .ToList();
        }

        private static FontInventoryItem CloneFontItem(FontInventoryItem item)
        {
            return item == null
                ? null
                : new FontInventoryItem
                {
                    FontName = item.FontName,
                    UsesCount = item.UsesCount,
                    EmbeddingStatus = item.EmbeddingStatus,
                    IsInstalled = item.IsInstalled,
                    IsEmbeddable = item.IsEmbeddable,
                    HasEmbeddableMetadata = item.HasEmbeddableMetadata,
                    HasPresentationMetadata = item.HasPresentationMetadata,
                    IsSubstituted = item.IsSubstituted,
                    HasSaveWarning = item.HasSaveWarning,
                    IsThemeFont = item.IsThemeFont,
                    Locations = CloneLocations(item.Locations)
                };
        }

        private static ColorScanResult CloneColorScanResult(ColorScanResult scanResult)
        {
            return scanResult == null
                ? null
                : new ColorScanResult
                {
                    Items = (scanResult.Items ?? Array.Empty<ColorInventoryItem>())
                        .Where(item => item != null)
                        .Select(CloneColorItem)
                        .ToList(),
                    ThemeColors = (scanResult.ThemeColors ?? Array.Empty<ThemeColorInfo>())
                        .Where(color => color != null)
                        .Select(CloneThemeColor)
                        .ToList()
                };
        }

        private static ColorInventoryItem CloneColorItem(ColorInventoryItem item)
        {
            return item == null
                ? null
                : new ColorInventoryItem
                {
                    UsageKind = item.UsageKind,
                    UsageKindLabel = item.UsageKindLabel,
                    HexValue = item.HexValue,
                    RgbValue = item.RgbValue,
                    UsesCount = item.UsesCount,
                    MatchesThemeColor = item.MatchesThemeColor,
                    MatchingThemeDisplayName = item.MatchingThemeDisplayName,
                    MatchingThemeSchemeName = item.MatchingThemeSchemeName,
                    Locations = CloneLocations(item.Locations)
                };
        }

        private static ThemeColorInfo CloneThemeColor(ThemeColorInfo color)
        {
            return color == null
                ? null
                : new ThemeColorInfo
                {
                    SchemeName = color.SchemeName,
                    DisplayName = color.DisplayName,
                    HexValue = color.HexValue
                };
        }

        private static IReadOnlyList<FontUsageLocation> CloneLocations(IReadOnlyList<FontUsageLocation> locations)
        {
            return (locations ?? Array.Empty<FontUsageLocation>())
                .Where(location => location != null)
                .Select(location => new FontUsageLocation
                {
                    Scope = location.Scope,
                    SlideIndex = location.SlideIndex,
                    ShapeId = location.ShapeId,
                    ScopeName = location.ScopeName,
                    ShapeName = location.ShapeName,
                    Label = location.Label,
                    IsSelectable = location.IsSelectable
                })
                .ToList();
        }

        private static void DeleteTemporaryFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            try
            {
                File.Delete(filePath);
            }
            catch
            {
            }
        }

        private static void CopyFileWithRetry(string sourcePath, string destinationPath)
        {
            Exception lastError = null;
            for (var attempt = 0; attempt < 80; attempt++)
            {
                try
                {
                    File.Copy(sourcePath, destinationPath, true);
                    return;
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    lastError = ex;
                    YieldToOffice();
                }
            }

            if (lastError != null)
            {
                throw lastError;
            }
        }

        private PowerPoint.Presentation OpenPresentationWithRetry(string filePath)
        {
            Exception lastError = null;
            for (var attempt = 0; attempt < 80; attempt++)
            {
                try
                {
                    WaitForFileAvailability(filePath, FileAccess.Read, FileShare.ReadWrite);
                    return _application.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
                }
                catch (Exception ex) when (ex is COMException || ex is IOException || ex is UnauthorizedAccessException)
                {
                    lastError = ex;
                    YieldToOffice();
                }
            }

            if (lastError != null)
            {
                throw lastError;
            }

            return null;
        }

        private static void ReplaceFileWithRetry(string sourcePath, string destinationPath, string backupPath)
        {
            Exception lastError = null;
            for (var attempt = 0; attempt < 80; attempt++)
            {
                try
                {
                    EnsureWritableFile(destinationPath);
                    EnsureWritableFile(sourcePath);
                    EnsureWritableFile(backupPath);
                    WaitForFileAvailability(sourcePath, FileAccess.ReadWrite, FileShare.None);
                    WaitForFileAvailability(destinationPath, FileAccess.ReadWrite, FileShare.None);

                    if (!string.IsNullOrWhiteSpace(backupPath) && File.Exists(backupPath))
                    {
                        File.Delete(backupPath);
                    }

                    File.Replace(sourcePath, destinationPath, backupPath, true);
                    return;
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    lastError = ex;
                    YieldToOffice();
                }
            }

            try
            {
                ReplaceFileByCopyWithRetry(sourcePath, destinationPath, backupPath);
                return;
            }
            catch (Exception fallbackEx) when (fallbackEx is IOException || fallbackEx is UnauthorizedAccessException)
            {
                lastError = fallbackEx;
            }

            if (lastError != null)
            {
                throw lastError;
            }
        }

        private static void WaitForFileAvailability(string filePath, FileAccess access, FileShare share)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            Exception lastError = null;
            for (var attempt = 0; attempt < 80; attempt++)
            {
                try
                {
                    using (new FileStream(filePath, FileMode.Open, access, share))
                    {
                        return;
                    }
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    lastError = ex;
                    YieldToOffice();
                }
            }

            if (lastError != null)
            {
                throw lastError;
            }
        }

        private static void WaitForPresentationReady(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return;
            }

            Exception lastError = null;
            for (var attempt = 0; attempt < 80; attempt++)
            {
                try
                {
                    var fullName = presentation.FullName;
                    var slideCount = presentation.Slides.Count;
                    _ = fullName;
                    _ = slideCount;
                    return;
                }
                catch (Exception ex) when (ex is COMException || ex is InvalidCastException)
                {
                    lastError = ex;
                    YieldToOffice();
                }
            }

            if (lastError != null)
            {
                throw lastError;
            }
        }

        private static PresentationWindowSnapshot CapturePresentationWindowSnapshot(PowerPoint.Presentation presentation)
        {
            var snapshot = new PresentationWindowSnapshot();
            if (presentation == null)
            {
                return snapshot;
            }

            PowerPoint.DocumentWindow window = null;
            try
            {
                if (presentation.Windows.Count > 0)
                {
                    window = presentation.Windows[1];
                }
            }
            catch
            {
            }

            if (window == null)
            {
                return snapshot;
            }

            try
            {
                snapshot.WindowState = window.WindowState;
                snapshot.HasWindowState = true;
            }
            catch
            {
            }

            try
            {
                snapshot.ViewType = window.ViewType;
                snapshot.HasViewType = true;
            }
            catch
            {
            }

            try
            {
                snapshot.SlideIndex = window.View.Slide.SlideIndex;
            }
            catch
            {
                try
                {
                    snapshot.SlideIndex = window.Selection.SlideRange[1].SlideIndex;
                }
                catch
                {
                }
            }

            return snapshot;
        }

        private static void ActivateReopenedFontPresentationWindow(
            PowerPoint.Presentation presentation,
            PresentationWindowSnapshot snapshot)
        {
            if (presentation == null)
            {
                return;
            }

            PowerPoint.DocumentWindow window = null;
            for (var attempt = 0; attempt < 80; attempt++)
            {
                try
                {
                    if (presentation.Windows.Count > 0)
                    {
                        window = presentation.Windows[1];
                        break;
                    }
                }
                catch
                {
                }

                try
                {
                    presentation.NewWindow();
                }
                catch
                {
                }

                YieldToOffice();
            }

            if (window == null)
            {
                return;
            }

            try
            {
                window.Activate();
            }
            catch
            {
            }

            if (snapshot != null && snapshot.HasWindowState)
            {
                try
                {
                    window.WindowState = snapshot.WindowState;
                }
                catch
                {
                }
            }

            var preferredViewType = snapshot != null && snapshot.HasViewType
                ? snapshot.ViewType
                : PowerPoint.PpViewType.ppViewNormal;

            try
            {
                window.ViewType = preferredViewType;
            }
            catch
            {
                try
                {
                    window.ViewType = PowerPoint.PpViewType.ppViewNormal;
                }
                catch
                {
                }
            }

            var targetSlideIndex = snapshot == null ? 0 : snapshot.SlideIndex;
            if (targetSlideIndex <= 0)
            {
                targetSlideIndex = 1;
            }

            try
            {
                if (presentation.Slides.Count >= targetSlideIndex)
                {
                    window.View.GotoSlide(targetSlideIndex);
                }
            }
            catch
            {
                try
                {
                    if (presentation.Slides.Count >= targetSlideIndex)
                    {
                        presentation.Slides[targetSlideIndex].Select();
                    }
                }
                catch
                {
                }
            }
        }

        private static bool IsFileMarkedReadOnly(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return false;
            }

            try
            {
                return (File.GetAttributes(filePath) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly;
            }
            catch
            {
                return false;
            }
        }

        private static void TryRestorePresentationFile(string destinationPath, string backupPath)
        {
            if (string.IsNullOrWhiteSpace(destinationPath)
                || string.IsNullOrWhiteSpace(backupPath)
                || !File.Exists(backupPath))
            {
                return;
            }

            try
            {
                EnsureWritableFile(destinationPath);
                WaitForFileAvailability(backupPath, FileAccess.Read, FileShare.Read);
                File.Copy(backupPath, destinationPath, true);
            }
            catch
            {
            }
        }

        private void TryOpenPresentation(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            try
            {
                _application.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            }
            catch
            {
            }
        }

        private void WaitForPresentationToClose(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            for (var attempt = 0; attempt < 80; attempt++)
            {
                if (!IsPresentationOpen(filePath))
                {
                    try
                    {
                        WaitForFileAvailability(filePath, FileAccess.ReadWrite, FileShare.None);
                        return;
                    }
                    catch (IOException)
                    {
                    }
                    catch (UnauthorizedAccessException)
                    {
                    }
                }

                YieldToOffice();
            }
        }

        private bool IsPresentationOpen(string filePath)
        {
            if (_application == null || string.IsNullOrWhiteSpace(filePath))
            {
                return false;
            }

            var normalizedPath = NormalizeFilePath(filePath);

            try
            {
                foreach (PowerPoint.Presentation openPresentation in _application.Presentations)
                {
                    if (string.Equals(NormalizeFilePath(TryGetPresentationPath(openPresentation)), normalizedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static void ReplaceFileByCopyWithRetry(string sourcePath, string destinationPath, string backupPath)
        {
            Exception lastError = null;
            for (var attempt = 0; attempt < 80; attempt++)
            {
                try
                {
                    EnsureWritableFile(destinationPath);
                    EnsureWritableFile(sourcePath);
                    EnsureWritableFile(backupPath);

                    if (!string.IsNullOrWhiteSpace(backupPath) && File.Exists(destinationPath))
                    {
                        File.Copy(destinationPath, backupPath, true);
                    }

                    if (File.Exists(destinationPath))
                    {
                        File.Delete(destinationPath);
                    }

                    File.Copy(sourcePath, destinationPath, true);
                    return;
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    lastError = ex;
                    YieldToOffice();
                }
            }

            if (lastError != null)
            {
                throw lastError;
            }
        }

        private static void EnsureWritableFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            try
            {
                var attributes = File.GetAttributes(filePath);
                if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    File.SetAttributes(filePath, attributes & ~FileAttributes.ReadOnly);
                }
            }
            catch
            {
            }
        }

        private static void RestoreReadOnlyAttribute(string filePath, bool isReadOnly)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            try
            {
                var attributes = File.GetAttributes(filePath);
                if (isReadOnly)
                {
                    File.SetAttributes(filePath, attributes | FileAttributes.ReadOnly);
                }
                else if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    File.SetAttributes(filePath, attributes & ~FileAttributes.ReadOnly);
                }
            }
            catch
            {
            }
        }

        private static void YieldToOffice()
        {
            try
            {
                WinForms.Application.DoEvents();
            }
            catch
            {
            }

            Thread.Sleep(RetryDelayMilliseconds);

            try
            {
                WinForms.Application.DoEvents();
            }
            catch
            {
            }
        }

        private static string NormalizeFilePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFullPath(filePath).Trim();
            }
            catch
            {
                return filePath.Trim();
            }
        }

        private static bool TryReadBooleanProperty(object target, string propertyName, out bool value)
        {
            value = false;
            object rawValue;
            if (!TryReadProperty(target, propertyName, out rawValue))
            {
                return false;
            }

            value = SafeMsoBoolean(rawValue);
            return true;
        }

        private static bool TryReadStringProperty(object target, string propertyName, out string value)
        {
            value = string.Empty;
            object rawValue;
            if (!TryReadProperty(target, propertyName, out rawValue))
            {
                return false;
            }

            value = Convert.ToString(rawValue);
            return true;
        }

        private static bool TryReadProperty(object target, string propertyName, out object value)
        {
            value = null;
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                value = target.GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    target,
                    null);
                return true;
            }
            catch
            {
                try
                {
                    value = ((dynamic)target)[propertyName];
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }

        private static bool TryWriteProperty(object target, string propertyName, object value)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                target.GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.SetProperty,
                    null,
                    target,
                    new[] { value });
                return true;
            }
            catch
            {
                try
                {
                    ((dynamic)target)[propertyName] = value;
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }

        private static bool SafeMsoBoolean(object value)
        {
            try
            {
                return Convert.ToInt32(value) == (int)MsoTriState.msoTrue;
            }
            catch
            {
                return false;
            }
        }

        private static int? SafeGetShapeId(PowerPoint.Shape shape)
        {
            try
            {
                return shape.Id;
            }
            catch
            {
                return null;
            }
        }

        private static string SafeGetShapeName(PowerPoint.Shape shape)
        {
            try
            {
                return shape.Name;
            }
            catch
            {
                return "Shape";
            }
        }

        private static string NormalizeFontName(string fontName)
        {
            return FontNameNormalizer.Normalize(fontName);
        }

        private static string BuildLocationIdKey(PresentationScope scope, int? slideIndex, int? shapeId, string shapeName, string scopeLabel)
        {
            return ((int)scope).ToString()
                + "|"
                + (slideIndex ?? 0).ToString()
                + "|"
                + (shapeId ?? 0).ToString()
                + "|"
                + NormalizeFontName(shapeName)
                + "|"
                + NormalizeFontName(scopeLabel);
        }

        private static string BuildLocationLabelKey(string label)
        {
            return NormalizeFontName(label);
        }

        private static bool IsThemeFontName(string fontName)
        {
            return !string.IsNullOrWhiteSpace(fontName)
                && fontName.StartsWith("+", StringComparison.OrdinalIgnoreCase);
        }

        private static FontEmbeddingStatus ResolveEmbeddingStatus(FontScanSnapshot snapshot, PresentationFontMetadata font)
        {
            if (snapshot == null || font == null || string.IsNullOrWhiteSpace(font.FontName))
            {
                return FontEmbeddingStatus.Unknown;
            }

            if (snapshot.HasEmbeddedFontDataKnown)
            {
                return snapshot.EmbeddedFontNames.Contains(font.FontName)
                    ? (snapshot.SaveSubsetFonts ? FontEmbeddingStatus.Subset : FontEmbeddingStatus.Yes)
                    : FontEmbeddingStatus.No;
            }

            if (snapshot.HasPackageEmbeddingSettings && !snapshot.RequestsEmbeddedFonts)
            {
                return FontEmbeddingStatus.No;
            }

            if (font.HasEmbeddedMetadata && font.IsEmbedded)
            {
                return snapshot.SaveSubsetFonts
                    ? FontEmbeddingStatus.Subset
                    : FontEmbeddingStatus.Yes;
            }

            return FontEmbeddingStatus.Unknown;
        }

        private static bool ShouldCreateAccumulatorFromPresentationFont(PresentationFontMetadata font)
        {
            return font != null
                && !string.IsNullOrWhiteSpace(font.FontName)
                && !IsThemeFontName(font.FontName);
        }

        private static bool ShouldSuppressThemeFontUsage(string fontName, FontUsageLocation location)
        {
            if (!IsThemeFontName(fontName) || location == null)
            {
                return false;
            }

            return location.Scope != PresentationScope.SlideMaster
                && location.Scope != PresentationScope.CustomLayout;
        }

        private static bool LooksLikeFontAvailabilityFailure(Exception ex)
        {
            var message = ex == null ? string.Empty : ex.Message ?? string.Empty;
            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            return message.IndexOf("font", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("embed", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("available", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string NormalizeReplacementFont(string fontName)
        {
            return FontNameNormalizer.NormalizeReplacementFont(fontName);
        }

        private static string TryGetPresentationPath(PowerPoint.Presentation presentation)
        {
            try
            {
                return presentation.FullName;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool TryIsPresentationSaved(PowerPoint.Presentation presentation)
        {
            try
            {
                return Convert.ToInt32(presentation.Saved) == (int)MsoTriState.msoTrue;
            }
            catch
            {
                return false;
            }
        }

        private static string BuildLocationLabel(PresentationScope scope, int? slideIndex, PowerPoint.Shape shape, string scopeLabel)
        {
            var shapeName = SafeGetShapeName(shape);

            if (scope == PresentationScope.Slide && slideIndex.HasValue)
            {
                return "Slide " + slideIndex.Value + " - " + shapeName;
            }

            return scopeLabel + " - " + shapeName;
        }

        private static PowerPoint.Shape FindShape(PowerPoint.Shapes shapes, int? shapeId, string shapeName)
        {
            if (shapes == null)
            {
                return null;
            }

            for (var i = 1; i <= shapes.Count; i++)
            {
                var match = FindShapeRecursive(shapes[i], shapeId, shapeName);
                if (match != null)
                {
                    return match;
                }
            }

            return null;
        }

        private static PowerPoint.Shape FindShapeRecursive(PowerPoint.Shape shape, int? shapeId, string shapeName)
        {
            if (shape == null)
            {
                return null;
            }

            try
            {
                if (shapeId.HasValue && shape.Id == shapeId.Value)
                {
                    return shape;
                }
            }
            catch
            {
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(shapeName) && string.Equals(shape.Name, shapeName, StringComparison.OrdinalIgnoreCase))
                {
                    return shape;
                }
            }
            catch
            {
            }

            try
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    for (var i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        var match = FindShapeRecursive(shape.GroupItems[i], shapeId, shapeName);
                        if (match != null)
                        {
                            return match;
                        }
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private FontReplacementResult ValidateEmbeddedSaveState(PowerPoint.Presentation presentation)
        {
            var result = new FontReplacementResult();
            if (presentation == null)
            {
                return result;
            }

            var installedFonts = GetInstalledFontsSet();
            var snapshot = _fontScanSessionCache.GetOrCreateSnapshot(presentation);
            IDictionary<string, FontAccumulator> accumulators = null;
            var usedPackageScan = false;

            using (var packageScanSource = AcquirePackageScanSource(presentation, snapshot))
            {
                if (packageScanSource != null)
                {
                    var packageUsages = PresentationPackageInspector.ReadFontUsages(packageScanSource.FilePath, null, CancellationToken.None);
                    if (packageUsages != null)
                    {
                        accumulators = BuildAccumulatorsFromPackage(packageUsages);
                        usedPackageScan = true;
                    }
                }

                if (accumulators == null)
                {
                    accumulators = new Dictionary<string, FontAccumulator>(StringComparer.OrdinalIgnoreCase);
                    foreach (var scope in BuildScanScopes(presentation))
                    {
                        ScanShapes(scope.Shapes, scope.Scope, scope.Index, scope.Label, accumulators);
                    }
                }
            }

            ApplyPresentationFontMetadata(snapshot, accumulators, !usedPackageScan);

            if (usedPackageScan && ShouldCaptureDisplayedFontMap(snapshot, accumulators))
            {
                ApplySubstitutionStateFromPowerPoint(accumulators, CaptureDisplayedFontMap(presentation));
            }

            var saveValidation = ShouldPerformEmbeddedSaveValidation(snapshot, accumulators)
                ? CreateEmbeddedSaveValidation(presentation)
                : null;
            ApplyValidatedEmbeddingState(snapshot, accumulators, saveValidation);
            ApplyWarningState(accumulators, saveValidation);

            result.RemainingSubstitutedFonts = accumulators.Values
                .Where(x => x.IsSubstituted)
                .Select(x => x.FontName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            result.RemainingNonEmbeddableFonts = accumulators.Values
                .Where(x => x.HasSaveWarning && !x.IsSubstituted)
                .Select(x => x.FontName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            ApplySaveValidationResult(saveValidation, result);
            return result;
        }

        private sealed class PackageScanSource : IDisposable
        {
            public PackageScanSource(string filePath, bool deleteAfterUse)
            {
                FilePath = filePath;
                DeleteAfterUse = deleteAfterUse;
            }

            public string FilePath { get; }

            public bool DeleteAfterUse { get; }

            public void Dispose()
            {
                if (!DeleteAfterUse || string.IsNullOrWhiteSpace(FilePath) || !File.Exists(FilePath))
                {
                    return;
                }

                try
                {
                    File.Delete(FilePath);
                }
                catch
                {
                }
            }
        }

        private sealed class EmbeddedSaveValidation
        {
            public EmbeddedSaveValidation()
            {
                EmbeddedFontNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            public bool CopySucceeded { get; set; }

            public bool HasEmbeddedFontNames { get; set; }

            public bool DetectedEmbeddedFontData { get; set; }

            public bool LooksLikeFontAvailabilityFailure { get; set; }

            public ISet<string> EmbeddedFontNames { get; set; }
        }

        private sealed class PresentationWindowSnapshot
        {
            public bool HasWindowState { get; set; }

            public PowerPoint.PpWindowState WindowState { get; set; }

            public bool HasViewType { get; set; }

            public PowerPoint.PpViewType ViewType { get; set; }

            public int SlideIndex { get; set; }
        }

        private sealed class DisplayedFontMap
        {
            public DisplayedFontMap()
            {
                ById = new Dictionary<string, ISet<string>>(StringComparer.OrdinalIgnoreCase);
                ByLabel = new Dictionary<string, ISet<string>>(StringComparer.OrdinalIgnoreCase);
            }

            public IDictionary<string, ISet<string>> ById { get; }

            public IDictionary<string, ISet<string>> ByLabel { get; }

            public void Add(string key, IEnumerable<string> fonts, bool useIdIndex)
            {
                if (string.IsNullOrWhiteSpace(key) || fonts == null)
                {
                    return;
                }

                var snapshot = new HashSet<string>(
                    fonts.Where(x => !string.IsNullOrWhiteSpace(x)).Select(NormalizeFontName),
                    StringComparer.OrdinalIgnoreCase);

                if (snapshot.Count == 0)
                {
                    return;
                }

                if (useIdIndex)
                {
                    ById[key] = snapshot;
                }
                else
                {
                    ByLabel[key] = snapshot;
                }
            }
        }

        private sealed class ScanScope
        {
            public ScanScope(PresentationScope scope, int? index, string label, PowerPoint.Shapes shapes)
            {
                Scope = scope;
                Index = index;
                Label = label;
                Shapes = shapes;
            }

            public PresentationScope Scope { get; }

            public int? Index { get; }

            public string Label { get; }

            public PowerPoint.Shapes Shapes { get; }
        }

        private sealed class ColorReplacementTarget
        {
            public ColorReplacementTarget(FontUsageLocation location)
            {
                SlideIndex = location == null ? null : location.SlideIndex;
                ShapeId = location == null ? null : location.ShapeId;
                ShapeName = location == null ? string.Empty : location.ShapeName;
                Instructions = new List<ColorReplacementInstruction>();
            }

            public int? SlideIndex { get; }

            public int? ShapeId { get; }

            public string ShapeName { get; }

            public List<ColorReplacementInstruction> Instructions { get; }
        }

        private sealed class FontAccumulator
        {
            public FontAccumulator(string fontName)
            {
                FontName = fontName;
                EmbeddingStatus = FontEmbeddingStatus.Unknown;
                Locations = new List<FontUsageLocation>();
            }

            public string FontName { get; }

            public int UsesCount { get; set; }

            public bool HasPresentationMetadata { get; set; }

            public bool IsInstalled { get; set; }

            public bool IsEmbeddable { get; set; }

            public bool HasEmbeddableMetadata { get; set; }

            public bool IsSubstituted { get; set; }

            public bool HasSaveWarning { get; set; }

            public FontEmbeddingStatus EmbeddingStatus { get; set; }

            public List<FontUsageLocation> Locations { get; }

            public bool IsThemeFont => IsThemeFontName(FontName);

            public FontInventoryItem ToInventoryItem(ISet<string> installedFonts)
            {
                var isInstalled = IsThemeFont
                    || (HasPresentationMetadata
                        ? IsInstalled
                        : installedFonts.Contains(FontName));

                return new FontInventoryItem
                {
                    FontName = FontName,
                    UsesCount = UsesCount,
                    EmbeddingStatus = EmbeddingStatus,
                    HasPresentationMetadata = HasPresentationMetadata,
                    IsEmbeddable = IsEmbeddable,
                    HasEmbeddableMetadata = HasEmbeddableMetadata,
                    IsSubstituted = IsSubstituted,
                    HasSaveWarning = HasSaveWarning,
                    IsInstalled = isInstalled,
                    IsThemeFont = IsThemeFont,
                    Locations = Locations
                        .GroupBy(x => x.Label, StringComparer.OrdinalIgnoreCase)
                        .Select(x => x.First())
                        .ToList()
                };
            }
        }

    }
}
