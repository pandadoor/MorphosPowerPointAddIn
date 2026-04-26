using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using MorphosPowerPointAddIn.Dialogs;
using MorphosPowerPointAddIn.Models;
using MorphosPowerPointAddIn.Services;
using MorphosPowerPointAddIn.Utilities;

namespace MorphosPowerPointAddIn.ViewModels
{
    public sealed class FontsPaneViewModel : BindableBase
    {
        private const int SearchDebounceMilliseconds = 160;
        private const int RpcECallRejected = unchecked((int)0x80010001);
        private const int RpcEServerCallRetryLater = unchecked((int)0x8001010A);
        private readonly PowerPointPresentationService _presentationService;
        private readonly SemaphoreSlim _scanGate = new SemaphoreSlim(1, 1);
        private CancellationTokenSource _scanCts;
        private DispatcherTimer _colorSearchDebounceTimer;
        private Dispatcher _dispatcher;
        private DispatcherTimer _fontSearchDebounceTimer;
        private bool _hasCompletedScan;
        private bool _hasColorResults;
        private bool _hasFontResults;
        private bool _hasSelectedColor;
        private bool _hasSelectedFont;
        private bool _hasSelectedNode;
        private string _colorSearchText = string.Empty;
        private string _embeddingSummary = "No fonts";
        private string _fontSearchText = string.Empty;
        private int _fontWarnings;
        private string _inspectorBody = "Select a font or color row to inspect its details.";
        private string _inspectorTitle = "Inspector";
        private bool _isScanning;
        private bool _lastScanSucceeded;
        private double _progressValue;
        private int _saveWarningFonts;
        private int _scanVersion;
        private TreeNodeViewModel _selectedNode;
        private int _substitutedFonts;
        private string _statusText = "Open the Morphos pane to scan the active presentation.";
        private int _themeMatchColors;
        private int _totalColorUses;
        private int _totalDirectColors;
        private int _totalFonts;
        private int _totalUses;
        private IReadOnlyList<FontInventoryItem> _fontItems = Array.Empty<FontInventoryItem>();
        private IReadOnlyList<FontNodeViewModel> _fontNodes = Array.Empty<FontNodeViewModel>();
        private readonly Dictionary<string, FontNodeViewModel> _fontNodeCache = new Dictionary<string, FontNodeViewModel>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ColorNodeViewModel> _colorNodeCache = new Dictionary<string, ColorNodeViewModel>(StringComparer.OrdinalIgnoreCase);
        private ColorScanResult _colorScanResult = new ColorScanResult();

        public FontsPaneViewModel(PowerPointPresentationService presentationService)
        {
            _presentationService = presentationService;
            RootNodes = new RangeObservableCollection<TreeNodeViewModel>();
            ColorGroups = new RangeObservableCollection<TreeNodeViewModel>();
        }

        public RangeObservableCollection<TreeNodeViewModel> RootNodes { get; }

        public RangeObservableCollection<TreeNodeViewModel> ColorGroups { get; }

        public void AttachDispatcher(Dispatcher dispatcher)
        {
            if (_dispatcher == null && dispatcher != null)
            {
                _dispatcher = dispatcher;
                EnsureSearchDebounceTimers();
            }
        }

        public bool IsScanning
        {
            get => _isScanning;
            private set => SetProperty(ref _isScanning, value);
        }

        public bool HasCompletedScan => _hasCompletedScan;

        public bool LastScanSucceeded => _lastScanSucceeded;

        public TreeNodeViewModel SelectedNode
        {
            get => _selectedNode;
            set
            {
                if (SetProperty(ref _selectedNode, value))
                {
                    UpdateSelectionState(value);
                    UpdateInspector(value);
                }
            }
        }

        public double ProgressValue
        {
            get => _progressValue;
            private set => SetProperty(ref _progressValue, value);
        }

        public string StatusText
        {
            get => _statusText;
            private set => SetProperty(ref _statusText, value);
        }

        public int TotalFonts
        {
            get => _totalFonts;
            private set => SetProperty(ref _totalFonts, value);
        }

        public int SubstitutedFonts
        {
            get => _substitutedFonts;
            private set => SetProperty(ref _substitutedFonts, value);
        }

        public int SaveWarningFonts
        {
            get => _saveWarningFonts;
            private set => SetProperty(ref _saveWarningFonts, value);
        }

        public int FontWarnings
        {
            get => _fontWarnings;
            private set => SetProperty(ref _fontWarnings, value);
        }

        public int TotalUses
        {
            get => _totalUses;
            private set => SetProperty(ref _totalUses, value);
        }

        public int TotalDirectColors
        {
            get => _totalDirectColors;
            private set => SetProperty(ref _totalDirectColors, value);
        }

        public int TotalColorUses
        {
            get => _totalColorUses;
            private set => SetProperty(ref _totalColorUses, value);
        }

        public int ThemeMatchColors
        {
            get => _themeMatchColors;
            private set => SetProperty(ref _themeMatchColors, value);
        }

        public bool HasSelectedNode
        {
            get => _hasSelectedNode;
            private set => SetProperty(ref _hasSelectedNode, value);
        }

        public bool HasSelectedFont
        {
            get => _hasSelectedFont;
            private set => SetProperty(ref _hasSelectedFont, value);
        }

        public bool HasSelectedColor
        {
            get => _hasSelectedColor;
            private set => SetProperty(ref _hasSelectedColor, value);
        }

        public bool HasFontResults
        {
            get => _hasFontResults;
            private set => SetProperty(ref _hasFontResults, value);
        }

        public bool HasColorResults
        {
            get => _hasColorResults;
            private set => SetProperty(ref _hasColorResults, value);
        }

        public string EmbeddingSummary
        {
            get => _embeddingSummary;
            private set => SetProperty(ref _embeddingSummary, value);
        }

        public string InspectorTitle
        {
            get => _inspectorTitle;
            private set => SetProperty(ref _inspectorTitle, value);
        }

        public string InspectorBody
        {
            get => _inspectorBody;
            private set => SetProperty(ref _inspectorBody, value);
        }

        public string FontSearchText
        {
            get => _fontSearchText;
            set
            {
                if (SetProperty(ref _fontSearchText, value))
                {
                    QueueFontNodeRebuild();
                }
            }
        }

        public string ColorSearchText
        {
            get => _colorSearchText;
            set
            {
                if (SetProperty(ref _colorSearchText, value))
                {
                    QueueColorGroupRebuild();
                }
            }
        }

        public async Task ScanAsync(bool showErrors = true)
        {
            using (OfficeBusyMessageFilter.Register())
            {
                var scanVersion = Interlocked.Increment(ref _scanVersion);
                var nextCts = new CancellationTokenSource();
                var previousCts = Interlocked.Exchange(ref _scanCts, nextCts);
                if (previousCts != null)
                {
                    TryCancel(previousCts);
                }

                await InvokeOnUiAsync(() =>
                {
                    _lastScanSucceeded = false;
                    IsScanning = true;
                    ProgressValue = 0;
                    StatusText = "Scanning presentation cleanup data...";
                }).ConfigureAwait(false);

                await _scanGate.WaitAsync().ConfigureAwait(false);

                try
                {
                    double fontPercentage = 0;
                    double colorPercentage = 0;

                    var fontProgress = new Progress<ScanProgressInfo>(info =>
                    {
                        if (scanVersion != Volatile.Read(ref _scanVersion))
                        {
                            return;
                        }

                        InvokeOnUi(() =>
                        {
                            fontPercentage = info == null ? 0 : info.Percentage;
                            StatusText = string.IsNullOrWhiteSpace(info?.Message) ? "Scanning fonts..." : "Fonts: " + info.Message;
                            ProgressValue = CombineProgress(fontPercentage, colorPercentage);
                        });
                    });

                    var colorProgress = new Progress<ScanProgressInfo>(info =>
                    {
                        if (scanVersion != Volatile.Read(ref _scanVersion))
                        {
                            return;
                        }

                        InvokeOnUi(() =>
                        {
                            colorPercentage = info == null ? 0 : info.Percentage;
                            StatusText = string.IsNullOrWhiteSpace(info?.Message) ? "Scanning colors..." : "Colors: " + info.Message;
                            ProgressValue = CombineProgress(fontPercentage, colorPercentage);
                        });
                    });

                    var scanResult = await ExecutePresentationActionWithRetryAsync(
                        token => InvokeOnUiAsync(() => _presentationService.AnalyzeActivePresentationAsync(fontProgress, colorProgress, token)),
                        nextCts.Token).ConfigureAwait(false);
                    if (scanVersion != Volatile.Read(ref _scanVersion))
                    {
                        return;
                    }

                    await InvokeOnUiAsync(() =>
                    {
                        _fontItems = scanResult == null || scanResult.FontItems == null
                            ? Array.Empty<FontInventoryItem>()
                            : scanResult.FontItems;
                        _colorScanResult = scanResult == null || scanResult.ColorScanResult == null
                            ? new ColorScanResult()
                            : scanResult.ColorScanResult;

                        var selectedKey = BuildNodeKey(SelectedNode);

                        ApplyFontSummary(_fontItems);
                        ApplyColorSummary(_colorScanResult.Items);

                        SyncFontNodeCache(_fontItems);
                        SyncColorNodeCache(_colorScanResult.Items);
                        RebuildFontNodes();
                        RebuildColorGroups();
                        RestoreSelection(selectedKey);

                        if (SelectedNode == null)
                        {
                            UpdateInspector(null);
                        }

                        _hasCompletedScan = true;
                        _lastScanSucceeded = true;
                        StatusText = BuildStatusSummary();
                        ProgressValue = 100;
                    }).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    if (scanVersion == Volatile.Read(ref _scanVersion))
                    {
                        await InvokeOnUiAsync(() =>
                        {
                            _lastScanSucceeded = false;
                            StatusText = "Scan cancelled.";
                        }).ConfigureAwait(false);
                    }
                }
                catch (Exception ex)
                {
                    await InvokeOnUiAsync(() =>
                    {
                        _lastScanSucceeded = false;
                        if (showErrors)
                        {
                            ErrorReporter.Show("Morphos could not scan the active presentation.", ex);
                        }

                        StatusText = "Scan failed.";
                    }).ConfigureAwait(false);
                }
                finally
                {
                    _scanGate.Release();

                    if (ReferenceEquals(_scanCts, nextCts))
                    {
                        _scanCts = null;
                    }

                    nextCts.Dispose();

                    if (scanVersion == Volatile.Read(ref _scanVersion))
                    {
                        await InvokeOnUiAsync(() => IsScanning = false).ConfigureAwait(false);
                    }
                }
            }
        }

        public void ShowInPowerPoint(TreeNodeViewModel node)
        {
            if (node is FontUsageNodeViewModel fontUsageNode)
            {
                _presentationService.ShowUsage(fontUsageNode.Location);
                return;
            }

            if (node is FontNodeViewModel fontNode)
            {
                var firstLocation = fontNode.Item == null || fontNode.Item.Locations == null
                    ? null
                    : fontNode.Item.Locations.FirstOrDefault();
                if (firstLocation != null)
                {
                    _presentationService.ShowUsage(firstLocation);
                }

                return;
            }

            if (node is ColorUsageNodeViewModel colorUsageNode)
            {
                _presentationService.ShowUsage(colorUsageNode.Location);
                return;
            }

            if (node is ColorNodeViewModel colorNode)
            {
                var firstLocation = colorNode.Item == null || colorNode.Item.Locations == null
                    ? null
                    : colorNode.Item.Locations.FirstOrDefault();
                if (firstLocation != null)
                {
                    _presentationService.ShowUsage(firstLocation);
                }
            }
        }

        public async Task ReplaceFontAsync(FontNodeViewModel fontNode)
        {
            if (fontNode == null)
            {
                return;
            }

            var sourceFonts = _fontItems
                .Select(node => node.FontName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var dialog = new ReplaceFontsDialog(
                sourceFonts,
                new[] { fontNode.FontName },
                _presentationService.GetReplacementTargets(new[] { fontNode.FontName }));
            DialogWindowHelper.AttachToPowerPoint(dialog, Globals.ThisAddIn.Application);

            if (dialog.ShowDialog() == true)
            {
                FontReplacementResult replacementResult = null;
                IDisposable refreshSuspension = null;
                var requiresTaskPaneRecovery = false;
                try
                {
                    refreshSuspension = Globals.ThisAddIn == null ? null : Globals.ThisAddIn.SuspendAutoRefresh();
                    replacementResult = await ExecutePresentationActionWithRetryAsync(
                        () => _presentationService.ReplaceFonts(dialog.SelectedSourceFontNames, dialog.SelectedFontName)).ConfigureAwait(true);

                    if (_presentationService.LastMutationReloadedPresentation)
                    {
                        requiresTaskPaneRecovery = true;
                    }
                    else
                    {
                        await RefreshFontsAsync(dialog.SelectedFontName).ConfigureAwait(true);
                    }

                    if (replacementResult != null && replacementResult.HasWarnings)
                    {
                        MessageBox.Show(
                            replacementResult.WarningMessage,
                            "Morphos",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                }
                finally
                {
                    refreshSuspension?.Dispose();

                    if (_presentationService.LastMutationReloadedPresentation && requiresTaskPaneRecovery)
                    {
                        Globals.ThisAddIn?.RecoverTaskPaneAfterPresentationMutation(true);
                    }
                }
            }
        }

        public async Task ReplaceColorsAsync(ColorNodeViewModel colorNode = null)
        {
            if (_colorScanResult == null || _colorScanResult.Items.Count == 0)
            {
                MessageBox.Show(
                    "No direct RGB colors were found in the active presentation.",
                    "Morphos",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var dialog = new ReplaceColorsDialog(
                _colorScanResult.Items,
                _colorScanResult.ThemeColors,
                colorNode == null ? null : colorNode.Item);
            DialogWindowHelper.AttachToPowerPoint(dialog, Globals.ThisAddIn.Application);

            if (dialog.ShowDialog() == true)
            {
                using (OfficeBusyMessageFilter.Register())
                {
                    IDisposable refreshSuspension = null;
                    try
                    {
                        refreshSuspension = Globals.ThisAddIn == null ? null : Globals.ThisAddIn.SuspendAutoRefresh();
                        var replacementResult = await ExecutePresentationActionWithRetryAsync(
                            () => _presentationService.ReplaceColors(dialog.SelectedInstructions)).ConfigureAwait(true);
                        if (replacementResult == null || replacementResult.ReplacementCount <= 0)
                        {
                            MessageBox.Show(
                                "Morphos did not find any saved direct RGB uses to replace for that color.",
                                "Morphos",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
                            return;
                        }

                        ApplyLocalColorReplacementPreview(dialog.SelectedInstructions);
                        StatusText = "Replaced " + replacementResult.ReplacementCount + (replacementResult.ReplacementCount == 1 ? " direct RGB use. " : " direct RGB uses. ") + BuildStatusSummary();
                        await RefreshColorsAsync().ConfigureAwait(true);
                    }
                    finally
                    {
                        refreshSuspension?.Dispose();
                        if (_presentationService.LastMutationReloadedPresentation)
                        {
                            Globals.ThisAddIn?.RecoverTaskPaneAfterPresentationMutation();
                        }
                    }
                }
            }
        }

        public async Task UpdateEmbeddingAsync(FontNodeViewModel fontNode, FontEmbeddingStatus status)
        {
            if (fontNode == null)
            {
                return;
            }

            try
            {
                await ExecutePresentationActionWithRetryAsync(() =>
                {
                    _presentationService.UpdateEmbedding(status);
                    return true;
                }).ConfigureAwait(true);
                await RefreshFontsAsync().ConfigureAwait(true);
            }
            finally
            {
                if (_presentationService.LastMutationReloadedPresentation)
                {
                    Globals.ThisAddIn?.RecoverTaskPaneAfterPresentationMutation();
                }
            }
        }

        private void ApplyFontSummary(IReadOnlyList<FontInventoryItem> items)
        {
            var safeItems = items ?? Array.Empty<FontInventoryItem>();
            TotalFonts = safeItems.Count;
            TotalUses = safeItems.Sum(x => x.UsesCount);
            SubstitutedFonts = safeItems.Count(x => x.IsSubstituted);
            SaveWarningFonts = safeItems.Count(x => x.HasSaveWarning);
            FontWarnings = safeItems.Count(x => x.IsSubstituted || x.HasSaveWarning);
            EmbeddingSummary = BuildEmbeddingSummary(safeItems);
            HasFontResults = safeItems.Count > 0;
        }

        private void ApplyColorSummary(IReadOnlyList<ColorInventoryItem> items)
        {
            var safeItems = items ?? Array.Empty<ColorInventoryItem>();
            TotalDirectColors = safeItems.Count(x => x.UsageKind != ColorUsageKind.ChartOverride);
            TotalColorUses = safeItems
                .Where(x => x.UsageKind != ColorUsageKind.ChartOverride)
                .Sum(x => x.UsesCount);
            ThemeMatchColors = safeItems.Count(x => x.MatchesThemeColor && x.UsageKind != ColorUsageKind.ChartOverride);
            HasColorResults = TotalDirectColors > 0;
        }

        private void ApplyLocalColorReplacementPreview(IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            if (_colorScanResult == null || instructions == null || instructions.Count == 0)
            {
                return;
            }

            var selectedKey = BuildNodeKey(SelectedNode);
            _colorScanResult = new ColorScanResult
            {
                Items = BuildUpdatedColorInventoryItems(_colorScanResult.Items, _colorScanResult.ThemeColors, instructions),
                ThemeColors = _colorScanResult.ThemeColors
            };

            ApplyColorSummary(_colorScanResult.Items);
            SyncColorNodeCache(_colorScanResult.Items);
            RebuildColorGroups();
            RestoreSelection(selectedKey);

            if (SelectedNode == null)
            {
                UpdateInspector(null);
            }

            ProgressValue = 100;
        }

        private static IReadOnlyList<ColorInventoryItem> BuildUpdatedColorInventoryItems(
            IReadOnlyList<ColorInventoryItem> existingItems,
            IReadOnlyList<ThemeColorInfo> themeColors,
            IReadOnlyList<ColorReplacementInstruction> instructions)
        {
            var safeItems = existingItems ?? Array.Empty<ColorInventoryItem>();
            if (safeItems.Count == 0 || instructions == null || instructions.Count == 0)
            {
                return safeItems;
            }

            var instructionLookup = instructions
                .Where(x => x != null)
                .GroupBy(x => BuildColorInstructionKey(x.UsageKind, x.SourceHexValue), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.Last(), StringComparer.OrdinalIgnoreCase);

            var themeLookup = (themeColors ?? Array.Empty<ThemeColorInfo>())
                .Where(x => x != null && !string.IsNullOrWhiteSpace(x.HexValue))
                .GroupBy(x => NormalizeColorHex(x.HexValue), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);

            var updatedItems = new List<ColorInventoryItem>(safeItems.Count);
            foreach (var item in safeItems)
            {
                if (item == null)
                {
                    continue;
                }

                ColorReplacementInstruction instruction;
                if (!instructionLookup.TryGetValue(BuildColorInstructionKey(item.UsageKind, item.HexValue), out instruction))
                {
                    updatedItems.Add(item);
                    continue;
                }

                if (instruction.UseThemeColor)
                {
                    continue;
                }

                var replacementHex = NormalizeColorHex(instruction.ReplacementHexValue);
                if (string.IsNullOrWhiteSpace(replacementHex))
                {
                    updatedItems.Add(item);
                    continue;
                }

                ThemeColorInfo matchingTheme;
                var hasThemeMatch = themeLookup.TryGetValue(replacementHex, out matchingTheme);
                updatedItems.Add(new ColorInventoryItem
                {
                    UsageKind = item.UsageKind,
                    UsageKindLabel = item.UsageKindLabel,
                    HexValue = replacementHex,
                    RgbValue = BuildRgbValue(replacementHex),
                    UsesCount = item.UsesCount,
                    MatchesThemeColor = hasThemeMatch,
                    MatchingThemeDisplayName = hasThemeMatch ? matchingTheme.DisplayName : string.Empty,
                    MatchingThemeSchemeName = hasThemeMatch ? matchingTheme.SchemeName : string.Empty,
                    Locations = item.Locations
                });
            }

            return updatedItems
                .GroupBy(x => BuildColorInstructionKey(x.UsageKind, x.HexValue), StringComparer.OrdinalIgnoreCase)
                .Select(group =>
                {
                    var first = group.First();
                    return new ColorInventoryItem
                    {
                        UsageKind = first.UsageKind,
                        UsageKindLabel = first.UsageKindLabel,
                        HexValue = first.HexValue,
                        RgbValue = first.RgbValue,
                        UsesCount = group.Sum(x => x.UsesCount),
                        MatchesThemeColor = first.MatchesThemeColor,
                        MatchingThemeDisplayName = first.MatchingThemeDisplayName,
                        MatchingThemeSchemeName = first.MatchingThemeSchemeName,
                        Locations = group
                            .SelectMany(x => x.Locations ?? Array.Empty<FontUsageLocation>())
                            .GroupBy(BuildLocationKey, StringComparer.OrdinalIgnoreCase)
                            .Select(x => x.First())
                            .ToList()
                    };
                })
                .OrderBy(x => x.UsageKindLabel, StringComparer.OrdinalIgnoreCase)
                .ThenByDescending(x => x.UsesCount)
                .ThenBy(x => x.HexValue, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string BuildColorInstructionKey(ColorUsageKind usageKind, string hexValue)
        {
            return ((int)usageKind).ToString() + "|" + NormalizeColorHex(hexValue);
        }

        private static string NormalizeColorHex(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Trim().TrimStart('#').ToUpperInvariant();
            return normalized.Length == 6 ? normalized : string.Empty;
        }

        private static string BuildRgbValue(string hexValue)
        {
            var normalized = NormalizeColorHex(hexValue);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return "0, 0, 0";
            }

            return Convert.ToInt32(normalized.Substring(0, 2), 16) + ", "
                + Convert.ToInt32(normalized.Substring(2, 2), 16) + ", "
                + Convert.ToInt32(normalized.Substring(4, 2), 16);
        }

        private static string BuildLocationKey(FontUsageLocation location)
        {
            if (location == null)
            {
                return string.Empty;
            }

            return (location.Scope.ToString() ?? string.Empty) + "|"
                + (location.SlideIndex.HasValue ? location.SlideIndex.Value.ToString() : string.Empty) + "|"
                + (location.ShapeId.HasValue ? location.ShapeId.Value.ToString() : string.Empty) + "|"
                + (location.ShapeName ?? string.Empty) + "|"
                + (location.Label ?? string.Empty);
        }

        private async Task RefreshColorsAsync()
        {
            using (OfficeBusyMessageFilter.Register())
            {
                var scanVersion = Interlocked.Increment(ref _scanVersion);
                var nextCts = new CancellationTokenSource();
                var previousCts = Interlocked.Exchange(ref _scanCts, nextCts);
                if (previousCts != null)
                {
                    TryCancel(previousCts);
                }

                await InvokeOnUiAsync(() =>
                {
                    IsScanning = true;
                    ProgressValue = 0;
                    StatusText = "Refreshing color inventory...";
                }).ConfigureAwait(false);

                await _scanGate.WaitAsync().ConfigureAwait(false);

                try
                {
                    var colorProgress = new Progress<ScanProgressInfo>(info =>
                    {
                        if (scanVersion != Volatile.Read(ref _scanVersion))
                        {
                            return;
                        }

                        InvokeOnUi(() =>
                        {
                            ProgressValue = info == null ? 0 : info.Percentage;
                            StatusText = string.IsNullOrWhiteSpace(info?.Message) ? "Refreshing colors..." : "Colors: " + info.Message;
                        });
                    });

                    var scanResult = await ExecutePresentationActionWithRetryAsync(
                        token => InvokeOnUiAsync(() => _presentationService.ScanActivePresentationColorsAsync(colorProgress, token)),
                        nextCts.Token).ConfigureAwait(false);
                    if (scanVersion != Volatile.Read(ref _scanVersion))
                    {
                        return;
                    }

                    await InvokeOnUiAsync(() =>
                    {
                        _colorScanResult = scanResult ?? new ColorScanResult();

                        var selectedKey = BuildNodeKey(SelectedNode);
                        ApplyColorSummary(_colorScanResult.Items);
                        SyncColorNodeCache(_colorScanResult.Items);
                        RebuildColorGroups();
                        RestoreSelection(selectedKey);

                        if (SelectedNode == null)
                        {
                            UpdateInspector(null);
                        }

                        StatusText = BuildStatusSummary();
                        ProgressValue = 100;
                    }).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    if (scanVersion == Volatile.Read(ref _scanVersion))
                    {
                        await InvokeOnUiAsync(() => StatusText = "Color refresh cancelled.").ConfigureAwait(false);
                    }
                }
                catch (Exception ex)
                {
                    await InvokeOnUiAsync(() =>
                    {
                        ErrorReporter.Show("Morphos could not refresh presentation colors.", ex);
                        StatusText = "Color refresh failed.";
                    }).ConfigureAwait(false);
                }
                finally
                {
                    _scanGate.Release();

                    if (ReferenceEquals(_scanCts, nextCts))
                    {
                        _scanCts = null;
                    }

                    nextCts.Dispose();

                    if (scanVersion == Volatile.Read(ref _scanVersion))
                    {
                        await InvokeOnUiAsync(() => IsScanning = false).ConfigureAwait(false);
                    }
                }
            }
        }

        private async Task RefreshFontsAsync(string preferredFontName = null)
        {
            using (OfficeBusyMessageFilter.Register())
            {
                var scanVersion = Interlocked.Increment(ref _scanVersion);
                var nextCts = new CancellationTokenSource();
                var previousCts = Interlocked.Exchange(ref _scanCts, nextCts);
                if (previousCts != null)
                {
                    TryCancel(previousCts);
                }

                await InvokeOnUiAsync(() =>
                {
                    IsScanning = true;
                    ProgressValue = 0;
                    StatusText = "Refreshing font inventory...";
                }).ConfigureAwait(false);

                await _scanGate.WaitAsync().ConfigureAwait(false);

                try
                {
                    var fontProgress = new Progress<ScanProgressInfo>(info =>
                    {
                        if (scanVersion != Volatile.Read(ref _scanVersion))
                        {
                            return;
                        }

                        InvokeOnUi(() =>
                        {
                            ProgressValue = info == null ? 0 : info.Percentage;
                            StatusText = string.IsNullOrWhiteSpace(info?.Message) ? "Refreshing fonts..." : "Fonts: " + info.Message;
                        });
                    });

                    var scanResult = await ExecutePresentationActionWithRetryAsync(
                        token => InvokeOnUiAsync(() => _presentationService.ScanActivePresentationAsync(fontProgress, token)),
                        nextCts.Token).ConfigureAwait(false);
                    if (scanVersion != Volatile.Read(ref _scanVersion))
                    {
                        return;
                    }

                    await InvokeOnUiAsync(() =>
                    {
                        _fontItems = scanResult ?? Array.Empty<FontInventoryItem>();

                        var selectedKey = BuildNodeKey(SelectedNode);
                        var resolvedSelectionKey = FontPaneSelectionResolver.ResolvePostReplaceSelectionKey(
                            preferredFontName,
                            _fontItems.Select(item => item == null ? string.Empty : item.FontName),
                            selectedKey);

                        ApplyFontSummary(_fontItems);
                        SyncFontNodeCache(_fontItems);
                        EnsurePreferredFontVisible(preferredFontName);
                        RebuildFontNodes(resolvedSelectionKey);

                        if (SelectedNode == null)
                        {
                            UpdateInspector(null);
                        }

                        _hasCompletedScan = true;
                        _lastScanSucceeded = true;
                        StatusText = BuildStatusSummary();
                        ProgressValue = 100;
                    }).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    if (scanVersion == Volatile.Read(ref _scanVersion))
                    {
                        await InvokeOnUiAsync(() => StatusText = "Font refresh cancelled.").ConfigureAwait(false);
                    }
                }
                catch (Exception ex)
                {
                    await InvokeOnUiAsync(() =>
                    {
                        ErrorReporter.Show("Morphos could not refresh presentation fonts.", ex);
                        StatusText = "Font refresh failed.";
                    }).ConfigureAwait(false);
                }
                finally
                {
                    _scanGate.Release();

                    if (ReferenceEquals(_scanCts, nextCts))
                    {
                        _scanCts = null;
                    }

                    nextCts.Dispose();

                    if (scanVersion == Volatile.Read(ref _scanVersion))
                    {
                        await InvokeOnUiAsync(() => IsScanning = false).ConfigureAwait(false);
                    }
                }
            }
        }

        private void SyncFontNodeCache(IReadOnlyList<FontInventoryItem> items)
        {
            var safeItems = items ?? Array.Empty<FontInventoryItem>();
            var activeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in safeItems)
            {
                var key = FontNameNormalizer.Normalize(item == null ? string.Empty : item.FontName);
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                FontNodeViewModel node;
                if (_fontNodeCache.TryGetValue(key, out node))
                {
                    node.UpdateItem(item);
                }
                else
                {
                    node = new FontNodeViewModel(item);
                    _fontNodeCache[key] = node;
                }

                activeKeys.Add(key);
            }

            var staleKeys = _fontNodeCache.Keys
                .Where(key => !activeKeys.Contains(key))
                .ToList();
            foreach (var staleKey in staleKeys)
            {
                _fontNodeCache.Remove(staleKey);
            }

            _fontNodes = safeItems
                .Select(item => _fontNodeCache[FontNameNormalizer.Normalize(item.FontName)])
                .OrderByDescending(node => node.Item.UsesCount)
                .ThenBy(node => node.FontName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void RebuildFontNodes(string selectionKey = null)
        {
            _fontSearchDebounceTimer?.Stop();
            selectionKey = selectionKey == null
                ? BuildNodeKey(SelectedNode)
                : selectionKey;
            var expandedKeys = CaptureExpandedNodeKeys(RootNodes);

            var nodes = _fontNodes
                .Where(MatchesFontSearch)
                .Cast<TreeNodeViewModel>()
                .ToList();

            RootNodes.ReplaceRange(nodes);

            RestoreExpandedState(RootNodes, expandedKeys);
            RestoreSelection(selectionKey);
        }

        private void RebuildColorGroups()
        {
            _colorSearchDebounceTimer?.Stop();
            var selectionKey = BuildNodeKey(SelectedNode);
            var expandedKeys = CaptureExpandedNodeKeys(ColorGroups);
            var hasSearchQuery = !string.IsNullOrWhiteSpace(NormalizeSearchText(ColorSearchText));

            var filteredItems = (_colorScanResult == null ? Array.Empty<ColorInventoryItem>() : _colorScanResult.Items)
                .Where(x => x.UsageKind != ColorUsageKind.ChartOverride)
                .Where(MatchesColorSearch)
                .GroupBy(x => x.UsageKind)
                .OrderBy(x => SortColorKind(x.Key));

            var groups = new List<TreeNodeViewModel>();
            foreach (var group in filteredItems)
            {
                var groupItems = group
                    .OrderByDescending(x => x.UsesCount)
                    .ThenBy(x => x.HexValue, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (groupItems.Count == 0)
                {
                    continue;
                }

                var groupNode = new ColorGroupNodeViewModel(
                    groupItems[0].UsageKindLabel,
                    groupItems.Sum(x => x.UsesCount),
                    GetColorNodes(groupItems),
                    ShouldInitiallyExpandColorGroup(group.Key, groupItems.Count, selectionKey, hasSearchQuery));

                groups.Add(groupNode);
            }

            ColorGroups.ReplaceRange(groups);

            RestoreExpandedState(ColorGroups, expandedKeys);
            RestoreSelection(selectionKey);
        }

        private void SyncColorNodeCache(IReadOnlyList<ColorInventoryItem> items)
        {
            var safeItems = (items ?? Array.Empty<ColorInventoryItem>())
                .Where(item => item != null && item.UsageKind != ColorUsageKind.ChartOverride)
                .ToList();
            var activeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in safeItems)
            {
                var key = BuildColorNodeCacheKey(item);
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                ColorNodeViewModel node;
                if (_colorNodeCache.TryGetValue(key, out node))
                {
                    node.UpdateItem(item);
                }
                else
                {
                    node = new ColorNodeViewModel(item);
                    _colorNodeCache[key] = node;
                }

                activeKeys.Add(key);
            }

            var staleKeys = _colorNodeCache.Keys
                .Where(key => !activeKeys.Contains(key))
                .ToList();
            foreach (var staleKey in staleKeys)
            {
                _colorNodeCache.Remove(staleKey);
            }
        }

        private IReadOnlyList<ColorNodeViewModel> GetColorNodes(IReadOnlyList<ColorInventoryItem> items)
        {
            var safeItems = items ?? Array.Empty<ColorInventoryItem>();
            return safeItems
                .Where(item => item != null)
                .Select(item => _colorNodeCache[BuildColorNodeCacheKey(item)])
                .ToList();
        }

        private static string BuildColorNodeCacheKey(ColorInventoryItem item)
        {
            return item == null ? string.Empty : BuildColorInstructionKey(item.UsageKind, item.HexValue);
        }

        private string BuildStatusSummary()
        {
            if (!HasFontResults && !HasColorResults)
            {
                return "No fonts or direct RGB colors found in the active presentation.";
            }

            if (!HasColorResults)
            {
                return "Scanned " + TotalFonts + " fonts. No direct RGB colors found.";
            }

            if (!HasFontResults)
            {
                return "Scanned " + TotalDirectColors + " direct RGB colors. No font inventory was found.";
            }

            return "Scanned " + TotalFonts + " fonts and " + TotalDirectColors + " direct RGB colors.";
        }

        private static string BuildEmbeddingSummary(IReadOnlyList<FontInventoryItem> items)
        {
            if (items == null || items.Count == 0)
            {
                return "No fonts";
            }

            if (items.All(x => x.EmbeddingStatus == FontEmbeddingStatus.Yes))
            {
                return "Embedded";
            }

            if (items.All(x => x.EmbeddingStatus == FontEmbeddingStatus.Subset))
            {
                return "Subset";
            }

            if (items.All(x => x.EmbeddingStatus == FontEmbeddingStatus.No))
            {
                return "Not embedded";
            }

            return "Mixed";
        }

        private bool MatchesFontSearch(FontNodeViewModel node)
        {
            return node != null && node.SearchIndex.Matches(NormalizeSearchText(FontSearchText));
        }

        private bool MatchesColorSearch(ColorInventoryItem item)
        {
            var query = NormalizeSearchText(ColorSearchText);
            if (string.IsNullOrWhiteSpace(query))
            {
                return true;
            }

            if (ContainsToken(item.HexValue, query)
                || ContainsToken(item.RgbValue, query)
                || ContainsToken(item.UsageKindLabel, query)
                || ContainsToken(item.MatchingThemeDisplayName, query))
            {
                return true;
            }

            return item.Locations != null && item.Locations.Any(x => ContainsToken(x.Label, query));
        }

        private static bool ContainsToken(string value, string token)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void QueueFontNodeRebuild()
        {
            if (_dispatcher == null)
            {
                RebuildFontNodes();
                return;
            }

            if (!_dispatcher.CheckAccess())
            {
                _dispatcher.BeginInvoke(new Action(QueueFontNodeRebuild), DispatcherPriority.Background);
                return;
            }

            EnsureSearchDebounceTimers();
            _fontSearchDebounceTimer.Stop();
            _fontSearchDebounceTimer.Start();
        }

        private void QueueColorGroupRebuild()
        {
            if (_dispatcher == null)
            {
                RebuildColorGroups();
                return;
            }

            if (!_dispatcher.CheckAccess())
            {
                _dispatcher.BeginInvoke(new Action(QueueColorGroupRebuild), DispatcherPriority.Background);
                return;
            }

            EnsureSearchDebounceTimers();
            _colorSearchDebounceTimer.Stop();
            _colorSearchDebounceTimer.Start();
        }

        private static string NormalizeSearchText(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private void EnsurePreferredFontVisible(string preferredFontName)
        {
            if (string.IsNullOrWhiteSpace(preferredFontName) || string.IsNullOrWhiteSpace(NormalizeSearchText(FontSearchText)))
            {
                return;
            }

            var preferredKey = FontPaneSelectionResolver.BuildFontSelectionKey(preferredFontName);
            var visibleKeys = new HashSet<string>(
                _fontNodes
                    .Where(MatchesFontSearch)
                    .Select(node => BuildNodeKey(node)),
                StringComparer.OrdinalIgnoreCase);

            if (visibleKeys.Contains(preferredKey))
            {
                return;
            }

            FontSearchText = string.Empty;
        }

        private void RestoreSelection(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                if (SelectedNode != null && !NodeExists(SelectedNode))
                {
                    SelectedNode = null;
                }

                return;
            }

            var restoredNode = TryFindFontNodeByKey(key) ?? FindNodeByKey(key, RootNodes) ?? FindNodeByKey(key, ColorGroups);
            if (restoredNode == null
                && key.StartsWith("font|", StringComparison.OrdinalIgnoreCase)
                && RootNodes.Count > 0)
            {
                restoredNode = RootNodes.OfType<FontNodeViewModel>().FirstOrDefault();
            }

            SelectedNode = restoredNode;
            ExpandPathToKey(key, RootNodes);
            ExpandPathToKey(key, ColorGroups);
        }

        private TreeNodeViewModel TryFindFontNodeByKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key) || !key.StartsWith("font|", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var fontName = key.Substring("font|".Length);
            FontNodeViewModel node;
            if (!_fontNodeCache.TryGetValue(FontNameNormalizer.Normalize(fontName), out node))
            {
                return null;
            }

            return RootNodes.Contains(node) ? node : null;
        }

        private bool NodeExists(TreeNodeViewModel node)
        {
            return FindNodeByKey(BuildNodeKey(node), RootNodes) != null
                || FindNodeByKey(BuildNodeKey(node), ColorGroups) != null;
        }

        private TreeNodeViewModel FindNodeByKey(string key, IEnumerable<TreeNodeViewModel> roots)
        {
            if (string.IsNullOrWhiteSpace(key) || roots == null)
            {
                return null;
            }

            foreach (var root in roots)
            {
                foreach (var node in EnumerateNodes(root))
                {
                    if (string.Equals(BuildNodeKey(node), key, StringComparison.OrdinalIgnoreCase))
                    {
                        return node;
                    }
                }
            }

            return null;
        }

        private static IEnumerable<TreeNodeViewModel> EnumerateNodes(TreeNodeViewModel node)
        {
            if (node == null)
            {
                yield break;
            }

            yield return node;

            foreach (var child in node.Children)
            {
                foreach (var descendant in EnumerateNodes(child))
                {
                    yield return descendant;
                }
            }
        }

        private static string BuildNodeKey(TreeNodeViewModel node)
        {
            if (node is FontNodeViewModel fontNode)
            {
                return "font|" + fontNode.FontName;
            }

            if (node is FontUsageNodeViewModel fontUsageNode)
            {
                return "font-usage|" + (fontUsageNode.Location == null ? string.Empty : fontUsageNode.Location.Label);
            }

            if (node is ColorGroupNodeViewModel colorGroupNode)
            {
                return "color-group|" + colorGroupNode.DisplayName;
            }

            if (node is ColorNodeViewModel colorNode)
            {
                return "color|" + colorNode.Item.UsageKind + "|" + colorNode.Item.HexValue;
            }

            if (node is ColorUsageNodeViewModel colorUsageNode)
            {
                return "color-usage|" + (colorUsageNode.Location == null ? string.Empty : colorUsageNode.Location.Label);
            }

            return string.Empty;
        }

        private static int SortColorKind(ColorUsageKind usageKind)
        {
            switch (usageKind)
            {
                case ColorUsageKind.ShapeFill:
                    return 0;
                case ColorUsageKind.TextFill:
                    return 1;
                case ColorUsageKind.Line:
                    return 2;
                case ColorUsageKind.Effect:
                    return 3;
                case ColorUsageKind.ChartOverride:
                    return 4;
                default:
                    return 10;
            }
        }

        private static double ScaleProgress(double percentage, double offset, double span)
        {
            var clamped = Math.Max(0, Math.Min(100, percentage));
            return offset + ((clamped / 100d) * span);
        }

        private static double CombineProgress(double fontPercentage, double colorPercentage)
        {
            return ScaleProgress(fontPercentage, 0, 60)
                + ScaleProgress(colorPercentage, 0, 40);
        }

        private static HashSet<string> CaptureExpandedNodeKeys(IEnumerable<TreeNodeViewModel> roots)
        {
            var expandedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (roots == null)
            {
                return expandedKeys;
            }

            foreach (var root in roots)
            {
                foreach (var node in EnumerateNodes(root))
                {
                    var key = BuildNodeKey(node);
                    if (node != null && node.IsExpanded && !string.IsNullOrWhiteSpace(key))
                    {
                        expandedKeys.Add(key);
                    }
                }
            }

            return expandedKeys;
        }

        private static void RestoreExpandedState(IEnumerable<TreeNodeViewModel> roots, ISet<string> expandedKeys)
        {
            if (roots == null || expandedKeys == null || expandedKeys.Count == 0)
            {
                return;
            }

            foreach (var root in roots)
            {
                foreach (var node in EnumerateNodes(root))
                {
                    var key = BuildNodeKey(node);
                    if (!string.IsNullOrWhiteSpace(key) && expandedKeys.Contains(key))
                    {
                        node.IsExpanded = true;
                    }
                }
            }
        }

        private static bool ExpandPathToKey(string key, IEnumerable<TreeNodeViewModel> roots)
        {
            if (string.IsNullOrWhiteSpace(key) || roots == null)
            {
                return false;
            }

            foreach (var node in roots)
            {
                if (ExpandNodePath(node, key))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool ExpandNodePath(TreeNodeViewModel node, string key)
        {
            if (node == null)
            {
                return false;
            }

            if (string.Equals(BuildNodeKey(node), key, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            foreach (var child in node.Children)
            {
                if (ExpandNodePath(child, key))
                {
                    node.IsExpanded = true;
                    return true;
                }
            }

            return false;
        }

        private void EnsureSearchDebounceTimers()
        {
            if (_dispatcher == null)
            {
                return;
            }

            if (_fontSearchDebounceTimer == null)
            {
                _fontSearchDebounceTimer = new DispatcherTimer(DispatcherPriority.Background, _dispatcher)
                {
                    Interval = TimeSpan.FromMilliseconds(SearchDebounceMilliseconds)
                };
                _fontSearchDebounceTimer.Tick += (sender, args) =>
                {
                    _fontSearchDebounceTimer.Stop();
                    RebuildFontNodes();
                };
            }

            if (_colorSearchDebounceTimer == null)
            {
                _colorSearchDebounceTimer = new DispatcherTimer(DispatcherPriority.Background, _dispatcher)
                {
                    Interval = TimeSpan.FromMilliseconds(SearchDebounceMilliseconds)
                };
                _colorSearchDebounceTimer.Tick += (sender, args) =>
                {
                    _colorSearchDebounceTimer.Stop();
                    RebuildColorGroups();
                };
            }
        }

        private static bool ShouldInitiallyExpandColorGroup(
            ColorUsageKind usageKind,
            int groupItemCount,
            string selectionKey,
            bool hasSearchQuery)
        {
            if (hasSearchQuery && groupItemCount <= 32)
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(selectionKey))
            {
                return false;
            }

            return selectionKey.StartsWith("color|" + usageKind + "|", StringComparison.OrdinalIgnoreCase)
                || string.Equals(
                    selectionKey,
                    "color-group|" + GetColorUsageLabel(usageKind),
                    StringComparison.OrdinalIgnoreCase);
        }

        private static string GetColorUsageLabel(ColorUsageKind usageKind)
        {
            switch (usageKind)
            {
                case ColorUsageKind.TextFill:
                    return "Text fill";
                case ColorUsageKind.Line:
                    return "Line";
                case ColorUsageKind.Effect:
                    return "Effect";
                case ColorUsageKind.ChartOverride:
                    return "Chart override";
                default:
                    return "Shape fill";
            }
        }

        private static void TryCancel(CancellationTokenSource cancellationTokenSource)
        {
            if (cancellationTokenSource == null)
            {
                return;
            }

            try
            {
                cancellationTokenSource.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }
        }

        private static bool IsTransientPresentationException(Exception exception)
        {
            while (exception != null)
            {
                if (exception is InvalidCastException)
                {
                    return true;
                }

                if (exception is COMException)
                {
                    var comException = (COMException)exception;
                    var message = exception.Message ?? string.Empty;
                    if (comException.ErrorCode == RpcEServerCallRetryLater
                        || comException.ErrorCode == RpcECallRejected
                        || message.IndexOf("specified cast is not valid", StringComparison.OrdinalIgnoreCase) >= 0
                        || message.IndexOf("application is busy", StringComparison.OrdinalIgnoreCase) >= 0
                        || message.IndexOf("retry later", StringComparison.OrdinalIgnoreCase) >= 0
                        || message.IndexOf("retrylater", StringComparison.OrdinalIgnoreCase) >= 0
                        || message.IndexOf("object does not exist", StringComparison.OrdinalIgnoreCase) >= 0
                        || message.IndexOf("unknown member", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }

                if (exception is ObjectDisposedException objectDisposedException)
                {
                    var disposedMessage = objectDisposedException.Message ?? string.Empty;
                    if (disposedMessage.IndexOf("CancellationTokenSource", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }

                exception = exception.InnerException;
            }

            return false;
        }

        private static async Task<T> ExecutePresentationActionWithRetryAsync<T>(
            Func<CancellationToken, Task<T>> action,
            CancellationToken cancellationToken)
        {
            if (action == null)
            {
                return default(T);
            }

            for (var attempt = 0; ; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    return await action(cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex) when (attempt < 6 && IsTransientPresentationException(ex))
                {
                    await Task.Delay(150, cancellationToken).ConfigureAwait(false);
                }
            }
        }

        private static async Task<T> ExecutePresentationActionWithRetryAsync<T>(Func<T> action)
        {
            if (action == null)
            {
                return default(T);
            }

            for (var attempt = 0; ; attempt++)
            {
                try
                {
                    return action();
                }
                catch (Exception ex) when (attempt < 12 && IsTransientPresentationException(ex))
                {
                    await Task.Delay(150).ConfigureAwait(true);
                }
            }
        }

        private void InvokeOnUi(Action action)
        {
            if (action == null)
            {
                return;
            }

            if (_dispatcher == null || _dispatcher.CheckAccess())
            {
                action();
                return;
            }

            _dispatcher.BeginInvoke(action);
        }

        private Task InvokeOnUiAsync(Action action)
        {
            if (action == null)
            {
                return Task.CompletedTask;
            }

            if (_dispatcher == null || _dispatcher.CheckAccess())
            {
                action();
                return Task.CompletedTask;
            }

            return _dispatcher.InvokeAsync(action).Task;
        }

        private Task<T> InvokeOnUiAsync<T>(Func<Task<T>> action)
        {
            if (action == null)
            {
                return Task.FromResult(default(T));
            }

            if (_dispatcher == null || _dispatcher.CheckAccess())
            {
                return action();
            }

            return _dispatcher.InvokeAsync(action).Task.Unwrap();
        }

        private void UpdateInspector(TreeNodeViewModel node)
        {
            if (node is FontNodeViewModel fontNode)
            {
                if (fontNode.Item.IsSubstituted)
                {
                    InspectorTitle = fontNode.FontName;
                    InspectorBody = "Stored font differs from what PowerPoint is currently rendering. Replace it if the deck should be portable.";
                    return;
                }

                if (fontNode.HasSaveWarning)
                {
                    InspectorTitle = fontNode.FontName;
                    InspectorBody = fontNode.Item.IsLocallyMissing
                        ? "This font is missing on this computer and PowerPoint reported save-validation risk."
                        : "PowerPoint reported an embed/save validation risk for this font.";
                    return;
                }

                InspectorTitle = fontNode.FontName;
                InspectorBody = "Uses: " + fontNode.UsesText
                    + Environment.NewLine
                    + "Embed: " + fontNode.EmbeddingText;
                return;
            }

            if (node is FontUsageNodeViewModel fontUsageNode)
            {
                InspectorTitle = "Font instance";
                InspectorBody = fontUsageNode.Location == null ? string.Empty : fontUsageNode.Location.Label;
                return;
            }

            if (node is ColorNodeViewModel colorNode)
            {
                InspectorTitle = colorNode.Item.UsageKindLabel;
                InspectorBody = "#" + colorNode.Item.HexValue
                    + Environment.NewLine
                    + "RGB " + colorNode.Item.RgbValue
                    + Environment.NewLine
                    + "Used in " + colorNode.Item.UsesCount + " shapes"
                    + (colorNode.Item.MatchesThemeColor
                        ? Environment.NewLine + "Exact theme match: " + colorNode.Item.MatchingThemeDisplayName
                        : string.Empty);
                return;
            }

            if (node is ColorUsageNodeViewModel colorUsageNode)
            {
                InspectorTitle = "Color instance";
                InspectorBody = colorUsageNode.Location == null ? string.Empty : colorUsageNode.Location.Label;
                return;
            }

            if (node is ColorGroupNodeViewModel colorGroupNode)
            {
                InspectorTitle = colorGroupNode.DisplayName;
                InspectorBody = colorGroupNode.ColorCountText
                    + Environment.NewLine
                    + colorGroupNode.UsesText + " uses";
                return;
            }

            InspectorTitle = "Inspector";
            InspectorBody = "Select a font or color row to inspect its details.";
        }

        private void UpdateSelectionState(TreeNodeViewModel node)
        {
            HasSelectedFont = node is FontNodeViewModel;
            HasSelectedColor = node is ColorNodeViewModel;
            HasSelectedNode = node is FontNodeViewModel
                || node is FontUsageNodeViewModel
                || node is ColorNodeViewModel
                || node is ColorUsageNodeViewModel;
        }
    }
}
