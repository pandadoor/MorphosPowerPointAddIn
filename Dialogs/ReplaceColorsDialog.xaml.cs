using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using MorphosPowerPointAddIn.Models;
using MorphosPowerPointAddIn.Utilities;

namespace MorphosPowerPointAddIn.Dialogs
{
    public partial class ReplaceColorsDialog : Window, INotifyPropertyChanged
    {
        private const double CompactPreviewWidthThreshold = 500;
        private readonly IReadOnlyList<ColorInventoryItem> _sourceColors;
        private readonly IReadOnlyList<ColorInventoryItem> _matchingSourceItems;
        private readonly string _sourceUsageSummary;
        private IReadOnlyList<ColorReplacementInstruction> _selectedInstructions = Array.Empty<ColorReplacementInstruction>();
        private string _customHexValue = string.Empty;
        private string _previewText;
        private ThemeChoiceOption _selectedThemeChoice = ThemeChoiceOption.None;

        public ReplaceColorsDialog(
            IReadOnlyList<ColorInventoryItem> sourceColors,
            IReadOnlyList<ThemeColorInfo> themeColors,
            ColorInventoryItem preferredItem)
        {
            InitializeComponent();

            _sourceColors = (sourceColors ?? Array.Empty<ColorInventoryItem>())
                .Where(x => x != null && x.UsageKind != ColorUsageKind.ChartOverride)
                .ToList();

            ThemeChoices = new ObservableCollection<ThemeChoiceOption>(
                new[] { ThemeChoiceOption.None }.Concat(
                    (themeColors ?? Array.Empty<ThemeColorInfo>())
                        .Where(x => !string.IsNullOrWhiteSpace(x.HexValue))
                        .GroupBy(x => x.SchemeName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                        .Select(x => x.First())
                        .Select(ThemeChoiceOption.FromThemeColor)));

            SourceItem = preferredItem
                ?? _sourceColors
                    .OrderByDescending(x => x.UsesCount)
                    .ThenBy(x => x.HexValue, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault();

            _matchingSourceItems = SourceItem == null
                ? (IReadOnlyList<ColorInventoryItem>)Array.Empty<ColorInventoryItem>()
                : _sourceColors
                    .Where(x => string.Equals(NormalizeHex(x.HexValue), NormalizeHex(SourceItem.HexValue), StringComparison.OrdinalIgnoreCase))
                    .GroupBy(x => x.UsageKind)
                    .Select(x => x.First())
                    .ToList();
            _sourceUsageSummary = SourceItem == null
                ? "No direct RGB color selected."
                : BuildSourceUsageSummary(_matchingSourceItems);
            _selectedInstructions = BuildInstructions();
            PreviewText = BuildPreviewText();
            DataContext = this;
            Loaded += (sender, args) => ApplyResponsiveLayout(ActualWidth);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public ColorInventoryItem SourceItem { get; }

        public ObservableCollection<ThemeChoiceOption> ThemeChoices { get; }

        public IReadOnlyList<ColorReplacementInstruction> SelectedInstructions
        {
            get => _selectedInstructions;
        }

        public string SourceBrush => SourceItem == null ? "#000000" : SourceItem.BrushValue;

        public string SourceHexText => "#" + (SourceItem == null ? "000000" : SourceItem.HexValue ?? "000000");

        public string SourceRgbText => "RGB " + (SourceItem == null ? "0, 0, 0" : SourceItem.RgbValue ?? "0, 0, 0");

        public string SourceUsageSummary => _sourceUsageSummary;

        public string ThemeMatchText => SourceItem != null && SourceItem.MatchesThemeColor
            ? "Exact theme match: " + SourceItem.MatchingThemeDisplayName
            : string.Empty;

        public Visibility ThemeMatchVisibility => SourceItem != null && SourceItem.MatchesThemeColor
            ? Visibility.Visible
            : Visibility.Collapsed;

        public ThemeChoiceOption SelectedThemeChoice
        {
            get => _selectedThemeChoice;
            set
            {
                var nextValue = value ?? ThemeChoiceOption.None;
                if (_selectedThemeChoice != null
                    && string.Equals(_selectedThemeChoice.Key, nextValue.Key, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                _selectedThemeChoice = nextValue;
                if (!_selectedThemeChoice.IsNone)
                {
                    _customHexValue = string.Empty;
                }

                RaiseReplacementChanged();
            }
        }

        public bool HasInstructions => _selectedInstructions.Count > 0;

        public string ReplacementBrush
        {
            get
            {
                var hexValue = GetReplacementHexValue();
                return string.IsNullOrWhiteSpace(hexValue) ? "#FFFDF8" : "#" + hexValue;
            }
        }

        public string PreviewText
        {
            get => _previewText;
            private set
            {
                if (_previewText == value)
                {
                    return;
                }

                _previewText = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PreviewText)));
            }
        }

        private void PickColor_Click(object sender, RoutedEventArgs e)
        {
            var initialHexValue = GetReplacementHexValue();
            if (string.IsNullOrWhiteSpace(initialHexValue))
            {
                initialHexValue = SourceItem == null ? string.Empty : SourceItem.HexValue;
            }

            using (var dialog = new System.Windows.Forms.ColorDialog
            {
                AllowFullOpen = true,
                FullOpen = true,
                Color = ParseColor(initialHexValue)
            })
            {
                var ownerWindow = DialogWindowHelper.TryGetOwnerWindow(this);
                var dialogResult = ownerWindow == null
                    ? dialog.ShowDialog()
                    : dialog.ShowDialog(ownerWindow);
                if (dialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    _customHexValue = ToHexValue(dialog.Color);
                    _selectedThemeChoice = ThemeChoiceOption.None;
                    RaiseReplacementChanged();
                }
            }
        }

        private void ResetRow_Click(object sender, RoutedEventArgs e)
        {
            _customHexValue = string.Empty;
            _selectedThemeChoice = ThemeChoiceOption.None;
            RaiseReplacementChanged();
        }

        private void Replace_Click(object sender, RoutedEventArgs e)
        {
            if (!HasInstructions)
            {
                MessageBox.Show("Choose a replacement color first.", "Morphos", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            ApplyResponsiveLayout(e.NewSize.Width);
        }

        private IReadOnlyList<ColorReplacementInstruction> BuildInstructions()
        {
            if (SourceItem == null || _matchingSourceItems.Count == 0)
            {
                return Array.Empty<ColorReplacementInstruction>();
            }

            if (_selectedThemeChoice != null && !_selectedThemeChoice.IsNone)
            {
                return _matchingSourceItems
                    .Select(x => new ColorReplacementInstruction
                    {
                        UsageKind = x.UsageKind,
                        SourceHexValue = x.HexValue,
                        UseThemeColor = true,
                        ThemeDisplayName = _selectedThemeChoice.DisplayName,
                        ThemeSchemeName = _selectedThemeChoice.SchemeName,
                        TargetLocations = x.Locations
                    })
                    .ToList();
            }

            if (!string.IsNullOrWhiteSpace(_customHexValue))
            {
                return _matchingSourceItems
                    .Select(x => new ColorReplacementInstruction
                    {
                        UsageKind = x.UsageKind,
                        SourceHexValue = x.HexValue,
                        ReplacementHexValue = _customHexValue,
                        TargetLocations = x.Locations
                    })
                    .ToList();
            }

            return Array.Empty<ColorReplacementInstruction>();
        }

        private string GetReplacementHexValue()
        {
            if (_selectedThemeChoice != null && !_selectedThemeChoice.IsNone)
            {
                return _selectedThemeChoice.HexValue;
            }

            return string.IsNullOrWhiteSpace(_customHexValue) ? string.Empty : _customHexValue;
        }

        private string BuildPreviewText()
        {
            if (SourceItem == null)
            {
                return "No color is available for replacement.";
            }

            if (_selectedThemeChoice != null && !_selectedThemeChoice.IsNone)
            {
                return SourceHexText + " will switch to theme color " + _selectedThemeChoice.DisplayName + ".";
            }

            if (!string.IsNullOrWhiteSpace(_customHexValue))
            {
                return SourceHexText + " will switch to custom color #" + _customHexValue + ".";
            }

            return "Choose one theme color or one custom color for " + SourceHexText + ".";
        }

        private string BuildSourceUsageSummary(IReadOnlyList<ColorInventoryItem> matchingItems)
        {
            if (matchingItems == null || matchingItems.Count == 0)
            {
                return SourceItem.UsageKindLabel + "   Used in " + SourceItem.UsesCount + (SourceItem.UsesCount == 1 ? " shape" : " shapes");
            }

            var totalUses = matchingItems.Sum(x => x.UsesCount);
            if (matchingItems.Count == 1)
            {
                return matchingItems[0].UsageKindLabel + "   Used in " + totalUses + (totalUses == 1 ? " shape" : " shapes");
            }

            return "Direct RGB uses across " + matchingItems.Count + " groups   Used in " + totalUses + (totalUses == 1 ? " shape" : " shapes");
        }

        private void RaiseReplacementChanged()
        {
            _selectedInstructions = BuildInstructions();
            PreviewText = BuildPreviewText();
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedThemeChoice)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedInstructions)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasInstructions)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ReplacementBrush)));
        }

        private void ApplyResponsiveLayout(double width)
        {
            if (PreviewSummaryPanel == null)
            {
                return;
            }

            var isCompact = width > 0 && width < CompactPreviewWidthThreshold;
            if (isCompact)
            {
                Grid.SetRow(PreviewSummaryPanel, 1);
                Grid.SetColumn(PreviewSummaryPanel, 0);
                Grid.SetColumnSpan(PreviewSummaryPanel, 4);
                PreviewSummaryPanel.Margin = new Thickness(0, 12, 0, 0);
                return;
            }

            Grid.SetRow(PreviewSummaryPanel, 0);
            Grid.SetColumn(PreviewSummaryPanel, 3);
            Grid.SetColumnSpan(PreviewSummaryPanel, 1);
            PreviewSummaryPanel.Margin = new Thickness(12, 0, 0, 0);
        }

        private static Color ParseColor(string hexValue)
        {
            var normalized = NormalizeHex(hexValue);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return Color.Black;
            }

            return ColorTranslator.FromHtml("#" + normalized);
        }

        private static string NormalizeHex(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Trim().TrimStart('#').ToUpperInvariant();
            return normalized.Length == 6 ? normalized : string.Empty;
        }

        private static string ToHexValue(Color color)
        {
            return color.R.ToString("X2")
                + color.G.ToString("X2")
                + color.B.ToString("X2");
        }

        public sealed class ThemeChoiceOption
        {
            private static readonly ThemeChoiceOption NoneOption = new ThemeChoiceOption
            {
                Key = "none",
                DisplayName = "No theme color",
                HexValue = string.Empty,
                IsNone = true
            };

            private ThemeChoiceOption()
            {
            }

            public string Key { get; private set; }

            public string DisplayName { get; private set; }

            public string SchemeName { get; private set; }

            public string HexValue { get; private set; }

            public bool IsNone { get; private set; }

            public static ThemeChoiceOption None => NoneOption;

            public static ThemeChoiceOption FromThemeColor(ThemeColorInfo themeColor)
            {
                return new ThemeChoiceOption
                {
                    Key = themeColor == null ? string.Empty : themeColor.SchemeName ?? string.Empty,
                    DisplayName = themeColor == null ? "Theme color" : themeColor.DisplayName,
                    SchemeName = themeColor == null ? string.Empty : themeColor.SchemeName,
                    HexValue = themeColor == null ? string.Empty : NormalizeHex(themeColor.HexValue),
                    IsNone = false
                };
            }
        }
    }
}
