using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using MorphosPowerPointAddIn.Models;

namespace MorphosPowerPointAddIn.Dialogs
{
    public partial class ReplaceFontsDialog : Window, INotifyPropertyChanged
    {
        private const double CompactLayoutWidthThreshold = 520;
        private FontReplacementTarget _selectedFontChoice;

        public ReplaceFontsDialog(
            System.Collections.Generic.IReadOnlyList<string> sourceFontNames,
            IEnumerable<string> initiallySelectedFontNames,
            System.Collections.Generic.IReadOnlyList<FontReplacementTarget> fontChoices)
        {
            InitializeComponent();

            SelectedSourceFontNames = new ObservableCollection<string>(
                (initiallySelectedFontNames ?? System.Array.Empty<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(System.StringComparer.OrdinalIgnoreCase)
                    .OrderBy(x => x, System.StringComparer.OrdinalIgnoreCase));

            if (SelectedSourceFontNames.Count == 0)
            {
                var fallbackSource = (sourceFontNames ?? System.Array.Empty<string>())
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
                if (!string.IsNullOrWhiteSpace(fallbackSource))
                {
                    SelectedSourceFontNames.Add(fallbackSource);
                }
            }

            FontChoices = new ObservableCollection<FontReplacementTarget>(
                (fontChoices ?? System.Array.Empty<FontReplacementTarget>())
                    .Where(choice => choice != null));

            SelectedFontChoice = FontChoices.FirstOrDefault(
                    choice => SelectedSourceFontNames.All(
                        source => !string.Equals(source, choice.NormalizedName, System.StringComparison.OrdinalIgnoreCase)))
                ?? FontChoices.FirstOrDefault();

            DataContext = this;
            Loaded += (sender, args) => ApplyResponsiveLayout(ActualWidth);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<string> SelectedSourceFontNames { get; }

        public ObservableCollection<FontReplacementTarget> FontChoices { get; }

        public string SelectedSourceSummary =>
            SelectedSourceFontNames.Count <= 1
                ? (SelectedSourceFontNames.FirstOrDefault() ?? "Selected font")
                : SelectedSourceFontNames.Count + " fonts selected";

        public FontReplacementTarget SelectedFontChoice
        {
            get => _selectedFontChoice;
            set
            {
                if (ReferenceEquals(_selectedFontChoice, value))
                {
                    return;
                }

                _selectedFontChoice = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedFontChoice)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedFontName)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedFontDisplayName)));
            }
        }

        public string SelectedFontName => SelectedFontChoice == null ? string.Empty : SelectedFontChoice.NormalizedName ?? string.Empty;

        public string SelectedFontDisplayName => SelectedFontChoice == null ? string.Empty : SelectedFontChoice.DisplayName ?? string.Empty;

        private void Replace_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedSourceFontNames.Count == 0)
            {
                MessageBox.Show("Select a source font first.", "Morphos", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (SelectedFontChoice == null || string.IsNullOrWhiteSpace(SelectedFontName))
            {
                MessageBox.Show("Select a replacement font first.", "Morphos", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var normalizedFontName = SelectedFontName.Trim();
            if (!FontChoices.Any(choice => string.Equals(choice.NormalizedName, normalizedFontName, System.StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(
                    "Choose a replacement from the installed Windows or theme fonts in this list.",
                    "Morphos",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (SelectedSourceFontNames.All(x => string.Equals(x, normalizedFontName, System.StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(
                    "Choose a different replacement font.",
                    "Morphos",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
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

        private void ApplyResponsiveLayout(double width)
        {
            var isCompact = width > 0 && width < CompactLayoutWidthThreshold;
            if (SelectedFontPanel == null
                || ReplacementFontPanel == null
                || SelectedFontColumn == null
                || FontPanelSpacerColumn == null
                || ReplacementFontColumn == null)
            {
                return;
            }

            if (isCompact)
            {
                SelectedFontColumn.Width = new GridLength(1, GridUnitType.Star);
                FontPanelSpacerColumn.Width = new GridLength(0);
                ReplacementFontColumn.Width = new GridLength(0);

                Grid.SetRow(SelectedFontPanel, 0);
                Grid.SetColumn(SelectedFontPanel, 0);
                Grid.SetColumnSpan(SelectedFontPanel, 3);

                Grid.SetRow(ReplacementFontPanel, 1);
                Grid.SetColumn(ReplacementFontPanel, 0);
                Grid.SetColumnSpan(ReplacementFontPanel, 3);
                ReplacementFontPanel.Margin = new Thickness(0, 12, 0, 0);
                return;
            }

            SelectedFontColumn.Width = new GridLength(1, GridUnitType.Star);
            FontPanelSpacerColumn.Width = new GridLength(12);
            ReplacementFontColumn.Width = new GridLength(1, GridUnitType.Star);

            Grid.SetRow(SelectedFontPanel, 0);
            Grid.SetColumn(SelectedFontPanel, 0);
            Grid.SetColumnSpan(SelectedFontPanel, 1);

            Grid.SetRow(ReplacementFontPanel, 0);
            Grid.SetColumn(ReplacementFontPanel, 2);
            Grid.SetColumnSpan(ReplacementFontPanel, 1);
            ReplacementFontPanel.Margin = new Thickness(0);
        }
    }
}
