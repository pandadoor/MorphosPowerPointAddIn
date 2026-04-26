using System;
using System.Windows;
using System.Windows.Controls;
using MorphosPowerPointAddIn.Models;
using MorphosPowerPointAddIn.Utilities;
using MorphosPowerPointAddIn.ViewModels;

namespace MorphosPowerPointAddIn.UI
{
    public partial class FontsUserControl : UserControl
    {
        private const double CompactPaneWidthThreshold = 500;
        private const double CompactTableWidthThreshold = 500;
        private const double WorkspaceSplitWidthThreshold = 760;
        private const double SingleColumnMetricsWidthThreshold = 390;
        private readonly FontsPaneViewModel _viewModel;

        public static readonly DependencyProperty IsNarrowTableLayoutProperty =
            DependencyProperty.Register(
                nameof(IsNarrowTableLayout),
                typeof(bool),
                typeof(FontsUserControl),
                new PropertyMetadata(false));

        public FontsUserControl(FontsPaneViewModel viewModel)
        {
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            InitializeComponent();
            _viewModel.AttachDispatcher(Dispatcher);
            DataContext = _viewModel;
            Loaded += (sender, args) => ApplyResponsiveLayout(ActualWidth);
        }

        public bool IsNarrowTableLayout
        {
            get => (bool)GetValue(IsNarrowTableLayoutProperty);
            set => SetValue(IsNarrowTableLayoutProperty, value);
        }

        private async void Refresh_Click(object sender, RoutedEventArgs e)
        {
            await SafeExecuteAsync(() => _viewModel.ScanAsync()).ConfigureAwait(true);
        }

        private async void ReplaceColors_Click(object sender, RoutedEventArgs e)
        {
            await SafeExecuteAsync(() => _viewModel.ReplaceColorsAsync()).ConfigureAwait(true);
        }

        private async void ReplaceColorNode_Click(object sender, RoutedEventArgs e)
        {
            var node = GetNodeFromSender(sender) as ColorNodeViewModel;
            if (node != null)
            {
                await SafeExecuteAsync(() => _viewModel.ReplaceColorsAsync(node)).ConfigureAwait(true);
            }
        }

        private async void ReplaceFonts_Click(object sender, RoutedEventArgs e)
        {
            var node = GetNodeFromSender(sender) as FontNodeViewModel;
            if (node != null)
            {
                await SafeExecuteAsync(() => _viewModel.ReplaceFontAsync(node)).ConfigureAwait(true);
            }
        }

        private async void ReplaceSelected_Click(object sender, RoutedEventArgs e)
        {
            var node = _viewModel.SelectedNode as FontNodeViewModel;
            if (node != null)
            {
                await SafeExecuteAsync(() => _viewModel.ReplaceFontAsync(node)).ConfigureAwait(true);
            }
        }

        private async void ReplaceSelectedColor_Click(object sender, RoutedEventArgs e)
        {
            var node = _viewModel.SelectedNode as ColorNodeViewModel;
            if (node != null)
            {
                await SafeExecuteAsync(() => _viewModel.ReplaceColorsAsync(node)).ConfigureAwait(true);
            }
        }

        private void ShowInPowerPoint_Click(object sender, RoutedEventArgs e)
        {
            var node = GetNodeFromSender(sender);
            if (node != null)
            {
                _viewModel.ShowInPowerPoint(node);
            }
        }

        private void ShowSelected_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel.SelectedNode != null)
            {
                _viewModel.ShowInPowerPoint(_viewModel.SelectedNode);
            }
        }

        private void GoToFontsTab_Click(object sender, RoutedEventArgs e)
        {
            SelectWorkspaceTab(FontsTab);
        }

        private void GoToColorsTab_Click(object sender, RoutedEventArgs e)
        {
            SelectWorkspaceTab(ColorsTab);
        }

        private async void EmbedSubset_Click(object sender, RoutedEventArgs e)
        {
            var node = GetNodeFromSender(sender) as FontNodeViewModel;
            if (node != null)
            {
                await SafeExecuteAsync(() => _viewModel.UpdateEmbeddingAsync(node, FontEmbeddingStatus.Subset)).ConfigureAwait(true);
            }
        }

        private async void EmbedFull_Click(object sender, RoutedEventArgs e)
        {
            var node = GetNodeFromSender(sender) as FontNodeViewModel;
            if (node != null)
            {
                await SafeExecuteAsync(() => _viewModel.UpdateEmbeddingAsync(node, FontEmbeddingStatus.Yes)).ConfigureAwait(true);
            }
        }

        private async void Unembed_Click(object sender, RoutedEventArgs e)
        {
            var node = GetNodeFromSender(sender) as FontNodeViewModel;
            if (node != null)
            {
                await SafeExecuteAsync(() => _viewModel.UpdateEmbeddingAsync(node, FontEmbeddingStatus.No)).ConfigureAwait(true);
            }
        }

        private static TreeNodeViewModel GetNodeFromSender(object sender)
        {
            var element = sender as FrameworkElement;
            if (element?.DataContext is TreeNodeViewModel directNode)
            {
                return directNode;
            }

            var contextMenu = element?.Parent as ContextMenu;
            return contextMenu?.PlacementTarget is FrameworkElement placementTarget
                ? placementTarget.DataContext as TreeNodeViewModel
                : null;
        }

        private void SelectWorkspaceTab(TabItem tab)
        {
            if (tab != null)
            {
                tab.IsSelected = true;
            }
        }

        private void FontTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            _viewModel.SelectedNode = e.NewValue as TreeNodeViewModel;
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(sender, e.OriginalSource))
            {
                return;
            }

            var selectedTab = (sender as TabControl)?.SelectedItem as TabItem;
            if (selectedTab == null)
            {
                return;
            }

            if (selectedTab.Name == "FontsTab"
                && (_viewModel.SelectedNode is ColorNodeViewModel
                    || _viewModel.SelectedNode is ColorUsageNodeViewModel
                    || _viewModel.SelectedNode is ColorGroupNodeViewModel))
            {
                _viewModel.SelectedNode = null;
                return;
            }

            if (selectedTab.Name == "ColorsTab"
                && (_viewModel.SelectedNode is FontNodeViewModel
                    || _viewModel.SelectedNode is FontUsageNodeViewModel))
            {
                _viewModel.SelectedNode = null;
                return;
            }

            if (selectedTab.Name == "HomeTab")
            {
                _viewModel.SelectedNode = null;
            }
        }

        private void FontsUserControl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            ApplyResponsiveLayout(e.NewSize.Width);
        }

        private void ApplyResponsiveLayout(double width)
        {
            var isCompact = width > 0 && width < CompactPaneWidthThreshold;
            IsNarrowTableLayout = width > 0 && width < CompactTableWidthThreshold;
            var useSplitWorkspace = width >= WorkspaceSplitWidthThreshold;
            var useSingleMetricColumn = width > 0 && width < SingleColumnMetricsWidthThreshold;

            if (HomeFontsMetricsGrid != null)
            {
                HomeFontsMetricsGrid.Columns = useSingleMetricColumn ? 1 : 2;
            }

            if (HomeColorsMetricsGrid != null)
            {
                HomeColorsMetricsGrid.Columns = useSingleMetricColumn ? 1 : 2;
            }

            if (FontTableHeaderRow != null)
            {
                FontTableHeaderRow.Visibility = IsNarrowTableLayout ? Visibility.Collapsed : Visibility.Visible;
            }

            if (ColorTableHeaderRow != null)
            {
                ColorTableHeaderRow.Visibility = IsNarrowTableLayout ? Visibility.Collapsed : Visibility.Visible;
            }

            ApplyWorkspaceLayout(
                FontsWorkspaceGapColumn,
                FontsDetailColumn,
                FontsDetailPanel,
                useSplitWorkspace);

            ApplyWorkspaceLayout(
                ColorsWorkspaceGapColumn,
                ColorsDetailColumn,
                ColorsDetailPanel,
                useSplitWorkspace);

            if (HeaderActionsPanel == null)
            {
                return;
            }

            if (isCompact)
            {
                Grid.SetRow(HeaderActionsPanel, 1);
                Grid.SetColumn(HeaderActionsPanel, 0);
                Grid.SetColumnSpan(HeaderActionsPanel, 2);
                HeaderActionsPanel.HorizontalAlignment = HorizontalAlignment.Left;
                HeaderActionsPanel.Margin = new Thickness(0, 10, 0, 0);

                if (RefreshButton != null)
                {
                    RefreshButton.MinWidth = 0;
                }

                return;
            }

            Grid.SetRow(HeaderActionsPanel, 0);
            Grid.SetColumn(HeaderActionsPanel, 1);
            Grid.SetColumnSpan(HeaderActionsPanel, 1);
            HeaderActionsPanel.HorizontalAlignment = HorizontalAlignment.Right;
            HeaderActionsPanel.Margin = new Thickness(0);

            if (RefreshButton != null)
            {
                RefreshButton.MinWidth = 80;
            }
        }

        private static void ApplyWorkspaceLayout(
            ColumnDefinition gapColumn,
            ColumnDefinition detailColumn,
            FrameworkElement detailPanel,
            bool useSplitWorkspace)
        {
            if (gapColumn == null || detailColumn == null || detailPanel == null)
            {
                return;
            }

            if (useSplitWorkspace)
            {
                gapColumn.Width = new GridLength(14);
                detailColumn.Width = new GridLength(0.92, GridUnitType.Star);
                Grid.SetRow(detailPanel, 0);
                Grid.SetColumn(detailPanel, 2);
                Grid.SetColumnSpan(detailPanel, 1);
                detailPanel.Margin = new Thickness(0);
                return;
            }

            gapColumn.Width = new GridLength(0);
            detailColumn.Width = new GridLength(0);
            Grid.SetRow(detailPanel, 1);
            Grid.SetColumn(detailPanel, 0);
            Grid.SetColumnSpan(detailPanel, 3);
            detailPanel.Margin = new Thickness(0, 12, 0, 0);
        }

        private static async System.Threading.Tasks.Task SafeExecuteAsync(Func<System.Threading.Tasks.Task> action)
        {
            try
            {
                await action().ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                ErrorReporter.Show("Morphos action failed.", ex);
            }
        }
    }
}
