param(
    [string]$AssemblyPath = (Join-Path $PSScriptRoot '..\bin\x64\Debug\MorphosPowerPointAddIn.dll')
)

$ErrorActionPreference = 'Stop'

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    & powershell -Sta -ExecutionPolicy Bypass -File $PSCommandPath -AssemblyPath $AssemblyPath
    exit $LASTEXITCODE
}

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml
Add-Type -AssemblyName WindowsFormsIntegration
Add-Type -AssemblyName System.Windows.Forms

$resolvedAssemblyPath = (Resolve-Path $AssemblyPath).Path
$assembly = [System.Reflection.Assembly]::LoadFrom($resolvedAssemblyPath)
$failures = [System.Collections.Generic.List[string]]::new()

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        $failures.Add($Message) | Out-Null
    }
}

function Wait-ForLayoutIdle {
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [Action] { },
        [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
}

$fontsPaneViewModelType = $assembly.GetType('MorphosPowerPointAddIn.ViewModels.FontsPaneViewModel', $true)
$fontsUserControlType = $assembly.GetType('MorphosPowerPointAddIn.UI.FontsUserControl', $true)
$fontsTaskPaneHostType = $assembly.GetType('MorphosPowerPointAddIn.UI.FontsTaskPaneHost', $true)
$replaceFontsDialogType = $assembly.GetType('MorphosPowerPointAddIn.Dialogs.ReplaceFontsDialog', $true)
$replaceColorsDialogType = $assembly.GetType('MorphosPowerPointAddIn.Dialogs.ReplaceColorsDialog', $true)
$fontReplacementTargetType = $assembly.GetType('MorphosPowerPointAddIn.Models.FontReplacementTarget', $true)
$themeColorInfoType = $assembly.GetType('MorphosPowerPointAddIn.Models.ThemeColorInfo', $true)
$colorInventoryItemType = $assembly.GetType('MorphosPowerPointAddIn.Models.ColorInventoryItem', $true)
$colorUsageKindType = $assembly.GetType('MorphosPowerPointAddIn.Models.ColorUsageKind', $true)
$fontUsageLocationType = $assembly.GetType('MorphosPowerPointAddIn.Models.FontUsageLocation', $true)

$viewModel = [Activator]::CreateInstance($fontsPaneViewModelType, @($null))
$taskPaneHost = [Activator]::CreateInstance($fontsTaskPaneHostType, @($viewModel))
Assert-True ($taskPaneHost.MinimumSize.Width -le 340) 'Task pane host minimum width should allow compact layouts.'

$control = [Activator]::CreateInstance($fontsUserControlType, @($viewModel))
$controlWindow = New-Object System.Windows.Window
$controlWindow.Width = 330
$controlWindow.Height = 860
$controlWindow.ShowInTaskbar = $false
$controlWindow.WindowStyle = [System.Windows.WindowStyle]::ToolWindow
$controlWindow.Content = $control
$controlWindow.Show()
$controlWindow.UpdateLayout()
Wait-ForLayoutIdle

$rootScrollViewer = $control.FindName('RootScrollViewer')
$headerActionsPanel = $control.FindName('HeaderActionsPanel')
$homeFontsMetricsGrid = $control.FindName('HomeFontsMetricsGrid')
$homeColorsMetricsGrid = $control.FindName('HomeColorsMetricsGrid')
$fontTableHeaderRow = $control.FindName('FontTableHeaderRow')
$colorTableHeaderRow = $control.FindName('ColorTableHeaderRow')

Assert-True ($null -ne $rootScrollViewer) 'Fonts user control should expose RootScrollViewer for layout verification.'
Assert-True ($null -ne $headerActionsPanel) 'Fonts user control should expose HeaderActionsPanel for compact layout verification.'
Assert-True ($null -ne $homeFontsMetricsGrid) 'Fonts user control should expose HomeFontsMetricsGrid for compact layout verification.'
Assert-True ($null -ne $homeColorsMetricsGrid) 'Fonts user control should expose HomeColorsMetricsGrid for compact layout verification.'
Assert-True ($null -ne $fontTableHeaderRow) 'Fonts user control should expose FontTableHeaderRow for responsive table verification.'
Assert-True ($null -ne $colorTableHeaderRow) 'Fonts user control should expose ColorTableHeaderRow for responsive table verification.'

if ($rootScrollViewer -ne $null) {
    Assert-True ($rootScrollViewer.HorizontalScrollBarVisibility -eq [System.Windows.Controls.ScrollBarVisibility]::Disabled) 'Root scroll viewer should suppress horizontal scrolling in compact mode.'
}

if ($headerActionsPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetRow($headerActionsPanel) -eq 1) 'Header actions should stack below the title in compact mode.'
    Assert-True ([System.Windows.Controls.Grid]::GetColumn($headerActionsPanel) -eq 0) 'Header actions should align with the content column in compact mode.'
}

if ($homeFontsMetricsGrid -ne $null) {
    Assert-True ($homeFontsMetricsGrid.Columns -eq 1) 'Font metric cards should collapse to one column in compact mode.'
}

if ($homeColorsMetricsGrid -ne $null) {
    Assert-True ($homeColorsMetricsGrid.Columns -eq 1) 'Color metric cards should collapse to one column in compact mode.'
}

if ($fontTableHeaderRow -ne $null) {
    Assert-True ($fontTableHeaderRow.Visibility -eq [System.Windows.Visibility]::Collapsed) 'Font table header should collapse in narrow layouts.'
}

if ($colorTableHeaderRow -ne $null) {
    Assert-True ($colorTableHeaderRow.Visibility -eq [System.Windows.Visibility]::Collapsed) 'Color table header should collapse in narrow layouts.'
}

$controlWindow.Width = 520
$controlWindow.UpdateLayout()
Wait-ForLayoutIdle

if ($headerActionsPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetRow($headerActionsPanel) -eq 0) 'Header actions should return to the title row in wide mode.'
    Assert-True ([System.Windows.Controls.Grid]::GetColumn($headerActionsPanel) -eq 1) 'Header actions should return to the action column in wide mode.'
}

if ($homeFontsMetricsGrid -ne $null) {
    Assert-True ($homeFontsMetricsGrid.Columns -eq 2) 'Font metric cards should expand back to two columns in wide mode.'
}

if ($homeColorsMetricsGrid -ne $null) {
    Assert-True ($homeColorsMetricsGrid.Columns -eq 2) 'Color metric cards should expand back to two columns in wide mode.'
}

if ($fontTableHeaderRow -ne $null) {
    Assert-True ($fontTableHeaderRow.Visibility -eq [System.Windows.Visibility]::Visible) 'Font table header should return in wide layouts.'
}

if ($colorTableHeaderRow -ne $null) {
    Assert-True ($colorTableHeaderRow.Visibility -eq [System.Windows.Visibility]::Visible) 'Color table header should return in wide layouts.'
}

$controlWindow.Close()

$fontChoiceOne = [Activator]::CreateInstance($fontReplacementTargetType)
$fontChoiceOne.DisplayName = 'Arial'
$fontChoiceOne.NormalizedName = 'Arial'

$fontChoiceTwo = [Activator]::CreateInstance($fontReplacementTargetType)
$fontChoiceTwo.DisplayName = 'Calibri'
$fontChoiceTwo.NormalizedName = 'Calibri'

$fontChoices = [Array]::CreateInstance($fontReplacementTargetType, 2)
$fontChoices.SetValue($fontChoiceOne, 0)
$fontChoices.SetValue($fontChoiceTwo, 1)

$replaceFontsDialog = [Activator]::CreateInstance(
    $replaceFontsDialogType,
    @([string[]]@('Arial', 'Calibri'), [string[]]@('Arial'), $fontChoices))

$replaceFontsDialog.Width = 390
$replaceFontsDialog.ShowInTaskbar = $false
$replaceFontsDialog.Show()
$replaceFontsDialog.UpdateLayout()
Wait-ForLayoutIdle

$selectedFontPanel = $replaceFontsDialog.FindName('SelectedFontPanel')
$replacementFontPanel = $replaceFontsDialog.FindName('ReplacementFontPanel')

Assert-True ($null -ne $selectedFontPanel) 'Replace fonts dialog should expose SelectedFontPanel for layout verification.'
Assert-True ($null -ne $replacementFontPanel) 'Replace fonts dialog should expose ReplacementFontPanel for layout verification.'

if ($selectedFontPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetColumnSpan($selectedFontPanel) -eq 3) 'Selected font panel should span the full width in compact mode.'
}

if ($replacementFontPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetRow($replacementFontPanel) -eq 1) 'Replacement font panel should stack below the selected font panel in compact mode.'
    Assert-True ([System.Windows.Controls.Grid]::GetColumnSpan($replacementFontPanel) -eq 3) 'Replacement font panel should span the full width in compact mode.'
}

$replaceFontsDialog.Width = 640
$replaceFontsDialog.UpdateLayout()
Wait-ForLayoutIdle

if ($selectedFontPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetColumnSpan($selectedFontPanel) -eq 1) 'Selected font panel should return to a single column in wide mode.'
}

if ($replacementFontPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetRow($replacementFontPanel) -eq 0) 'Replacement font panel should return to the top row in wide mode.'
    Assert-True ([System.Windows.Controls.Grid]::GetColumn($replacementFontPanel) -eq 2) 'Replacement font panel should return to the second content column in wide mode.'
}

$replaceFontsDialog.Close()

$themeColor = [Activator]::CreateInstance($themeColorInfoType)
$themeColor.SchemeName = 'accent1'
$themeColor.DisplayName = 'Accent 1'
$themeColor.HexValue = '2255AA'

$themeColors = [Array]::CreateInstance($themeColorInfoType, 1)
$themeColors.SetValue($themeColor, 0)

$colorItem = [Activator]::CreateInstance($colorInventoryItemType)
$colorItem.UsageKind = [System.Enum]::Parse($colorUsageKindType, 'ShapeFill')
$colorItem.UsageKindLabel = 'Shape fill'
$colorItem.HexValue = 'AA3300'
$colorItem.RgbValue = '170, 51, 0'
$colorItem.UsesCount = 6
$colorItem.Locations = [Array]::CreateInstance($fontUsageLocationType, 0)

$sourceColors = [Array]::CreateInstance($colorInventoryItemType, 1)
$sourceColors.SetValue($colorItem, 0)

$replaceColorsDialog = [Activator]::CreateInstance(
    $replaceColorsDialogType,
    @($sourceColors, $themeColors, $colorItem))

$replaceColorsDialog.Width = 430
$replaceColorsDialog.ShowInTaskbar = $false
$replaceColorsDialog.Show()
$replaceColorsDialog.UpdateLayout()
Wait-ForLayoutIdle

$previewSummaryPanel = $replaceColorsDialog.FindName('PreviewSummaryPanel')
Assert-True ($null -ne $previewSummaryPanel) 'Replace colors dialog should expose PreviewSummaryPanel for layout verification.'

if ($previewSummaryPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetRow($previewSummaryPanel) -eq 1) 'Queued replacement summary should move below the swatches in compact mode.'
    Assert-True ([System.Windows.Controls.Grid]::GetColumnSpan($previewSummaryPanel) -eq 4) 'Queued replacement summary should span the preview width in compact mode.'
}

$replaceColorsDialog.Width = 620
$replaceColorsDialog.UpdateLayout()
Wait-ForLayoutIdle

if ($previewSummaryPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetRow($previewSummaryPanel) -eq 0) 'Queued replacement summary should return to the first row in wide mode.'
    Assert-True ([System.Windows.Controls.Grid]::GetColumn($previewSummaryPanel) -eq 3) 'Queued replacement summary should return to the details column in wide mode.'
}

$replaceColorsDialog.Close()

if ($failures.Count -gt 0) {
    Write-Host 'Responsive layout checks failed:' -ForegroundColor Red
    $failures | ForEach-Object { Write-Host (' - ' + $_) -ForegroundColor Red }
    exit 1
}

Write-Host 'Responsive layout checks passed.' -ForegroundColor Green
