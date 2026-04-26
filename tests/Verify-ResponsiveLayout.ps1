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

function Get-RelativeBounds {
    param(
        [System.Windows.FrameworkElement]$Element,
        [System.Windows.FrameworkElement]$RelativeTo
    )

    $origin = $Element.TranslatePoint((New-Object System.Windows.Point(0, 0)), $RelativeTo)
    return [pscustomobject]@{
        Left = $origin.X
        Top = $origin.Y
        Right = $origin.X + $Element.ActualWidth
        Bottom = $origin.Y + $Element.ActualHeight
        Width = $Element.ActualWidth
        Height = $Element.ActualHeight
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
$mainTabControl = $control.FindName('MainTabControl')
$headerActionsPanel = $control.FindName('HeaderActionsPanel')
$headerVisualPanel = $control.FindName('HeaderVisualPanel')
$headerGraphicShell = $control.FindName('HeaderGraphicShell')
$homeGraphicShell = $control.FindName('HomeGraphicShell')
$homeSignalsBoard = $control.FindName('HomeSignalsBoard')
$fontsTab = $control.FindName('FontsTab')
$colorsTab = $control.FindName('ColorsTab')
$homeTab = $control.FindName('HomeTab')
$fontTableHeaderRow = $control.FindName('FontTableHeaderRow')
$colorTableHeaderRow = $control.FindName('ColorTableHeaderRow')
$tabHeaderStrip = $mainTabControl.Template.FindName('TabHeaderStrip', $mainTabControl)

Assert-True ($null -ne $rootScrollViewer) 'Fonts user control should expose RootScrollViewer for layout verification.'
Assert-True ($null -ne $mainTabControl) 'Fonts user control should expose MainTabControl for layout verification.'
Assert-True ($null -ne $headerActionsPanel) 'Fonts user control should expose HeaderActionsPanel for compact layout verification.'
Assert-True ($null -ne $headerVisualPanel) 'Fonts user control should expose HeaderVisualPanel for compact layout verification.'
Assert-True ($null -ne $headerGraphicShell) 'Fonts user control should expose HeaderGraphicShell for responsive layout verification.'
Assert-True ($null -ne $homeGraphicShell) 'Fonts user control should expose HomeGraphicShell for responsive layout verification.'
Assert-True ($null -ne $homeSignalsBoard) 'Fonts user control should expose HomeSignalsBoard for home layout verification.'
Assert-True ($null -ne $tabHeaderStrip) 'Fonts user control should expose a custom tab header strip.'
Assert-True ($null -ne $fontsTab -and $null -ne $colorsTab -and $null -ne $homeTab) 'Fonts user control should expose all workspace tabs.'
Assert-True ($null -ne $fontTableHeaderRow) 'Fonts user control should expose FontTableHeaderRow for responsive table verification.'
Assert-True ($null -ne $colorTableHeaderRow) 'Fonts user control should expose ColorTableHeaderRow for responsive table verification.'

if ($rootScrollViewer -ne $null) {
    Assert-True ($rootScrollViewer.HorizontalScrollBarVisibility -eq [System.Windows.Controls.ScrollBarVisibility]::Disabled) 'Root scroll viewer should suppress horizontal scrolling in compact mode.'
}

if ($headerActionsPanel -ne $null) {
    Assert-True ($headerActionsPanel.HorizontalAlignment -eq [System.Windows.HorizontalAlignment]::Left) 'Header actions should align left in compact mode.'
}

if ($headerVisualPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetRow($headerVisualPanel) -eq 1) 'Header visuals should stack below the copy in compact mode.'
    Assert-True ([System.Windows.Controls.Grid]::GetColumn($headerVisualPanel) -eq 0) 'Header visuals should align with the content column in compact mode.'
}

if ($headerGraphicShell -ne $null) {
    Assert-True ($headerGraphicShell.Visibility -eq [System.Windows.Visibility]::Collapsed) 'Header graphic should hide in very narrow layouts.'
}

if ($homeGraphicShell -ne $null) {
    Assert-True ($homeGraphicShell.Visibility -eq [System.Windows.Visibility]::Collapsed) 'Home graphic should hide in very narrow layouts.'
}

if ($tabHeaderStrip -ne $null -and $fontsTab -ne $null -and $colorsTab -ne $null -and $homeTab -ne $null) {
    $stripBounds = Get-RelativeBounds -Element $tabHeaderStrip -RelativeTo $mainTabControl
    $fontsBounds = Get-RelativeBounds -Element $fontsTab -RelativeTo $mainTabControl
    $colorsBounds = Get-RelativeBounds -Element $colorsTab -RelativeTo $mainTabControl
    $homeBounds = Get-RelativeBounds -Element $homeTab -RelativeTo $mainTabControl

    Assert-True ($fontsBounds.Right -le ($colorsBounds.Left + 0.5)) 'Fonts and Colors tab headers should not overlap in compact mode.'
    Assert-True ($colorsBounds.Right -le ($homeBounds.Left + 0.5)) 'Colors and Home tab headers should not overlap in compact mode.'
    Assert-True ($homeBounds.Right -le ($stripBounds.Right + 1)) 'Home tab header should stay within the custom tab strip.'
}

if ($fontTableHeaderRow -ne $null) {
    Assert-True ($fontTableHeaderRow.Visibility -eq [System.Windows.Visibility]::Collapsed) 'Font table header should collapse in narrow layouts.'
}

if ($colorTableHeaderRow -ne $null) {
    Assert-True ($colorTableHeaderRow.Visibility -eq [System.Windows.Visibility]::Collapsed) 'Color table header should collapse in narrow layouts.'
}

$controlWindow.Width = 620
$controlWindow.UpdateLayout()
Wait-ForLayoutIdle

if ($headerActionsPanel -ne $null) {
    Assert-True ($headerActionsPanel.HorizontalAlignment -eq [System.Windows.HorizontalAlignment]::Right) 'Header actions should align right in wide mode.'
}

if ($headerVisualPanel -ne $null) {
    Assert-True ([System.Windows.Controls.Grid]::GetRow($headerVisualPanel) -eq 0) 'Header visuals should return to the first row in wide mode.'
    Assert-True ([System.Windows.Controls.Grid]::GetColumn($headerVisualPanel) -eq 1) 'Header visuals should return to the rail column in wide mode.'
}

if ($headerGraphicShell -ne $null) {
    Assert-True ($headerGraphicShell.Visibility -eq [System.Windows.Visibility]::Visible) 'Header graphic should return in wide layouts.'
}

if ($homeGraphicShell -ne $null) {
    Assert-True ($homeGraphicShell.Visibility -eq [System.Windows.Visibility]::Visible) 'Home graphic should return in wide layouts.'
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
