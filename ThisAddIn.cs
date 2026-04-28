using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using MorphosPowerPointAddIn.Ribbon;
using MorphosPowerPointAddIn.Services;
using MorphosPowerPointAddIn.UI;
using MorphosPowerPointAddIn.Utilities;
using MorphosPowerPointAddIn.ViewModels;

namespace MorphosPowerPointAddIn
{
    public partial class ThisAddIn
    {
        private const int MinimumTaskPaneWidth = 300;
        private CustomTaskPane _fontsTaskPane;
        private FontsTaskPaneHost _taskPaneHost;
        private MorphosRibbon _ribbon;
        private PowerPointPresentationService _presentationService;
        private FontsPaneViewModel _paneViewModel;
        private readonly object _refreshSync = new object();
        private Task<bool> _activeRefreshTask;
        private string _activeRefreshPresentationKey = string.Empty;
        private int _activeRefreshVersion;
        private string _lastScannedPresentationKey = string.Empty;
        private bool _isTaskPaneRequestedVisible;
        private IntPtr _currentTaskPaneWindowKey = IntPtr.Zero;
        private int _refreshSuppressionCount;
        private int _warmRefreshVersion;
        private int _windowActivateDebounceVersion;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                _presentationService = new PowerPointPresentationService(this.Application);
                _paneViewModel = new FontsPaneViewModel(_presentationService);

                this.Application.PresentationOpen += Application_PresentationOpen;
                this.Application.WindowActivate += Application_WindowActivate;
                this.Application.PresentationCloseFinal += Application_PresentationCloseFinal;
            }
            catch (Exception ex)
            {
                ErrorReporter.Show("Morphos failed to start.", ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (this.Application != null)
                {
                    this.Application.PresentationOpen -= Application_PresentationOpen;
                    this.Application.WindowActivate -= Application_WindowActivate;
                    this.Application.PresentationCloseFinal -= Application_PresentationCloseFinal;
                }

                RemoveCurrentTaskPane();
            }
            catch
            {
            }
        }

        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new MorphosRibbon(this);
            return _ribbon;
        }

        internal Task<bool> RefreshAsync(bool force = false, bool showErrors = true)
        {
            if (_paneViewModel == null)
            {
                return Task.FromResult(false);
            }

            if (Volatile.Read(ref _refreshSuppressionCount) > 0)
            {
                return Task.FromResult(false);
            }

            var activePresentation = GetActivePresentation();
            var activePresentationKey = GetPresentationKey(activePresentation);
            if (string.IsNullOrWhiteSpace(activePresentationKey))
            {
                return Task.FromResult(false);
            }

            lock (_refreshSync)
            {
                if (_activeRefreshTask != null
                    && !_activeRefreshTask.IsCompleted
                    && string.Equals(_activeRefreshPresentationKey, activePresentationKey, StringComparison.OrdinalIgnoreCase))
                {
                    return _activeRefreshTask;
                }

                if (!force && !HasPresentationChanged(activePresentation))
                {
                    return Task.FromResult(true);
                }

                var refreshVersion = ++_activeRefreshVersion;
                _activeRefreshPresentationKey = activePresentationKey;
                _activeRefreshTask = RefreshCoreAsync(activePresentationKey, showErrors, refreshVersion);
                return _activeRefreshTask;
            }
        }

        private async Task<bool> RefreshCoreAsync(string requestedPresentationKey, bool showErrors, int refreshVersion)
        {
            try
            {
                await _paneViewModel.ScanAsync(showErrors).ConfigureAwait(true);
                if (!_paneViewModel.LastScanSucceeded)
                {
                    return false;
                }

                var activePresentationKey = GetPresentationKey(GetActivePresentation());
                _lastScannedPresentationKey = string.IsNullOrWhiteSpace(activePresentationKey)
                    ? requestedPresentationKey
                    : activePresentationKey;
                return true;
            }
            catch (Exception ex)
            {
                if (showErrors)
                {
                    ErrorReporter.Show("Morphos could not scan the active presentation.", ex);
                }

                return false;
            }
            finally
            {
                lock (_refreshSync)
                {
                    if (refreshVersion == _activeRefreshVersion
                        && _activeRefreshTask != null
                        && _activeRefreshTask.IsCompleted)
                    {
                        _activeRefreshTask = null;
                        _activeRefreshPresentationKey = string.Empty;
                    }
                }
            }
        }

        internal void ToggleTaskPane(bool visible)
        {
            _isTaskPaneRequestedVisible = visible;

            try
            {
                if (!visible)
                {
                    if (_fontsTaskPane != null)
                    {
                        _fontsTaskPane.Visible = false;
                    }

                    NotifyRibbonStateChanged();
                    return;
                }

                var taskPane = EnsureTaskPaneForActiveWindow();
                if (taskPane == null)
                {
                    _isTaskPaneRequestedVisible = false;
                    NotifyRibbonStateChanged();
                    ErrorReporter.Show("Open a presentation window before using Morphos.");
                    return;
                }

                taskPane.Width = Math.Max(taskPane.Width, MinimumTaskPaneWidth);
                taskPane.Visible = true;
                NotifyRibbonStateChanged();
                RequestPaneRefresh(!_paneViewModel.HasCompletedScan || !_paneViewModel.LastScanSucceeded || HasPresentationChanged());
            }
            catch (Exception ex)
            {
                _isTaskPaneRequestedVisible = false;
                NotifyRibbonStateChanged();
                ErrorReporter.Show("Morphos could not open its inspector pane.", ex);
            }
        }

        internal bool IsTaskPaneVisible
        {
            get
            {
                if (_fontsTaskPane == null || !_fontsTaskPane.Visible)
                {
                    return false;
                }

                var activeWindowKey = GetWindowKey(GetActiveDocumentWindow());
                return activeWindowKey != IntPtr.Zero && activeWindowKey == _currentTaskPaneWindowKey;
            }
        }

        internal PowerPointPresentationService PresentationService => _presentationService;

        internal IDisposable SuspendAutoRefresh()
        {
            Interlocked.Increment(ref _refreshSuppressionCount);
            return new RefreshSuppressionScope(this);
        }

        internal void RecoverTaskPaneAfterPresentationMutation(bool forceRefresh = false)
        {
            if (!_isTaskPaneRequestedVisible)
            {
                return;
            }

            try
            {
                RemoveCurrentTaskPane();

                var taskPane = EnsureTaskPaneForActiveWindow();
                if (taskPane == null)
                {
                    NotifyRibbonStateChanged();
                    return;
                }

                taskPane.Width = Math.Max(taskPane.Width, MinimumTaskPaneWidth);
                taskPane.Visible = true;
                NotifyRibbonStateChanged();

                if (forceRefresh)
                {
                    RequestPaneRefresh(true);
                }
            }
            catch
            {
            }
        }

        private CustomTaskPane EnsureTaskPaneForActiveWindow()
        {
            var activeWindow = GetActiveDocumentWindow();
            var activeWindowKey = GetWindowKey(activeWindow);
            if (activeWindow == null || activeWindowKey == IntPtr.Zero)
            {
                return null;
            }

            if (_fontsTaskPane != null && activeWindowKey == _currentTaskPaneWindowKey)
            {
                return _fontsTaskPane;
            }

            RemoveCurrentTaskPane();

            _taskPaneHost = new FontsTaskPaneHost(_paneViewModel);
            _fontsTaskPane = this.CustomTaskPanes.Add(_taskPaneHost, "Morphos", activeWindow);
            _fontsTaskPane.Width = MinimumTaskPaneWidth;
            _fontsTaskPane.VisibleChanged += FontsTaskPane_VisibleChanged;
            _currentTaskPaneWindowKey = activeWindowKey;
            return _fontsTaskPane;
        }

        private void RemoveCurrentTaskPane()
        {
            if (_fontsTaskPane != null)
            {
                try
                {
                    _fontsTaskPane.VisibleChanged -= FontsTaskPane_VisibleChanged;
                }
                catch
                {
                }

                try
                {
                    this.CustomTaskPanes.Remove(_fontsTaskPane);
                }
                catch
                {
                }

                _fontsTaskPane = null;
            }

            if (_taskPaneHost != null)
            {
                try
                {
                    _taskPaneHost.Dispose();
                }
                catch
                {
                }

                _taskPaneHost = null;
            }

            _currentTaskPaneWindowKey = IntPtr.Zero;
        }

        private void FontsTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            _isTaskPaneRequestedVisible = _fontsTaskPane != null && _fontsTaskPane.Visible;
            NotifyRibbonStateChanged();

            if (_isTaskPaneRequestedVisible)
            {
                RequestPaneRefresh(!_paneViewModel.HasCompletedScan || !_paneViewModel.LastScanSucceeded || HasPresentationChanged());
            }
        }

        private void Application_PresentationOpen(Presentation pres)
        {
            SyncTaskPaneToActiveWindow(false);
            if (_isTaskPaneRequestedVisible)
            {
                QueueWarmRefresh(true);
            }
        }

        private void Application_WindowActivate(Presentation pres, DocumentWindow wn)
        {
            var activatedWindowKey = GetWindowKey(wn);
            var debounceVersion = Interlocked.Increment(ref _windowActivateDebounceVersion);
            _ = HandleWindowActivateAsync(activatedWindowKey, debounceVersion);
        }

        private async Task HandleWindowActivateAsync(IntPtr activatedWindowKey, int debounceVersion)
        {
            await Task.Delay(100).ConfigureAwait(true);
            if (debounceVersion != Volatile.Read(ref _windowActivateDebounceVersion))
            {
                return;
            }

            if (_isTaskPaneRequestedVisible)
            {
                var forceRefresh = activatedWindowKey != IntPtr.Zero && activatedWindowKey != _currentTaskPaneWindowKey;
                SyncTaskPaneToActiveWindow(forceRefresh);
                return;
            }

            if (_fontsTaskPane != null && activatedWindowKey != IntPtr.Zero && activatedWindowKey != _currentTaskPaneWindowKey)
            {
                RemoveCurrentTaskPane();
            }

            NotifyRibbonStateChanged();
        }

        private void Application_PresentationCloseFinal(Presentation pres)
        {
            _lastScannedPresentationKey = string.Empty;
            Interlocked.Increment(ref _warmRefreshVersion);
            SyncTaskPaneToActiveWindow(true);
        }

        private void NotifyRibbonStateChanged()
        {
            try
            {
                _ribbon?.Invalidate();
            }
            catch
            {
            }
        }

        private PowerPoint.Presentation GetActivePresentation()
        {
            try
            {
                return this.Application?.ActivePresentation;
            }
            catch
            {
                return null;
            }
        }

        private DocumentWindow GetActiveDocumentWindow()
        {
            try
            {
                return this.Application?.ActiveWindow;
            }
            catch
            {
                return null;
            }
        }

        private bool HasPresentationChanged(Presentation presentation = null)
        {
            var key = GetPresentationKey(presentation ?? GetActivePresentation());
            return !string.Equals(key, _lastScannedPresentationKey, StringComparison.OrdinalIgnoreCase);
        }

        private void SyncTaskPaneToActiveWindow(bool forceRefresh)
        {
            if (!_isTaskPaneRequestedVisible)
            {
                NotifyRibbonStateChanged();
                return;
            }

            var taskPane = EnsureTaskPaneForActiveWindow();
            if (taskPane == null)
            {
                RemoveCurrentTaskPane();
                NotifyRibbonStateChanged();
                return;
            }

            taskPane.Width = Math.Max(taskPane.Width, MinimumTaskPaneWidth);
            if (!taskPane.Visible)
            {
                taskPane.Visible = true;
            }

            NotifyRibbonStateChanged();

            if (forceRefresh)
            {
                RequestPaneRefresh(true);
                return;
            }

            if (HasPresentationChanged() || !_paneViewModel.HasCompletedScan || !_paneViewModel.LastScanSucceeded)
            {
                RequestPaneRefresh(false);
            }
        }

        private void RequestPaneRefresh(bool force)
        {
            _ = RefreshAsync(force, false);
        }

        private void QueueWarmRefresh(bool force)
        {
            if (!_isTaskPaneRequestedVisible && _paneViewModel != null && _paneViewModel.HasCompletedScan)
            {
                return;
            }

            var refreshVersion = Interlocked.Increment(ref _warmRefreshVersion);
            _ = WarmRefreshAsync(force, refreshVersion);
        }

        private async Task WarmRefreshAsync(bool force, int refreshVersion)
        {
            var delays = new[] { 0, 300, 1200 };
            for (var attempt = 0; attempt < delays.Length; attempt++)
            {
                if (refreshVersion != Volatile.Read(ref _warmRefreshVersion))
                {
                    return;
                }

                var delay = delays[attempt];
                if (delay > 0)
                {
                    await Task.Delay(delay).ConfigureAwait(true);
                }

                if (refreshVersion != Volatile.Read(ref _warmRefreshVersion))
                {
                    return;
                }

                if (!_isTaskPaneRequestedVisible && _paneViewModel != null && _paneViewModel.HasCompletedScan)
                {
                    return;
                }

                var activePresentation = GetActivePresentation();
                if (string.IsNullOrWhiteSpace(GetPresentationKey(activePresentation)))
                {
                    continue;
                }

                var shouldForceRefresh = force
                    || (attempt == 0 && (!_paneViewModel.HasCompletedScan || !_paneViewModel.LastScanSucceeded || HasPresentationChanged(activePresentation)));
                var refreshed = await RefreshAsync(shouldForceRefresh, false).ConfigureAwait(true);
                if (refreshed)
                {
                    return;
                }
            }
        }

        private static string GetPresentationKey(Presentation presentation)
        {
            if (presentation == null)
            {
                return string.Empty;
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(presentation.FullName))
                {
                    return presentation.FullName;
                }
            }
            catch
            {
            }

            try
            {
                return presentation.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static IntPtr GetWindowKey(DocumentWindow window)
        {
            if (window == null)
            {
                return IntPtr.Zero;
            }

            try
            {
                var unknown = Marshal.GetIUnknownForObject(window);
                try
                {
                    return unknown;
                }
                finally
                {
                    Marshal.Release(unknown);
                }
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        private void ReleaseAutoRefreshSuppression()
        {
            var nextValue = Interlocked.Decrement(ref _refreshSuppressionCount);
            if (nextValue < 0)
            {
                Interlocked.Exchange(ref _refreshSuppressionCount, 0);
            }
        }

        private sealed class RefreshSuppressionScope : IDisposable
        {
            private ThisAddIn _owner;

            public RefreshSuppressionScope(ThisAddIn owner)
            {
                _owner = owner;
            }

            public void Dispose()
            {
                var owner = _owner;
                if (owner == null)
                {
                    return;
                }

                _owner = null;
                owner.ReleaseAutoRefreshSuppression();
            }
        }
    }
}
