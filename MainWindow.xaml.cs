using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using NetIsolatePlus.Models;
using NetIsolatePlus.Services;

namespace NetIsolatePlus
{
    public partial class MainWindow : Window
    {
        private readonly NicManager _nicManager = new();
        private readonly SettingsStore _settings = new("Tensioncore Administration Services", "NetIsolatePlus");
        private readonly WindowPlacementService _placement;

        private bool _isIsolationActive = false;
        private Dictionary<string, bool> _previousStates = new();

        private bool _showVirtual;
        private bool _startWithWindows;
        private string _sortMode = "Enabled"; // "Enabled" or "Name"
        private bool _suspendUi;

        private ManagementEventWatcher? _nicModifyWatcher;
        private ManagementEventWatcher? _nicCreateWatcher;
        private ManagementEventWatcher? _nicDeleteWatcher;

        private readonly DispatcherTimer _refreshDebounce = new() { Interval = TimeSpan.FromMilliseconds(800) };
        private volatile bool _suspendAutoRefresh;   // true while we perform bulk changes

        private DateTime _ignoreExternalEventsUntilUtc = DateTime.MinValue; // cooldown after bulk ops

        private string _lastStateToken = "";                 // snapshot of adapter states
        private bool _wmiWatchersOk = false;                 // only use NetworkChange if WMI fails

        // --- Bulk hardening (no behavior changes unless user spams bulk toggles) ---
        private readonly SemaphoreSlim _bulkOpLock = new(1, 1);
        private DateTime _bulkCooldownUntilUtc = DateTime.MinValue;

        // Shutdown guard for watchers/dispatcher calls
        private volatile bool _isClosing = false;

        public MainWindow()
        {
            InitializeComponent();
            _placement = new WindowPlacementService("MainWindow", _settings);
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _placement.Restore(this);
            InitSpinner();

            // Load prefs
            _showVirtual = _settings.Load("UI.ShowVirtual", false);
            StartupManager.Reconcile();
            _startWithWindows = StartupManager.IsEnabled();
            _sortMode = _settings.Load("UI.SortMode", "Enabled");

            // Set UI switches without firing handlers
            _suspendUi = true;
            ToggleVirtual.IsChecked = _showVirtual;
            ToggleStartup.IsChecked = _startWithWindows;
            SelectSortByTag(_sortMode is "Enabled" or "Disabled" or "Name" or "NameDesc" ? _sortMode : "Enabled");
            _suspendUi = false;

            // Debounce handler FIRST, so watchers use it
            _refreshDebounce.Tick += async (_, __) =>
            {
                if (_isClosing) return;
                _refreshDebounce.Stop();
                await RefreshUiAsync();   // external changes: normal (not forced)
            };

            // Start auto-refresh watchers (they call ScheduleExternalRefresh)
            SetupAutoRefreshWatchers();

            // Initial paint — force full rebuild so filter/sort/UI state applies even if token is unchanged
            await RefreshUiAsync(force: true);

            ApplySizeConstraints();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            _isClosing = true;

            try { _refreshDebounce.Stop(); } catch { }

            _placement.Save(this);
            try { _nicModifyWatcher?.Stop(); _nicModifyWatcher?.Dispose(); } catch { }
            try { _nicCreateWatcher?.Stop(); _nicCreateWatcher?.Dispose(); } catch { }
            try { _nicDeleteWatcher?.Stop(); _nicDeleteWatcher?.Dispose(); } catch { }

            try
            {
                NetworkChange.NetworkAddressChanged -= OnNetworkChanged;
                NetworkChange.NetworkAvailabilityChanged -= OnNetworkAvailabilityChanged;
            }
            catch { }
        }

        // Helpers
        private static string ComputeStateToken(IEnumerable<NicInfo> adapters)
        {
            var ordered = adapters.OrderBy(a => a.Id, StringComparer.OrdinalIgnoreCase);
            System.Text.StringBuilder sb = new();
            foreach (var a in ordered)
                sb.Append(a.Id)
                  .Append('|').Append(a.Enabled ? '1' : '0')
                  .Append('|').Append(a.NetConnectionStatus?.ToString() ?? "")
                  .Append(';');
            return sb.ToString();
        }

        private void RequestRefresh(TimeSpan? delay = null)
        {
            if (_isClosing) return;

            if (delay.HasValue) _refreshDebounce.Interval = delay.Value;
            _refreshDebounce.Stop();
            _refreshDebounce.Start();
        }

        private void ScheduleExternalRefresh(TimeSpan? delay = null)
        {
            if (_isClosing) return;
            if (_suspendAutoRefresh) return;
            if (DateTime.UtcNow < _ignoreExternalEventsUntilUtc) return;
            RequestRefresh(delay);
        }

        private void SetupAutoRefreshWatchers()
        {
            _wmiWatchersOk = false;

            try
            {
                var scope = new ManagementScope(@"\\.\root\CIMV2");
                scope.Connect();

                var modifyQ = new WqlEventQuery(
                    "__InstanceModificationEvent",
                    TimeSpan.FromSeconds(2),
                    "TargetInstance ISA 'Win32_NetworkAdapter' AND " +
                    "(TargetInstance.NetEnabled != PreviousInstance.NetEnabled OR " +
                    " TargetInstance.NetConnectionStatus != PreviousInstance.NetConnectionStatus)");

                _nicModifyWatcher = new ManagementEventWatcher(scope, modifyQ);
                _nicModifyWatcher.EventArrived += (_, __) =>
                {
                    if (_isClosing) return;
                    try
                    {
                        Dispatcher.BeginInvoke(new Action(() => ScheduleExternalRefresh()));
                    }
                    catch { }
                };
                _nicModifyWatcher.Start();
                _wmiWatchersOk = true;

                try
                {
                    var createQ = new WqlEventQuery("__InstanceCreationEvent", TimeSpan.FromSeconds(2),
                        "TargetInstance ISA 'Win32_NetworkAdapter'");
                    _nicCreateWatcher = new ManagementEventWatcher(scope, createQ);
                    _nicCreateWatcher.EventArrived += (_, __) =>
                    {
                        if (_isClosing) return;
                        try
                        {
                            Dispatcher.BeginInvoke(new Action(() => ScheduleExternalRefresh(TimeSpan.FromMilliseconds(1200))));
                        }
                        catch { }
                    };
                    _nicCreateWatcher.Start();
                }
                catch { }

                try
                {
                    var deleteQ = new WqlEventQuery("__InstanceDeletionEvent", TimeSpan.FromSeconds(2),
                        "TargetInstance ISA 'Win32_NetworkAdapter'");
                    _nicDeleteWatcher = new ManagementEventWatcher(scope, deleteQ);
                    _nicDeleteWatcher.EventArrived += (_, __) =>
                    {
                        if (_isClosing) return;
                        try
                        {
                            Dispatcher.BeginInvoke(new Action(() => ScheduleExternalRefresh(TimeSpan.FromMilliseconds(1200))));
                        }
                        catch { }
                    };
                    _nicDeleteWatcher.Start();
                }
                catch { }
            }
            catch
            {
                _wmiWatchersOk = false;
            }

            if (!_wmiWatchersOk)
            {
                try
                {
                    // IMPORTANT: subscribe using named handlers so we can unsubscribe cleanly on exit.
                    NetworkChange.NetworkAddressChanged += OnNetworkChanged;
                    NetworkChange.NetworkAvailabilityChanged += OnNetworkAvailabilityChanged;
                }
                catch { }
            }
        }

        private void OnNetworkChanged(object? s, EventArgs e)
        {
            if (_isClosing) return;
            try { Dispatcher.BeginInvoke(new Action(() => ScheduleExternalRefresh())); } catch { }
        }

        private void OnNetworkAvailabilityChanged(object? s, NetworkAvailabilityEventArgs e)
        {
            if (_isClosing) return;
            try { Dispatcher.BeginInvoke(new Action(() => ScheduleExternalRefresh())); } catch { }
        }

        private DoubleAnimation? _spinAnim;

        private void InitSpinner()
        {
            _spinAnim = new DoubleAnimation(0, 360, TimeSpan.FromSeconds(0.8))
            {
                RepeatBehavior = RepeatBehavior.Forever,
                SpeedRatio = 0.5
            };
        }

        private void ShowBusy(bool show)
        {
            BusyOverlay.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            BusyOverlay.IsHitTestVisible = show;

            if (show)
                SpinnerRotate.BeginAnimation(RotateTransform.AngleProperty, _spinAnim);
            else
                SpinnerRotate.BeginAnimation(RotateTransform.AngleProperty, null);
        }

        // NEW: run work on a dedicated STA thread (keeps WPF UI thread free so spinner can animate)
        private static Task RunStaAsync(Action action)
        {
            var tcs = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);

            Thread thread = new Thread(() =>
            {
                try
                {
                    action();
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });

            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }

        private static readonly string[] _virtualKeywords = new[]
        {
            "hyper-v", "vethernet",
            "vmware", "virtualbox",
            "loopback", "isatap", "teredo",
            "bluetooth",
            "tap", "tun", "tunnel",
            "miniport",
            "host-only",
            "vpn"
        };

        // FIX: Do NOT let PhysicalAdapter==true override obvious keyword matches.
        // Some virtual adapters report PhysicalAdapter=true depending on provider/driver.
        private static bool IsVirtualAdapter(NicInfo nic)
        {
            static string Safe(string? s) => s ?? "";

            bool Has(string s) => _virtualKeywords.Any(k => s.Contains(k, StringComparison.OrdinalIgnoreCase));

            bool keywordMatch =
                Has(Safe(nic.Name)) ||
                Has(Safe(nic.Description));

            if (nic.PhysicalAdapter.HasValue)
            {
                // PhysicalAdapter=false is a strong signal for virtual.
                if (nic.PhysicalAdapter.Value == false)
                    return true;

                // PhysicalAdapter=true may still be virtual; use keyword fallback.
                return keywordMatch;
            }

            return keywordMatch;
        }

        private static string GetStatusText(NicInfo nic)
        {
            if (!nic.Enabled) return "Disabled";

            return nic.NetConnectionStatus switch
            {
                2 => "Connected",
                7 => "Not Connected",
                1 => "Connecting",
                3 => "Disconnecting",
                4 => "Hardware not present",
                5 => "Hardware disabled",
                6 => "Hardware malfunction",
                8 => "Authenticating",
                9 => "Authentication succeeded",
                10 => "Authentication failed",
                11 => "Invalid address",
                12 => "Credentials required",
                0 => "Disconnected",
                null => "Enabled",
                _ => "Enabled"
            };
        }

        private void UpdateBulkToggleState(IEnumerable<NicInfo> adapters)
        {
            // Respect the current "Virtual adapters" view for bulk toggle behavior.
            IEnumerable<NicInfo> target = adapters;
            if (!_showVirtual)
                target = target.Where(n => !IsVirtualAdapter(n));

            var list = target.ToList();
            bool any = list.Count > 0;
            bool allEnabled = any && list.All(n => n.Enabled);

            _suspendUi = true;
            BulkToggle.IsChecked = allEnabled;
            BulkToggle.IsEnabled = any && !_isIsolationActive;
            _suspendUi = false;
        }

        private async Task RefreshUiAsync(bool force = false)
        {
            if (_isClosing) return;

            var adapters = await Task.Run(() => _nicManager.ListAdapters());

            if (!force)
            {
                string token = ComputeStateToken(adapters);
                if (token == _lastStateToken) return;
                _lastStateToken = token;
            }

            ListHost.Children.Clear();

            IEnumerable<NicInfo> q = adapters;

            if (!_showVirtual)
                q = q.Where(n => !IsVirtualAdapter(n));

            q = _sortMode switch
            {
                "Disabled" => q.OrderBy(n => n.Enabled).ThenBy(n => n.Name, StringComparer.OrdinalIgnoreCase),
                "Name" => q.OrderBy(n => n.Name, StringComparer.OrdinalIgnoreCase),
                "NameDesc" => q.OrderByDescending(n => n.Name, StringComparer.OrdinalIgnoreCase),
                _ => q.OrderByDescending(n => n.Enabled).ThenBy(n => n.Name, StringComparer.OrdinalIgnoreCase),
            };

            foreach (var nic in q)
                ListHost.Children.Add(BuildCard(nic));

            if (ListHost.Children.Count > 0 && ListHost.Children[^1] is FrameworkElement last)
                last.Margin = new Thickness(last.Margin.Left, last.Margin.Top, last.Margin.Right, 0);

            UpdateBulkToggleState(adapters);
            ApplySizeConstraints();
        }

        private async void ToggleVirtual_Checked(object sender, RoutedEventArgs e)
        {
            if (_suspendUi) return;
            _showVirtual = ToggleVirtual.IsChecked == true;
            _settings.Save("UI.ShowVirtual", _showVirtual);
            await RefreshUiAsync(force: true);
        }

        private void ToggleStartup_Checked(object sender, RoutedEventArgs e)
        {
            if (_suspendUi) return;
            try
            {
                _startWithWindows = ToggleStartup.IsChecked == true;
                StartupManager.SetEnabled(_startWithWindows);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Could not update startup setting:\n\n" + ex.Message,
                    "NetIsolate+", MessageBoxButton.OK, MessageBoxImage.Error);
                _suspendUi = true;
                ToggleStartup.IsChecked = StartupManager.IsEnabled();
                _suspendUi = false;
            }
        }

        private async void BulkToggle_Checked(object sender, RoutedEventArgs e)
        {
            if (_suspendUi) return;

            bool enable = BulkToggle.IsChecked == true;

            if (!enable && IsRdpSession())
            {
                var res = MessageBox.Show(this,
                    "You appear to be in a Remote Desktop session. Disabling all adapters may disconnect you.\n\nContinue?",
                    "NetIsolate+", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (res != MessageBoxResult.Yes)
                {
                    // revert UI to previous computed state
                    await RefreshUiAsync(force: true);
                    return;
                }
            }

            await BulkSetAdaptersAsync(enable: enable);
        }

        private async void SortCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suspendUi) return;
            if (SortCombo.SelectedItem is ComboBoxItem it)
            {
                _sortMode = (string)it.Tag;
                _settings.Save("UI.SortMode", _sortMode);
                await RefreshUiAsync(force: true);
            }
        }

        private static bool IsRdpSession()
            => (Environment.GetEnvironmentVariable("SESSIONNAME") ?? "").StartsWith("RDP-", StringComparison.OrdinalIgnoreCase);

        private async void IsolateBtn_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button btn || btn.Tag is not NicInfo nic) return;

            if (IsRdpSession())
            {
                var res = MessageBox.Show(this,
                    "You appear to be in a Remote Desktop session. Isolating may disconnect you.\n\nContinue?",
                    "NetIsolate+", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (res != MessageBoxResult.Yes) return;
            }

            btn.IsEnabled = false;
            ShowBusy(true);
            _suspendAutoRefresh = true;

            try
            {
                await Task.Delay(1);

                await Task.Run(() =>
                {
                    if (!_isIsolationActive)
                    {
                        _previousStates = _nicManager.CaptureStates();
                        _nicManager.IsolateTo(nic.Id);
                        _isIsolationActive = true;
                    }
                    else
                    {
                        if (_nicManager.CurrentIsolated?.Id == nic.Id)
                        {
                            _nicManager.RestoreStates(_previousStates);
                            _previousStates.Clear();
                            _nicManager.CurrentIsolated = null;
                            _isIsolationActive = false;
                        }
                        else
                        {
                            _nicManager.IsolateTo(nic.Id);
                        }
                    }
                });

                _refreshDebounce.Stop();
                await RefreshUiAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Isolation failed:\n\n" + ex.Message,
                    "NetIsolate+", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _ignoreExternalEventsUntilUtc = DateTime.UtcNow.AddSeconds(2.5);
                _suspendAutoRefresh = false;
                ShowBusy(false);
                btn.IsEnabled = true;
            }
        }

        private void SelectSortByTag(string tag)
        {
            foreach (var item in SortCombo.Items)
                if (item is ComboBoxItem it && (string)it.Tag == tag)
                {
                    SortCombo.SelectedItem = it;
                    return;
                }
        }

        private void ApplySizeConstraints()
        {
            var wa = SystemParameters.WorkArea;
            MaxHeight = wa.Height - 40;
            MaxWidth = wa.Width - 120;

            ControlBar.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            double controlsH = ControlBar.DesiredSize.Height;

            Scroller.MaxHeight = MaxHeight - 44 - controlsH - 2 - 10;
        }

        private Border BuildCard(NicInfo nic)
        {
            var border = new Border();
            border.Style = (Style)FindResource("CardBorder");

            bool isIsolated = _isIsolationActive && _nicManager.CurrentIsolated?.Id == nic.Id;
            if (isIsolated)
            {
                border.BorderBrush = (Brush)FindResource("AccentBrush");
                border.BorderThickness = new Thickness(2);
                border.Background = new SolidColorBrush(Color.FromRgb(0x18, 0x1E, 0x22));
            }

            var grid = new Grid { Margin = new Thickness(2) };
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var stack = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(2, 0, 0, 0) };
            var title = new TextBlock { Text = nic.Name, FontSize = 13, FontWeight = FontWeights.SemiBold };
            var subtitle = new TextBlock { Text = nic.Description, Foreground = (Brush)FindResource("SubTextBrush"), FontSize = 11, TextTrimming = TextTrimming.CharacterEllipsis };

            var statusText = GetStatusText(nic);
            var status = new TextBlock
            {
                Text = statusText,
                Foreground = (!nic.Enabled) ? (Brush)FindResource("SubTextBrush") : (Brush)FindResource("TextBrush"),
                FontSize = 11
            };

            stack.Children.Add(title);
            stack.Children.Add(subtitle);
            stack.Children.Add(status);

            var actions = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0, 0, 0) };

            var toggle = new CheckBox
            {
                IsChecked = nic.Enabled,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0),
                ToolTip = "Enable/Disable this adapter",
                Style = (Style)FindResource("ToggleSwitchStyle"),

                // If isolation is active, only the isolated adapter's toggle remains usable.
                IsEnabled = !_isIsolationActive || isIsolated
            };
            toggle.Checked += async (_, __) => { await ToggleAdapterAsync(nic, true); };
            toggle.Unchecked += async (_, __) => { await ToggleAdapterAsync(nic, false); };

            var iso = MakeButton(isIsolated ? "Isolated" : "Isolate");
            iso.Tag = nic;
            iso.Click += IsolateBtn_Click;

            var statusBtn = MakeButton("Status");
            statusBtn.Tag = nic;
            statusBtn.Margin = new Thickness(8, 0, 0, 0);
            statusBtn.Click += StatusBtn_Click;

            actions.Children.Add(toggle);
            actions.Children.Add(iso);
            actions.Children.Add(statusBtn);

            grid.Children.Add(stack);
            Grid.SetColumn(actions, 1);
            grid.Children.Add(actions);

            border.Child = grid;
            return border;
        }

        private async void StatusBtn_Click(object? sender, RoutedEventArgs e)
        {
            if (_isClosing) return;
            if (sender is not Button b || b.Tag is not NicInfo nic) return;

            b.IsEnabled = false;
            ShowBusy(true);

            try
            {
                // Ensure overlay paints before doing any work
                await Dispatcher.Yield(DispatcherPriority.Background);
                if (_isClosing) return;

                // Run Shell COM on a dedicated STA thread so the UI thread stays responsive (spinner rotates)
                await RunStaAsync(() => _nicManager.OpenStatus(nic));
                if (_isClosing) return;

                // Centering/polling happens async; keep it on UI thread
                await ExternalWindowPlacer.CenterStatusWindowAsync(this, nic.Name);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    $"Failed to open status for '{nic.Name}':\n\n{ex.Message}",
                    "NetIsolate+", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ShowBusy(false);
                if (!_isClosing) b.IsEnabled = true;
            }
        }

        private Button MakeButton(string text) => new Button
        {
            Content = text,
            Margin = new Thickness(0),
            Padding = new Thickness(14, 8, 14, 8),
            Foreground = (Brush)FindResource("TextBrush"),
            Background = (Brush)FindResource("BgBrush"),
            BorderBrush = (Brush)FindResource("AccentBrush"),
            BorderThickness = new Thickness(1),
            Cursor = Cursors.Hand,
            FontSize = 12,
            FontWeight = FontWeights.SemiBold,
            Style = (Style)FindResource("PillButtonStyle")
        };

        private async Task ToggleAdapterAsync(NicInfo nic, bool enable)
        {
            try
            {
                ShowBusy(true);
                await Task.Delay(1);

                await Task.Run(() =>
                {
                    if (enable) _nicManager.EnablePublic(nic.Id);
                    else _nicManager.DisablePublic(nic.Id);
                });

                await RefreshUiAsync();

                if (!enable && _nicManager.CurrentIsolated?.Id == nic.Id)
                {
                    _isIsolationActive = false;
                    _nicManager.CurrentIsolated = null;
                    _previousStates.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    $"Failed to {(enable ? "enable" : "disable")} '{nic.Name}':\n\n{ex.Message}",
                    "NetIsolate+", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ShowBusy(false);
            }
        }

        // Bulk hardening helpers (local to UI; no changes to NicManager)
        private static bool IsTransientBulkException(Exception ex)
        {
            // We keep this intentionally broad: WMI often throws generic exceptions during state transitions.
            // Retrying once after a short delay handles most "device busy/state changing" scenarios.
            return ex is InvalidOperationException
                || ex is ManagementException
                || ex is System.Runtime.InteropServices.COMException;
        }

        private void BulkSetOne(NicInfo nic, bool enable)
        {
            if (enable) _nicManager.EnablePublic(nic.Id);
            else _nicManager.DisablePublic(nic.Id);
        }

        private void BulkSetOneWithRetry(NicInfo nic, bool enable)
        {
            try
            {
                BulkSetOne(nic, enable);
                return;
            }
            catch (Exception ex) when (IsTransientBulkException(ex))
            {
                // One retry after a short delay (adapter is probably transitioning)
                Thread.Sleep(350);
                BulkSetOne(nic, enable);
            }
        }

        private async Task BulkSetAdaptersAsync(bool enable)
        {
            // Soft cooldown: avoids immediate opposite bulk while Windows is still transitioning.
            if (DateTime.UtcNow < _bulkCooldownUntilUtc)
                return;

            // Prevent overlapping bulk operations (re-entrancy, fast clicks, etc.)
            if (!await _bulkOpLock.WaitAsync(0))
                return;

            try
            {
                // prevent accidental toggling while busy
                BulkToggle.IsEnabled = false;

                ShowBusy(true);
                _suspendAutoRefresh = true;

                await Task.Delay(1);

                var failures = new List<string>();

                await Task.Run(() =>
                {
                    var all = _nicManager.ListAdapters();

                    // Respect the current "Virtual adapters" view for bulk actions.
                    IEnumerable<NicInfo> target = all;
                    if (!_showVirtual)
                        target = target.Where(n => !IsVirtualAdapter(n));

                    var list = target.ToList();

                    foreach (var nic in list)
                    {
                        try
                        {
                            if (enable)
                            {
                                if (!nic.Enabled) BulkSetOneWithRetry(nic, enable: true);
                            }
                            else
                            {
                                if (nic.Enabled) BulkSetOneWithRetry(nic, enable: false);
                            }
                        }
                        catch (Exception ex)
                        {
                            failures.Add($"{nic.Name}: {ex.Message}");
                        }
                    }
                });

                // Bulk operations invalidate isolation state — clear it cleanly
                _isIsolationActive = false;
                _nicManager.CurrentIsolated = null;
                _previousStates.Clear();

                _refreshDebounce.Stop();
                await RefreshUiAsync(force: true);

                if (failures.Count > 0)
                {
                    int show = Math.Min(4, failures.Count);
                    var msg = $"Bulk toggle completed with {failures.Count} warning(s).\n\n"
                            + string.Join("\n", failures.Take(show));
                    if (failures.Count > show)
                        msg += $"\n… and {failures.Count - show} more.";

                    MessageBox.Show(this, msg, "NetIsolate+", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            finally
            {
                _bulkCooldownUntilUtc = DateTime.UtcNow.AddSeconds(1.5);

                _ignoreExternalEventsUntilUtc = DateTime.UtcNow.AddSeconds(2.5);
                _suspendAutoRefresh = false;
                ShowBusy(false);

                // Re-enable based on current state (and isolation state)
                try
                {
                    var adapters = _nicManager.ListAdapters();
                    UpdateBulkToggleState(adapters);
                }
                catch
                {
                    // best effort
                    BulkToggle.IsEnabled = !_isIsolationActive;
                }

                _bulkOpLock.Release();
            }
        }

        // NEW: Minimize support
        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void About_Click(object sender, RoutedEventArgs e)
        {
            var w = new AboutWindow { Owner = this, WindowStartupLocation = WindowStartupLocation.CenterOwner };
            w.ShowDialog();
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2) return;
            DragMove();
        }
    }
}
