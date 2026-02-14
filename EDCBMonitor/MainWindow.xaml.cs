using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Interop;
using System.Diagnostics;

// 衝突回避の別名
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;
using Point = System.Windows.Point;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using FontFamily = System.Windows.Media.FontFamily;

namespace EDCBMonitor
{
    public partial class MainWindow : Window
    {
        #region Constants & Win32 API
        private const int MAX_RETRY_COUNT = 3;
        private const int MOUSE_SNAP_DIST = 20;
        private const int WM_MOUSEHWHEEL = 0x020E;
        #endregion

        #region Fields
        private readonly DispatcherTimer _updateTimer;
        private FileSystemWatcher? _fileWatcher;
        private WinForms.NotifyIcon? _notifyIcon;
        private DispatcherTimer? _reloadDebounceTimer;
        
        private int _retryCount = 0;
        private bool _isShowingTempMessage = false;
        private bool _isDragging = false;
        private Point _startMousePoint;
        private Rect? _restoreBounds = null;

        private readonly ReservationService _reservationService;
        private readonly GridColumnManager _columnManager;
        
        // 監視状態を表示するためのフィールド
        private string _watcherStatusMessage = "";
        #endregion

        public MainWindow()
        {
            InitializeComponent();
            InitializeNotifyIcon();
            LoadAppIcon();
            
            _reservationService = new ReservationService();
            // ここでエラーになっていたメソッドを渡せるようにクラス内にメソッド定義が必要
            _columnManager = new GridColumnManager(LstReservations, ReservationCheckBox_Click);

            _updateTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _updateTimer.Tick += UpdateTimer_Tick;
            _updateTimer.Start();

            ApplySettings(true);

            Loaded += async (s, e) => 
            {
                EnsureWindowIsVisible();
                SetupFileWatcher();
                await UpdateReservations();
            };

            Closing += Window_Closing;
        }

        private void LoadAppIcon()
        {
            try
            {
                string exePath = Process.GetCurrentProcess().MainModule?.FileName ?? "";
                if (!string.IsNullOrEmpty(exePath) && _notifyIcon != null && File.Exists(exePath))
                {
                    _notifyIcon.Icon = Drawing.Icon.ExtractAssociatedIcon(exePath);
                }
            }
            catch
            {
                if (_notifyIcon != null) _notifyIcon.Icon = Drawing.SystemIcons.Application;
            }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            if (PresentationSource.FromVisual(this) is HwndSource source)
            {
                source.AddHook(WndProc);
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_MOUSEHWHEEL)
            {
                int tilt = (short)((wParam.ToInt64() >> 16) & 0xFFFF);
                var sv = GetScrollViewer(LstReservations);
                
                if (sv != null)
                {
                    // 設定された行数分スクロール
                    int lines = Config.Data.ScrollAmountHorizontal;
                    if (tilt > 0) for (int i = 0; i < lines; i++) sv.LineRight();
                    else if (tilt < 0) for (int i = 0; i < lines; i++) sv.LineLeft();
                    handled = true;
                }
            }
            return IntPtr.Zero;
        }
        
        private void Window_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            var sv = GetScrollViewer(LstReservations);
            if (sv != null)
            {
                int lines = Config.Data.ScrollAmountVertical;
                if (e.Delta > 0)
                {
                    for (int i = 0; i < lines; i++) sv.LineUp();
                }
                else
                {
                    for (int i = 0; i < lines; i++) sv.LineDown();
                }
                e.Handled = true;
            }
        }

        private async void UpdateTimer_Tick(object? sender, EventArgs e)
        {
            if (LstReservations.ItemsSource is not List<ReserveItem> list)
            {
                _retryCount++;
                if (_retryCount >= MAX_RETRY_COUNT)
                {
                    _retryCount = 0;
                    await UpdateReservations();
                }
                return;
            }
            foreach (var item in list) item.UpdateProgress(); 
        }

        private void InitializeNotifyIcon()
        {
            _notifyIcon = new WinForms.NotifyIcon
            {
                Text = "EDCB Monitor",
                Visible = true
            };

            var menu = new WinForms.ContextMenuStrip();
            menu.Items.Add("設定...", null, (s, e) => MenuSettings_Click(null, null));
            menu.Items.Add("再読み込み", null, (s, e) => { _ = UpdateReservations(); });
            menu.Items.Add("終了", null, (s, e) => {
                if (_notifyIcon != null) _notifyIcon.Visible = false;
                System.Windows.Application.Current.Shutdown(); 
            });
            _notifyIcon.ContextMenuStrip = menu;

            _notifyIcon.DoubleClick += (s, e) => 
            {
                if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
                Activate();
                Topmost = true;
                Topmost = Config.Data.Topmost;
            };
        }

        #region Window Interaction
        private void Window_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject obj && !IsInteractiveControl(obj))
            {
                LstReservations.UnselectAll();
                if (e.ButtonState == MouseButtonState.Pressed)
                {
                    _isDragging = true;
                    _startMousePoint = e.GetPosition(this);
                    CaptureMouse();
                }
            }
        }

        private void Window_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging) return;

            try
            {
                var cursorScreenPos = WinForms.Cursor.Position;
                var source = PresentationSource.FromVisual(this);
                
                double dpiX = source?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
                double dpiY = source?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;

                var currentScreen = WinForms.Screen.FromPoint(cursorScreenPos);
                var workArea = currentScreen.WorkingArea;

                double newLeft = (cursorScreenPos.X / dpiX) - _startMousePoint.X;
                double newTop = (cursorScreenPos.Y / dpiY) - _startMousePoint.Y;

                double screenLeft = workArea.Left / dpiX;
                double screenTop = workArea.Top / dpiY;
                double screenRight = workArea.Right / dpiX;
                double screenBottom = workArea.Bottom / dpiY;

                if (Math.Abs(newLeft - screenLeft) < MOUSE_SNAP_DIST) newLeft = screenLeft;
                else if (Math.Abs(newLeft + ActualWidth - screenRight) < MOUSE_SNAP_DIST) newLeft = screenRight - ActualWidth;

                if (Math.Abs(newTop - screenTop) < MOUSE_SNAP_DIST) newTop = screenTop;
                else if (Math.Abs(newTop + ActualHeight - screenBottom) < MOUSE_SNAP_DIST) newTop = screenBottom - ActualHeight;

                Left = newLeft;
                Top = newTop;
            }
            catch (Exception ex) { Logger.Write($"Mouse Move Error: {ex.Message}"); }
        }

        private void Window_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDragging)
            {
                _isDragging = false;
                ReleaseMouseCapture();
                SaveCurrentState();
            }
        }

        private void Window_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            DependencyObject? obj = e.OriginalSource as DependencyObject;
            while (obj != null)
            {
                if (obj is System.Windows.Controls.Primitives.ButtonBase || obj is System.Windows.Controls.Primitives.Thumb || obj is System.Windows.Controls.TextBox) return;
                if (obj is System.Windows.Controls.ListViewItem)
                {
                    ExternalAppHelper.ActivateOrLaunchEpgTimer();
                    e.Handled = true;
                    return;
                }
                obj = VisualTreeHelper.GetParent(obj);
            }
            ExternalAppHelper.ActivateOrLaunchEpgTimer();
        }

        private void listView_MouseDoubleClick(object sender, MouseButtonEventArgs e) => ExternalAppHelper.ActivateOrLaunchEpgTimer();

        private bool IsInteractiveControl(DependencyObject? obj)
        {
            while (obj != null)
            {
                if (obj is System.Windows.Controls.Primitives.ButtonBase || obj is System.Windows.Controls.Primitives.Thumb || obj is System.Windows.Controls.TextBox || obj is System.Windows.Controls.ListViewItem) return true;
                obj = VisualTreeHelper.GetParent(obj);
            }
            return false;
        }
        #endregion

        #region Window State Management
        private void EnsureWindowIsVisible()
        {
            try
            {
                bool isVisible = WinForms.Screen.AllScreens.Any(screen =>
                    Left > -10000 && Left < 30000 && 
                    Top > -10000 && Top < 30000 && 
                    Width > 10 && Height > 10);

                if (!isVisible)
                {
                    Left = 100;
                    Top = 100;
                    Width = Config.Data.Width; 
                    Height = Config.Data.Height;
                }
            }
            catch { }
        }

        private void SaveCurrentState()
        {
            if (_restoreBounds.HasValue)
            {
                Config.Data.IsVerticalMaximized = true;
                Config.Data.Top = Top;
                Config.Data.Height = Height;
                Config.Data.Left = Left;
                Config.Data.Width = Width;
                Config.Data.RestoreTop = _restoreBounds.Value.Top;
                Config.Data.RestoreHeight = _restoreBounds.Value.Height;
            }
            else if (WindowState == WindowState.Normal)
            {
                Config.Data.IsVerticalMaximized = false;
                Config.Data.Top = Top;
                Config.Data.Left = Left;
                Config.Data.Width = Width;
                Config.Data.Height = Height;
            }
            else
            {
                Config.Data.IsVerticalMaximized = false;
                Config.Data.Top = RestoreBounds.Top;
                Config.Data.Left = RestoreBounds.Left;
                Config.Data.Width = RestoreBounds.Width;
                Config.Data.Height = RestoreBounds.Height;
            }

            _columnManager.SaveColumnState();
        }

        private void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }
            _fileWatcher?.Dispose();
            _fileWatcher = null;
            SaveCurrentState();
            Config.Save();
        }
        #endregion

        #region Apply Settings
        public void ApplySettings(bool updateSize = false)
        {
            try
            {
                if (updateSize)
                {
                    if (Config.Data.Width > 0) Width = Config.Data.Width;
                    if (Config.Data.Height > 0) Height = Config.Data.Height;
                    Top = Config.Data.Top;
                    Left = Config.Data.Left;
                }
                
                if (Config.Data.IsVerticalMaximized)
                {
                    _restoreBounds = new Rect(Left, Config.Data.RestoreTop, Width, Config.Data.RestoreHeight);
                }

                Topmost = Config.Data.Topmost;

                try { FontFamily = new FontFamily(Config.Data.FontFamily); } catch { }
                FontSize = Config.Data.FontSize;
                
                if (LblStatus != null) LblStatus.FontSize = Config.Data.FooterFontSize;

                Resources["HeaderFontSize"] = Config.Data.HeaderFontSize;
                
                var brushConverter = new System.Windows.Media.BrushConverter();
                var bgBrush = (SolidColorBrush?)brushConverter.ConvertFromString(Config.Data.BackgroundColor) ?? System.Windows.Media.Brushes.Black;
                bgBrush.Opacity = Config.Data.Opacity;
                MainBorder.Background = bgBrush;

                var fgBrush = (SolidColorBrush?)brushConverter.ConvertFromString(Config.Data.ForegroundColor) ?? System.Windows.Media.Brushes.White;
                LstReservations.Foreground = fgBrush;

                if (brushConverter.ConvertFromString(Config.Data.ScrollBarColor) is SolidColorBrush sb)
                {
                    Resources["ScrollBarBrush"] = sb;
                }
                
                var colBorderBrush = (SolidColorBrush?)brushConverter.ConvertFromString(Config.Data.ColumnBorderColor) ?? new SolidColorBrush(System.Windows.Media.Color.FromRgb(68, 68, 68));
                Resources["ColumnBorderBrush"] = colBorderBrush;
                if (MainBorder.Resources.Contains("ColumnBorderBrush"))
                {
                    MainBorder.Resources["ColumnBorderBrush"] = colBorderBrush;
                }

                if (brushConverter.ConvertFromString(Config.Data.FooterColor) is SolidColorBrush fBrush)
                {
                    if (LblStatus != null) LblStatus.Foreground = fBrush;
                }

                HeaderTitle.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(136, 204, 255));
                
                try
                {
                    var mainBorderBrush = (SolidColorBrush?)brushConverter.ConvertFromString(Config.Data.MainBorderColor) 
                                          ?? new SolidColorBrush(System.Windows.Media.Color.FromRgb(85, 85, 85));
                    mainBorderBrush.Opacity = Config.Data.Opacity;
                    MainBorder.BorderBrush = mainBorderBrush;
                }
                catch
                {
                    var defBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(85, 85, 85)) { Opacity = Config.Data.Opacity };
                    MainBorder.BorderBrush = defBrush;
                }
                
                var toolTipStyle = new Style(typeof(System.Windows.Controls.ToolTip));
                toolTipStyle.Setters.Add(new Setter(FrameworkElement.MaxWidthProperty, Config.Data.ToolTipWidth));
                try
                {
                    toolTipStyle.Setters.Add(new Setter(System.Windows.Controls.Control.FontSizeProperty, Config.Data.ToolTipFontSize));
                    if (brushConverter.ConvertFromString(Config.Data.ToolTipBackColor) is SolidColorBrush ttBg)
                        toolTipStyle.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, ttBg));
                    if (brushConverter.ConvertFromString(Config.Data.ToolTipForeColor) is SolidColorBrush ttFg)
                        toolTipStyle.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, ttFg));
                    if (brushConverter.ConvertFromString(Config.Data.ToolTipBorderColor) is SolidColorBrush ttBorder)
                        toolTipStyle.Setters.Add(new Setter(System.Windows.Controls.Control.BorderBrushProperty, ttBorder));
                }
                catch 
                {
                    toolTipStyle.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 30, 30))));
                    toolTipStyle.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, System.Windows.Media.Brushes.White));
                    toolTipStyle.Setters.Add(new Setter(System.Windows.Controls.Control.BorderBrushProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100))));
                }
                var ttTemplate = new DataTemplate();
                var ttText = new FrameworkElementFactory(typeof(TextBlock));
                ttText.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding());
                ttText.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
                ttTemplate.VisualTree = ttText;
                toolTipStyle.Setters.Add(new Setter(System.Windows.Controls.ContentControl.ContentTemplateProperty, ttTemplate));
                Resources[typeof(System.Windows.Controls.ToolTip)] = toolTipStyle;

                var itemStyle = new Style(typeof(System.Windows.Controls.ListViewItem));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, System.Windows.Media.Brushes.Transparent));
                if (Config.Data.ShowToolTip) itemStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, new System.Windows.Data.Binding("ToolTipText")));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.BorderThicknessProperty, new Thickness(0)));
                itemStyle.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(-1, 0, 0, 0)));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.PaddingProperty, new Thickness(0)));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.MinHeightProperty, 0.0));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.HorizontalContentAlignmentProperty, System.Windows.HorizontalAlignment.Stretch));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                itemStyle.Setters.Add(new EventSetter(System.Windows.Controls.Control.MouseDoubleClickEvent, new MouseButtonEventHandler(listView_MouseDoubleClick)));

                try
                {
                    if (brushConverter.ConvertFromString(Config.Data.RecColor) is SolidColorBrush recBrush)
                    {
                        MainBorder.Resources["RecBrush"] = recBrush;
                        var recWeight = Config.Data.RecBold ? FontWeights.Bold : FontWeights.Normal;
                        MainBorder.Resources["RecWeight"] = recWeight;
                        var recTrigger = new DataTrigger { Binding = new System.Windows.Data.Binding("IsRecording"), Value = true };
                        recTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, recBrush));
                        recTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.FontWeightProperty, recWeight));
                        itemStyle.Triggers.Add(recTrigger);
                    }
                    if (brushConverter.ConvertFromString(Config.Data.DisabledColor) is SolidColorBrush disabledBrush)
                    {
                        MainBorder.Resources["DisabledBrush"] = disabledBrush;
                        var disabledTrigger = new DataTrigger { Binding = new System.Windows.Data.Binding("IsDisabled"), Value = true };
                        disabledTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, disabledBrush));
                        itemStyle.Triggers.Add(disabledTrigger);
                    }
                    if (brushConverter.ConvertFromString(Config.Data.ReserveErrorColor) is SolidColorBrush errBrush)
                    {
                        MainBorder.Resources["ErrorBrush"] = errBrush;
                        var errorTrigger = new DataTrigger { Binding = new System.Windows.Data.Binding("HasError"), Value = true };
                        errorTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, errBrush));
                        itemStyle.Triggers.Add(errorTrigger);
                    }
                }
                catch { }

                double brightness = bgBrush.Color.R * 0.299 + bgBrush.Color.G * 0.587 + bgBrush.Color.B * 0.114;
                bool isLight = brightness > 128; 
                var selectedTrigger = new Trigger { Property = System.Windows.Controls.ListViewItem.IsSelectedProperty, Value = true };
                var selColor = System.Windows.Media.Color.FromArgb(100, 0, 100, 200);
                if (!isLight) selColor = System.Windows.Media.Color.FromArgb(80, 255, 255, 255);
                selectedTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(selColor)));
                itemStyle.Triggers.Add(selectedTrigger);

                var mouseOverTrigger = new Trigger { Property = System.Windows.UIElement.IsMouseOverProperty, Value = true };
                var hoverColor = isLight ? System.Windows.Media.Color.FromArgb(50, 0, 100, 200) : System.Windows.Media.Color.FromArgb(50, 255, 255, 255);
                mouseOverTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(hoverColor)));
                itemStyle.Triggers.Add(mouseOverTrigger);

                LstReservations.ItemContainerStyle = itemStyle;
                LstReservations.Margin = new Thickness(Config.Data.ListMarginLeft, Config.Data.ListMarginTop, Config.Data.ListMarginRight, Config.Data.ListMarginBottom);

                RowHeader.Height = Config.Data.ShowHeader ? GridLength.Auto : new GridLength(0);
                RowFooter.Height = Config.Data.ShowFooter ? GridLength.Auto : new GridLength(0);
                MainBorder.BorderThickness = new Thickness(1);

                _columnManager.UpdateColumns();
                _columnManager.UpdateHeaderStyle(bgBrush, fgBrush, colBorderBrush);
                LstReservations.Items.Refresh();
                ApplyVisualSettings();
            }
            catch (Exception ex) { Logger.Write($"ApplySettings Error: {ex.Message}"); }
        }
        
        private void ApplyVisualSettings()
        {
            try
            {
                if (FindName("BtnVerticalMax") is System.Windows.Controls.Button btn)
                {
                    string colorCode = Config.Data.FooterBtnColor ?? "#555555";
                    var brush = new System.Windows.Media.BrushConverter().ConvertFromString(colorCode) as System.Windows.Media.Brush;
                    btn.Background = brush;
                }
            }
            catch
            {
                if (FindName("BtnVerticalMax") is System.Windows.Controls.Button btn) btn.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(68, 68, 68));
            }
        }
        #endregion

        #region Actions & Events
        private async void ReservationCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox cb && cb.DataContext is ReserveItem item)
            {
                cb.IsEnabled = false;
                bool isChecked = cb.IsChecked == true;
                if (await _reservationService.ToggleReservationStatusAsync(new List<uint> { item.ID }))
                {
                    await UpdateReservations();
                }
                else 
                {
                    cb.IsChecked = !isChecked;
                    cb.IsEnabled = true;
                }
            }
        }

        private void BtnVerticalMaximize_Click(object sender, RoutedEventArgs e)
        {
            if (_restoreBounds.HasValue)
            {
                Top = _restoreBounds.Value.Top;
                Height = _restoreBounds.Value.Height;
                _restoreBounds = null; 
            }
            else
            {
                _restoreBounds = new Rect(Left, Top, Width, Height);
                try
                {
                    var centerPoint = new Drawing.Point((int)(Left + Width / 2), (int)(Top + Height / 2));
                    var screen = WinForms.Screen.FromPoint(centerPoint);
                    var dpiY = PresentationSource.FromVisual(this)?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;

                    Top = screen.WorkingArea.Top / dpiY;
                    Height = screen.WorkingArea.Height / dpiY;
                }
                catch
                {
                    Top = SystemParameters.WorkArea.Top;
                    Height = SystemParameters.WorkArea.Height;
                }
            }
        }
        #endregion

        #region File Watcher (Robust Implementation)
        private void SetupFileWatcher()
        {
            _fileWatcher?.Dispose();
            _fileWatcher = null;
            _watcherStatusMessage = "";

            if (_reloadDebounceTimer == null)
            {
                // リロード時の待機時間を設定（EDCB側の書き込み完了待ち）
                _reloadDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                _reloadDebounceTimer.Tick += async (s, e) => 
                {
                    if (_reloadDebounceTimer != null) _reloadDebounceTimer.Stop();
                    // ログ抑制: Logger.Write("File change detected. Updating...");
                    await UpdateReservations();
                };
            }

            // 1. パスの決定ロジック（ユーザー入力を柔軟に解釈）
            string reservePath = "";
            string configPath = Config.Data.EdcbInstallPath;

            if (!string.IsNullOrEmpty(configPath))
            {
                if (File.Exists(configPath) && Path.GetFileName(configPath).Equals("Reserve.txt", StringComparison.OrdinalIgnoreCase))
                {
                    reservePath = configPath;
                }
                else if (Directory.Exists(configPath) && File.Exists(Path.Combine(configPath, "Reserve.txt")))
                {
                    reservePath = Path.Combine(configPath, "Reserve.txt");
                }
                else if (Directory.Exists(configPath) && File.Exists(Path.Combine(configPath, "Setting", "Reserve.txt")))
                {
                    reservePath = Path.Combine(configPath, "Setting", "Reserve.txt");
                }
            }

            if (string.IsNullOrEmpty(reservePath))
            {
                 string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                 if (File.Exists(Path.Combine(baseDir, "Setting", "Reserve.txt"))) 
                     reservePath = Path.Combine(baseDir, "Setting", "Reserve.txt");
                 else 
                 {
                     string? parent = Directory.GetParent(baseDir)?.FullName;
                     if (parent != null && File.Exists(Path.Combine(parent, "Setting", "Reserve.txt"))) 
                         reservePath = Path.Combine(parent, "Setting", "Reserve.txt");
                 }
            }

            if (!string.IsNullOrEmpty(reservePath) && File.Exists(reservePath))
            {
                try
                {
                    string dir = Path.GetDirectoryName(reservePath)!;
                    string fileName = Path.GetFileName(reservePath);
                    
                    _fileWatcher = new FileSystemWatcher(dir, fileName)
                    {
                        NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName | NotifyFilters.CreationTime,
                        InternalBufferSize = 65536
                    };

                    _fileWatcher.Changed += (s, e) => HandleFileChange();
                    _fileWatcher.Created += (s, e) => HandleFileChange();
                    _fileWatcher.Renamed += (s, e) => HandleFileChange();
                    _fileWatcher.EnableRaisingEvents = true;
                    
                    _watcherStatusMessage = ""; 
                }
                catch (Exception ex)
                {
                    Logger.Write($"FileWatcher Start Error: {ex.Message}");
                    _watcherStatusMessage = "監視エラー";
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(configPath))
                {
                    Logger.Write($"Reserve.txt not found at configured path: {configPath}");
                    _watcherStatusMessage = "Reserve.txt 不明";
                }
            }
            
            UpdateFooterStatus();
        }

        private void HandleFileChange()
        {
            Dispatcher.Invoke(() => 
            {
                _reloadDebounceTimer?.Stop();
                _reloadDebounceTimer?.Start();
            });
        }
        #endregion

        #region Data Update
        public async Task RefreshDataAsync() => await UpdateReservations();

        private async Task<bool> UpdateReservations(bool updateFooter = true)
        {
            try
            {
                PresetManager.Instance.Load();
                var selectedIds = new HashSet<uint>();
                if (LstReservations.SelectedItems != null)
                {
                    foreach (var item in LstReservations.SelectedItems.OfType<ReserveItem>()) selectedIds.Add(item.ID);
                }

                var sv = GetScrollViewer(LstReservations);
                double vOffset = sv?.VerticalOffset ?? 0;

                var list = await _reservationService.GetReservationsAsync();
                
                if (list == null)
                {
                    if (updateFooter && !_isShowingTempMessage) LblStatus.Text = "接続待機中...";
                    return false;
                }

                if (Config.Data.EnableTitleRemove && !string.IsNullOrEmpty(Config.Data.TitleRemovePattern))
                {
                    try
                    {
                        var regex = new System.Text.RegularExpressions.Regex(Config.Data.TitleRemovePattern);
                        foreach (var item in list) item.Data.Title = regex.Replace(item.Data.Title, "");
                    }
                    catch { }
                }

                if (Config.Data.HideDisabled) list = list.Where(x => !x.IsDisabled).ToList();

                LstReservations.ItemsSource = list;
                
                if (selectedIds.Count > 0)
                {
                    foreach (var item in list.Where(x => selectedIds.Contains(x.ID))) LstReservations?.SelectedItems.Add(item);
                }
                if (sv != null) sv.ScrollToVerticalOffset(vOffset);

                if (updateFooter) UpdateFooterStatus();
                return true;
            }
            catch (Exception ex)
            {
                Logger.Write($"Update Error: {ex}");
                if (updateFooter && !_isShowingTempMessage && LblStatus != null) LblStatus.Text = "エラー発生";
                return false;
            }
        }

        private void UpdateFooterStatus()
        {
            if (_isShowingTempMessage || LblStatus == null) return;
            
            string status = "";
            if (LstReservations.ItemsSource is List<ReserveItem> list)
            {
                int validCount = list.Count(x => !x.IsDisabled);
                status = $"更新: {DateTime.Now:HH:mm} ({validCount}件)";
            }
            
            if (!string.IsNullOrEmpty(_watcherStatusMessage))
            {
                status += $" [{_watcherStatusMessage}]";
            }

            LblStatus.Text = status;
        }

        private async Task ShowTemporaryMessage(string message)
        {
            _isShowingTempMessage = true; 
            if (LblStatus != null) LblStatus.Text = message;
            try { await Task.Delay(2000); }
            finally
            {
                _isShowingTempMessage = false; 
                UpdateFooterStatus(); 
            }
        }
        #endregion

        #region Menu Handlers
        private void MenuReload_Click(object sender, RoutedEventArgs e) => _ = UpdateReservations();
        private void MenuExit_Click(object sender, RoutedEventArgs e) => Close();
        
        private void MenuSettings_Click(object? sender, RoutedEventArgs? e)
        {
            SaveCurrentState();
            var win = new SettingsWindow { Owner = this };
            if (win.ShowDialog() == true)
            {
                ApplySettings(true);
                SetupFileWatcher();
                _ = UpdateReservations();
            }
        }

        private void ContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.ContextMenu menu)
            {
                var selectedItem = LstReservations.SelectedItem as ReserveItem;

                if (MenuItemPlay != null && SepPlay != null)
                {
                    bool showPlay = selectedItem != null && selectedItem.IsRecording && !string.IsNullOrEmpty(Config.Data.TvTestPath);
                    MenuItemPlay.Visibility = showPlay ? Visibility.Visible : Visibility.Collapsed;
                    SepPlay.Visibility = showPlay ? Visibility.Visible : Visibility.Collapsed;
                }

                if (MenuItemOpenFolder != null)
                {
                    MenuItemOpenFolder.Items.Clear();
                    if (selectedItem != null && string.IsNullOrEmpty(selectedItem.RecFolder))
                    {
                        // Common.ini から共通録画フォルダを取得してサブメニューに表示
                        var commonFolders = _reservationService.GetCommonRecFolders();
                        if (commonFolders.Count > 0)
                        {
                            foreach (var folder in commonFolders)
                            {
                                var subItem = new System.Windows.Controls.MenuItem { Header = folder };
                                subItem.Click += (s, args) => ExternalAppHelper.OpenFolder(folder);
                                MenuItemOpenFolder.Items.Add(subItem);
                            }
                        }
                    }
                }
                
                foreach (var item in menu.Items.OfType<MenuItem>())
                {
                    if (item.Name == "MenuItemHideDisabled") item.Header = Config.Data.HideDisabled ? "無効予約を表示する" : "無効予約を表示しない";
                    else if (item.Name == "MenuItemVerticalMax") item.Header = _restoreBounds.HasValue ? "上下最大化を復元する" : "上下最大化にする";
                }
            }
        }

        private void MenuPlay_Click(object sender, RoutedEventArgs e)
        {
            if (LstReservations.SelectedItem is not ReserveItem item) return;
            string? recPath = _reservationService.OpenTimeShift(item.ID, out uint ctrlId);

            if (recPath == null)
            {
                Logger.Write("SendNwTimeShiftOpen failed. Fallback to path guessing.");
                if (!string.IsNullOrEmpty(item.RecFolder) && !string.IsNullOrEmpty(item.RecFileName))
                {
                    try { recPath = Path.Combine(item.RecFolder, item.RecFileName); } catch { }
                }
            }
            else
            {
                _reservationService.CloseNwPlay(ctrlId);
            }

            ExternalAppHelper.OpenTvTest(recPath ?? "");
        }

        private void MenuHideDisabled_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi)
            {
                Config.Data.HideDisabled = !Config.Data.HideDisabled;
                Config.Save();
                mi.Header = Config.Data.HideDisabled ? "無効予約を表示する" : "無効予約を表示しない";
                _ = RefreshDataAsync();
            }
        }

        private async void MenuChgOnOff_Click(object sender, RoutedEventArgs e)
        {
            var targetIDs = LstReservations.SelectedItems.OfType<ReserveItem>().Select(x => x.ID).ToList();
            if (targetIDs.Count == 0) return;

            _isShowingTempMessage = true;
            bool success = await _reservationService.ToggleReservationStatusAsync(targetIDs);

            if (success)
            {
                await UpdateReservations(false);
                await ShowTemporaryMessage("予約状態を変更しました");
            }
            else
            {
                System.Windows.MessageBox.Show("変更に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                await ShowTemporaryMessage("変更に失敗しました");
            }
        }
        
        private async void MenuDel_Click(object sender, RoutedEventArgs e)
        {
            var targetIDs = LstReservations.SelectedItems.OfType<ReserveItem>().Select(x => x.ID).ToList();
            if (targetIDs.Count == 0) return;

            if (System.Windows.MessageBox.Show($"{targetIDs.Count}件の予約を削除しますか？", "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;

            _isShowingTempMessage = true;
            bool success = await _reservationService.DeleteReservationsAsync(targetIDs);

            if (success)
            {
                await UpdateReservations(false);
                await ShowTemporaryMessage("予約を削除しました");
            }
            else
            {
                System.Windows.MessageBox.Show("削除に失敗しました。\r\n詳細はログを確認してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                await ShowTemporaryMessage("削除に失敗しました");
            }
        }
        
        private void MenuOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            if (LstReservations.SelectedItem is ReserveItem item)
            {
                // フォルダ指定がある場合は直接開く
                // 指定がない(デフォルト)場合は、ContextMenu_Opened で追加されたサブメニューから選択する運用になる
                if (!string.IsNullOrEmpty(item.RecFolder))
                {
                    ExternalAppHelper.OpenFolder(item.RecFolder);
                }
            }
        }
        #endregion

        #region Helpers
        private ScrollViewer? GetScrollViewer(DependencyObject o)
        {
            if (o is ScrollViewer sv) return sv;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(o); i++)
            {
                var child = VisualTreeHelper.GetChild(o, i);
                if (child != null)
                {
                    if (GetScrollViewer(child) is ScrollViewer result) return result;
                }
            }
            return null;
        }
        #endregion
    }
}