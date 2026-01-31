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
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Interop;
using System.Diagnostics;
using System.Runtime.InteropServices;
using WinForms = System.Windows.Forms;
using EpgTimer;

// Windows Formsとの名前衝突を避けるためのエイリアス定義
using Application = System.Windows.Application;
using Binding = System.Windows.Data.Binding;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using Control = System.Windows.Controls.Control;
using ListViewItem = System.Windows.Controls.ListViewItem;
using TextBox = System.Windows.Controls.TextBox;
using FontFamily = System.Windows.Media.FontFamily;
using Brushes = System.Windows.Media.Brushes;
using Point = System.Windows.Point;
using Brush = System.Windows.Media.Brush;
using Color = System.Windows.Media.Color;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;

namespace EDCBMonitor
{
    public partial class MainWindow : Window
    {
        private class ColumnDef
        {
            public string Header { get; set; } = "";
            public string BindingPath { get; set; } = "";
            public Func<bool> GetShow { get; set; } = () => true;
            public Func<double> GetWidth { get; set; } = () => 100;
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private const int SW_RESTORE = 9;

        private DispatcherTimer _updateTimer;
        private FileSystemWatcher _fileWatcher;
        private WinForms.NotifyIcon _notifyIcon;
        private DateTime _lastReloadTime = DateTime.MinValue;

        // スナップ機能用の変数
        private bool _isDragging = false;
        private Point _startMousePoint; 

        public MainWindow()
        {
            InitializeComponent();
            InitializeNotifyIcon();

            try
            {
                string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? "";
                if (!string.IsNullOrEmpty(exePath) && _notifyIcon != null)
                {
                    _notifyIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
                }
            }
            catch 
            {
                if (_notifyIcon != null) _notifyIcon.Icon = System.Drawing.SystemIcons.Application;
            }

            _updateTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
            _updateTimer.Tick += async (s, e) => 
            {
                bool success = await UpdateReservations();
                _updateTimer.Interval = TimeSpan.FromSeconds(success ? 60 : 5);
            };
            _updateTimer.Start();

            ApplySettings();

            this.Loaded += async (s, e) => 
            {
                EnsureWindowIsVisible();
                await UpdateReservations();
            };

            this.Closing += Window_Closing;

            SetupFileWatcher();
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var source = PresentationSource.FromVisual(this) as HwndSource;
            source?.AddHook(WndProc);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_MOUSEHWHEEL = 0x020E;

            if (msg == WM_MOUSEHWHEEL)
            {
                int tilt = (short)((wParam.ToInt64() >> 16) & 0xFFFF);
                
                var sv = GetScrollViewer(LstReservations);
                if (sv != null)
                {
                    if (tilt > 0)
                    {
                        for(int i=0; i<15; i++) sv.LineRight();
                    }
                    else if (tilt < 0)
                    {
                        for(int i=0; i<15; i++) sv.LineLeft();
                    }
                    handled = true;
                }
            }

            return IntPtr.Zero;
        }

        private ScrollViewer GetScrollViewer(DependencyObject o)
        {
            if (o is ScrollViewer sv) return sv;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(o); i++)
            {
                var child = VisualTreeHelper.GetChild(o, i);
                var result = GetScrollViewer(child);
                if (result != null) return result;
            }
            return null;
        }

        private void InitializeNotifyIcon()
        {
            _notifyIcon = new WinForms.NotifyIcon();
            _notifyIcon.Text = "EDCB Monitor";
            _notifyIcon.Visible = true;

            var menu = new WinForms.ContextMenuStrip();
            menu.Items.Add("設定...", null, (s, e) => MenuSettings_Click(null, null));
            menu.Items.Add("再読み込み", null, (s, e) => { var t = UpdateReservations(); });
            menu.Items.Add("終了", null, (s, e) => {
                if (_notifyIcon != null) _notifyIcon.Visible = false;
                Application.Current.Shutdown(); 
            });
            _notifyIcon.ContextMenuStrip = menu;

            _notifyIcon.DoubleClick += (s, e) => 
            {
                if (this.WindowState == WindowState.Minimized)
                    this.WindowState = WindowState.Normal;
                this.Activate();
                this.Topmost = true;
                this.Topmost = Config.Data.Topmost;
            };
        }

        // --- スナップ機能 ---
        private void Window_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!IsInteractiveControl(e.OriginalSource as DependencyObject))
            {
                LstReservations.UnselectAll();
            }
            if (IsInteractiveControl(e.OriginalSource as DependencyObject)) return;
            
            try
            {
                if (e.ButtonState == MouseButtonState.Pressed)
                {
                    _isDragging = true;
                    _startMousePoint = e.GetPosition(this);
                    this.CaptureMouse();
                }
            }
            catch { }
        }

        private void Window_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging) return;

            try
            {
                System.Drawing.Point cursorScreenPos = WinForms.Cursor.Position;

                var source = PresentationSource.FromVisual(this);
                double dpiX = 1.0, dpiY = 1.0;
                if (source != null && source.CompositionTarget != null)
                {
                    dpiX = source.CompositionTarget.TransformToDevice.M11;
                    dpiY = source.CompositionTarget.TransformToDevice.M22;
                }

                WinForms.Screen currentScreen = WinForms.Screen.FromPoint(cursorScreenPos);
                System.Drawing.Rectangle workArea = currentScreen.WorkingArea;

                double newLeft = (cursorScreenPos.X / dpiX) - _startMousePoint.X;
                double newTop = (cursorScreenPos.Y / dpiY) - _startMousePoint.Y;

                double snapDist = 20.0;

                double screenLeft = workArea.Left / dpiX;
                double screenTop = workArea.Top / dpiY;
                double screenRight = workArea.Right / dpiX;
                double screenBottom = workArea.Bottom / dpiY;

                if (Math.Abs(newLeft - screenLeft) < snapDist)
                    newLeft = screenLeft;
                else if (Math.Abs(newLeft + this.ActualWidth - screenRight) < snapDist)
                    newLeft = screenRight - this.ActualWidth;

                if (Math.Abs(newTop - screenTop) < snapDist)
                    newTop = screenTop;
                else if (Math.Abs(newTop + this.ActualHeight - screenBottom) < snapDist)
                    newTop = screenBottom - this.ActualHeight;

                this.Left = newLeft;
                this.Top = newTop;
            }
            catch { }
        }

        private void Window_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDragging)
            {
                _isDragging = false;
                this.ReleaseMouseCapture();
                SaveCurrentState();
            }
        }

        private void Window_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            DependencyObject obj = e.OriginalSource as DependencyObject;
            while (obj != null)
            {
                if (obj is ButtonBase || obj is Thumb || obj is TextBox) return;
                if (obj is ListViewItem)
                {
                    ActivateOrLaunchEpgTimer();
                    e.Handled = true;
                    return;
                }
                obj = VisualTreeHelper.GetParent(obj);
            }
            ActivateOrLaunchEpgTimer();
        }
        
        private void listView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ActivateOrLaunchEpgTimer();
        }
        
        private void ActivateOrLaunchEpgTimer()
        {
            try
            {
                var proc = Process.GetProcessesByName("EpgTimer").FirstOrDefault();
                if (proc != null)
                {
                    IntPtr hwnd = proc.MainWindowHandle;
                    if (hwnd == IntPtr.Zero)
                    {
                        hwnd = FindWindow(null, "EpgTimer");
                    }

                    if (hwnd != IntPtr.Zero)
                    {
                        ShowWindow(hwnd, SW_RESTORE);
                        SetForegroundWindow(hwnd);
                    }
                    return;
                }

                string exePath = GetEpgTimerExePath();
                if (!string.IsNullOrEmpty(exePath) && File.Exists(exePath))
                {
                    Process.Start(new ProcessStartInfo(exePath)
                    {
                        UseShellExecute = true,
                        WorkingDirectory = Path.GetDirectoryName(exePath)
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.Write("EpgTimer Launch Error: " + ex.Message);
            }
        }

        private string GetEpgTimerExePath()
        {
            string txtPath = ReserveTextReader.GetReserveFilePath();
            if (string.IsNullOrEmpty(txtPath)) return "";

            string dir = Path.GetDirectoryName(txtPath) ?? "";
            
            string parentDir = Path.GetDirectoryName(dir) ?? "";
            string pathA = Path.Combine(parentDir, "EpgTimer.exe");
            if (File.Exists(pathA)) return pathA;

            string pathB = Path.Combine(dir, "EpgTimer.exe");
            if (File.Exists(pathB)) return pathB;

            return "";
        }

        private bool IsInteractiveControl(DependencyObject obj)
        {
            while (obj != null)
            {
                if (obj is ButtonBase ||
                    obj is Thumb ||
                    obj is TextBox ||
                    obj is ListViewItem) 
                {
                    return true;
                }
                obj = VisualTreeHelper.GetParent(obj);
            }
            return false;
        }
        
        private void EnsureWindowIsVisible()
        {
            try
            {
                bool isVisible = false;
                foreach (var screen in WinForms.Screen.AllScreens)
                {
                    if (this.Left > -10000 && this.Left < 30000 && 
                        this.Top > -10000 && this.Top < 30000 && 
                        this.Width > 10 && this.Height > 10)
                    {
                        isVisible = true;
                        break;
                    }
                }

                if (!isVisible)
                {
                    this.Left = 100;
                    this.Top = 100;
                    this.Width = Config.Data.Width; 
                    this.Height = Config.Data.Height;
                }
            }
            catch { }
        }

        private void SaveCurrentState()
        {
            if (this.WindowState == WindowState.Normal)
            {
                Config.Data.Top = this.Top;
                Config.Data.Left = this.Left;
                Config.Data.Width = this.Width;
                Config.Data.Height = this.Height;
            }
            else
            {
                Config.Data.Top = this.RestoreBounds.Top;
                Config.Data.Left = this.RestoreBounds.Left;
                Config.Data.Width = this.RestoreBounds.Width;
                Config.Data.Height = this.RestoreBounds.Height;
            }

            var gv = (GridView)LstReservations.View;
            if (gv != null)
            {
                Config.Data.ColumnHeaderOrder.Clear();
                foreach (var col in gv.Columns)
                {
                    if (col.Header is string headerText)
                    {
                        Config.Data.ColumnHeaderOrder.Add(headerText);
                    }
                }

                foreach (var col in gv.Columns)
                {
                    string header = col.Header as string;
                    if (header == null) continue;

                    double w = col.ActualWidth;

                    if (header == "ID") Config.Data.WidthColID = w;
                    else if (header == "状態") Config.Data.WidthColStatus = w;
                    else if (header == "日時") Config.Data.WidthColDateTime = w;
                    else if (header == "長さ") Config.Data.WidthColDuration = w;
                    else if (header == "ネットワーク") Config.Data.WidthColNetwork = w;
                    else if (header == "サービス名") Config.Data.WidthColServiceName = w;
                    else if (header == "番組名") Config.Data.WidthColTitle = w;
                    else if (header == "番組内容") Config.Data.WidthColDesc = w;
                    else if (header == "ジャンル") Config.Data.WidthColGenre = w;
                    else if (header == "付属情報") Config.Data.WidthColExtraInfo = w;
                    else if (header == "有効") Config.Data.WidthColEnabled = w;
                    else if (header == "プログラム予約") Config.Data.WidthColProgramType = w;
                    else if (header == "予約状況") Config.Data.WidthColComment = w;
                    else if (header == "エラー状況") Config.Data.WidthColError = w;
                    else if (header == "予定ファイル名") Config.Data.WidthColRecFileName = w;
                    else if (header == "予定ファイル名リスト") Config.Data.WidthColRecFileNameList = w;
                    else if (header == "使用予定チューナー") Config.Data.WidthColTuner = w;
                    else if (header == "予想サイズ") Config.Data.WidthColEstSize = w;
                    else if (header == "プリセット") Config.Data.WidthColPreset = w;
                    else if (header == "録画モード") Config.Data.WidthColRecMode = w;
                    else if (header == "優先度") Config.Data.WidthColPriority = w;
                    else if (header == "追従") Config.Data.WidthColTuijyuu = w;
                    else if (header == "ぴったり") Config.Data.WidthColPittari = w;
                    else if (header == "チューナー強制") Config.Data.WidthColTunerForce = w;
                    else if (header == "録画後動作") Config.Data.WidthColRecEndMode = w;
                    else if (header == "復帰後再起動") Config.Data.WidthColReboot = w;
                    else if (header == "録画後実行bat") Config.Data.WidthColBat = w;
                    else if (header == "録画タグ") Config.Data.WidthColRecTag = w;
                    else if (header == "録画フォルダ") Config.Data.WidthColRecFolder = w;
                    else if (header == "開始") Config.Data.WidthColStartMargin = w;
                    else if (header == "終了") Config.Data.WidthColEndMargin = w;
                }
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }

            if (_fileWatcher != null)
            {
                _fileWatcher.Dispose();
                _fileWatcher = null;
            }

            SaveCurrentState();
            Config.Save();
        }

        public void ApplySettings()
        {
            try
            {
                if (Config.Data.Width > 0) this.Width = Config.Data.Width;
                if (Config.Data.Height > 0) this.Height = Config.Data.Height;
                
                this.Top = Config.Data.Top;
                this.Left = Config.Data.Left;
                
                this.Topmost = Config.Data.Topmost;

                try { this.FontFamily = new FontFamily(Config.Data.FontFamily); } catch { }
                this.FontSize = Config.Data.FontSize;
                
                if (LblStatus != null)
                {
                    LblStatus.FontSize = Config.Data.FooterFontSize;
                }

                this.Resources["HeaderFontSize"] = Config.Data.HeaderFontSize;
                
                var brushConverter = new BrushConverter();
                var bgBrush = new SolidColorBrush(Colors.Black);
                try
                {
                    object converted = brushConverter.ConvertFromString(Config.Data.BackgroundColor);
                    if (converted is SolidColorBrush b)
                    {
                        bgBrush = b;
                        bgBrush.Opacity = Config.Data.Opacity;
                        MainBorder.Background = bgBrush;
                    }
                }
                catch { }

                var fgBrush = new SolidColorBrush(Colors.White);
                try
                {
                    object converted = brushConverter.ConvertFromString(Config.Data.ForegroundColor);
                    if (converted is SolidColorBrush b)
                    {
                        fgBrush = b;
                        LstReservations.Foreground = fgBrush;
                    }
                }
                catch { }

                try 
                {
                    object converted = brushConverter.ConvertFromString(Config.Data.ScrollBarColor);
                    if (converted is SolidColorBrush sb)
                    {
                        this.Resources["ScrollBarBrush"] = sb;
                    }
                }
                catch { }
                
                var colBorderBrush = new SolidColorBrush(Color.FromRgb(68, 68, 68));
                try 
                {
                    var converter = new BrushConverter();
                    var converted = (SolidColorBrush)converter.ConvertFromString(Config.Data.ColumnBorderColor);
                    if (converted != null)
                    {
                        colBorderBrush = converted;
                    }
                    if (MainBorder.Resources.Contains("ColumnBorderBrush"))
                    {
                        MainBorder.Resources["ColumnBorderBrush"] = colBorderBrush;
                    }
                    this.Resources["ColumnBorderBrush"] = colBorderBrush;

                    var fBrush = (SolidColorBrush)converter.ConvertFromString(Config.Data.FooterColor);
                    LblStatus.Foreground = fBrush;
                } 
                catch { }

                HeaderTitle.Foreground = new SolidColorBrush(Color.FromRgb(136, 204, 255));

                var borderBrush = new SolidColorBrush(Color.FromRgb(85, 85, 85));
                borderBrush.Opacity = Config.Data.Opacity;
                MainBorder.BorderBrush = borderBrush;
                
                LstReservations.Margin = new Thickness(
                    Config.Data.ListMarginLeft, 
                    Config.Data.ListMarginTop, 
                    Config.Data.ListMarginRight, 
                    Config.Data.ListMarginBottom
                );
                
                // ツールチップのスタイル設定（幅固定・折り返し・配色）
                var toolTipStyle = new Style(typeof(System.Windows.Controls.ToolTip));
                toolTipStyle.Setters.Add(new Setter(FrameworkElement.MaxWidthProperty, Config.Data.ToolTipWidth));
                
                try
                {
                    toolTipStyle.Setters.Add(new Setter(Control.FontSizeProperty, Config.Data.ToolTipFontSize));

                    var bgObj = brushConverter.ConvertFromString(Config.Data.ToolTipBackColor);
                    if (bgObj is SolidColorBrush ttBgBrush)
                        toolTipStyle.Setters.Add(new Setter(Control.BackgroundProperty, ttBgBrush));

                    var fgObj = brushConverter.ConvertFromString(Config.Data.ToolTipForeColor);
                    if (fgObj is SolidColorBrush ttFgBrush)
                        toolTipStyle.Setters.Add(new Setter(Control.ForegroundProperty, ttFgBrush));

                    var borderObj = brushConverter.ConvertFromString(Config.Data.ToolTipBorderColor);
                    if (borderObj is SolidColorBrush ttBorderBrush)
                        toolTipStyle.Setters.Add(new Setter(Control.BorderBrushProperty, ttBorderBrush));
                }
                catch 
                {
                    // デフォルトフォールバック（設定読み込み失敗時）
                    toolTipStyle.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(30, 30, 30))));
                    toolTipStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
                    toolTipStyle.Setters.Add(new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(100, 100, 100))));
                }
                
                // テキストの折り返し設定
                var ttTemplate = new DataTemplate();
                var ttText = new FrameworkElementFactory(typeof(System.Windows.Controls.TextBlock));
                ttText.SetBinding(System.Windows.Controls.TextBlock.TextProperty, new Binding());
                ttText.SetValue(System.Windows.Controls.TextBlock.TextWrappingProperty, TextWrapping.Wrap);
                ttTemplate.VisualTree = ttText;
                
                toolTipStyle.Setters.Add(new Setter(ContentControl.ContentTemplateProperty, ttTemplate));
                
                // ウィンドウのリソースとして登録
                this.Resources[typeof(System.Windows.Controls.ToolTip)] = toolTipStyle;

                var itemStyle = new Style(typeof(ListViewItem));
                itemStyle.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.Transparent));
                if (Config.Data.ShowToolTip)
                {
                    itemStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, new Binding("ToolTipText")));
                }
                itemStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0)));
                itemStyle.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(-1, 0, 0, 0)));
                itemStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(0)));
                itemStyle.Setters.Add(new Setter(Control.MinHeightProperty, 0.0));
                itemStyle.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, System.Windows.HorizontalAlignment.Stretch));
                itemStyle.Setters.Add(new Setter(Control.VerticalContentAlignmentProperty, System.Windows.VerticalAlignment.Center));

                itemStyle.Setters.Add(new EventSetter(Control.MouseDoubleClickEvent, new MouseButtonEventHandler(listView_MouseDoubleClick)));

                try
                {
                    object recConv = brushConverter.ConvertFromString(Config.Data.RecColor);
                    if (recConv is SolidColorBrush recBrush)
                    {
                        MainBorder.Resources["RecBrush"] = recBrush;
                        var recWeight = Config.Data.RecBold ? FontWeights.Bold : FontWeights.Normal;
                        MainBorder.Resources["RecWeight"] = recWeight;

                        var recTrigger = new DataTrigger { Binding = new Binding("IsRecording"), Value = true };
                        recTrigger.Setters.Add(new Setter(Control.ForegroundProperty, recBrush));
                        recTrigger.Setters.Add(new Setter(Control.FontWeightProperty, recWeight));
                        itemStyle.Triggers.Add(recTrigger);
                    }
                    
                    object disConv = brushConverter.ConvertFromString(Config.Data.DisabledColor);
                    if (disConv is SolidColorBrush disabledBrush)
                    {
                        MainBorder.Resources["DisabledBrush"] = disabledBrush;
                        var disabledTrigger = new DataTrigger { Binding = new Binding("IsDisabled"), Value = true };
                        disabledTrigger.Setters.Add(new Setter(Control.ForegroundProperty, disabledBrush));
                        itemStyle.Triggers.Add(disabledTrigger);
                    }

                    object errConv = brushConverter.ConvertFromString(Config.Data.ReserveErrorColor);
                    if (errConv is SolidColorBrush errBrush)
                    {
                        MainBorder.Resources["ErrorBrush"] = errBrush;
                        
                        var errorTrigger = new DataTrigger { Binding = new Binding("HasError"), Value = true };
                        errorTrigger.Setters.Add(new Setter(Control.BackgroundProperty, errBrush));
                        itemStyle.Triggers.Add(errorTrigger);
                    }
                }
                catch { }

                double brightness = bgBrush.Color.R * 0.299 + bgBrush.Color.G * 0.587 + bgBrush.Color.B * 0.114;
                bool isLight = brightness > 128; 

                var selectedTrigger = new Trigger { Property = ListViewItem.IsSelectedProperty, Value = true };
                var selColor = isLight ? Color.FromArgb(100, 0, 100, 200) : Color.FromArgb(80, 255, 255, 255);
                selectedTrigger.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(selColor)));
                itemStyle.Triggers.Add(selectedTrigger);

                var mouseOverTrigger = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
                var hoverColor = isLight ? Color.FromArgb(50, 0, 100, 200) : Color.FromArgb(50, 255, 255, 255);
                mouseOverTrigger.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush(hoverColor)));
                itemStyle.Triggers.Add(mouseOverTrigger);

                LstReservations.ItemContainerStyle = itemStyle;

                RowHeader.Height = Config.Data.ShowHeader ? GridLength.Auto : new GridLength(0);
                RowFooter.Height = Config.Data.ShowFooter ? GridLength.Auto : new GridLength(0);
                MainBorder.BorderThickness = new Thickness(1);

                UpdateListColumns();
                UpdateColumnHeaderStyle(bgBrush, fgBrush, colBorderBrush);
                LstReservations.Items.Refresh();
            }
            catch (Exception ex) { Logger.Write("ApplySettings Error: " + ex.Message); }
        }

        private void UpdateListColumns()
        {
            var gv = (GridView)LstReservations.View;
            gv.Columns.Clear();

            var defs = new List<ColumnDef>();
            defs.Add(new ColumnDef { Header = "状態", BindingPath = "Status", GetShow = () => Config.Data.ShowColStatus, GetWidth = () => Config.Data.WidthColStatus });
            defs.Add(new ColumnDef { Header = "日時", BindingPath = "DateTimeInfo", GetShow = () => Config.Data.ShowColDateTime, GetWidth = () => Config.Data.WidthColDateTime });
            defs.Add(new ColumnDef { Header = "長さ", BindingPath = "Duration", GetShow = () => Config.Data.ShowColDuration, GetWidth = () => Config.Data.WidthColDuration });
            defs.Add(new ColumnDef { Header = "ネットワーク", BindingPath = "NetworkName", GetShow = () => Config.Data.ShowColNetwork, GetWidth = () => Config.Data.WidthColNetwork });
            defs.Add(new ColumnDef { Header = "サービス名", BindingPath = "ServiceName", GetShow = () => Config.Data.ShowColServiceName, GetWidth = () => Config.Data.WidthColServiceName });
            defs.Add(new ColumnDef { Header = "番組名", BindingPath = "Title", GetShow = () => Config.Data.ShowColTitle, GetWidth = () => Config.Data.WidthColTitle });
            defs.Add(new ColumnDef { Header = "番組内容", BindingPath = "Desc", GetShow = () => Config.Data.ShowColDesc, GetWidth = () => Config.Data.WidthColDesc });
            defs.Add(new ColumnDef { Header = "ジャンル", BindingPath = "Genre", GetShow = () => Config.Data.ShowColGenre, GetWidth = () => Config.Data.WidthColGenre });
            defs.Add(new ColumnDef { Header = "付属情報", BindingPath = "ExtraInfo", GetShow = () => Config.Data.ShowColExtraInfo, GetWidth = () => Config.Data.WidthColExtraInfo });
            
            defs.Add(new ColumnDef { Header = "有効", BindingPath = "IsEnabledBool", GetShow = () => Config.Data.ShowColEnabled, GetWidth = () => Config.Data.WidthColEnabled });
            
            defs.Add(new ColumnDef { Header = "プログラム予約", BindingPath = "ProgramType", GetShow = () => Config.Data.ShowColProgramType, GetWidth = () => Config.Data.WidthColProgramType });
            defs.Add(new ColumnDef { Header = "予約状況", BindingPath = "Comment", GetShow = () => Config.Data.ShowColComment, GetWidth = () => Config.Data.WidthColComment });
            defs.Add(new ColumnDef { Header = "エラー状況", BindingPath = "ErrorInfo", GetShow = () => Config.Data.ShowColError, GetWidth = () => Config.Data.WidthColError });
            defs.Add(new ColumnDef { Header = "予定ファイル名", BindingPath = "RecFileName", GetShow = () => Config.Data.ShowColRecFileName, GetWidth = () => Config.Data.WidthColRecFileName });
            defs.Add(new ColumnDef { Header = "予定ファイル名リスト", BindingPath = "RecFileNameList", GetShow = () => Config.Data.ShowColRecFileNameList, GetWidth = () => Config.Data.WidthColRecFileNameList });
            defs.Add(new ColumnDef { Header = "使用予定チューナー", BindingPath = "Tuner", GetShow = () => Config.Data.ShowColTuner, GetWidth = () => Config.Data.WidthColTuner });
            defs.Add(new ColumnDef { Header = "予想サイズ", BindingPath = "EstSize", GetShow = () => Config.Data.ShowColEstSize, GetWidth = () => Config.Data.WidthColEstSize });
            defs.Add(new ColumnDef { Header = "プリセット", BindingPath = "Preset", GetShow = () => Config.Data.ShowColPreset, GetWidth = () => Config.Data.WidthColPreset });
            defs.Add(new ColumnDef { Header = "録画モード", BindingPath = "RecMode", GetShow = () => Config.Data.ShowColRecMode, GetWidth = () => Config.Data.WidthColRecMode });
            defs.Add(new ColumnDef { Header = "優先度", BindingPath = "Priority", GetShow = () => Config.Data.ShowColPriority, GetWidth = () => Config.Data.WidthColPriority });
            defs.Add(new ColumnDef { Header = "追従", BindingPath = "Tuijyuu", GetShow = () => Config.Data.ShowColTuijyuu, GetWidth = () => Config.Data.WidthColTuijyuu });
            defs.Add(new ColumnDef { Header = "ぴったり", BindingPath = "Pittari", GetShow = () => Config.Data.ShowColPittari, GetWidth = () => Config.Data.WidthColPittari });
            defs.Add(new ColumnDef { Header = "チューナー強制", BindingPath = "TunerForce", GetShow = () => Config.Data.ShowColTunerForce, GetWidth = () => Config.Data.WidthColTunerForce });
            defs.Add(new ColumnDef { Header = "録画後動作", BindingPath = "RecEndMode", GetShow = () => Config.Data.ShowColRecEndMode, GetWidth = () => Config.Data.WidthColRecEndMode });
            defs.Add(new ColumnDef { Header = "復帰後再起動", BindingPath = "Reboot", GetShow = () => Config.Data.ShowColReboot, GetWidth = () => Config.Data.WidthColReboot });
            defs.Add(new ColumnDef { Header = "録画後実行bat", BindingPath = "Bat", GetShow = () => Config.Data.ShowColBat, GetWidth = () => Config.Data.WidthColBat });
            defs.Add(new ColumnDef { Header = "録画タグ", BindingPath = "RecTag", GetShow = () => Config.Data.ShowColRecTag, GetWidth = () => Config.Data.WidthColRecTag });
            defs.Add(new ColumnDef { Header = "録画フォルダ", BindingPath = "RecFolder", GetShow = () => Config.Data.ShowColRecFolder, GetWidth = () => Config.Data.WidthColRecFolder });
            defs.Add(new ColumnDef { Header = "開始", BindingPath = "StartMargin", GetShow = () => Config.Data.ShowColStartMargin, GetWidth = () => Config.Data.WidthColStartMargin });
            defs.Add(new ColumnDef { Header = "終了", BindingPath = "EndMargin", GetShow = () => Config.Data.ShowColEndMargin, GetWidth = () => Config.Data.WidthColEndMargin });
            defs.Add(new ColumnDef { Header = "ID", BindingPath = "ID", GetShow = () => Config.Data.ShowColID, GetWidth = () => Config.Data.WidthColID });


            if (Config.Data.ColumnHeaderOrder != null && Config.Data.ColumnHeaderOrder.Count > 0)
            {
                foreach (string header in Config.Data.ColumnHeaderOrder)
                {
                    var d = defs.FirstOrDefault(x => x.Header == header);
                    if (d != null && d.GetShow())
                    {
                        if (d.Header == "有効") AddCheckBoxColumn(gv, d);
                        else if (d.Header == "日時") AddDateTimeColumn(gv, d);
                        else AddColumn(gv, d);
                    }
                }

                foreach (var d in defs)
                {
                    if (d.GetShow() && !Config.Data.ColumnHeaderOrder.Contains(d.Header))
                    {
                        if (d.Header == "有効") AddCheckBoxColumn(gv, d);
                        else if (d.Header == "日時") AddDateTimeColumn(gv, d);
                        else AddColumn(gv, d);
                    }
                }
            }
            else
            {
                foreach(var d in defs)
                {
                    if (d.GetShow())
                    {
                        if (d.Header == "有効") AddCheckBoxColumn(gv, d);
                        else if (d.Header == "日時") AddDateTimeColumn(gv, d);
                        else AddColumn(gv, d);
                    }
                }
            }
        }

        private void AddColumn(GridView gv, ColumnDef d)
        {
            var col = new GridViewColumn();
            col.Header = d.Header;
            col.Width = d.GetWidth();
            
            var dataTemplate = new DataTemplate();
            var factory = new FrameworkElementFactory(typeof(TextBlock));
            factory.SetBinding(TextBlock.TextProperty, new Binding(d.BindingPath));
            
            double left = 2;
            factory.SetValue(TextBlock.MarginProperty, new Thickness(left, Config.Data.ItemPadding, -6, Config.Data.ItemPadding));
            
            dataTemplate.VisualTree = factory;
            
            col.CellTemplate = dataTemplate;
            gv.Columns.Add(col);
        }

        private void AddCheckBoxColumn(GridView gv, ColumnDef d)
        {
            var col = new GridViewColumn();
            col.Header = d.Header;
            col.Width = d.GetWidth();

            var dataTemplate = new DataTemplate();
            var factory = new FrameworkElementFactory(typeof(System.Windows.Controls.CheckBox));
            factory.SetBinding(ToggleButton.IsCheckedProperty, new Binding(d.BindingPath) { Mode = BindingMode.OneWay });
            factory.SetValue(Control.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
            factory.SetValue(Control.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
            factory.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(ReservationCheckBox_Click));

            dataTemplate.VisualTree = factory;
            col.CellTemplate = dataTemplate;
            gv.Columns.Add(col);
        }
        
        private void AddDateTimeColumn(GridView gv, ColumnDef d)
        {
            var col = new GridViewColumn();
            col.Header = d.Header;
            col.Width = d.GetWidth();

            var dataTemplate = new DataTemplate();
            var gridFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.Grid));

            // 「進捗バーを省略」していない場合のみバーを追加
            if (!Config.Data.OmitProgress)
            {
                var progressFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.ProgressBar));
                progressFactory.SetBinding(System.Windows.Controls.Primitives.RangeBase.ValueProperty, new Binding("ProgressValue"));
                progressFactory.SetValue(System.Windows.Controls.ProgressBar.MinimumProperty, 0.0);
                progressFactory.SetValue(System.Windows.Controls.ProgressBar.MaximumProperty, 100.0);
                
                // デザイン調整（枠なし）
                progressFactory.SetValue(Control.BorderThicknessProperty, new Thickness(0));
                progressFactory.SetValue(Control.BackgroundProperty, Brushes.Transparent);
                
                try
                {
                    // 設定された色をそのまま適用（不透明）
                    var color = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(Config.Data.ProgressBarColor);
                    var brush = new SolidColorBrush(color);
                    progressFactory.SetValue(Control.ForegroundProperty, brush);
                }
                catch
                {
                    // パース失敗時のデフォルト色（緑）
                    progressFactory.SetValue(Control.ForegroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 255, 0)));
                }

                progressFactory.SetValue(FrameworkElement.MarginProperty, new Thickness(1, 0, -6, 0));
                progressFactory.SetValue(FrameworkElement.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Stretch);
                progressFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Stretch);
                
                // 下約20%の高さになるように変形
                progressFactory.SetValue(UIElement.RenderTransformOriginProperty, new Point(0.5, 1.0));
                progressFactory.SetValue(UIElement.RenderTransformProperty, new System.Windows.Media.ScaleTransform(1.0, 0.20));
                
                gridFactory.AppendChild(progressFactory);
            }

            // テキスト表示（前面）
            var textFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.TextBlock));
            textFactory.SetBinding(System.Windows.Controls.TextBlock.TextProperty, new Binding(d.BindingPath));
            textFactory.SetValue(FrameworkElement.MarginProperty, new Thickness(2, Config.Data.ItemPadding, -6, Config.Data.ItemPadding));
            
            textFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
            
            gridFactory.AppendChild(textFactory);

            dataTemplate.VisualTree = gridFactory;
            col.CellTemplate = dataTemplate;
            gv.Columns.Add(col);
        }

        // --- ロジックを統合したチェックボックス処理 ---
        private async void ReservationCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox cb && cb.DataContext is ReserveItem item)
            {
                cb.IsEnabled = false; 
                bool isChecked = cb.IsChecked == true;
                uint id = item.ID;

                bool success = await ChangeReservationStatus(id, isChecked);
                
                if (success) await UpdateReservations();
                else 
                {
                    cb.IsChecked = !isChecked; 
                    cb.IsEnabled = true;
                }
            }
        }

        // --- ロジックを統合した予約状態変更 ---
        private async Task<bool> ChangeReservationStatus(uint targetID, bool enable)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var cmd = new CtrlCmdUtil();
                    var list = new List<ReserveData>();
                    if (cmd.SendEnumReserve(ref list) == ErrCode.CMD_SUCCESS)
                    {
                        var target = list.FirstOrDefault(r => r.ReserveID == targetID);
                        if (target != null)
                        {
                            byte current = target.RecSetting.RecMode;
                            byte next = current;
                            bool isCurrentlyEnabled = current <= 4;

                            if (enable && !isCurrentlyEnabled)
                            {
                                // 無効(5-9) -> 有効(0-4)
                                // RecModeを維持するロジック（current%5）
                                // 0(全サービス)より1(指定サービス)が無難なため調整
                                next = (byte)((current % 5)); 
                                if (next == 0) next = 1; 
                            }
                            else if (!enable && isCurrentlyEnabled)
                            {
                                // 有効(0-4) -> 無効(5-9)
                                next = (byte)(current + 5);
                            }
                            else
                            {
                                return true;
                            }

                            target.RecSetting.RecMode = next;
                            var changeList = new List<ReserveData> { target };
                            return cmd.SendChgReserve(changeList) == ErrCode.CMD_SUCCESS;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Write("ChangeReservationStatus Error: " + ex.Message);
                }
                return false;
            });
        }

        private ContextMenu CreateHeaderContextMenu()
        {
            var menu = new ContextMenu();
            
            Action<string, bool, Action<bool>> addItem = (header, current, setAction) => {
                var item = new MenuItem { Header = header, IsCheckable = true, IsChecked = current };
                item.Click += (s, e) => {
                    setAction(item.IsChecked);
                    SaveCurrentState();
                    ApplySettings();
                };
                menu.Items.Add(item);
            };

            addItem("状態", Config.Data.ShowColStatus, v => Config.Data.ShowColStatus = v);
            addItem("日時", Config.Data.ShowColDateTime, v => Config.Data.ShowColDateTime = v);
            addItem("長さ", Config.Data.ShowColDuration, v => Config.Data.ShowColDuration = v);
            addItem("ネットワーク", Config.Data.ShowColNetwork, v => Config.Data.ShowColNetwork = v);
            addItem("サービス名", Config.Data.ShowColServiceName, v => Config.Data.ShowColServiceName = v);
            addItem("番組名", Config.Data.ShowColTitle, v => Config.Data.ShowColTitle = v);
            
            menu.Items.Add(new Separator());

            addItem("番組内容", Config.Data.ShowColDesc, v => Config.Data.ShowColDesc = v);
            addItem("ジャンル", Config.Data.ShowColGenre, v => Config.Data.ShowColGenre = v);
            addItem("付属情報", Config.Data.ShowColExtraInfo, v => Config.Data.ShowColExtraInfo = v);
            addItem("有効/無効", Config.Data.ShowColEnabled, v => Config.Data.ShowColEnabled = v);
            addItem("プログラム予約", Config.Data.ShowColProgramType, v => Config.Data.ShowColProgramType = v);
            
            menu.Items.Add(new Separator());

            addItem("予約状況", Config.Data.ShowColComment, v => Config.Data.ShowColComment = v);
            addItem("エラー状況", Config.Data.ShowColError, v => Config.Data.ShowColError = v);
            addItem("予定ファイル名", Config.Data.ShowColRecFileName, v => Config.Data.ShowColRecFileName = v);
            addItem("予定ファイル名リスト", Config.Data.ShowColRecFileNameList, v => Config.Data.ShowColRecFileNameList = v);
            
            menu.Items.Add(new Separator());

            addItem("使用予定チューナー", Config.Data.ShowColTuner, v => Config.Data.ShowColTuner = v);
            addItem("予想サイズ", Config.Data.ShowColEstSize, v => Config.Data.ShowColEstSize = v);
            addItem("プリセット", Config.Data.ShowColPreset, v => Config.Data.ShowColPreset = v);
            addItem("録画モード", Config.Data.ShowColRecMode, v => Config.Data.ShowColRecMode = v);
            addItem("優先度", Config.Data.ShowColPriority, v => Config.Data.ShowColPriority = v);
            addItem("追従", Config.Data.ShowColTuijyuu, v => Config.Data.ShowColTuijyuu = v);
            addItem("ぴったり", Config.Data.ShowColPittari, v => Config.Data.ShowColPittari = v);
            addItem("チューナー強制", Config.Data.ShowColTunerForce, v => Config.Data.ShowColTunerForce = v);
            
            menu.Items.Add(new Separator());

            addItem("録画後動作", Config.Data.ShowColRecEndMode, v => Config.Data.ShowColRecEndMode = v);
            addItem("復帰後再起動", Config.Data.ShowColReboot, v => Config.Data.ShowColReboot = v);
            addItem("録画後実行bat", Config.Data.ShowColBat, v => Config.Data.ShowColBat = v);
            addItem("録画タグ", Config.Data.ShowColRecTag, v => Config.Data.ShowColRecTag = v);
            addItem("録画フォルダ", Config.Data.ShowColRecFolder, v => Config.Data.ShowColRecFolder = v);
            addItem("開始", Config.Data.ShowColStartMargin, v => Config.Data.ShowColStartMargin = v);
            addItem("終了", Config.Data.ShowColEndMargin, v => Config.Data.ShowColEndMargin = v);
            addItem("ID", Config.Data.ShowColID, v => Config.Data.ShowColID = v);

            return menu;
        }

        // --- ヘッダーフォントサイズの反映 ---
        private void UpdateColumnHeaderStyle(Brush bgBrush, Brush fgBrush, Brush borderBrush)
        {
            try
            {
                var gv = (GridView)LstReservations.View;
                if (gv == null || bgBrush == null || fgBrush == null) return;

                bool isVisible = Config.Data.ShowListHeader;

                var headerStyle = new Style(typeof(GridViewColumnHeader));
                
                // 設定値のフォントサイズを適用
                headerStyle.Setters.Add(new Setter(Control.FontSizeProperty, Config.Data.HeaderFontSize));

                headerStyle.Setters.Add(new Setter(UIElement.VisibilityProperty, isVisible ? Visibility.Visible : Visibility.Collapsed));
                headerStyle.Setters.Add(new Setter(Control.BackgroundProperty, bgBrush));
                headerStyle.Setters.Add(new Setter(Control.ForegroundProperty, fgBrush));
                
                headerStyle.Setters.Add(new Setter(Control.BorderBrushProperty, borderBrush));
                
                headerStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1, 1, 1, 1)));
                headerStyle.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(-1, 0, 0, 1)));

                headerStyle.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, System.Windows.HorizontalAlignment.Left));
                headerStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(4, 2, 0, 2)));
                
                headerStyle.Setters.Add(new Setter(FrameworkElement.ContextMenuProperty, CreateHeaderContextMenu()));

                var paddingTrigger = new Trigger { Property = GridViewColumnHeader.RoleProperty, Value = GridViewColumnHeaderRole.Padding };
                paddingTrigger.Setters.Add(new Setter(UIElement.VisibilityProperty, Visibility.Collapsed));
                headerStyle.Triggers.Add(paddingTrigger);

                // XAMLリソースからテンプレートを取得
                var template = this.FindResource("HeaderTemplate") as ControlTemplate;
                if (template != null)
                {
                    headerStyle.Setters.Add(new Setter(Control.TemplateProperty, template));
                }
                
                gv.ColumnHeaderContainerStyle = headerStyle;
            }
            catch { }
        }

        private void SetupFileWatcher()
        {
            if (_fileWatcher != null) 
            { 
                _fileWatcher.Dispose(); 
                _fileWatcher = null; 
            }

            string path = ReserveTextReader.GetReserveFilePath();
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                try
                {
                    string dir = Path.GetDirectoryName(path);
                    string fileName = Path.GetFileName(path);

                    if (!string.IsNullOrEmpty(dir))
                    {
                        _fileWatcher = new FileSystemWatcher(dir, fileName);
                        _fileWatcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName | NotifyFilters.CreationTime;
                        _fileWatcher.InternalBufferSize = 65536;

                        FileSystemEventHandler handler = async (s, e) => { await Dispatcher.InvokeAsync(async () => await HandleFileChange()); };
                        RenamedEventHandler renamedHandler = async (s, e) => { await Dispatcher.InvokeAsync(async () => await HandleFileChange()); };

                        _fileWatcher.Changed += handler;
                        _fileWatcher.Created += handler;
                        _fileWatcher.Renamed += renamedHandler;

                        _fileWatcher.EnableRaisingEvents = true;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Write("FileWatcher Error: " + ex.Message);
                }
            }
        }

        private async Task HandleFileChange()
        {
            if ((DateTime.Now - _lastReloadTime).TotalMilliseconds < 500)
            {
                return;
            }
            _lastReloadTime = DateTime.Now;
            await UpdateReservations();
        }

        private async Task<bool> UpdateReservations()
        {
            try
            {
                var list = await Task.Run(() => ReserveTextReader.Load());
                
                if (list == null)
                {
                    LblStatus.Text = "接続待機中...";
                    return false;
                }

                LstReservations.ItemsSource = list;
                LblStatus.Text = string.Format("更新: {0:HH:mm} ({1}件)", DateTime.Now, list.Count);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Write("Update Error: " + ex.ToString());
                LblStatus.Text = "エラー発生";
                return false;
            }
        }
        
        public async Task RefreshDataAsync()
        {
            await UpdateReservations();
        }

        private void MenuReload_Click(object sender, RoutedEventArgs e) { var t = UpdateReservations(); }
        
        private void MenuExit_Click(object sender, RoutedEventArgs e) { this.Close(); }
        
        private void MenuSettings_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrentState();

            var win = new SettingsWindow { Owner = this };
            if (win.ShowDialog() == true)
            {
                ApplySettings();
                SetupFileWatcher();
                var t = UpdateReservations();
            }
        }

        private void ContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            if (sender is ContextMenu menu)
            {
                foreach (var item in menu.Items)
                {
                    if (item is MenuItem mi && mi.Name == "MenuItemHideDisabled")
                    {
                        mi.Header = Config.Data.HideDisabled ? "無効予約を表示する" : "無効予約を表示しない";
                        break;
                    }
                }
            }
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

        // --- ロジックを統合したメニュー操作 ---
        private async void MenuChgOnOff_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = LstReservations.SelectedItems;
            if (selectedItems == null || selectedItems.Count == 0) return;

            var targetIDs = new List<uint>();
            foreach (var item in selectedItems)
            {
                try
                {
                    if (item != null)
                    {
                        dynamic dItem = item;
                        targetIDs.Add((uint)dItem.ID);
                    }
                }
                catch { }
            }

            if (targetIDs.Count == 0) return;

            LblStatus.Text = "予約情報を変更中...";

            bool success = await Task.Run(() =>
            {
                try
                {
                    var cmd = new CtrlCmdUtil();
                    var list = new List<ReserveData>();

                    if (cmd.SendEnumReserve(ref list) == ErrCode.CMD_SUCCESS)
                    {
                        var changeList = new List<ReserveData>();

                        foreach (var data in list)
                        {
                           if (targetIDs.Contains(data.ReserveID))
                           {
                               byte current = data.RecSetting.RecMode;
                               byte next;

                               if (current <= 4) 
                               {
                                   next = (byte)(5 + (current + 4) % 5);
                               }
                               else if (current <= 9) 
                               {
                                    next = (byte)((current + 1) % 5);
                               }
                               else 
                               {
                                   next = 1; 
                               }

                               data.RecSetting.RecMode = next;
                               changeList.Add(data);
                           }
                       }

                        if (changeList.Count > 0)
                        {
                            return cmd.SendChgReserve(changeList) == ErrCode.CMD_SUCCESS;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Write("MenuChgOnOff Error: " + ex.Message);
                }
                return false;
            });

            if (success)
            {
                LblStatus.Text = "予約状態を変更しました";
                await UpdateReservations();
            }
            else
            {
                System.Windows.MessageBox.Show("変更に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                LblStatus.Text = "変更に失敗しました";
            }
        }
        
        private async void MenuDel_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = LstReservations.SelectedItems;
            if (selectedItems == null || selectedItems.Count == 0) return;

            var targetIDs = new List<uint>();
            foreach (var item in selectedItems)
            {
                try
                {
                    if (item != null)
                    {
                        dynamic dItem = item;
                        targetIDs.Add((uint)dItem.ID);
                    }
                }
                catch { }
            }

            if (targetIDs.Count == 0) return;

            var res = System.Windows.MessageBox.Show(
                string.Format("{0}件の予約を削除しますか？", targetIDs.Count),
                "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (res != MessageBoxResult.Yes) return;

            LblStatus.Text = "予約を削除中...";

            bool success = await Task.Run(() =>
            {
                try
                {
                    var cmd = new CtrlCmdUtil();
                    var err = cmd.SendDelReserve(targetIDs);
                    return err == ErrCode.CMD_SUCCESS;
                }
                catch (Exception ex)
                {
                    Logger.Write("MenuDel: 例外発生\r\n" + ex.ToString()); 
                    return false;
                }
            });

            if (success)
            {
                LblStatus.Text = "予約を削除しました";
                await UpdateReservations(); 
            }
            else
            {
                System.Windows.MessageBox.Show("削除に失敗しました。\r\n詳細なエラー原因は EDCBMonitor_Log.txt を確認してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                LblStatus.Text = "削除に失敗しました";
            }
        }
    }
}