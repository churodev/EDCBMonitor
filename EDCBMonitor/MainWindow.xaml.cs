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
        private const int MAX_RETRY_COUNT = 3;
        private const int MOUSE_SNAP_DIST = 20;
        private const int WM_MOUSEHWHEEL = 0x020E;

        // 正確なウィンドウ位置を取得するためのAPI定義
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private readonly DispatcherTimer _updateTimer;
        private FileSystemWatcher? _fileWatcher;
        private WinForms.NotifyIcon? _notifyIcon;
        private DispatcherTimer? _reloadDebounceTimer;
        
        private int _retryCount = 0;
        private bool _isShowingTempMessage = false;
        private bool _isDragging = false;
        private Point _startMousePoint;
        private Rect? _restoreBounds = null;

        // ミニモード管理用
        private bool _isMiniMode = false;
        private Rect _fullWindowRect;
        private DispatcherTimer _miniModeTimer;
        private DispatcherTimer _miniModeExpandTimer;
        private DispatcherTimer _saveDebounceTimer;
        private bool _isProgrammaticMove = false;

        private readonly ReservationService _reservationService;
        private readonly GridColumnManager _columnManager;
        
        // 監視状態を表示するためのフィールド
        private string _watcherStatusMessage = "";
        private bool _isContextMenuOpen = false;
        // 前回録画中だった予約のIDを保持する
        private HashSet<uint> _lastRecordingIds = new HashSet<uint>();

        public MainWindow()
        {
            InitializeComponent();
            InitializeNotifyIcon();
            LoadAppIcon();
            
            // 保存用タイマーの初期化 (操作終了後 0.5秒で保存)
            _saveDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            _saveDebounceTimer.Tick += SaveDebounceTimer_Tick;
            
            _reservationService = new ReservationService();
            _columnManager = new GridColumnManager(LstReservations, ReservationCheckBox_Click);

            _updateTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _updateTimer.Tick += UpdateTimer_Tick;
            _updateTimer.Start();
            
            // ミニモード用タイマー初期化
            _miniModeTimer = new DispatcherTimer();
            _miniModeTimer.Tick += MiniModeTimer_Tick;
            
            // 展開用タイマー初期化
            _miniModeExpandTimer = new DispatcherTimer();
            _miniModeExpandTimer.Tick += MiniModeExpandTimer_Tick;

            ApplySettings(true);

            Loaded += async (s, e) => 
            {
                EnsureWindowIsVisible();
                SetupFileWatcher();
                await UpdateReservations();
            };

            Closing += Window_Closing;
            ContextMenuOpening += Window_ContextMenuOpening;
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

        private void InitializeNotifyIcon()
        {
            _notifyIcon = new WinForms.NotifyIcon
            {
                Text = "EDCB Monitor",
                Visible = Config.Data.ShowTrayIcon
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

    }
}