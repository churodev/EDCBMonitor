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
            // ミニモード中は展開サイズを正として保存する
            if (_isMiniMode)
            {
                // ドラッグ等で位置が変わっている場合を考慮してフルサイズRectを更新してから保存
                // 現在位置から逆算するのは複雑なため、ここでは_fullWindowRectの値を優先する
                Config.Data.Top = _fullWindowRect.Top;
                Config.Data.Left = _fullWindowRect.Left;
                Config.Data.Width = _fullWindowRect.Width;
                Config.Data.Height = _fullWindowRect.Height;
                if (_restoreBounds.HasValue)
                {
                    Config.Data.IsVerticalMaximized = true;
                    Config.Data.RestoreTop = _restoreBounds.Value.Top;
                    Config.Data.RestoreHeight = _restoreBounds.Value.Height;
                }
                else
                {
                    Config.Data.IsVerticalMaximized = false;
                }

                _columnManager.SaveColumnState();
                return;
            }
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

        private void BtnVerticalMaximize_Click(object sender, RoutedEventArgs e)
        {
            // ミニモード中なら先に展開して正しいフルサイズ座標に戻す
            if (_isMiniMode)
            {
                UpdateMiniModeState(false);
            }

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

        // OSの機能を使って正確にマウス位置を判定し設定保存と縮小を行う
        private void SaveDebounceTimer_Tick(object? sender, EventArgs e)
        {
            if (_saveDebounceTimer == null) return;
            _saveDebounceTimer.Stop();

            // ミニモード展開中（通常表示）のときのみ実行
            if (!_isMiniMode && WindowState == WindowState.Normal)
            {
                // 1. 設定を確実に保存
                _fullWindowRect = new Rect(Left, Top, Width, Height);
                SaveCurrentState();
                Config.Save();

                // リサイズ・移動操作が終わった後のマウス位置確認
                // 操作中は MouseLeave をガードしているので操作終了時のここで改めて判定して縮小へ繋げる
                try
                {
                    var interopHelper = new WindowInteropHelper(this);
                    RECT winRect;
                    
                    // OSから正確なウィンドウ位置を取得
                    if (GetWindowRect(interopHelper.Handle, out winRect))
                    {
                        var cursor = WinForms.Cursor.Position;
                        bool isOver = (cursor.X >= winRect.Left && cursor.X <= winRect.Right &&
                                       cursor.Y >= winRect.Top && cursor.Y <= winRect.Bottom);
                        
                        // マウスが外にあるなら、縮小処理(MouseLeave)を開始
                        if (!isOver)
                        {
                            // この時点では _saveDebounceTimer は停止済みなので
                            // MouseLeave 内のガード条件を通過して正しく縮小タイマーが始動する
                            Window_MouseLeave(this, null);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Write($"Mouse Check Error: {ex.Message}");
                }
            }
        }

        // --- ミニモード制御ロジック ---
        private void Window_MouseEnter(object sender, MouseEventArgs e)
        {
            // マウスが戻ってきたら縮小タイマーをキャンセル
            _miniModeTimer.Stop();
            
            if (_isMiniMode)
            {
                // 設定された遅延時間で展開
                if (Config.Data.MiniModeExpandDelay > 0)
                {
                    _miniModeExpandTimer.Interval = TimeSpan.FromMilliseconds(Config.Data.MiniModeExpandDelay);
                    _miniModeExpandTimer.Start();
                }
                else
                {
                    UpdateMiniModeState(false);
                }
            }
        }

        private void Window_MouseLeave(object sender, MouseEventArgs e)
        {
            // 展開待ちの間にマウスが外れたら展開をキャンセル
            _miniModeExpandTimer.Stop();
            
            // コンテキストメニューや設定画面などが開いているときは縮小しない
            if (_isContextMenuOpen) return;

            // 設定画面が開いているか確認（簡易チェック）
            if (OwnedWindows.Count > 0) return;
            // リサイズや移動の操作中（保存タイマー稼働中）は、マウスが外れても縮小処理を開始しない
            if (_saveDebounceTimer.IsEnabled) return;

            // 上下最大化中は縮小しない
            if (_restoreBounds.HasValue) return;

            if (Config.Data.EnableMiniMode && !_isMiniMode)
            {
                // 即座に縮小せず、タイマーを開始する
                _miniModeTimer.Interval = TimeSpan.FromMilliseconds(Config.Data.MiniModeDelay);
                _miniModeTimer.Start();
            }
        }

        // 展開用タイマーの処理
        private void MiniModeExpandTimer_Tick(object? sender, EventArgs e)
        {
            _miniModeExpandTimer.Stop();
            // タイマー発火時にまだマウスが乗っていれば展開
            if (IsMouseOver)
            {
                UpdateMiniModeState(false);
            }
        }

        private void MiniModeTimer_Tick(object? sender, EventArgs e)
        {
            _miniModeTimer.Stop();

            // メニューが開いている場合は縮小しない
            if (_isContextMenuOpen) return;
            if ((ContextMenu != null && ContextMenu.IsOpen) ||
                (LstReservations.ContextMenu != null && LstReservations.ContextMenu.IsOpen)) return;

            // WPFの判定は不安定なため、OSのAPIを使って「本当にマウスが外れたか」を厳密にチェックする
            bool isReallyOver = false;
            try
            {
                var interopHelper = new WindowInteropHelper(this);
                RECT winRect;
                // ウィンドウの正確な位置を取得
                if (GetWindowRect(interopHelper.Handle, out winRect))
                {
                    var cursor = WinForms.Cursor.Position;
                    // マウス座標がウィンドウ矩形の中に含まれているか判定
                    if (cursor.X >= winRect.Left && cursor.X <= winRect.Right &&
                        cursor.Y >= winRect.Top && cursor.Y <= winRect.Bottom)
                    {
                        isReallyOver = true;
                    }
                }
            }
            catch 
            {
                // エラー時は安全側に倒して「乗っている」とみなす（縮小ループ防止）
                isReallyOver = true; 
            }

            // まだマウスが乗っている(OS判定)、またはWPFが乗っていると言っている場合
            if (isReallyOver || IsMouseOver)
            {
                // 誤判定でここに来た可能性があるため、縮小は中止する。
                // ただし、その後本当にマウスが外れるのを検知するため、少し間隔を空けてタイマーを再始動(監視)する。
                // (0秒設定だと高負荷になるため、監視ループは最低100ms空ける)
                _miniModeTimer.Interval = TimeSpan.FromMilliseconds(Math.Max(100, Config.Data.MiniModeDelay));
                _miniModeTimer.Start();
                return;
            }

            // 本当に外れていれば縮小実行
            UpdateMiniModeState(true);
        }

        private void UpdateMiniModeState(bool enterMiniMode)
        {
            if (enterMiniMode)
            {
                // ミニモードへ移行（縮小）
                if (!_isMiniMode)
                {
                    // 現在の状態をフルサイズとして保存
                    _fullWindowRect = new Rect(Left, Top, Width, Height);
                }

                double scaleX = Math.Max(10, Config.Data.MiniModeScaleX) / 100.0;
                double scaleY = Math.Max(10, Config.Data.MiniModeScaleY) / 100.0;

                double newW = _fullWindowRect.Width * scaleX;
                double newH = _fullWindowRect.Height * scaleY;

                // 方向に応じた位置調整
                // 0:左上固定, 1:右上固定, 2:左下固定, 3:右下固定
                double newL = _fullWindowRect.Left;
                double newT = _fullWindowRect.Top;

                switch (Config.Data.MiniModeDirection)
                {
                    case 0: // 右下に伸ばす (左上固定) -> 縮小時も左上維持
                        // newL, newT は初期値のまま
                        break;
                    case 1: // 左下に伸ばす (右上固定) -> 縮小時は右端維持
                        newL = _fullWindowRect.Right - newW;
                        // newT はTop維持
                        break;
                    case 2: // 右上に伸ばす (左下固定) -> 縮小時は左下維持
                        // newL はLeft維持
                        newT = _fullWindowRect.Bottom - newH;
                        break;
                    case 3: // 左上に伸ばす (右下固定) -> 縮小時は右下維持
                        newL = _fullWindowRect.Right - newW;
                        newT = _fullWindowRect.Bottom - newH;
                        break;
                }

                // --- 画面外に出ないように補正 ---
                double vLeft = SystemParameters.VirtualScreenLeft;
                double vTop = SystemParameters.VirtualScreenTop;
                double vRight = vLeft + SystemParameters.VirtualScreenWidth;
                double vBottom = vTop + SystemParameters.VirtualScreenHeight;

                if (newL < vLeft) newL = vLeft;
                if (newT < vTop) newT = vTop;
                if (newL + newW > vRight) newL = vRight - newW;
                if (newT + newH > vBottom) newT = vBottom - newH;

                // フラグを立ててから移動・サイズ変更（OnLocationChangedでの誤更新防止）
                _isProgrammaticMove = true;
                try
                {
                    Left = newL;
                    Top = newT;
                    Width = newW;
                    Height = newH;
                }
                finally
                {
                    _isProgrammaticMove = false;
                }
                
                // スクロールバーや余白が見切れるのを防ぐため必要ならここでListViewの見た目を変える等の処理も可能
                // スクロールバーを非表示にする
                if (LstReservations != null)
                {
                    ScrollViewer.SetVerticalScrollBarVisibility(LstReservations, ScrollBarVisibility.Hidden);
                    ScrollViewer.SetHorizontalScrollBarVisibility(LstReservations, ScrollBarVisibility.Hidden);
                }
                
                _isMiniMode = true;
            }
            else
            {
                // 通常モードへ復帰（展開）
                // スクロールバーを自動に戻す
                if (LstReservations != null)
                {
                    ScrollViewer.SetVerticalScrollBarVisibility(LstReservations, ScrollBarVisibility.Auto);
                    ScrollViewer.SetHorizontalScrollBarVisibility(LstReservations, ScrollBarVisibility.Auto);
                }

                // 展開時も画面外補正を行う（ミニモード中に移動して、展開したらはみ出るケース防止）
                double restoreL = _fullWindowRect.Left;
                double restoreT = _fullWindowRect.Top;
                double restoreW = _fullWindowRect.Width;
                double restoreH = _fullWindowRect.Height;

                double vLeft = SystemParameters.VirtualScreenLeft;
                double vTop = SystemParameters.VirtualScreenTop;
                double vRight = vLeft + SystemParameters.VirtualScreenWidth;
                double vBottom = vTop + SystemParameters.VirtualScreenHeight;

                if (restoreL < vLeft) restoreL = vLeft;
                if (restoreT < vTop) restoreT = vTop;
                if (restoreL + restoreW > vRight) restoreL = vRight - restoreW;
                if (restoreT + restoreH > vBottom) restoreT = vBottom - restoreH;

                // フラグを立てて復元
                _isProgrammaticMove = true;
                try
                {
                    Left = restoreL;
                    Top = restoreT;
                    Width = restoreW;
                    Height = restoreH;
                }
                finally
                {
                    _isProgrammaticMove = false;
                }
                
                _isMiniMode = false;
            }
        }

        // ドラッグ移動時のフルサイズ位置更新（ミニモード中に移動した場合の補正）
        protected override void OnLocationChanged(EventArgs e)
        {
            base.OnLocationChanged(e);
            
            // プログラムによる移動または最大化/最小化中は無視
            if (_isProgrammaticMove || WindowState != WindowState.Normal) return;
            
            if (_isMiniMode)
            {
                // ミニモード中に移動した場合、_fullWindowRect も追従させる
                // 方向に応じて基準点が変わる
                
                double currentL = Left;
                double currentT = Top;

                switch (Config.Data.MiniModeDirection)
                {
                    case 0: // 左上固定
                        _fullWindowRect.X = currentL;
                        _fullWindowRect.Y = currentT;
                        break;
                    case 1: // 右上固定
                        _fullWindowRect.X = (currentL + Width) - _fullWindowRect.Width;
                        _fullWindowRect.Y = currentT;
                        break;
                    case 2: // 左下固定
                        _fullWindowRect.X = currentL;
                        _fullWindowRect.Y = (currentT + Height) - _fullWindowRect.Height;
                        break;
                    case 3: // 右下固定
                        _fullWindowRect.X = (currentL + Width) - _fullWindowRect.Width;
                        _fullWindowRect.Y = (currentT + Height) - _fullWindowRect.Height;
                        break;
                }
            }
            else
            {
                // 通常モード（ミニモード展開中含む）の移動を即座に反映
                // これによりドラッグ終了時や縮小時に正しい位置が計算されるようになります
                _fullWindowRect.X = Left;
                _fullWindowRect.Y = Top;
                
                _saveDebounceTimer.Stop();
                _saveDebounceTimer.Start();
                
                _miniModeTimer.Stop();
            }
        }
        
        // サイズ変更を即座に反映
        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);

            // プログラムによる変更や最大化/最小化中は無視
            if (_isProgrammaticMove || WindowState != WindowState.Normal) return;

            // 通常モード（展開中）の場合
            if (!_isMiniMode)
            {
                // 現在のサイズをフルサイズとして記録
                _fullWindowRect.Width = sizeInfo.NewSize.Width;
                _fullWindowRect.Height = sizeInfo.NewSize.Height;

                // サイズ変更中はタイマーをリセットし、変更終了後に保存させる
                _saveDebounceTimer.Stop();
                _saveDebounceTimer.Start();
                
                _miniModeTimer.Stop();
            }
        }
    }
}