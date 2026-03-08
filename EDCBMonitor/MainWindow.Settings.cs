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
        public void ApplySettings(bool updateSize = false)
        {
            try
            {
                if (updateSize)
                {
                    if (_isMiniMode)
                    {
                        // ミニモード中は実際のウィンドウサイズ(Width/Height)を直接書き換えない
                        // フルサイズに戻った時のために _fullWindowRect だけ更新しておく
                        _fullWindowRect = new Rect(Config.Data.Left, Config.Data.Top, Config.Data.Width, Config.Data.Height);
                    }
                    else
                    {
                        // ミニモードでない場合は通常通り設定から復元
                        if (Config.Data.Width > 0) Width = Config.Data.Width;
                        if (Config.Data.Height > 0) Height = Config.Data.Height;
                        Top = Config.Data.Top;
                        Left = Config.Data.Left;

                        // 初期Rectを保存
                        _fullWindowRect = new Rect(Left, Top, Width, Height);
                    }

                    // 設定フラグを見て状態を更新する
                    if (Config.Data.EnableMiniMode)
                    {
                        // 有効かつマウス外ならミニモードへ移行するが、
                        // 「上下最大化(IsVerticalMaximized)」で復元する場合は、
                        // 起動時に勝手に縮小せず、展開状態(最大化)のまま開始する。
                        if (!IsMouseOver && !Config.Data.IsVerticalMaximized)
                        {
                            UpdateMiniModeState(true);
                        }
                    }
                    else
                    {
                        // 無効化された場合、現在ミニモードなら復帰させる
                        if (_isMiniMode)
                        {
                            UpdateMiniModeState(false);
                        }
                    }
                }
                
                if (Config.Data.IsVerticalMaximized)
                {
                    _restoreBounds = new Rect(Left, Config.Data.RestoreTop, Width, Config.Data.RestoreHeight);
                }

                Topmost = Config.Data.Topmost;

                if (Config.Data.ShowTrayIcon)
                {
                    ShowInTaskbar = false;
                    if (_notifyIcon != null) _notifyIcon.Visible = true;
                }
                else
                {
                    ShowInTaskbar = true;
                    if (_notifyIcon != null) _notifyIcon.Visible = false;
                }

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
                
                if (Config.Data.ShowToolTip)
                {
                    itemStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, new System.Windows.Data.Binding("ToolTipText")));

                    // WPF標準の「勝手に表示する機能」を完全に機能停止させる
                    itemStyle.Setters.Add(new Setter(ToolTipService.IsEnabledProperty, false));

                    // 表示タイミングはすべて自前の統合メソッドに任せる
                    itemStyle.Setters.Add(new EventSetter(System.Windows.Controls.ListViewItem.SelectedEvent, new RoutedEventHandler(Item_Selected)));
                    itemStyle.Setters.Add(new EventSetter(System.Windows.Controls.ListViewItem.UnselectedEvent, new RoutedEventHandler(Item_Unselected)));
                    itemStyle.Setters.Add(new EventSetter(System.Windows.UIElement.MouseEnterEvent, new System.Windows.Input.MouseEventHandler(Item_MouseEnter)));
                    itemStyle.Setters.Add(new EventSetter(System.Windows.UIElement.MouseLeaveEvent, new System.Windows.Input.MouseEventHandler(Item_MouseLeave)));
                }

                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.BorderThicknessProperty, new Thickness(0)));
                itemStyle.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(-1, 0, 0, 0)));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.PaddingProperty, new Thickness(0)));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.MinHeightProperty, 0.0));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.HorizontalContentAlignmentProperty, System.Windows.HorizontalAlignment.Stretch));
                itemStyle.Setters.Add(new Setter(System.Windows.Controls.Control.VerticalContentAlignmentProperty, VerticalAlignment.Center));

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
                
                double mLeft = _isMiniMode ? Config.Data.MiniListMarginLeft : Config.Data.ListMarginLeft;
                double mTop = _isMiniMode ? Config.Data.MiniListMarginTop : Config.Data.ListMarginTop;
                double mRight = _isMiniMode ? Config.Data.MiniListMarginRight : Config.Data.ListMarginRight;
                double mBottom = _isMiniMode ? Config.Data.MiniListMarginBottom : Config.Data.ListMarginBottom;
                LstReservations.Margin = new Thickness(mLeft, mTop, mRight, mBottom);

                bool showHeader = _isMiniMode ? Config.Data.MiniShowHeader : Config.Data.ShowHeader;
                bool showFooter = _isMiniMode ? Config.Data.MiniShowFooter : Config.Data.ShowFooter;
                bool showListHeader = _isMiniMode ? Config.Data.MiniShowListHeader : Config.Data.ShowListHeader;

                RowHeader.Height = showHeader ? GridLength.Auto : new GridLength(0);
                RowFooter.Height = showFooter ? GridLength.Auto : new GridLength(0);

                // 枠組み自体を完全に非表示にし、ボタン等のはみ出し描画を防ぐ
                if (FindName("HeaderGrid") is Grid hGrid) hGrid.Visibility = showHeader ? Visibility.Visible : Visibility.Collapsed;
                if (FindName("FooterGrid") is Grid fGrid) fGrid.Visibility = showFooter ? Visibility.Visible : Visibility.Collapsed;

                MainBorder.BorderThickness = new Thickness(1);

                _columnManager.UpdateColumns();
                _columnManager.UpdateHeaderStyle(bgBrush, fgBrush, colBorderBrush, showListHeader);
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

        // 各イベントの入り口（すべて1つのコア処理に流す）
        private void Item_Selected(object sender, RoutedEventArgs e) => UpdateItemToolTip(sender as System.Windows.Controls.ListViewItem, e);
        private void Item_Unselected(object sender, RoutedEventArgs e) => UpdateItemToolTip(sender as System.Windows.Controls.ListViewItem, e);
        private void Item_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e) => UpdateItemToolTip(sender as System.Windows.Controls.ListViewItem, e);
        private void Item_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e) => UpdateItemToolTip(sender as System.Windows.Controls.ListViewItem, e);

        // ツールチップの表示状態を決定する統合メソッド
        private void UpdateItemToolTip(System.Windows.Controls.ListViewItem? item, RoutedEventArgs e)
        {
            if (item == null || !Config.Data.ShowToolTip || item.DataContext is not ReserveItem res) return;

            // リスト自動更新時の誤作動を完全にシャットアウトするストッパー
            // マウスが乗っていないのに発生した「選択／選択解除」イベントは、リストの裏側での更新処理によるものなので完全に無視する
            if (!item.IsMouseOver && (e.RoutedEvent == System.Windows.Controls.ListViewItem.SelectedEvent || e.RoutedEvent == System.Windows.Controls.ListViewItem.UnselectedEvent))
            {
                return;
            }

            if (!(item.ToolTip is System.Windows.Controls.ToolTip tt))
            {
                tt = new System.Windows.Controls.ToolTip { Style = Resources[typeof(System.Windows.Controls.ToolTip)] as Style };
                item.ToolTip = tt;
            }

            tt.DataContext = res;
            tt.Content = res.ToolTipText;

            bool shouldOpen = item.IsSelected && item.IsMouseOver;

            if (e.RoutedEvent == System.Windows.UIElement.MouseLeaveEvent || e.RoutedEvent == System.Windows.Controls.ListViewItem.UnselectedEvent)
            {
                shouldOpen = false;
            }

            if (System.Windows.Input.Mouse.RightButton == System.Windows.Input.MouseButtonState.Pressed || _isContextMenuOpen)
            {
                shouldOpen = false;
            }

            // 無駄に false を代入して WPF のフォーカスを乱さないように、すでに閉じているなら何もしない
            if (!shouldOpen)
            {
                if (tt.IsOpen) tt.IsOpen = false;
            }
            else
            {
                // 開く場合のみ、UIの準備を待ってから開く
                Dispatcher.BeginInvoke(new Action(() => 
                {
                    if (System.Windows.Input.Mouse.RightButton == System.Windows.Input.MouseButtonState.Released && !_isContextMenuOpen && item.IsMouseOver && item.IsSelected)
                    {
                        tt.IsOpen = true;
                    }
                }), DispatcherPriority.Render);
            }
        }

        // コピーやURLクリックができる専用ウィンドウを生成する
        private void ShowDetailWindow(ReserveItem item)
        {
            var brushConverter = new System.Windows.Media.BrushConverter();
            
            // ツールチップの設定から幅を取得（万が一無効な値だった場合は500を代入）
            double targetWidth = Config.Data.ToolTipWidth > 0 ? Config.Data.ToolTipWidth : 500;

            var window = new Window
            {
                Title = "番組詳細 - " + item.Data.Title,
                Width = targetWidth,                                // 幅をツールチップ設定に合わせる
                SizeToContent = SizeToContent.Height,               // 縦は文章量に合わせて自動で伸ばす
                MaxHeight = SystemParameters.WorkArea.Height * 0.9, // 画面外にはみ出さないようにストッパーをかける
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Owner = this,
                WindowStyle = WindowStyle.ToolWindow
            };

            // ツールチップの設定から背景色を読み込んで適用
            try
            {
                if (brushConverter.ConvertFromString(Config.Data.ToolTipBackColor) is SolidColorBrush bgBrush)
                    window.Background = bgBrush;
                else
                    window.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 30, 30));
            }
            catch { window.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 30, 30)); }

            var doc = new System.Windows.Documents.FlowDocument();
            var paragraph = new System.Windows.Documents.Paragraph();
            
            // ツールチップの設定からフォント、文字サイズ、文字色を読み込んで適用
            try { paragraph.FontFamily = new System.Windows.Media.FontFamily(Config.Data.FontFamily); } catch { }
            paragraph.FontSize = Config.Data.ToolTipFontSize > 0 ? Config.Data.ToolTipFontSize : 12;

            try
            {
                if (brushConverter.ConvertFromString(Config.Data.ToolTipForeColor) is SolidColorBrush fgBrush)
                    paragraph.Foreground = fgBrush;
                else
                    paragraph.Foreground = System.Windows.Media.Brushes.White;
            }
            catch { paragraph.Foreground = System.Windows.Media.Brushes.White; }

            string text = item.ToolTipText ?? "";
            
            // URLっぽい文字列を探し出すための正規表現
            var regex = new System.Text.RegularExpressions.Regex(@"(https?://[\w\-./%?:@&=+]+)");
            int lastPos = 0;
            
            foreach (System.Text.RegularExpressions.Match match in regex.Matches(text))
            {
                if (match.Index > lastPos)
                {
                    paragraph.Inlines.Add(new System.Windows.Documents.Run(text.Substring(lastPos, match.Index - lastPos)));
                }
                
                var link = new System.Windows.Documents.Hyperlink(new System.Windows.Documents.Run(match.Value))
                {
                    NavigateUri = new Uri(match.Value)
                };
                
                // リンクの色
                link.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 200, 255));

                link.RequestNavigate += (s, ev) => 
                {
                    try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(ev.Uri.AbsoluteUri) { UseShellExecute = true }); } catch { }
                    ev.Handled = true;
                };
                paragraph.Inlines.Add(link);
                
                lastPos = match.Index + match.Length;
            }
            
            if (lastPos < text.Length)
            {
                paragraph.Inlines.Add(new System.Windows.Documents.Run(text.Substring(lastPos)));
            }

            doc.Blocks.Add(paragraph);

            var viewer = new System.Windows.Controls.FlowDocumentScrollViewer 
            { 
                Document = doc,
                Margin = new Thickness(5),
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Background = System.Windows.Media.Brushes.Transparent
            };

            window.Content = viewer;
            window.Show();
        }

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
    }
}