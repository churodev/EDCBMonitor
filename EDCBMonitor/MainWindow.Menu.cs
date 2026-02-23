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

        // メニューが表示される直前に呼ばれる。ここでフラグを立てれば縮小を防げる。
        private void Window_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            _isContextMenuOpen = true;
        }

        private void ContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.ContextMenu menu)
            {
                menu.Closed -= ContextMenu_Closed;
                menu.Closed += ContextMenu_Closed;

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

        // メニューが閉じた時の処理
        private void ContextMenu_Closed(object sender, RoutedEventArgs e)
        {
            _isContextMenuOpen = false;
            if (sender is System.Windows.Controls.ContextMenu menu)
            {
                menu.Closed -= ContextMenu_Closed;
            }

            // メニューを閉じた時点でマウスが既に外にあるなら縮小処理を開始する
            // (これをしないと、メニューを閉じた後に縮小されなくなる)
            if (!IsMouseOver)
            {
                Window_MouseLeave(this, null);
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

        // 番組詳細ウィンドウを開くメニュー
        private void MenuShowDetail_Click(object sender, RoutedEventArgs e)
        {
            if (LstReservations.SelectedItem is ReserveItem res)
            {
                ShowDetailWindow(res);
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

    }
}