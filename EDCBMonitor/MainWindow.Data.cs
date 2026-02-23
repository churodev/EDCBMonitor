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

            bool hasNewRecording = false;
            var currentRecordingIds = new HashSet<uint>();

            foreach (var item in list)
            {
                item.UpdateProgress(); 
                if (item.IsRecording)
                {
                    currentRecordingIds.Add(item.ID);
                    if (!_lastRecordingIds.Contains(item.ID))
                    {
                        hasNewRecording = true;
                    }
                }
            }

            _lastRecordingIds = currentRecordingIds;

            if (hasNewRecording)
            {
                var sv = GetScrollViewer(LstReservations);
                if (sv != null && LstReservations.Items.Count > 0)
                {
                    LstReservations.ScrollIntoView(LstReservations.Items[0]);
                    sv.ScrollToLeftEnd();
                }
            }
        }

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
                    await UpdateReservations();
                };
            }

            // パスの決定ロジック（ユーザー入力を柔軟に解釈）
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

        public async Task RefreshDataAsync() => await UpdateReservations();

        private async Task<bool> UpdateReservations(bool updateFooter = true)
        {
            try
            {
                PresetManager.Instance.Load();
                
                // ここでawaitするため時間がかかる可能性がある
                var list = await _reservationService.GetReservationsAsync();
                
                if (list == null)
                {
                    if (updateFooter && !_isShowingTempMessage) LblStatus.Text = "接続待機中...";
                    return false;
                }

                var selectedIds = new HashSet<uint>();
                if (LstReservations.SelectedItems != null)
                {
                    foreach (var item in LstReservations.SelectedItems.OfType<ReserveItem>()) selectedIds.Add(item.ID);
                }

                // 非同期処理中にユーザーやタイマーがスクロールした結果を打ち消さないよう
                // UI(ItemsSource)を差し替える「直前」の正確な座標を取得する
                var sv = GetScrollViewer(LstReservations);
                double currentVOffset = sv?.VerticalOffset ?? 0;
                double currentHOffset = sv?.HorizontalOffset ?? 0;

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

                // --- 録画中判定とスクロール制御（即時録画・リスト更新時用） ---
                bool hasNewRecording = false;
                var currentRecordingIds = new HashSet<uint>();
                
                foreach (var item in list)
                {
                    if (item.IsRecording)
                    {
                        currentRecordingIds.Add(item.ID);
                        if (!_lastRecordingIds.Contains(item.ID))
                        {
                            hasNewRecording = true;
                        }
                    }
                }
                _lastRecordingIds = currentRecordingIds;

                if (sv != null)
                {
                    LstReservations.UpdateLayout();

                    if (hasNewRecording)
                    {
                        if (LstReservations.Items.Count > 0)
                        {
                            LstReservations.ScrollIntoView(LstReservations.Items[0]);
                        }
                        sv.ScrollToLeftEnd();
                    }
                    else
                    {
                        sv.ScrollToVerticalOffset(currentVOffset);
                        sv.ScrollToHorizontalOffset(currentHOffset);
                    }
                }

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
    }
}