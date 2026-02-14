using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Text;
using System.Linq;
using EpgTimer;

namespace EDCBMonitor
{
    public class ReserveItem : INotifyPropertyChanged
    {
        public ReserveData Data { get; private set; }
        // サーバーから取得した詳細情報を格納する場所
        public EpgTimer.EpgEventInfo? EventInfo { get; set; }
        
        // 全チューナーのIDと名前の辞書（ReservationServiceから渡される）
        public Dictionary<uint, string>? TunerNameMap { get; set; }

        public ReserveItem(ReserveData data)
        {
            Data = data;
            UpdateProgress();
        }

        // --- 基本プロパティ ---
        public uint ID => Data.ReserveID;
        public string Title => Data.Title ?? "";
        public string ServiceName => Data.StationName ?? "";
        public string NetworkName
        {
            get
            {
                var onid = Data.OriginalNetworkID;
                if (onid == 0x0004) return "BS";
                if (onid == 0x0006) return "CS1";
                if (onid == 0x0007) return "CS2";
                if (onid == 0x000A) return "スカパー";
                if (onid >= 0x7880 && onid <= 0x7FE8) return "地デジ";
                return onid == 0 ? "" : $"ONID:0x{onid:X4}";
            }
        }
        
        public string Comment
        {
            get
            {
                if (!string.IsNullOrEmpty(Data.Comment))
                {
                    return Data.Comment.Replace("EPG自動予約", "KW");
                }

                return Data.EventID == 0xFFFF ? "個別予約(プログラム)" : "個別予約(EPG)";
            }
        }

        // --- 番組内容 ---
        public string Desc => EventInfo?.ShortInfo?.text_char?.Replace("\r\n", " ").Replace("\n", " ") ?? "";

        // --- ジャンル ---
        public string Genre
        {
            get
            {
                if (EventInfo?.ContentInfo?.nibbleList == null) return "";
                var genreList = new List<string>();
                foreach (var nibble in EventInfo.ContentInfo.nibbleList)
                {
                    // 0x0E は属性情報なのでジャンルには含めない
                    if (nibble.content_nibble_level_1 == 0x0E && nibble.content_nibble_level_2 == 0x00) continue;

                    uint lv1 = (uint)nibble.content_nibble_level_1 << 24;
                    uint lv2 = (uint)nibble.content_nibble_level_2 << 16;
                    
                    if (GenreDefinition.Map.TryGetValue(lv1 | lv2, out string? detailName))
                    {
                        genreList.Add(detailName);
                    }
                    else if (GenreDefinition.Map.TryGetValue(lv1 | 0x00FF0000, out string? parentName))
                    {
                        genreList.Add(parentName);
                    }
                }
                return string.Join(", ", genreList.Distinct());
            }
        }

        // --- 付属情報 (ExtraInfo/Attrib) ---
        public string ExtraInfo
        {
            get
            {
                if (EventInfo == null) return "";
                var list = new List<string>();

                // イベントリレー
                if (EventInfo.EventRelayInfo != null && EventInfo.EventRelayInfo.eventDataList.Count > 0)
                {
                    list.Add("[イベントリレー]");
                }

                // 放送局付属情報 (0x0E)
                if (EventInfo.ContentInfo?.nibbleList != null)
                {
                    foreach (var nibble in EventInfo.ContentInfo.nibbleList)
                    {
                        if (nibble.content_nibble_level_1 == 0x0E && nibble.content_nibble_level_2 == 0x00)
                        {
                            string text = "";
                            switch (nibble.user_nibble_1)
                            {
                                case 0x00: text = "[編成情報]"; break;
                                case 0x01: text = "[中断情報]"; break;
                                case 0x02: text = "[3D映像]"; break;
                                case 0x0F: text = "[その他付属情報]"; break;
                            }
                            if (!string.IsNullOrEmpty(text) && !list.Contains(text))
                            {
                                list.Add(text);
                            }
                        }
                    }
                }
                return string.Join(",", list);
            }
        }

        // --- 有効/無効 ---
        public bool IsEnabled => Data.RecSetting?.IsEnable ?? false;
        public string IsEnabledBool => IsEnabled ? "有効" : "無効";
        public bool IsDisabled => !IsEnabled;
        
        // --- プログラム予約判定 ---
        public string ProgramType => Data.EventID == 0xFFFF ? "はい" : "いいえ";

        // --- 録画設定情報 ---
        public string RecMode
        {
            get
            {
                if (Data.RecSetting == null) return "";
                int recMode = Data.RecSetting.RecMode;
                int m = (recMode + (recMode / 5)) % 5;

                return m switch
                {
                    0 => "全サービス",
                    1 => "指定サービス",
                    2 => "全サービス(デコード処理なし)",
                    3 => "指定サービス(デコード処理なし)",
                    4 => "視聴",
                    _ => m.ToString()
                };
            }
        }

        public string Priority => Data.RecSetting?.Priority.ToString() ?? "";
        public string Tuijyuu => Data.RecSetting?.TuijyuuFlag == 1 ? "する" : "しない";
        public string Pittari => Data.RecSetting?.PittariFlag == 1 ? "する" : "しない";
        
        // --- 割り当てチューナー情報 ---
        public uint? AllocatedTunerID { get; set; }
        public string? AllocatedTunerName { get; set; }

        public string Tuner
        {
            get
            {
                // 予約が無効の場合は ID:FFFFFFFF (無効予約) と表示
                if (!IsEnabled) return "ID:FFFFFFFF (無効予約)";
                
                // 実際に割り当てられているチューナーがあれば、そのIDと名前(BonDriver)を表示
                if (AllocatedTunerID.HasValue)
                {
                    // 名前が空の場合はIDのみ
                    if (string.IsNullOrEmpty(AllocatedTunerName))
                        return $"ID:{AllocatedTunerID.Value:X8}";

                    // IDと名前を表示 (例: ID:00000001 (BonDriver_PT3-T...))
                    return $"ID:{AllocatedTunerID.Value:X8} ({AllocatedTunerName})";
                }

                // 割り当て情報が取れなかった場合
                if (Data.RecSetting == null) return "";
                if (Data.RecSetting.TunerID == 0) return "自動";
                return $"ID:{Data.RecSetting.TunerID:X8}";
            }
        }
        
        // チューナー強制: TunerNameMapを使用してIDから名前を解決する
        public string TunerForce
        {
            get
            {
                if (Data.RecSetting == null) return "";
                uint id = Data.RecSetting.TunerID;
                if (id == 0) return "自動";

                string text = $"ID:{id:X8}";

                // 辞書から名前を検索して付与する
                if (TunerNameMap != null && TunerNameMap.TryGetValue(id, out string? name) && !string.IsNullOrEmpty(name))
                {
                    text += $" ({name})";
                }
                // 辞書がない場合のフォールバック
                else if (!string.IsNullOrEmpty(AllocatedTunerName) && AllocatedTunerID == id)
                {
                    text += $" ({AllocatedTunerName})";
                }

                return text;
            }
        }
        
        // 録画後動作
        public string RecEndMode
        {
            get
            {
                if (Data.RecSetting == null) return "";
                if (Data.RecSetting.SuspendMode == 0)
                {
                    int defMode = GetCommonRecEndMode(); 
                    return "*" + ConvertSuspendModeText(defMode);
                }
                return ConvertSuspendModeText(Data.RecSetting.SuspendMode);
            }
        }

        // 復帰後再起動
        public string Reboot
        {
            get
            {
                if (Data.RecSetting == null) return "";
                string text = Data.RecSetting.RebootFlag == 1 ? "する" : "しない";
                return Data.RecSetting.SuspendMode == 0 ? "*" + text : text;
            }
        }

        // --- ヘルパーメソッド ---
        private string ConvertSuspendModeText(int mode) => mode switch
        {
            1 => "スタンバイ",
            2 => "休止",
            3 => "シャットダウン",
            4 => "何もしない",
            _ => "何もしない"
        };

        private int GetCommonRecEndMode()
        {
            string iniPath = PresetManager.Instance.GetIniPath();
            if (string.IsNullOrEmpty(iniPath)) return 1;
            return IniFileHandler.GetPrivateProfileInt("SET", "RecEndMode", 1, iniPath);
        }

        public string RecFolder => (Data.RecSetting?.RecFolderList?.Count > 0) ? Data.RecSetting.RecFolderList[0].RecFolder : "";
        public string RecFileName => (Data.RecFileNameList?.Count > 0) ? Data.RecFileNameList[0] : "";
        public string RecFileNameList => Data.RecFileNameList != null ? string.Join(", ", Data.RecFileNameList) : "";
        public string Preset => Data.RecSetting != null ? PresetManager.Instance.GetPresetName(Data.RecSetting) : "";
        public string Bat => Data.RecSetting?.BatFilePath ?? "";
        public string RecTag => Data.RecSetting?.RecTag ?? "";
        
        // 外部から代入されるためのプロパティ
        public int DefaultStartMargin { get; set; } = 5;
        public int DefaultEndMargin { get; set; } = 2;

        public string StartMargin
        {
            get
            {
                if (Data.RecSetting == null) return "";
                bool isDefault = Data.RecSetting.UseMargineFlag == 0;
                int val = isDefault ? DefaultStartMargin : Data.RecSetting.StartMargine;
                return FormatOffsetTime(-1 * val) + (isDefault ? "*" : "");
            }
        }

        public string EndMargin
        {
            get
            {
                if (Data.RecSetting == null) return "";
                bool isDefault = Data.RecSetting.UseMargineFlag == 0;
                int val = isDefault ? DefaultEndMargin : Data.RecSetting.EndMargine;
                return FormatOffsetTime(val) + (isDefault ? "*" : "");
            }
        }

        private string FormatOffsetTime(int seconds)
        {
            int abs = Math.Abs(seconds);
            return $"{(seconds >= 0 ? "+" : "-")}{abs / 60}:{abs % 60:D2}";
        }
        
        // --- エラー判定 ---
        public bool HasError => Data.OverlapMode != 0;
        public string ErrorInfo
        {
            get
            {
                if (Data.OverlapMode == 1) return "チューナー不足(一部)";
                if (Data.OverlapMode == 2) return "チューナー不足(不可)";
                return "";
            }
        }

        // --- ステータス (EDCB互換ロジック) ---
        static readonly string[] wiewString = { "", "", "無", "予+", "予+", "無+", "録*", "視*", "無*" };

        public string Status
        {
            get
            {
                int index = 0;
                if (Data != null)
                {
                    if (IsOnAir()) index = 3;
                    if (IsOnRec()) index = 6; 

                    if (!IsEnabled) index += 2;
                    else if (IsWatchMode) index += 1;
                }
                if (index >= 0 && index < wiewString.Length)
                    return wiewString[index];
                return "";
            }
        }

        public bool IsOnAir()
        {
            var now = DateTime.Now;
            return now >= Data.StartTime && now < Data.StartTime.AddSeconds(Data.DurationSecond);
        }

        public bool IsOnRec()
        {
            if (Data.RecSetting == null) return false;
            var now = DateTime.Now;
            int sm = (Data.RecSetting.UseMargineFlag == 0) ? DefaultStartMargin : Data.RecSetting.StartMargine;
            int em = (Data.RecSetting.UseMargineFlag == 0) ? DefaultEndMargin : Data.RecSetting.EndMargine;

            var start = Data.StartTime.AddSeconds(-sm);
            var end = Data.StartTime.AddSeconds(Data.DurationSecond + em);
            return now >= start && now < end;
        }
        
        public bool IsWatchMode => Data.RecSetting?.RecMode == 4;

        // --- 表示用プロパティ（UpdateProgressで更新） ---
        
        public void UpdateProgress()
        {
            OnPropertyChanged(nameof(Status));
            
            if (Data.RecSetting == null) return;
            var now = DateTime.Now;
            
            int sm = (Data.RecSetting.UseMargineFlag == 0) ? DefaultStartMargin : Data.RecSetting.StartMargine;
            int em = (Data.RecSetting.UseMargineFlag == 0) ? DefaultEndMargin : Data.RecSetting.EndMargine;

            var start = Data.StartTime.AddSeconds(-sm);
            var totalSec = Data.DurationSecond + sm + em;
            var end = start.AddSeconds(totalSec);

            bool isRec = (now >= start && now < end) && IsEnabled;
            IsRecording = isRec;

            if (isRec)
            {
                ProgressValue = totalSec > 0 ? ((now - start).TotalSeconds / totalSec) * 100.0 : 0;

                // 残り時間表示ロジック（1分未満切り上げ）
                if (!Config.Data.ShowRemainingTime)
                {
                    double remainSeconds = (end - now).TotalSeconds;
                    
                    // 切り上げ処理
                    int displayMinutes = (int)Math.Ceiling(remainSeconds / 60.0);
                    if (displayMinutes == 0 && remainSeconds > 0) displayMinutes = 1;

                    int h = displayMinutes / 60;
                    int m = displayMinutes % 60;

                    DurationHour = h.ToString();
                    DurationMinute = m.ToString("D2");
                    ColonOpacity = (now.Second % 2 == 0 ? 1.0 : 0.0);
                    
                    // プロパティ名を変更
                    DurationText = $"{h}:{m:D2}";
                }
                else
                {
                    TimeSpan ts = TimeSpan.FromSeconds(Data.DurationSecond);
                    SetStaticDuration(ts);
                }
            }
            
            else
            {
                // 終了時刻を過ぎていても有効ならリストから消えるまで録画中扱いにする
                if (now >= end && IsEnabled)
                {
                    IsRecording = true;
                    ProgressValue = 100;
                }
                else
                {
                    IsRecording = false;
                    ProgressValue = 0;
                }

                TimeSpan ts = TimeSpan.FromSeconds(Data.DurationSecond);
                SetStaticDuration(ts);
            }
            
        }

        private void SetStaticDuration(TimeSpan ts)
        {
            DurationHour = ((int)ts.TotalHours).ToString();
            DurationMinute = ts.Minutes.ToString("D2");
            ColonOpacity = 1.0;
            DurationText = $"{(int)ts.TotalHours}:{ts.Minutes:D2}";
        }

        private bool _isRecording;
        public bool IsRecording { get => _isRecording; set => SetProperty(ref _isRecording, value); }

        private double _progressValue;
        public double ProgressValue { get => _progressValue; set => SetProperty(ref _progressValue, value); }

        private string _durationHour = "0";
        public string DurationHour { get => _durationHour; set => SetProperty(ref _durationHour, value); }

        private string _durationMinute = "00";
        public string DurationMinute { get => _durationMinute; set => SetProperty(ref _durationMinute, value); }

        private double _colonOpacity = 1.0;
        public double ColonOpacity { get => _colonOpacity; set => SetProperty(ref _colonOpacity, value); }

        private string _durationText = "0:00";
        public string DurationText { get => _durationText; set => SetProperty(ref _durationText, value); }

        public string DateTimeInfo
        {
            get
            {
                var start = Data.StartTime;
                string dateFmt = "yyyy/MM/dd(ddd)";
                string timeFmt = "HH:mm";
                if (Config.Data.OmitYear) dateFmt = dateFmt.Replace("yyyy/", "");
                if (Config.Data.OmitMonth) dateFmt = dateFmt.Replace("MM/", "");
                string startStr = $"{start.ToString(dateFmt)} {start.ToString(timeFmt)}";
                string endStr = !Config.Data.OmitEndTime ? $"～{start.AddSeconds(Data.DurationSecond).ToString(timeFmt)}" : "";
                return startStr + endStr;
            }
        }

        // --- ツールチップ ---
        public string ToolTipText
        {
            get
            {
                var r = Data;
                var pgInfo = EventInfo;
                var endTime = r.StartTime.AddSeconds(r.DurationSecond);
                var sb = new StringBuilder();

                sb.AppendLine($"【番組名】 {r.Title ?? ""}");
                sb.AppendLine($"【日時】 {r.StartTime:MM/dd(ddd) HH:mm}～{endTime:HH:mm} ({FormatDuration((int)r.DurationSecond)})");
                sb.AppendLine($"【放送局】 {r.StationName ?? ""}");
                sb.AppendLine();

                string descText = pgInfo?.ShortInfo?.text_char ?? "";
                if (!string.IsNullOrEmpty(descText)) { sb.AppendLine(descText); sb.AppendLine(); }

                if (!string.IsNullOrEmpty(pgInfo?.ExtInfo?.text_char))
                {
                    sb.AppendLine("--------------------------------------------------");
                    sb.AppendLine(pgInfo.ExtInfo.text_char.Trim());
                    sb.AppendLine();
                }

                string genreStr = Genre;
                if (!string.IsNullOrEmpty(genreStr)) sb.AppendLine($"【ジャンル】 {genreStr}");

                return sb.ToString().TrimEnd();
            }
        }
        
        // 予想サイズ
        public string EstimatedSize
        {
            get
            {
                if (Data.RecSetting == null || Data.RecSetting.RecMode == 4 || Data.DurationSecond == 0) return "";

                int sm = (Data.RecSetting.UseMargineFlag == 0) ? DefaultStartMargin : Data.RecSetting.StartMargine;
                int em = (Data.RecSetting.UseMargineFlag == 0) ? DefaultEndMargin : Data.RecSetting.EndMargine;
                
                long totalSec = Data.DurationSecond + sm + em;
                if (totalSec <= 0) return ""; 

                double mbps = (Data.OriginalNetworkID >= 0x7880 && Data.OriginalNetworkID <= 0x7FEF) ? 17.0 : 24.0;
                long sizeBytes = (long)(mbps * 125000 * totalSec);
                
                return (sizeBytes > 1073741824)
                    ? $"{sizeBytes / 1073741824.0:F1} GB" 
                    : $"{sizeBytes / 1048576.0:F0} MB";
            }
        }
        private string FormatDuration(int totalSeconds)
        {
            TimeSpan ts = TimeSpan.FromSeconds(totalSeconds);
            return ts.TotalHours >= 1 ? $"{(int)ts.TotalHours}時間{ts.Minutes}分" : $"{ts.Minutes}分";
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void SetProperty<T>(ref T storage, T value, [CallerMemberName] string? name = null)
        {
            if (!Equals(storage, value))
            {
                storage = value;
                OnPropertyChanged(name);
            }
        }
        protected void OnPropertyChanged(string? name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}