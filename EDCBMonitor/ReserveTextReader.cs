#nullable disable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
using System.Reflection;
using EpgTimer;

namespace EDCBMonitor
{
    public class ReserveItem
    {
        public uint ID { get; set; }
        public string Status { get; set; } = "";
        public ushort EventID { get; set; } 
        public string StartTime { get; set; }
        public string DateTimeInfo { get; set; } = "";
        public string Duration { get; set; } = "";
        public string NetworkName { get; set; } = "";
        public string ServiceName { get; set; } = "";
        public string Title { get; set; } = "";
        public string Desc { get; set; } = "";
        public string Genre { get; set; } = "";
        public string ExtraInfo { get; set; } = "";
        public string Enabled { get; set; } = "";
        public string ProgramType { get; set; } = "";
        public string Comment { get; set; } = "";
        public string ErrorInfo { get; set; } = "";
        public bool HasError => !string.IsNullOrEmpty(ErrorInfo);
        public string RecFileName { get; set; } = "";
        public string RecFileNameList { get; set; } = "";
        public string Tuner { get; set; } = "";
        public string EstSize { get; set; } = "";
        public string Preset { get; set; } = "";
        public string RecMode { get; set; } = "";
        public string Priority { get; set; } = "";
        public string Tuijyuu { get; set; } = "";
        public string Pittari { get; set; } = "";
        public string TunerForce { get; set; } = "";
        public string RecEndMode { get; set; } = "";
        public string Reboot { get; set; } = "";
        public string Bat { get; set; } = "";
        public string RecTag { get; set; } = "";
        public string RecFolder { get; set; } = "";
        public string StartMargin { get; set; } = "";
        public string EndMargin { get; set; } = "";

        public bool IsEnabledBool { get; set; }
        public DateTime StartTimeRaw { get; set; }
        public bool IsRecording { get; set; }
        public bool IsDisabled { get; set; }
    }

    public class RecPresetItem
    {
        public string DisplayName { get; set; }
        public int ID { get; set; }
        public RecSettingData Data { get; set; }
    }

    public class TunerReserveInfo : ICtrlCmdReadWrite
    {
        public uint TunerID;
        public string TunerName = "";
        public List<uint> ReserveList = new List<uint>();

        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref TunerID);
            r.Read(ref TunerName);
            r.Read(ref ReserveList);
            r.End();
        }
        public void Write(MemoryStream s, ushort version) { }
    }

    public static class ReserveTextReader
    {
        static ReserveTextReader()
        {
            try
            {
                var type = Type.GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages");
                if (type != null)
                {
                    var instanceProp = type.GetProperty("Instance", BindingFlags.Public | BindingFlags.Static);
                    var provider = instanceProp?.GetValue(null) as EncodingProvider;
                    if (provider != null) Encoding.RegisterProvider(provider);
                }
            }
            catch { }
        }

        public static string GetReserveFilePath()
        {
            if (!string.IsNullOrEmpty(Config.Data.EdcbInstallPath))
            {
                if (File.Exists(Config.Data.EdcbInstallPath) && Config.Data.EdcbInstallPath.EndsWith("Reserve.txt", StringComparison.OrdinalIgnoreCase)) return Config.Data.EdcbInstallPath;
                string p1 = Path.Combine(Config.Data.EdcbInstallPath, "Setting", "Reserve.txt"); if (File.Exists(p1)) return p1;
                string p2 = Path.Combine(Config.Data.EdcbInstallPath, "Reserve.txt"); if (File.Exists(p2)) return p2;
            }
            string local = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reserve.txt"); if (File.Exists(local)) return local;
            string setDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Setting", "Reserve.txt"); if (File.Exists(setDir)) return setDir;
            return "";
        }

        private static string GetEpgTimerSrvIniPath()
        {
            string reservePath = GetReserveFilePath();
            if (string.IsNullOrEmpty(reservePath)) return "";
            string dir = Path.GetDirectoryName(reservePath);
            if (string.IsNullOrEmpty(dir)) return "";

            DirectoryInfo parentInfo = Directory.GetParent(dir);
            if (parentInfo != null)
            {
                string pParent = Path.Combine(parentInfo.FullName, "EpgTimerSrv.ini");
                if (File.Exists(pParent)) return pParent;
            }
            string pSame = Path.Combine(dir, "EpgTimerSrv.ini");
            if (File.Exists(pSame)) return pSame;
            return "";
        }

        private static int GetDefaultRecEndMode()
        {
            try
            {
                string iniPath = GetEpgTimerSrvIniPath();
                if (string.IsNullOrEmpty(iniPath) || !File.Exists(iniPath)) return 1;
                var ini = LoadIni(iniPath);
                if (ini.ContainsKey("SET") && ini["SET"].ContainsKey("RecEndMode"))
                {
                    if (int.TryParse(ini["SET"]["RecEndMode"], out int val)) return val;
                }
            } catch { }
            return 1;
        }

        private static void GetDefaultMargins(out int startMargin, out int endMargin)
        {
            startMargin = 5; 
            endMargin = 2;
            
            string iniPath = GetEpgTimerSrvIniPath();
            if (string.IsNullOrEmpty(iniPath) || !File.Exists(iniPath)) 
            {
                return;
            }
            
            try
            {
                var ini = LoadIni(iniPath);
                if (ini.ContainsKey("SET"))
                {
                    if (ini["SET"].ContainsKey("StartMargin")) int.TryParse(ini["SET"]["StartMargin"], out startMargin);
                    if (ini["SET"].ContainsKey("EndMargin")) int.TryParse(ini["SET"]["EndMargin"], out endMargin);
                }
            }
            catch (Exception ex)
            {
                Logger.Write("INI読み込みエラー: " + ex.Message);
            }
        }

        private static Dictionary<string, Dictionary<string, string>> LoadIni(string path)
        {
            var data = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(path)) return data;

            try
            {
                Encoding enc = Encoding.Default;
                try 
                { 
                    enc = Encoding.GetEncoding(932); 
                } 
                catch { }

                string currentSection = "";
                foreach (var line in File.ReadLines(path, enc))
                {
                    string t = line.Trim();
                    if (t.Length == 0 || t.StartsWith(";") || t.StartsWith("//")) continue;

                    if (t.StartsWith("[") && t.EndsWith("]"))
                    {
                        currentSection = t.Substring(1, t.Length - 2).Trim();
                        if (!data.ContainsKey(currentSection)) data[currentSection] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    }
                    else if (!string.IsNullOrEmpty(currentSection))
                    {
                        int idx = t.IndexOf('=');
                        if (idx > 0)
                        {
                            string key = t.Substring(0, idx).Trim();
                            string val = t.Substring(idx + 1).Trim();
                            data[currentSection][key] = val;
                        }
                    }
                }
            }
            catch { }
            return data;
        }

        private static List<RecPresetItem> LoadPresets()
        {
            var list = new List<RecPresetItem>();
            string iniPath = GetEpgTimerSrvIniPath();
            if (string.IsNullOrEmpty(iniPath)) return list;

            var ini = LoadIni(iniPath);
            list.Add(LoadPresetItem(ini, 0));

            if (ini.ContainsKey("SET") && ini["SET"].ContainsKey("PresetID"))
            {
                string[] ids = ini["SET"]["PresetID"].Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string sId in ids)
                {
                    if (int.TryParse(sId, out int id) && id != 0)
                    {
                        list.Add(LoadPresetItem(ini, id));
                    }
                }
            }
            return list;
        }

        private static RecPresetItem LoadPresetItem(Dictionary<string, Dictionary<string, string>> ini, int id)
        {
            string ids = id == 0 ? "" : id.ToString();
            string secName = "REC_DEF" + ids;
            string secFolder = "REC_DEF_FOLDER" + ids;
            string sec1Seg = "REC_DEF_FOLDER_1SEG" + ids;

            var item = new RecPresetItem { ID = id, Data = new RecSettingData() };
            var d = item.Data;

            Func<string, string, string, string> GetStr = (sec, key, def) => 
                (ini.ContainsKey(sec) && ini[sec].ContainsKey(key)) ? ini[sec][key] : def;
            Func<string, string, int, int> GetInt = (sec, key, def) => 
                (ini.ContainsKey(sec) && ini[sec].ContainsKey(key) && int.TryParse(ini[sec][key], out int v)) ? v : def;

            item.DisplayName = GetStr(secName, "SetName", "デフォルト");

            int rawRecMode = GetInt(secName, "RecMode", 1);
            bool isEnable = (rawRecMode / 5 % 2 == 0);
            if (isEnable) d.RecMode = (byte)(rawRecMode % 5);
            else { int noRecMode = GetInt(secName, "NoRecMode", 1); d.RecMode = (byte)((noRecMode % 5) + 5); }

            d.Priority = (byte)GetInt(secName, "Priority", 2);
            d.TuijyuuFlag = (byte)GetInt(secName, "TuijyuuFlag", 1);
            d.ServiceMode = (uint)GetInt(secName, "ServiceMode", 16);
            d.PittariFlag = (byte)GetInt(secName, "PittariFlag", 0);
            
            string packedBat = GetStr(secName, "BatFilePath", "");
            int sep = packedBat.IndexOf('*');
            if (sep >= 0) { d.BatFilePath = packedBat.Substring(0, sep); d.RecTag = packedBat.Substring(sep + 1); }
            else { d.BatFilePath = packedBat; d.RecTag = ""; }

            Action<string, List<RecFileSetInfo>> LoadFolders = (sec, targetList) => {
                int count = GetInt(sec, "Count", 0);
                for (int i = 0; i < count; i++)
                {
                    var info = new RecFileSetInfo();
                    info.RecFolder = GetStr(sec, i.ToString(), "");
                    if (!string.IsNullOrEmpty(info.RecFolder) && !info.RecFolder.EndsWith("\\")) info.RecFolder += "\\";
                    
                    info.WritePlugIn = GetStr(sec, "WritePlugIn" + i, "Write_Default.dll");
                    info.RecNamePlugIn = GetStr(sec, "RecNamePlugIn" + i, "");
                    targetList.Add(info);
                }
            };

            LoadFolders(secFolder, d.RecFolderList);
            LoadFolders(sec1Seg, d.PartialRecFolder);

            d.SuspendMode = (byte)GetInt(secName, "SuspendMode", 0);
            d.RebootFlag = (byte)GetInt(secName, "RebootFlag", 0);
            d.UseMargineFlag = (byte)GetInt(secName, "UseMargineFlag", 0);
            d.StartMargine = GetInt(secName, "StartMargine", 5);
            d.EndMargine = GetInt(secName, "EndMargine", 2);
            d.ContinueRecFlag = (byte)GetInt(secName, "ContinueRec", 0);
            d.PartialRecFlag = (byte)GetInt(secName, "PartialRec", 0);
            d.TunerID = (uint)GetInt(secName, "TunerID", 0);

            return item;
        }

        private static string FormatOffsetTime(int seconds)
        {
            string sign = seconds >= 0 ? "+" : "-";
            int abs = Math.Abs(seconds);
            return $"{sign}{abs / 60}:{abs % 60:D2}";
        }

        private static string FormatDuration(int seconds) { TimeSpan ts = TimeSpan.FromSeconds(seconds); return $"{ts.Hours + ts.Days * 24}:{ts.Minutes:D2}"; }
        
        private static string GetNetworkName(int onid)
        {
            return onid switch {
                >= 0x7880 and <= 0x7FEF => "地デジ", 4 => "BS", 6 or 7 => "CS", 10 => "CS", 1 or 3 => "スカパー！", _ => $"ONID:{onid}"
            };
        }

        private static string GetRecEndModeText(int value, int defaultMode)
        {
            if (value == 0) {
                string defText = defaultMode switch { 0 => "何もしない", 1 => "スタンバイ", 2 => "休止", 3 => "シャットダウン", 4 => "何もしない", _ => "不明" };
                return $"*{defText}";
            }
            return value switch { 1 => "スタンバイ", 2 => "休止", 3 => "シャットダウン", 4 => "何もしない", _ => value.ToString() };
        }

        private static string GetRecModeText(int recMode)
        {
            int m = (recMode + (recMode / 5)) % 5;
            return m switch { 0 => "全サービス", 1 => "指定サービス", 2 => "全サービス(デコード処理なし)", 3 => "指定サービス(デコード処理なし)", 4 => "視聴", _ => m.ToString() };
        }

        private static string GetYesNo(byte flag) => flag == 1 ? "する" : "しない";

        private static string SanitizeText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();
        }

        private static string CleanTitle(string title)
        {
            if (string.IsNullOrEmpty(title)) return "";
            string t = title;
            if (Config.Data.EnableTitleRemove)
            {
                try { t = Regex.Replace(t, Config.Data.TitleRemovePattern, ""); } catch { }
            }
            return SanitizeText(t);
        }

        private static string GetGenreText(EpgContentInfo info)
        {
            if (info == null || info.nibbleList == null || info.nibbleList.Count == 0) return "";

            var list = new List<string>();
            foreach (var data in info.nibbleList)
            {
                if (data.content_nibble_level_1 == 0x0E && data.content_nibble_level_2 == 0x00) continue;

                uint keyFull = ((uint)data.content_nibble_level_1 << 24) | ((uint)data.content_nibble_level_2 << 16) | ((uint)data.user_nibble_1 << 8) | data.user_nibble_2;
                
                // 変更: GenreDict -> GenreDefinition.Map
                if (GenreDefinition.Map.ContainsKey(keyFull))
                {
                    list.Add(GenreDefinition.Map[keyFull]);
                    continue;
                }

                if (data.content_nibble_level_1 != 0x0E)
                {
                    uint keySub = ((uint)data.content_nibble_level_1 << 24) | ((uint)data.content_nibble_level_2 << 16);
                    if (GenreDefinition.Map.ContainsKey(keySub))
                    {
                        list.Add(GenreDefinition.Map[keySub]);
                        continue;
                    }
                }

                uint keyMajor = ((uint)data.content_nibble_level_1 << 24) | 0x00FF0000;
                if (GenreDefinition.Map.ContainsKey(keyMajor))
                {
                    list.Add(GenreDefinition.Map[keyMajor]);
                }
            }
            return string.Join(" ", list.Distinct());
        }

        public static List<ReserveItem> Load()
        {
            var list = new List<ReserveItem>();
            var cmd = new CtrlCmdUtil();
            cmd.SetSendMode(false);
            cmd.SetPipeSetting("Global\\EpgTimerSrvConnect", "EpgTimerSrvPipe");

            var reserveDataList = new List<ReserveData>();
            ErrCode err = cmd.SendEnumReserve(ref reserveDataList);
            // 通信エラー時はnullを返す
            if (err != ErrCode.CMD_SUCCESS) return null;
            
            var tunerReserveList = new List<TunerReserveInfo>();
            SendEnumTunerReserve(cmd, ref tunerReserveList);

            var resIdToTuner = new Dictionary<uint, TunerReserveInfo>();
            foreach(var tri in tunerReserveList)
            {
                foreach(var resId in tri.ReserveList)
                {
                    if (!resIdToTuner.ContainsKey(resId)) resIdToTuner[resId] = tri;
                }
            }

            int defRecEndMode = GetDefaultRecEndMode();
            GetDefaultMargins(out int defStartMargin, out int defEndMargin);
            var presetList = LoadPresets();

            if (err == ErrCode.CMD_SUCCESS)
            {
                foreach (var r in reserveDataList)
                {
                    bool isDisabled = false;
                    bool isRecording = false;
                    string status = "";

                    if (r.RecSetting != null) isDisabled = r.RecSetting.IsNoRec();
                    
                    DateTime endTime = r.StartTime.AddSeconds(r.DurationSecond);
                    if (DateTime.Now >= r.StartTime && DateTime.Now < endTime && !isDisabled) { isRecording = true; status = "録*"; }
                    else if (isDisabled) { status = "無"; }

                    if (Config.Data.HideDisabled && isDisabled) continue;

                    string dateTimeInfo = $"{r.StartTime:MM/dd(ddd) HH:mm}～{endTime:HH:mm}";
                    string recMode="", priority="", tuijyuu="", pittari="", tunerForce="", tunerAppointed="", recEndMode="", reboot="", bat="", tag="", folder="", startMargin="", endMargin="", estSize="", presetName="";

                    if (r.RecSetting != null)
                    {
                        recMode = GetRecModeText(r.RecSetting.GetRecMode());
                        priority = r.RecSetting.Priority.ToString();
                        tuijyuu = GetYesNo(r.RecSetting.TuijyuuFlag);
                        pittari = GetYesNo(r.RecSetting.PittariFlag);
                        
                        if (r.RecSetting.TunerID == 0)
                        {
                            tunerForce = "自動";
                        }
                        else
                        {
                            var matchedTuner = tunerReserveList.FirstOrDefault(t => t.TunerID == r.RecSetting.TunerID);
                            if (matchedTuner != null && !string.IsNullOrEmpty(matchedTuner.TunerName))
                            {
                                tunerForce = $"ID:{r.RecSetting.TunerID:X8} ({matchedTuner.TunerName})";
                            }
                            else
                            {
                                tunerForce = $"ID:{r.RecSetting.TunerID:X8}";
                            }
                        }

                        if (isDisabled)
                        {
                            tunerAppointed = "ID:FFFFFFFF (無効予約)";
                        }
                        else if (resIdToTuner.TryGetValue(r.ReserveID, out var assignedTuner))
                        {
                            tunerAppointed = $"ID:{assignedTuner.TunerID:X8} ({assignedTuner.TunerName})";
                        }
                        else
                        {
                            if (r.RecSetting.TunerID != 0)
                            {
                                var matchedTuner = tunerReserveList.FirstOrDefault(t => t.TunerID == r.RecSetting.TunerID);
                                if (matchedTuner != null && !string.IsNullOrEmpty(matchedTuner.TunerName))
                                {
                                    tunerAppointed = $"ID:{r.RecSetting.TunerID:X8} ({matchedTuner.TunerName})";
                                }
                                else
                                {
                                    tunerAppointed = $"ID:{r.RecSetting.TunerID:X8}";
                                }
                            }
                            else
                            {
                                tunerAppointed = "自動";
                            }
                        }

                        recEndMode = GetRecEndModeText(r.RecSetting.SuspendMode, defRecEndMode);
                        reboot = GetYesNo(r.RecSetting.RebootFlag);
                        bat = r.RecSetting.BatFilePath;
                        tag = r.RecSetting.RecTag;
                        if (r.RecSetting.RecFolderList != null && r.RecSetting.RecFolderList.Count > 0) folder = r.RecSetting.RecFolderList[0].RecFolder;
                        
                        int sm = 0, em = 0;
                        bool isDefault = false;
                        if (r.RecSetting.UseMargineFlag == 0) { sm = defStartMargin; em = defEndMargin; isDefault = true; }
                        else { sm = r.RecSetting.StartMargine; em = r.RecSetting.EndMargine; }
                        startMargin = FormatOffsetTime(-1 * sm) + (isDefault ? "*" : "");
                        endMargin = FormatOffsetTime(em) + (isDefault ? "*" : "");
                        
                        double mbps = (r.OriginalNetworkID >= 0x7880 && r.OriginalNetworkID <= 0x7FEF) ? 17.0 : 24.0;
                        long sizeBytes = (long)(mbps * 1000 * 1000 / 8 * r.DurationSecond);
                        estSize = (sizeBytes > 1024 * 1024 * 1024) ? $"{sizeBytes / (1024.0 * 1024.0 * 1024.0):F1} GB" : $"{sizeBytes / (1024.0 * 1024.0):F0} MB";

                        var matched = presetList.FirstOrDefault(p => r.RecSetting.EqualsSettingTo(p.Data));
                        presetName = matched != null ? matched.DisplayName : "登録時";
                    }

                    // --- エラー判定ロジック ---
                    var errList = new List<string>();
                    if (r.OverlapMode == 1) errList.Add("チューナー不足(一部録画)");
                    else if (r.OverlapMode == 2) errList.Add("チューナー不足(録画不可)");

                    if (r.RecSetting != null && r.RecSetting.UseMargineFlag == 1)
                    {
                        int actualDuration = (int)r.DurationSecond + r.RecSetting.StartMargine + r.RecSetting.EndMargine;
                        if (actualDuration <= 0) errList.Add("不可マージン設定(録画不可)");
                    }
                    string errorInfo = string.Join(" ", errList);
                    // --------------------------

                    string programType = (r.EventID == 0xFFFF) ? "はい" : "いいえ";
                    string recFileName = (r.RecFileNameList != null && r.RecFileNameList.Count > 0) ? r.RecFileNameList[0] : "";
                    string recFileNameList = (r.RecFileNameList != null) ? string.Join(", ", r.RecFileNameList) : "";
                    string desc = "";
                    string genre = "";
                    string extraInfo = "";

                    try
                    {
                        ulong key = ((ulong)r.OriginalNetworkID << 48) | ((ulong)r.TransportStreamID << 32) | ((ulong)r.ServiceID << 16) | (ulong)r.EventID;
                        var pgInfo = new EpgEventInfo();
                        if (cmd.SendGetPgInfo(key, ref pgInfo) == ErrCode.CMD_SUCCESS)
                        {
                            if (pgInfo.ShortInfo != null) desc = pgInfo.ShortInfo.text_char;
                            if (pgInfo.ContentInfo != null) genre = GetGenreText(pgInfo.ContentInfo);
                            
                            // 付属情報 (ConvertAttribText相当)
                            var extList = new List<string>();
                            if (pgInfo.EventRelayInfo != null && pgInfo.EventRelayInfo.eventDataList.Count > 0) extList.Add("[イベントリレー]");
                            
                            if (pgInfo.ContentInfo != null && pgInfo.ContentInfo.nibbleList != null)
                            {
                                var attrs = pgInfo.ContentInfo.nibbleList
                                    .Where(x => x.content_nibble_level_1 == 0x0E && x.content_nibble_level_2 == 0x00)
                                    .Select(x => x.user_nibble_1)
                                    .Distinct();

                                foreach (var u1 in attrs)
                                {
                                    switch (u1)
                                    {
                                        case 0x00: extList.Add("[編成情報]"); break;
                                        case 0x01: extList.Add("[中断情報]"); break;
                                        case 0x02: extList.Add("[3D映像]"); break;
                                        case 0x0F: extList.Add("[その他付属情報]"); break;
                                        default:   extList.Add("[不明な情報]"); break;
                                    }
                                }
                            }
                            extraInfo = string.Join(",", extList);
                        }
                    }
                    catch { }

                    // 1. 予約種類の判定
                    bool isEpgReserve = (r.EventID != 0xFFFF);
                    bool isAutoAdded = (!string.IsNullOrEmpty(r.Comment) && !r.Comment.EndsWith("$"));

                    string typeStr = "";
                    if (isAutoAdded) { if (isEpgReserve) typeStr = "KW"; else typeStr = "自動(PG)"; }
                    else { if (isEpgReserve) typeStr = "個別(EPG)"; else typeStr = "個別(PG)"; }

                    // 2. 表示用文字列の生成
                    string rawComment = SanitizeText(r.Comment);
                    string bodyComment = rawComment
                        .Replace("EPG自動予約", "")
                        .Replace("キーワード予約", "")
                        .Replace("自動予約", "")
                        .Replace("EPG予約", "")
                        .Replace("予約", "")
                        .Trim();

                    string comment;
                    if (!isAutoAdded) comment = typeStr;
                    else comment = string.IsNullOrEmpty(bodyComment) ? typeStr : $"{typeStr}{bodyComment}";

                    list.Add(new ReserveItem {
                        ID = r.ReserveID,
                        Status = status,
                        DateTimeInfo = dateTimeInfo,
                        Duration = FormatDuration((int)r.DurationSecond),
                        NetworkName = GetNetworkName(r.OriginalNetworkID),
                        ServiceName = r.StationName,
                        Title = CleanTitle(r.Title),
                        Desc = SanitizeText(desc),
                        Genre = SanitizeText(genre),
                        ExtraInfo = SanitizeText(extraInfo),
                        Enabled = isDisabled ? "無効" : "有効",
                        IsEnabledBool = !isDisabled,
                        ProgramType = programType,
                        Comment = comment,
                        ErrorInfo = errorInfo, // エラー情報
                        RecFileName = SanitizeText(recFileName),
                        RecFileNameList = SanitizeText(recFileNameList),
                        Tuner = tunerAppointed,
                        EstSize = estSize,
                        Preset = presetName,
                        RecMode = recMode,
                        Priority = priority,
                        Tuijyuu = tuijyuu,
                        Pittari = pittari,
                        TunerForce = tunerForce,
                        RecEndMode = recEndMode,
                        Reboot = reboot,
                        Bat = SanitizeText(bat),
                        RecTag = SanitizeText(tag),
                        RecFolder = SanitizeText(folder),
                        StartMargin = startMargin,
                        EndMargin = endMargin,
                        StartTimeRaw = r.StartTime,
                        IsRecording = isRecording,
                        IsDisabled = isDisabled,
                        EventID = r.EventID // 重複チェック用
                    });
                }

                // 重複予約のチェックと追記
                var duplicates = list
                    .Where(x => x.EventID != 0xFFFF)
                    .GroupBy(x => x.EventID)
                    .Where(g => g.Count() > 1)
                    .SelectMany(g => g)
                    .ToList();

                foreach (var item in duplicates)
                {
                    if (!string.IsNullOrEmpty(item.ErrorInfo)) item.ErrorInfo += " 重複したEPG予約";
                    else item.ErrorInfo = "重複したEPG予約";
                }

                list.Sort((a, b) => a.StartTimeRaw.CompareTo(b.StartTimeRaw));
            }
            return list;
        }

        private static ErrCode SendEnumTunerReserve(CtrlCmdUtil cmd, ref List<TunerReserveInfo> val)
        {
            try
            {
                var method = typeof(CtrlCmdUtil).GetMethod("SendCmdStream", BindingFlags.NonPublic | BindingFlags.Instance);
                if (method != null)
                {
                    using (var ms = new MemoryStream()) 
                    {
                        object[] args = new object[] { (CtrlCmd)1016, ms, null };
                        ErrCode ret = (ErrCode)method.Invoke(cmd, args);
                        var res = args[2] as MemoryStream;
                        if (ret == ErrCode.CMD_SUCCESS && res != null)
                        {
                            var r = new CtrlCmdReader(res);
                            r.Read(ref val);
                        }
                        return ret;
                    }
                }
            }
            catch { }
            return ErrCode.CMD_ERR;
        }
    }

    public static class RecSettingDataExtensions
    {
        public static bool EqualsTo(this List<RecFileSetInfo> list, List<RecFileSetInfo> other)
        {
            if (list == null || other == null) return list == other;
            if (list.Count != other.Count) return false;
            for (int i = 0; i < list.Count; i++)
            {
                var a = list[i];
                var b = other[i];
                if (!string.Equals(a.RecFolder, b.RecFolder, StringComparison.OrdinalIgnoreCase)) return false;
                if (!string.Equals(a.WritePlugIn, b.WritePlugIn, StringComparison.OrdinalIgnoreCase)) return false;
                if (!string.Equals(a.RecNamePlugIn, b.RecNamePlugIn, StringComparison.OrdinalIgnoreCase)) return false;
            }
            return true;
        }

        public static bool EqualsSettingTo(this RecSettingData self, RecSettingData other)
        {
            if (other == null) return false;
            int selfMode = (self.RecMode + (self.RecMode / 5)) % 5;
            int otherMode = (other.RecMode + (other.RecMode / 5)) % 5;            
            if (selfMode != otherMode) return false;
            if (self.Priority != other.Priority) return false;
            if (self.TuijyuuFlag != other.TuijyuuFlag) return false;
            if (!(self.ServiceMode == other.ServiceMode || ((self.ServiceMode | other.ServiceMode) & 0x0F) == 0)) return false;
            if (self.PittariFlag != other.PittariFlag) return false;
            if (!string.Equals(self.BatFilePath, other.BatFilePath, StringComparison.OrdinalIgnoreCase)) return false;
            if (!self.RecFolderList.EqualsTo(other.RecFolderList)) return false;
            if (!self.PartialRecFolder.EqualsTo(other.PartialRecFolder)) return false;
            if (self.SuspendMode != other.SuspendMode) return false;
            if (self.SuspendMode != 0 && self.RebootFlag != other.RebootFlag) return false;
            if (self.UseMargineFlag != other.UseMargineFlag) return false;
            if (self.UseMargineFlag != 0)
            {
                if (self.StartMargine != other.StartMargine) return false;
                if (self.EndMargine != other.EndMargine) return false;
            }
            if (self.ContinueRecFlag != other.ContinueRecFlag) return false;
            if (self.PartialRecFlag != other.PartialRecFlag) return false;
            if (self.TunerID != other.TunerID) return false;
            return true;
        }
    }
}