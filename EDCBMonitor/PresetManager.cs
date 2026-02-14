using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using EpgTimer;

namespace EDCBMonitor
{
    public class PresetManager
    {
        public static PresetManager Instance { get; } = new PresetManager();

        private class PresetInfo
        {
            public int ID { get; set; }
            public string Name { get; set; } = "";
            public RecSettingData Setting { get; set; } = new RecSettingData();
        }

        private List<PresetInfo> _presets = new List<PresetInfo>();
        private bool _isLoaded = false;

        public void Load()
        {
            _presets.Clear();
            _isLoaded = false;

            string iniPath = GetIniPath();
            if (string.IsNullOrEmpty(iniPath)) return;

            try
            {
                // 1. デフォルトプリセット (ID=0) の読み込み
                _presets.Add(LoadPresetItem(iniPath, 0));

                // 2. [SET] PresetID から有効なプリセットIDリストを取得
                string presetIdsStr = IniFileHandler.GetPrivateProfileString("SET", "PresetID", "", iniPath);
                if (!string.IsNullOrEmpty(presetIdsStr))
                {
                    string[] ids = presetIdsStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string sId in ids)
                    {
                        if (int.TryParse(sId, out int id) && id != 0)
                        {
                            _presets.Add(LoadPresetItem(iniPath, id));
                        }
                    }
                }
                
                _isLoaded = true;
            }
            catch (Exception ex)
            {
                Logger.Write($"Preset Load Error: {ex.Message}");
            }
        }

        public string GetPresetName(RecSettingData data)
        {
            // 読み込みに失敗した場合やリストが空の場合も登録時」を返す
            if (!_isLoaded || _presets.Count == 0) return "登録時";

            // 設定内容が一致するものを探す
            foreach (var preset in _presets)
            {
                if (IsEqualSetting(data, preset.Setting)) return preset.Name;
            }

            // 一致するプリセットがない場合は「登録時」を表示
            return "登録時";
        }

        private PresetInfo LoadPresetItem(string path, int id)
        {
            string suffix = id == 0 ? "" : id.ToString();
            string secName = "REC_DEF" + suffix;
            string secFolder = "REC_DEF_FOLDER" + suffix;
            string sec1Seg = "REC_DEF_FOLDER_1SEG" + suffix;

            var info = new PresetInfo { ID = id };
            info.Name = IniFileHandler.GetPrivateProfileString(secName, "SetName", "デフォルト", path);
            
            var d = info.Setting;

            // RecMode
            int rawRecMode = IniFileHandler.GetPrivateProfileInt(secName, "RecMode", 1, path);
            if ((rawRecMode / 5 % 2) == 0)
            {
                d.RecMode = (byte)(rawRecMode % 5);
            }
            else
            {
                int noRecMode = IniFileHandler.GetPrivateProfileInt(secName, "NoRecMode", 1, path);
                d.RecMode = (byte)((noRecMode % 5) + 5);
            }

            d.Priority = (byte)IniFileHandler.GetPrivateProfileInt(secName, "Priority", 2, path);
            d.TuijyuuFlag = (byte)IniFileHandler.GetPrivateProfileInt(secName, "TuijyuuFlag", 1, path);
            d.ServiceMode = (uint)IniFileHandler.GetPrivateProfileInt(secName, "ServiceMode", 16, path);
            d.PittariFlag = (byte)IniFileHandler.GetPrivateProfileInt(secName, "PittariFlag", 0, path);

            string packedBat = IniFileHandler.GetPrivateProfileString(secName, "BatFilePath", "", path);
            int sep = packedBat.IndexOf('*');
            if (sep >= 0)
            {
                d.BatFilePath = packedBat.Substring(0, sep);
                d.RecTag = packedBat.Substring(sep + 1);
            }
            else
            {
                d.BatFilePath = packedBat;
                d.RecTag = "";
            }

            d.RecFolderList = LoadFolders(path, secFolder);
            d.PartialRecFolder = LoadFolders(path, sec1Seg);

            d.SuspendMode = (byte)IniFileHandler.GetPrivateProfileInt(secName, "SuspendMode", 0, path);
            d.RebootFlag = (byte)IniFileHandler.GetPrivateProfileInt(secName, "RebootFlag", 0, path);
            d.UseMargineFlag = (byte)IniFileHandler.GetPrivateProfileInt(secName, "UseMargineFlag", 0, path);
            d.StartMargine = IniFileHandler.GetPrivateProfileInt(secName, "StartMargine", 5, path);
            d.EndMargine = IniFileHandler.GetPrivateProfileInt(secName, "EndMargine", 2, path);
            d.ContinueRecFlag = (byte)IniFileHandler.GetPrivateProfileInt(secName, "ContinueRec", 0, path);
            d.PartialRecFlag = (byte)IniFileHandler.GetPrivateProfileInt(secName, "PartialRec", 0, path);
            d.TunerID = (uint)IniFileHandler.GetPrivateProfileInt(secName, "TunerID", 0, path);

            return info;
        }

        private List<RecFileSetInfo> LoadFolders(string path, string sec)
        {
            var list = new List<RecFileSetInfo>();
            int count = IniFileHandler.GetPrivateProfileInt(sec, "Count", 0, path);
            for (int i = 0; i < count; i++)
            {
                var f = new RecFileSetInfo();
                f.RecFolder = IniFileHandler.GetPrivateProfileString(sec, i.ToString(), "", path);
                
                if (!string.IsNullOrEmpty(f.RecFolder) && !f.RecFolder.EndsWith("\\"))
                {
                    f.RecFolder += "\\";
                }

                f.WritePlugIn = IniFileHandler.GetPrivateProfileString(sec, "WritePlugIn" + i, "Write_Default.dll", path);
                f.RecNamePlugIn = IniFileHandler.GetPrivateProfileString(sec, "RecNamePlugIn" + i, "", path);
                list.Add(f);
            }
            return list;
        }

        private bool IsEqualSetting(RecSettingData self, RecSettingData other)
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
            
            if (!IsEqualFolderList(self.RecFolderList, other.RecFolderList)) return false;
            if (!IsEqualFolderList(self.PartialRecFolder, other.PartialRecFolder)) return false;

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

        private bool IsEqualFolderList(List<RecFileSetInfo> a, List<RecFileSetInfo> b)
        {
            if (a == null || b == null) return a == b;
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
            {
                if (!string.Equals(a[i].RecFolder, b[i].RecFolder, StringComparison.OrdinalIgnoreCase)) return false;
                if (!string.Equals(a[i].WritePlugIn, b[i].WritePlugIn, StringComparison.OrdinalIgnoreCase)) return false;
                if (!string.Equals(a[i].RecNamePlugIn, b[i].RecNamePlugIn, StringComparison.OrdinalIgnoreCase)) return false;
            }
            return true;
        }

        public string GetIniPath()
        {
            string path = Config.Data.EdcbInstallPath;
            if (string.IsNullOrEmpty(path)) return "";

            if (File.Exists(path) && !File.GetAttributes(path).HasFlag(FileAttributes.Directory))
            {
                path = Path.GetDirectoryName(path);
            }

            if (string.IsNullOrEmpty(path)) return "";

            var parent = Directory.GetParent(path);
            if (parent != null)
            {
                string p1 = Path.Combine(parent.FullName, "EpgTimerSrv.ini");
                if (File.Exists(p1)) return p1;
            }
            
            string p2 = Path.Combine(path, "EpgTimerSrv.ini");
            if (File.Exists(p2)) return p2;

            return "";
        }
    }
}