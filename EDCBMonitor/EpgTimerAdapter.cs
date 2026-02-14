using System;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
using System.Net;
using EDCBMonitor;

namespace EpgTimer
{
    public class CommonManager
    {
        public static CommonManager Instance { get; } = new CommonManager();
        
        public bool NWMode { get; set; } = false;
        
        public IPAddress NWConnectedIP { get; set; } = IPAddress.Loopback;
        public uint NWConnectedPort { get; set; } = 5678;

        public static CtrlCmdUtil CreateSrvCtrl()
        {
            var cmd = new CtrlCmdUtil();
            cmd.SetPipeSetting("Global\\EpgTimerSrvConnect", "EpgTimerSrvPipe");
            return cmd;
        }

        public static ulong Create64Key(ushort onid, ushort tsid, ushort sid)
        {
            return ((ulong)onid << 32) | ((ulong)tsid << 16) | sid;
        }

        public static string GetErrCodeText(ErrCode err) => err.ToString();
    }

    // Settings: TVTest関連の設定
    public static class Settings
    {
        public static ConfigProxy Instance { get; } = new ConfigProxy();

        public class ConfigProxy
        {
            public string TvTestExe => Config.Data.TvTestPath;
            public string TvTestCmd => Config.Data.TvTestCmd;
            public int TvTestOpenWait { get; set; } = 2000;
            public int TvTestChgBonWait { get; set; } = 2000;

            public bool NwTvMode { get; set; } = false;
            public bool NwTvModeUDP { get; set; } = false;
            public bool NwTvModeTCP { get; set; } = false;
            public bool NwTvModePipe { get; set; } = false;
        }
    }

    // SettingPath: 設定ファイルのパス定義
    public static class SettingPath
    {
        public static string ModulePath => AppDomain.CurrentDomain.BaseDirectory;
        public static string SettingDir => Path.Combine(ModulePath, "Setting");
        public static string TimerSrvIniPath => Path.Combine(SettingDir, "Common.ini");
        public static string EpgTimerSrvIniPath => Path.Combine(SettingDir, "EpgTimerSrv.ini");
    }

    // IniFileHandler: Win32APIのラッパー (重複回避版)
    public static class IniFileHandler
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetPrivateProfileIntW")]
        private static extern uint GetPrivateProfileIntAPI(string lpAppName, string lpKeyName, int nDefault, string lpFileName);

        // 文字列読み込み用API
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetPrivateProfileStringW")]
        private static extern uint GetPrivateProfileStringAPI(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, uint nSize, string lpFileName);

        public static int GetPrivateProfileInt(string app, string key, int def, string path)
        {
            return (int)GetPrivateProfileIntAPI(app, key, def, path);
        }

        // 文字列読み込みメソッド
        public static string GetPrivateProfileString(string app, string key, string def, string path)
        {
            var sb = new StringBuilder(1024);
            GetPrivateProfileStringAPI(app, key, def, sb, (uint)sb.Capacity, path);
            return sb.ToString();
        }

        public static void WritePrivateProfileString(string app, string key, string val, string path) { }
    }
}