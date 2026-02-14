using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows; // WPFのMessageBox用

namespace EDCBMonitor
{
    public class ExternalAppHelper
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

        private const int SW_RESTORE = 9;

        public static void ActivateOrLaunchEpgTimer()
        {
            try
            {
                var proc = Process.GetProcessesByName("EpgTimer").FirstOrDefault();
                if (proc != null)
                {
                    IntPtr hwnd = proc.MainWindowHandle;
                    if (hwnd == IntPtr.Zero) hwnd = FindWindow(null, "EpgTimer");

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
                Logger.Write($"EpgTimer Launch Error: {ex.Message}");
            }
        }

        public static void OpenTvTest(string recPath)
        {
            string exePath = Config.Data.TvTestPath;
            if (string.IsNullOrEmpty(exePath) || !File.Exists(exePath))
            {
                System.Windows.MessageBox.Show("設定画面で TVTest のパスを指定してください。", "設定未完了", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrEmpty(recPath))
            {
                System.Windows.MessageBox.Show("録画ファイルの場所を取得できませんでした。\n録画が開始されていない可能性があります。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                string cmdTemplate = Config.Data.TvTestCmd;
                string exeDir = Path.GetDirectoryName(exePath) ?? "";

                if (string.IsNullOrWhiteSpace(cmdTemplate) || cmdTemplate.Contains("/d /nd"))
                {
                    string driverName = "";
                    if (File.Exists(Path.Combine(exeDir, "BonDriver_Pipe.dll"))) driverName = "BonDriver_Pipe.dll";
                    else if (File.Exists(Path.Combine(exeDir, "BonDriver_TCP.dll"))) driverName = "BonDriver_TCP.dll";
                    else if (File.Exists(Path.Combine(exeDir, "BonDriver_UDP.dll"))) driverName = "BonDriver_UDP.dll";

                    if (!string.IsNullOrEmpty(driverName))
                    {
                        cmdTemplate = $"/d \"{driverName}\" /p 1 \"$FilePath$\"";
                    }
                    else
                    {
                        cmdTemplate = "/nd /p 1 \"$FilePath$\"";
                    }
                }

                string fileNameExt = Path.GetFileName(recPath);
                string args = cmdTemplate.Replace("$FileNameExt$", fileNameExt).Replace("$FilePath$", recPath);

                // ProcessStartInfo を変数で受けて、WorkingDirectory をセットする
                var psi = new ProcessStartInfo(exePath, args);
                psi.UseShellExecute = true;
                psi.WorkingDirectory = Path.GetDirectoryName(exePath); // ← これを追加！

                Process.Start(psi);
            }
            catch (Exception ex)
            {
                Logger.Write($"TVTest Launch Error: {ex.Message}");
                System.Windows.MessageBox.Show($"外部アプリの起動に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public static void OpenFolder(string path)
        {
            if (string.IsNullOrEmpty(path)) return;

            if (Directory.Exists(path))
            {
                try
                {
                    Process.Start("explorer.exe", path);
                }
                catch (Exception ex)
                {
                    Logger.Write($"Folder Open Error: {ex.Message}");
                    System.Windows.MessageBox.Show("フォルダを開けませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                System.Windows.MessageBox.Show($"フォルダが見つかりません。\n{path}", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private static string GetEpgTimerExePath()
        {
            if (!string.IsNullOrEmpty(Config.Data.EdcbInstallPath))
            {
                string p = Path.Combine(Config.Data.EdcbInstallPath, "EpgTimer.exe");
                if (File.Exists(p)) return p;
            }
            
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            if (File.Exists(Path.Combine(baseDir, "EpgTimer.exe"))) return Path.Combine(baseDir, "EpgTimer.exe");
            string parent = Directory.GetParent(baseDir)?.FullName ?? "";
            if (File.Exists(Path.Combine(parent, "EpgTimer.exe"))) return Path.Combine(parent, "EpgTimer.exe");

            return "";
        }
    }
}