using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using EpgTimer;

namespace EDCBMonitor
{
    public class ReservationService
    {
        private readonly CtrlCmdUtil _ctrlCmd;
        
        // 排他制御用のセマフォ (同時に1つだけ実行許可)
        private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);

        // パイプ名とイベント名の定義
        private const string EventName = "Global\\EpgTimerSrvConnect";
        private const string PipeName = "EpgTimerSrvPipe";

        public ReservationService()
        {
            _ctrlCmd = new CtrlCmdUtil();
            _ctrlCmd.SetPipeSetting(EventName, PipeName);
            _ctrlCmd.SetConnectTimeOut(3000); 
        }

        // 予約一覧の取得（リトライ機能付き）
        public async Task<List<ReserveItem>?> GetReservationsAsync()
        {
            await _semaphore.WaitAsync();
            try
            {
                return await Task.Run(async () =>
                {
                    const int MaxRetries = 5;
                    const int RetryDelay = 500;
                    
                    GetDefaultMargins(out int defStartMargin, out int defEndMargin);

                    for (int i = 0; i < MaxRetries; i++)
                    {
                        try
                        {
                            // --- チューナーごとの予約割り当て状況を取得 ---
                            var tunerReserveList = new List<EpgTimer.TunerReserveInfo>();
                            var reserveIdToTunerMap = new Dictionary<uint, EpgTimer.TunerReserveInfo>();
                            
                            // 全チューナーIDと名前のマップ
                            Dictionary<uint, string>? allTunerNameMap = null;

                            if (_ctrlCmd.SendEnumTunerReserve(ref tunerReserveList) == ErrCode.CMD_SUCCESS)
                            {
                                allTunerNameMap = tunerReserveList.ToDictionary(t => t.tunerID, t => t.tunerName);
                                
                                foreach (var t in tunerReserveList)
                                {
                                    foreach (var rid in t.reserveList)
                                    {
                                        if (!reserveIdToTunerMap.ContainsKey(rid))
                                        {
                                            reserveIdToTunerMap[rid] = t;
                                        }
                                    }
                                }
                            }
                            // -----------------------------------------------------

                            var res = new List<ReserveData>();
                            ErrCode status = _ctrlCmd.SendEnumReserve(ref res);

                            if (status == ErrCode.CMD_SUCCESS)
                            {
                                var items = new List<ReserveItem>();
                                var serviceEpgCache = new Dictionary<ulong, List<EpgTimer.EpgEventInfo>>();

                                foreach (var r in res)
                                {
                                    var item = new ReserveItem(r);
                                    
                                    // デフォルトマージンをセットして再計算
                                    item.DefaultStartMargin = defStartMargin;
                                    item.DefaultEndMargin = defEndMargin;
                                    item.UpdateProgress();

                                    item.TunerNameMap = allTunerNameMap;

                                    if (reserveIdToTunerMap.TryGetValue(r.ReserveID, out var allocatedTuner))
                                    {
                                        item.AllocatedTunerID = allocatedTuner.tunerID;
                                        item.AllocatedTunerName = allocatedTuner.tunerName;
                                    }
                                    // -----------------------------------------------------

                                    if (r.EventID != 0xFFFF)
                                    {
                                        ulong pgID = ((ulong)r.OriginalNetworkID << 48) |
                                                     ((ulong)r.TransportStreamID << 32) |
                                                     ((ulong)r.ServiceID << 16) |
                                                     (ulong)r.EventID;

                                        var eventInfo = new EpgTimer.EpgEventInfo();
                                        if (_ctrlCmd.SendGetPgInfo(pgID, ref eventInfo) == ErrCode.CMD_SUCCESS)
                                        {
                                            item.EventInfo = eventInfo;
                                        }
                                    }
                                    else
                                    {
                                        // プログラム予約
                                        ulong serviceKey = CommonManager.Create64Key(r.OriginalNetworkID, r.TransportStreamID, r.ServiceID);
                                        if (!serviceEpgCache.ContainsKey(serviceKey))
                                        {
                                            var pgList = new List<EpgTimer.EpgEventInfo>();
                                            if (_ctrlCmd.SendEnumPgInfo(serviceKey, ref pgList) == ErrCode.CMD_SUCCESS)
                                                serviceEpgCache[serviceKey] = pgList;
                                            else
                                                serviceEpgCache[serviceKey] = null;
                                        }
                                        
                                        var cachedList = serviceEpgCache[serviceKey];
                                        if (cachedList != null)
                                        {
                                            var foundEvent = cachedList.FirstOrDefault(e => 
                                                e.start_time <= r.StartTime && 
                                                r.StartTime < e.start_time.AddSeconds(e.durationSec));
                                            if (foundEvent != null) item.EventInfo = foundEvent;
                                        }
                                    }

                                    items.Add(item);
                                }
                                return items.OrderBy(item => item.Data.StartTime).ToList();
                            }

                            // 接続エラーはログに出さない（起動待ちの可能性があるため）
                            if (status != ErrCode.CMD_ERR_CONNECT)
                            {
                                Logger.Write($"GetReservations Failed (Attempt {i + 1}/{MaxRetries}). Status: {status}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Write($"GetReservations Exception (Attempt {i + 1}/{MaxRetries}): {ex.Message}");
                        }

                        if (i < MaxRetries - 1) await Task.Delay(RetryDelay);
                    }

                    return null;
                });
            }
            finally
            {
                _semaphore.Release();
            }
        }

        // 予約の有効/無効切り替え
        public async Task<bool> ToggleReservationStatusAsync(List<uint> targetIDs)
        {
            await _semaphore.WaitAsync();
            try
            {
                return await Task.Run(async () =>
                {
                    const int MaxRetries = 3;
                    const int RetryDelay = 200;

                    for (int i = 0; i < MaxRetries; i++)
                    {
                        try
                        {
                            var list = new List<ReserveData>();
                            if (_ctrlCmd.SendEnumReserve(ref list) != ErrCode.CMD_SUCCESS) 
                            {
                                throw new Exception("Failed to get reserve list for toggle");
                            }

                            var changeList = new List<ReserveData>();
                            foreach (var data in list.Where(d => targetIDs.Contains(d.ReserveID)))
                            {
                                data.RecSetting.IsEnable = !data.RecSetting.IsEnable;
                                changeList.Add(data);
                            }

                            if (changeList.Count > 0)
                            {
                                if (_ctrlCmd.SendChgReserve(changeList) == ErrCode.CMD_SUCCESS)
                                {
                                    return true;
                                }
                            }
                            else
                            {
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Write($"ToggleReservationStatus Error (Attempt {i+1}): {ex.Message}");
                        }

                        if (i < MaxRetries - 1) await Task.Delay(RetryDelay);
                    }
                    return false;
                });
            }
            finally
            {
                _semaphore.Release();
            }
        }

        // 予約の削除
        public async Task<bool> DeleteReservationsAsync(List<uint> targetIDs)
        {
            await _semaphore.WaitAsync();
            try
            {
                return await Task.Run(async () =>
                {
                    // 削除処理もリトライを行うように修正 (MaxRetries=3)
                    const int MaxRetries = 3;
                    const int RetryDelay = 200;

                    for (int i = 0; i < MaxRetries; i++)
                    {
                        try
                        {
                            if (_ctrlCmd.SendDelReserve(targetIDs) == ErrCode.CMD_SUCCESS)
                            {
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Write($"DeleteReservations Error (Attempt {i+1}): {ex.Message}");
                        }
                        
                        if (i < MaxRetries - 1) await Task.Delay(RetryDelay);
                    }
                    return false;
                });
            }
            finally
            {
                _semaphore.Release();
            }
        }

        // 共通録画設定の取得（録画保存フォルダのパス取得などに使用）
        public RecSettingData? GetRecSetting()
        {
            try
            {
                var set = new RecSettingData();
                if (_ctrlCmd.SendGetRecSetting(ref set) == ErrCode.CMD_SUCCESS)
                {
                    return set;
                }
            }
            catch (Exception ex)
            {
                Logger.Write($"GetRecSetting Error: {ex.Message}");
            }
            return null;
        }

        // 追っかけ再生用のファイルパス取得
        public string? OpenTimeShift(uint reserveId, out uint ctrlId)
        {
            ctrlId = 0;
            try
            {
                var info = new NWPlayTimeShiftInfo();
                if (_ctrlCmd.SendNwTimeShiftOpen(reserveId, ref info) == ErrCode.CMD_SUCCESS)
                {
                    ctrlId = info.ctrlID;
                    return info.filePath;
                }
            }
            catch (Exception ex)
            {
                Logger.Write($"OpenTimeShift Error: {ex.Message}");
            }
            return null;
        }

        // 追っかけ再生終了通知
        public void CloseNwPlay(uint ctrlId)
        {
            try
            {
                _ctrlCmd.SendNwPlayClose(ctrlId);
            }
            catch { }
        }
    
        private const string RESERVE_TXT_NAME = "Reserve.txt";
        private const string EPG_TIMER_SRV_INI = "EpgTimerSrv.ini";

        private void GetDefaultMargins(out int startMargin, out int endMargin)
        {
            startMargin = 5;
            endMargin = 2;
            string iniPath = GetEpgTimerSrvIniPath();
            if (string.IsNullOrEmpty(iniPath)) return;

            var ini = LoadIni(iniPath);
            if (ini.TryGetValue("SET", out var setSec))
            {
                if (setSec.TryGetValue("StartMargin", out var sm)) int.TryParse(sm, out startMargin);
                if (setSec.TryGetValue("EndMargin", out var em)) int.TryParse(em, out endMargin);
            }
        }

        private string GetEpgTimerSrvIniPath()
        {
            string reservePath = GetReserveFilePath();
            if (string.IsNullOrEmpty(reservePath)) return "";

            string dir = System.IO.Path.GetDirectoryName(reservePath);
            if (string.IsNullOrEmpty(dir)) return "";

            var parent = System.IO.Directory.GetParent(dir);
            if (parent != null)
            {
                string pParent = System.IO.Path.Combine(parent.FullName, EPG_TIMER_SRV_INI);
                if (System.IO.File.Exists(pParent)) return pParent;
            }

            string pSame = System.IO.Path.Combine(dir, EPG_TIMER_SRV_INI);
            return System.IO.File.Exists(pSame) ? pSame : "";
        }

        private string GetReserveFilePath()
        {
            if (!string.IsNullOrEmpty(Config.Data.EdcbInstallPath))
            {
                var pathsToCheck = new[]
                {
                    Config.Data.EdcbInstallPath,
                    System.IO.Path.Combine(Config.Data.EdcbInstallPath, "Setting", RESERVE_TXT_NAME),
                    System.IO.Path.Combine(Config.Data.EdcbInstallPath, RESERVE_TXT_NAME)
                };

                foreach (var p in pathsToCheck)
                {
                    if (IsReserveFile(p)) return p;
                }
            }

            var localPaths = new[]
            {
                System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RESERVE_TXT_NAME),
                System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Setting", RESERVE_TXT_NAME)
            };

            return localPaths.FirstOrDefault(p => IsReserveFile(p)) ?? "";
        }

        private bool IsReserveFile(string path) 
            => !string.IsNullOrEmpty(path) && System.IO.File.Exists(path) && path.EndsWith(RESERVE_TXT_NAME, StringComparison.OrdinalIgnoreCase);

        private Dictionary<string, Dictionary<string, string>> LoadIni(string path)
        {
            var data = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            if (!System.IO.File.Exists(path)) return data;

            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                var enc = System.Text.Encoding.GetEncoding(932);
                
                string currentSection = "";

                foreach (var line in System.IO.File.ReadLines(path, enc))
                {
                    string t = line.Trim();
                    if (t.Length == 0 || t.StartsWith(";") || t.StartsWith("//")) continue;

                    if (t.StartsWith("[") && t.EndsWith("]"))
                    {
                        currentSection = t.Substring(1, t.Length - 2).Trim();
                        if (!data.ContainsKey(currentSection))
                            data[currentSection] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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
            catch (Exception ex)
            {
                Logger.Write($"LoadIni Error: {ex.Message}");
            }
            return data;
        }
        
        private string GetCommonIniPath()
        {
            string reservePath = GetReserveFilePath();
            if (string.IsNullOrEmpty(reservePath)) return "";

            string dir = System.IO.Path.GetDirectoryName(reservePath) ?? "";
            if (string.IsNullOrEmpty(dir)) return "";

            var parent = System.IO.Directory.GetParent(dir);
            if (parent != null)
            {
                string pParent = System.IO.Path.Combine(parent.FullName, "Common.ini");
                if (System.IO.File.Exists(pParent)) return pParent;
            }

            string pSame = System.IO.Path.Combine(dir, "Common.ini");
            return System.IO.File.Exists(pSame) ? pSame : "";
        }

        public List<string> GetCommonRecFolders()
        {
            var list = new List<string>();
            string iniPath = GetCommonIniPath();
            if (string.IsNullOrEmpty(iniPath)) return list;

            var ini = LoadIni(iniPath);
            if (ini.TryGetValue("SET", out var setSec))
            {
                if (setSec.TryGetValue("RecFolderNum", out string numStr) && int.TryParse(numStr, out int num))
                {
                    for (int i = 0; i < num; i++)
                    {
                        if (setSec.TryGetValue("RecFolderPath" + i, out string path) && !string.IsNullOrEmpty(path))
                        {
                            list.Add(path);
                        }
                    }
                }
            }
            return list;
        }
    }
}