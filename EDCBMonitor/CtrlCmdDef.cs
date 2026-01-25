using System;
using System.Collections.Generic;
using System.IO;

namespace EpgTimer
{
    public enum ErrCode : uint { CMD_SUCCESS = 1, CMD_ERR = 0, CMD_ERR_CONNECT = 204, CMD_ERR_TIMEOUT = 205, CMD_ERR_DISCONNECT = 206 }

    // コマンドID定義
    public enum CtrlCmd : uint
    {
        CMD_EPG_SRV_ENUM_RESERVE2 = 2011,
        CMD_EPG_SRV_CHG_RESERVE2 = 2015,
        CMD_EPG_SRV_GET_PG_INFO = 1023,
        CMD_EPG_SRV_DEL_RESERVE = 1014,
    }

    // --- EPG情報関連 ---

    public class EpgShortEventInfo : ICtrlCmdReadWrite
    {
        public string event_name = "";
        public string text_char = "";
        public void Write(MemoryStream s, ushort version) { }
        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref event_name);
            r.Read(ref text_char);
            r.End();
        }
    }

    public class EpgExtendedEventInfo : ICtrlCmdReadWrite
    {
        public string text_char = "";
        public void Write(MemoryStream s, ushort version) { }
        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref text_char);
            r.End();
        }
    }

    public class EpgContentData : ICtrlCmdReadWrite
    {
        public byte content_nibble_level_1;
        public byte content_nibble_level_2;
        public byte user_nibble_1;
        public byte user_nibble_2;
        public void Write(MemoryStream s, ushort version) { }
        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref content_nibble_level_1);
            r.Read(ref content_nibble_level_2);
            r.Read(ref user_nibble_1);
            r.Read(ref user_nibble_2);
            r.End();
        }
    }

    public class EpgContentInfo : ICtrlCmdReadWrite
    {
        public List<EpgContentData> nibbleList = new List<EpgContentData>();
        
        public void Write(MemoryStream s, ushort version) { }
        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref nibbleList);
            r.End();
        }
    }

    // 読み飛ばし用ダミークラス（既存）
    public class EpgDummyInfo : ICtrlCmdReadWrite
    {
        public void Write(MemoryStream s, ushort version) { }
        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin(); r.End(); 
        }
    }

    // イベントデータ詳細
    public class EpgEventData : ICtrlCmdReadWrite
    {
        public ushort original_network_id;
        public ushort transport_stream_id;
        public ushort service_id;
        public ushort event_id;
        public void Write(MemoryStream s, ushort version) { }
        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref original_network_id);
            r.Read(ref transport_stream_id);
            r.Read(ref service_id);
            r.Read(ref event_id);
            r.End();
        }
    }

    // イベントグループ情報（イベントリレー等で必要）
    public class EpgEventGroupInfo : ICtrlCmdReadWrite
    {
        public List<EpgEventData> eventDataList = new List<EpgEventData>();
        public byte group_type;
        public void Write(MemoryStream s, ushort version) { }
        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref group_type);
            r.Read(ref eventDataList);
            r.End();
        }
    }

    public class EpgEventInfo : ICtrlCmdReadWrite
    {
        public ushort original_network_id;
        public ushort transport_stream_id;
        public ushort service_id;
        public ushort event_id;
        public byte StartTimeFlag;
        public DateTime StartTime;
        public byte DurationFlag;
        public uint DurationSec;
        
        public EpgShortEventInfo? ShortInfo;
        public EpgExtendedEventInfo? ExtInfo;
        public EpgContentInfo? ContentInfo;
        public EpgDummyInfo? ComponentInfo;
        public EpgDummyInfo? AudioInfo;
        public EpgEventGroupInfo? EventGroupInfo;
        public EpgEventGroupInfo? EventRelayInfo;
        
        public byte FreeCAFlag;

        public void Write(MemoryStream s, ushort version) { }
        
        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            
            r.Read(ref original_network_id);
            r.Read(ref transport_stream_id);
            r.Read(ref service_id);
            r.Read(ref event_id);
            r.Read(ref StartTimeFlag);
            try { r.Read(ref StartTime); } catch { }
            r.Read(ref DurationFlag);
            r.Read(ref DurationSec);

            T? ReadOpt<T>() where T : ICtrlCmdReadWrite, new()
            {
                int size = 0;
                r.Read(ref size);
                if (size != 4)
                {
                    r.Stream.Seek(-4, SeekOrigin.Current);
                    var obj = new T();
                    obj.Read(s, version);
                    return obj;
                }
                return default;
            }

            ShortInfo = ReadOpt<EpgShortEventInfo>();
            ExtInfo = ReadOpt<EpgExtendedEventInfo>();
            ContentInfo = ReadOpt<EpgContentInfo>();
            ComponentInfo = ReadOpt<EpgDummyInfo>();
            AudioInfo = ReadOpt<EpgDummyInfo>();
            EventGroupInfo = ReadOpt<EpgEventGroupInfo>();
            EventRelayInfo = ReadOpt<EpgEventGroupInfo>();
            
            r.Read(ref FreeCAFlag);
            r.End();
        }
    }

    // --- 予約情報 ---

    public class RecSettingData : ICtrlCmdReadWrite
    {
        private static readonly char SEPARATOR = '*';
        public byte RecMode;
        public byte Priority;
        public byte TuijyuuFlag;
        public uint ServiceMode;
        public byte PittariFlag;
        public string BatFilePath;
        public string RecTag;
        public List<RecFileSetInfo> RecFolderList; 
        public byte SuspendMode;
        public byte RebootFlag;
        public byte UseMargineFlag;
        public int StartMargine;
        public int EndMargine;
        public byte ContinueRecFlag;
        public byte PartialRecFlag;
        public uint TunerID;
        public List<RecFileSetInfo> PartialRecFolder;

        public RecSettingData()
        {
            RecMode = 1; Priority = 1; TuijyuuFlag = 1; ServiceMode = 0; PittariFlag = 0;
            BatFilePath = ""; RecTag = ""; RecFolderList = new List<RecFileSetInfo>();
            SuspendMode = 0; RebootFlag = 0; UseMargineFlag = 0; StartMargine = 10; EndMargine = 5;
            ContinueRecFlag = 0; PartialRecFlag = 0; TunerID = 0; PartialRecFolder = new List<RecFileSetInfo>();
        }
        public byte GetRecMode() { return RecMode; }

        public void Write(MemoryStream s, ushort version)
        {
            var w = new CtrlCmdWriter(s, version);
            w.Begin();
            w.Write(RecMode);
            w.Write(Priority);
            w.Write(TuijyuuFlag);
            w.Write(ServiceMode);
            w.Write(PittariFlag);
            
            string bat = BatFilePath;
            if (!string.IsNullOrEmpty(RecTag))
            {
                bat += SEPARATOR + RecTag;
            }
            w.Write(bat);
            
            w.Write(RecFolderList);
            w.Write(SuspendMode);
            w.Write(RebootFlag);
            w.Write(UseMargineFlag);
            w.Write(StartMargine);
            w.Write(EndMargine);
            w.Write(ContinueRecFlag);
            w.Write(PartialRecFlag);
            w.Write(TunerID);
            if (version >= 2) w.Write(PartialRecFolder);
            w.End();
        }

        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref RecMode);
            r.Read(ref Priority);
            r.Read(ref TuijyuuFlag);
            r.Read(ref ServiceMode);
            r.Read(ref PittariFlag);
            {
                string batFilePathAndRecTag = "";
                r.Read(ref batFilePathAndRecTag);
                int pos = batFilePathAndRecTag.IndexOf(SEPARATOR);
                if (pos < 0) { BatFilePath = batFilePathAndRecTag; RecTag = ""; }
                else { BatFilePath = batFilePathAndRecTag.Substring(0, pos); RecTag = batFilePathAndRecTag.Substring(pos + 1); }
            }
            r.Read(ref RecFolderList);
            r.Read(ref SuspendMode);
            r.Read(ref RebootFlag);
            r.Read(ref UseMargineFlag);
            r.Read(ref StartMargine);
            r.Read(ref EndMargine);
            r.Read(ref ContinueRecFlag);
            r.Read(ref PartialRecFlag);
            r.Read(ref TunerID);
            if (version >= 2) r.Read(ref PartialRecFolder);
            r.End();
        }

        public bool IsNoRec() 
        { 
            return RecMode / 5 % 2 != 0; 
        }
    }

    public class ReserveData : ICtrlCmdReadWrite
    {
        public string Title;
        public DateTime StartTime;
        public uint DurationSecond;
        public string StationName;
        public ushort OriginalNetworkID;
        public ushort TransportStreamID;
        public ushort ServiceID;
        public ushort EventID;
        public string Comment;
        public uint ReserveID;
        public byte UnusedRecWaitFlag;
        public byte OverlapMode;
        public string UnusedRecFilePath;
        public DateTime StartTimeEpg;
        public RecSettingData RecSetting;
        public uint ReserveStatus;
        public List<string> RecFileNameList;
        public uint UnusedParam1;
        
        public List<EpgAutoAddBasicInfo> AutoAddInfo;

        public ReserveData()
        {
            Title = ""; StartTime = new DateTime(); DurationSecond = 0; StationName = "";
            OriginalNetworkID = 0; TransportStreamID = 0; ServiceID = 0; EventID = 0;
            Comment = ""; ReserveID = 0; UnusedRecWaitFlag = 0; OverlapMode = 0;
            UnusedRecFilePath = "";
            StartTimeEpg = new DateTime();
            RecSetting = new RecSettingData(); ReserveStatus = 0;
            RecFileNameList = new List<string>(); 
            UnusedParam1 = 0;
            AutoAddInfo = new List<EpgAutoAddBasicInfo>();
        }

        public void Write(MemoryStream s, ushort version)
        {
            var w = new CtrlCmdWriter(s, version);
            w.Begin();
            w.Write(Title);
            w.Write(StartTime);
            w.Write(DurationSecond);
            w.Write(StationName);
            w.Write(OriginalNetworkID);
            w.Write(TransportStreamID);
            w.Write(ServiceID);
            w.Write(EventID);
            w.Write(Comment);
            w.Write(ReserveID);
            w.Write(UnusedRecWaitFlag);
            w.Write(OverlapMode);
            w.Write(UnusedRecFilePath);
            w.Write(StartTimeEpg);
            w.Write(RecSetting);
            w.Write(ReserveStatus);
            
            if (version >= 5) 
            { 
                w.Write(RecFileNameList); 
                w.Write(UnusedParam1);
            }
            if (version >= 6) w.Write(AutoAddInfo);
            w.End();
        }

        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref Title);
            r.Read(ref StartTime);
            r.Read(ref DurationSecond);
            r.Read(ref StationName);
            r.Read(ref OriginalNetworkID);
            r.Read(ref TransportStreamID);
            r.Read(ref ServiceID);
            r.Read(ref EventID);
            r.Read(ref Comment);
            r.Read(ref ReserveID);
            r.Read(ref UnusedRecWaitFlag);
            r.Read(ref OverlapMode);
            r.Read(ref UnusedRecFilePath);
            r.Read(ref StartTimeEpg);
            r.Read(ref RecSetting);
            r.Read(ref ReserveStatus);
            
            if (version >= 5) 
            { 
                r.Read(ref RecFileNameList); 
                r.Read(ref UnusedParam1);
            }
            if (version >= 6 && r.RemainSize() > 0) r.Read(ref AutoAddInfo);
            r.End();
        }
    }

    public class RecFileSetInfo : ICtrlCmdReadWrite
    {
        public string RecFolder = "", WritePlugIn = "", RecNamePlugIn = "", RecFileName = "";
        
        public void Write(MemoryStream s, ushort version) 
        {
            var w = new CtrlCmdWriter(s, version);
            w.Begin();
            w.Write(RecFolder); w.Write(WritePlugIn); w.Write(RecNamePlugIn); w.Write(RecFileName);
            w.End();
        }

        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref RecFolder); r.Read(ref WritePlugIn); r.Read(ref RecNamePlugIn); r.Read(ref RecFileName);
            r.End(); 
        }
    }

    public class EpgAutoAddBasicInfo : ICtrlCmdReadWrite
    {
        public uint DataID;
        public string Key = "";

        public void Write(MemoryStream s, ushort version)
        {
            var w = new CtrlCmdWriter(s, version);
            w.Begin();
            w.Write(DataID);
            w.Write(Key);
            w.End(); 
        }

        public void Read(MemoryStream s, ushort version)
        {
            var r = new CtrlCmdReader(s, version);
            r.Begin();
            r.Read(ref DataID);
            r.Read(ref Key);
            r.End(); 
        }
    }
}