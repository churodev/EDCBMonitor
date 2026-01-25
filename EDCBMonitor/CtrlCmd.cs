using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EpgTimer
{
    /// <summary>CtrlCmdバイナリ形式との相互変換インターフェイス</summary>
    public interface ICtrlCmdReadWrite
    {
        /// <summary>ストリームをCtrlCmdバイナリ形式で読み込む</summary>
        void Read(MemoryStream s, ushort version);
        /// <summary>ストリームにCtrlCmdバイナリ形式で書き込む</summary>
        void Write(MemoryStream s, ushort version);
    }

    public class CtrlCmdWriter
    {
        public MemoryStream Stream { get; private set; }
        public ushort Version { get; set; }
        private long lastPos;
        public CtrlCmdWriter(MemoryStream stream, ushort version = 0)
        {
            Stream = stream;
            Version = version;
            lastPos = 0;
        }
        /// <summary>変換可能なオブジェクトをストリームに書き込む</summary>
        public void Write(object? v)
        {
            if (v == null) return;

            if (v is byte) Stream.WriteByte((byte)v);
            else if (v is ushort) Stream.Write(BitConverter.GetBytes((ushort)v), 0, 2);
            else if (v is int) Stream.Write(BitConverter.GetBytes((int)v), 0, 4);
            else if (v is uint) Stream.Write(BitConverter.GetBytes((uint)v), 0, 4);
            else if (v is float) Stream.Write(BitConverter.GetBytes((float)v), 0, 4);
            else if (v is long) Stream.Write(BitConverter.GetBytes((long)v), 0, 8);
            else if (v is ulong) Stream.Write(BitConverter.GetBytes((ulong)v), 0, 8);
            else if (v is ICtrlCmdReadWrite) ((ICtrlCmdReadWrite)v).Write(Stream, Version);
            else if (v is DateTime)
            {
                var t = (DateTime)v;
                Write((ushort)t.Year);
                Write((ushort)t.Month);
                Write((ushort)t.DayOfWeek);
                Write((ushort)t.Day);
                Write((ushort)t.Hour);
                Write((ushort)t.Minute);
                Write((ushort)t.Second);
                Write((ushort)t.Millisecond);
            }
            else if (v is string)
            {
                byte[] a = Encoding.Unicode.GetBytes((string)v);
                Write(a.Length + 6);
                Stream.Write(a, 0, a.Length);
                Write((ushort)0);
            }
            else if (v is System.Collections.IList)
            {
                long lpos = Stream.Position;
                Write(0);
                Write(((System.Collections.IList)v).Count);
                foreach (object o in ((System.Collections.IList)v))
                {
                    Write(o);
                }
                long pos = Stream.Position;
                Stream.Seek(lpos, SeekOrigin.Begin);
                Write((int)(pos - lpos));
                Stream.Seek(pos, SeekOrigin.Begin);
            }
            else
            {
                throw new ArgumentException();
            }
        }
        /// <summary>C++構造体型オブジェクトの書き込みを開始する</summary>
        public void Begin()
        {
            lastPos = Stream.Position;
            Write(0);
        }
        /// <summary>C++構造体型オブジェクトの書き込みを終了する</summary>
        public void End()
        {
            long pos = Stream.Position;
            Stream.Seek(lastPos, SeekOrigin.Begin);
            Write((int)(pos - lastPos));
            Stream.Seek(pos, SeekOrigin.Begin);
        }
    }

    public class CtrlCmdReader
    {
        public MemoryStream Stream { get; private set; }
        public ushort Version { get; set; }
        private long tailPos;
        private byte[]? buff; // Null許容に変更
        public CtrlCmdReader(MemoryStream stream, ushort version = 0)
        {
            Stream = stream;
            Version = version;
            tailPos = long.MaxValue;
        }
        public byte[] ReadBytes(int size)
        {
            if (buff == null || buff.Length < size)
            {
                buff = new byte[size];
            }
            if (Stream.Read(buff, 0, size) != size)
            {
                throw new EndOfStreamException();
            }
            return buff;
        }
        /// <summary>配列のサイズだけストリームから読み込む</summary>
        public void ReadToArray(byte[] v)
        {
            if (Stream.Read(v, 0, v.Length) != v.Length)
            {
                throw new EndOfStreamException();
            }
        }
        /// <summary>ストリームから読み込む</summary>
        public void Read(ref byte v) { v = ReadBytes(1)[0]; }
        /// <summary>ストリームから読み込む</summary>
        public void Read(ref ushort v) { v = BitConverter.ToUInt16(ReadBytes(2), 0); }
        /// <summary>ストリームから読み込む</summary>
        public void Read(ref int v) { v = BitConverter.ToInt32(ReadBytes(4), 0); }
        /// <summary>ストリームから読み込む</summary>
        public void Read(ref uint v) { v = BitConverter.ToUInt32(ReadBytes(4), 0); }
        /// <summary>ストリームから読み込む</summary>
        public void Read(ref float v) { v = BitConverter.ToSingle(ReadBytes(4), 0); }
        /// <summary>ストリームから読み込む</summary>
        public void Read(ref long v) { v = BitConverter.ToInt64(ReadBytes(8), 0); }
        /// <summary>ストリームから読み込む</summary>
        public void Read(ref ulong v) { v = BitConverter.ToUInt64(ReadBytes(8), 0); }
        /// <summary>ストリームから読み込む</summary>
        public void Read(ref DateTime v)
        {
            byte[] a = ReadBytes(16);
            v = new DateTime(BitConverter.ToUInt16(a, 0), BitConverter.ToUInt16(a, 2), BitConverter.ToUInt16(a, 6), BitConverter.ToUInt16(a, 8),
                             BitConverter.ToUInt16(a, 10), BitConverter.ToUInt16(a, 12), BitConverter.ToUInt16(a, 14));
        }
        /// <summary>オブジェクトの型に従ってストリームから読み込む</summary>
        public void Read<T>(ref T v) where T : class
        {
            // Null免除演算子(!)を使用して警告を抑制
            if (v is byte) v = (T)(object)ReadBytes(1)[0];
            else if (v is ushort) v = (T)(object)BitConverter.ToUInt16(ReadBytes(2), 0);
            else if (v is int) v = (T)(object)BitConverter.ToInt32(ReadBytes(4), 0);
            else if (v is uint) v = (T)(object)BitConverter.ToUInt32(ReadBytes(4), 0);
            else if (v is float) v = (T)(object)BitConverter.ToSingle(ReadBytes(4), 0);
            else if (v is long) v = (T)(object)BitConverter.ToInt64(ReadBytes(8), 0);
            else if (v is ulong) v = (T)(object)BitConverter.ToUInt64(ReadBytes(8), 0);
            else if (v is ICtrlCmdReadWrite) ((ICtrlCmdReadWrite)v).Read(Stream, Version);
            else if (v is DateTime)
            {
                byte[] a = ReadBytes(16);
                v = (T)(object)new DateTime(BitConverter.ToUInt16(a, 0), BitConverter.ToUInt16(a, 2), BitConverter.ToUInt16(a, 6), BitConverter.ToUInt16(a, 8),
                                            BitConverter.ToUInt16(a, 10), BitConverter.ToUInt16(a, 12), BitConverter.ToUInt16(a, 14));
            }
            else if (v is string)
            {
                int size = 0;
                Read(ref size);
                if (size < 4 || Stream.Length - Stream.Position < size - 4)
                {
                    throw new EndOfStreamException("サイズフィールドの値が異常です");
                }
                v = (T)(object)"";
                if (size > 4)
                {
                    byte[] a = ReadBytes(size - 4);
                    if (size > 6)
                    {
                        v = (T)(object)Encoding.Unicode.GetString(a, 0, size - 6);
                    }
                }
            }
            else if (v is System.Collections.IList)
            {
                int size = 0;
                Read(ref size);
                if (size < 4 || Stream.Length - Stream.Position < size - 4)
                {
                    throw new EndOfStreamException("サイズフィールドの値が異常です");
                }
                long tPos = Stream.Position + size - 4;
                int count = 0;
                Read(ref count);
                var list = (System.Collections.IList)v;
                Type t = list.GetType();
                if (t.IsGenericType == false || t.GetGenericTypeDefinition() != typeof(List<>))
                {
                    throw new ArgumentException();
                }
                t = t.GetGenericArguments()[0];
                for (int i = 0; i < count; i++)
                {
                    // CreateInstanceのnull戻り値を抑制
                    object e = (t == typeof(string) ? "" : Activator.CreateInstance(t))!;
                    Read(ref e);
                    list.Add(e);
                }
                if (Stream.Position > tPos)
                {
                    throw new EndOfStreamException("サイズフィールドの値を超えて読み込みました");
                }
                Stream.Seek(tPos, SeekOrigin.Begin);
            }
            else
            {
                throw new ArgumentException();
            }
        }
        /// <summary>C++構造体型オブジェクトの読み込みを開始する</summary>
        public void Begin()
        {
            int size = 0;
            Read(ref size);
            if (size < 4 || Stream.Length - Stream.Position < size - 4)
            {
                throw new EndOfStreamException("サイズフィールドの値が異常です");
            }
            tailPos = Stream.Position + size - 4;
        }
        /// <summary>C++構造体型オブジェクトの読み込みを終了する</summary>
        public void End()
        {
            if (Stream.Position > tailPos)
            {
                throw new EndOfStreamException("サイズフィールドの値を超えて読み込みました");
            }
            Stream.Seek(tailPos, SeekOrigin.Begin);
            tailPos = long.MaxValue;
        }
        /// <summary>C++構造体型オブジェクトの読み込みに利用できる残サイズを取得する</summary>
        public long RemainSize()
        {
            return tailPos - Stream.Position;
        }
    }

    public class CtrlCmdUtil
    {
        private const ushort CMD_VER = 5;
        private bool tcpFlag = false;
        private int connectTimeOut = 15000;
        private string eventName = "Global\\EpgTimerSrvConnect";
        private string pipeName = "EpgTimerSrvPipe";
        private System.Net.IPAddress ip = System.Net.IPAddress.Loopback;
        private uint port = 5678;

        public void SetSendMode(bool tcpFlag) { this.tcpFlag = tcpFlag; }
        public void SetPipeSetting(string eventName, string pipeName) { this.eventName = eventName; this.pipeName = pipeName; }
        public bool PipeExists()
        {
            try { using (System.Threading.EventWaitHandle.OpenExisting(eventName)) { return true; } }
            catch { }
            return false;
        }
        public void SetNWSetting(System.Net.IPAddress ip, uint port) { this.ip = ip; this.port = port; }
        public void SetConnectTimeOut(int timeOut) { connectTimeOut = timeOut; }

        public ErrCode SendEnumReserve(ref List<ReserveData> val) { object o = val; return ReceiveCmdData2(CtrlCmd.CMD_EPG_SRV_ENUM_RESERVE2, ref o); }
        public ErrCode SendChgReserve(List<ReserveData> val) { return SendCmdData2(CtrlCmd.CMD_EPG_SRV_CHG_RESERVE2, val); }
        public ErrCode SendDelReserve(List<uint> val) { return SendCmdData(CtrlCmd.CMD_EPG_SRV_DEL_RESERVE, val); }

        private ErrCode SendPipe(CtrlCmd param, MemoryStream? send, ref MemoryStream? res)
        {
            try
            {
                using (var waitEvent = System.Threading.EventWaitHandle.OpenExisting(eventName))
                {
                    if (waitEvent.WaitOne(connectTimeOut) == false) return ErrCode.CMD_ERR_TIMEOUT;
                }
            }
            catch (System.Threading.WaitHandleCannotBeOpenedException) { return ErrCode.CMD_ERR_CONNECT; }

            using (var pipe = new System.IO.Pipes.NamedPipeClientStream(pipeName))
            {
                pipe.Connect(0);
                var head = new byte[8];
                BitConverter.GetBytes((uint)param).CopyTo(head, 0);
                BitConverter.GetBytes((uint)(send == null ? 0 : send.Length)).CopyTo(head, 4);
                pipe.Write(head, 0, 8);
                if (send != null && send.Length != 0)
                {
                    send.Close();
                    byte[] data = send.ToArray();
                    pipe.Write(data, 0, data.Length);
                }
                if (ReadAll(pipe, head, 0, 8) != 8) return ErrCode.CMD_ERR_DISCONNECT;
                uint resParam = BitConverter.ToUInt32(head, 0);
                var resData = new byte[BitConverter.ToUInt32(head, 4)];
                if (ReadAll(pipe, resData, 0, resData.Length) != resData.Length) return ErrCode.CMD_ERR_DISCONNECT;
                res = new MemoryStream(resData, false);
                return Enum.IsDefined(typeof(ErrCode), resParam) ? (ErrCode)resParam : ErrCode.CMD_ERR;
            }
        }

        private static int ReadAll(Stream s, byte[] buffer, int offset, int size)
        {
            int n = 0;
            for (int m; n < size && (m = s.Read(buffer, offset + n, size - n)) > 0; n += m) ;
            return n;
        }

        private ErrCode SendTCP(CtrlCmd param, MemoryStream? send, ref MemoryStream? res)
        {
            if (ip == null) return ErrCode.CMD_ERR_CONNECT;
            using (var tcp = new System.Net.Sockets.TcpClient(ip.AddressFamily))
            {
                try { tcp.Connect(ip, (int)port); }
                catch (System.Net.Sockets.SocketException) { return ErrCode.CMD_ERR_CONNECT; }
                using (System.Net.Sockets.NetworkStream ns = tcp.GetStream())
                {
                    var head = new byte[8 + (send == null ? 0 : send.Length)];
                    BitConverter.GetBytes((uint)param).CopyTo(head, 0);
                    BitConverter.GetBytes((uint)(send == null ? 0 : send.Length)).CopyTo(head, 4);
                    if (send != null && send.Length != 0)
                    {
                        send.Close();
                        send.ToArray().CopyTo(head, 8);
                    }
                    ns.Write(head, 0, head.Length);
                    if (ReadAll(ns, head, 0, 8) != 8) return ErrCode.CMD_ERR_DISCONNECT;
                    uint resParam = BitConverter.ToUInt32(head, 0);
                    var resData = new byte[BitConverter.ToUInt32(head, 4)];
                    if (ReadAll(ns, resData, 0, resData.Length) != resData.Length) return ErrCode.CMD_ERR_DISCONNECT;
                    res = new MemoryStream(resData, false);
                    return Enum.IsDefined(typeof(ErrCode), resParam) ? (ErrCode)resParam : ErrCode.CMD_ERR;
                }
            }
        }

        private ErrCode SendCmdStream(CtrlCmd param, MemoryStream? send, ref MemoryStream? res)
        {
            return tcpFlag ? SendTCP(param, send, ref res) : SendPipe(param, send, ref res);
        }
        private ErrCode SendCmdData(CtrlCmd param, object val)
        {
            var w = new CtrlCmdWriter(new MemoryStream());
            w.Write(val);
            MemoryStream? res = null;
            return SendCmdStream(param, w.Stream, ref res);
        }
        private ErrCode SendCmdData2(CtrlCmd param, object val)
        {
            var w = new CtrlCmdWriter(new MemoryStream(), CMD_VER);
            w.Write(CMD_VER);
            w.Write(val);
            MemoryStream? res = null;
            return SendCmdStream(param, w.Stream, ref res);
        }
        private ErrCode ReceiveCmdData2(CtrlCmd param, ref object val)
        {
            var w = new CtrlCmdWriter(new MemoryStream(), CMD_VER);
            w.Write(CMD_VER);
            MemoryStream? res = null;
            ErrCode ret = SendCmdStream(param, w.Stream, ref res);
            if (ret == ErrCode.CMD_SUCCESS && res != null)
            {
                var r = new CtrlCmdReader(res);
                ushort version = 0;
                r.Read(ref version);
                r.Version = version;
                r.Read(ref val);
            }
            return ret;
        }
        private ErrCode SendAndReceiveCmdData2(CtrlCmd param, object val, ref object resVal)
        {
            var w = new CtrlCmdWriter(new MemoryStream(), CMD_VER);
            w.Write(CMD_VER);
            w.Write(val);
            MemoryStream? res = null;
            ErrCode ret = SendCmdStream(param, w.Stream, ref res);
            if (ret == ErrCode.CMD_SUCCESS && res != null)
            {
                var r = new CtrlCmdReader(res);
                ushort version = 0;
                r.Read(ref version);
                r.Version = version;
                r.Read(ref resVal);
            }
            return ret;
        }

        // バージョンヘッダ無しで送受信を行うメソッド（番組情報取得用）
        private ErrCode SendAndReceiveCmdData(CtrlCmd param, object val, ref object resVal)
        {
            var w = new CtrlCmdWriter(new MemoryStream(), CMD_VER);
            // w.Write(CMD_VER); // バージョンは書かない
            w.Write(val);
            MemoryStream? res = null;
            ErrCode ret = SendCmdStream(param, w.Stream, ref res);
            if (ret == ErrCode.CMD_SUCCESS && res != null)
            {
                var r = new CtrlCmdReader(res);
                // バージョンは読まない（データ本体のみ返ってくるため）
                r.Version = CMD_VER; 
                r.Read(ref resVal);
            }
            return ret;
        }

        public ErrCode SendGetPgInfo(ulong id, ref EpgEventInfo val)
        {
            object o = new EpgEventInfo();
            ErrCode ret = SendAndReceiveCmdData(CtrlCmd.CMD_EPG_SRV_GET_PG_INFO, id, ref o);
            if (ret == ErrCode.CMD_SUCCESS)
            {
                val = (EpgEventInfo)o;
            }
            return ret;
        }
    }
}