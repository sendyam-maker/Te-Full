using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;



namespace DEMO
{
    public class Dump96
    {



        //inter variable
        private byte[] cSysData = new byte[100];

        private string sDeviceVersion = "";
        private string _appname = "";
        private bool _bisagent = false;
        private string dumppath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

        public delegate void callbackmsg(string s);
        public event callbackmsg Callbackmsg;
        public delegate void callbackpubevent(clPubevent cEvent);
        public event callbackpubevent CallbackPubEvent;
        public UInt32 RAC2000Handle;
        public int RAC2000Id;
        private Int32 iDumpMemoryLen = 0;
        public readonly int MaxMemorySize = 4096000; //最大記憶體大小
        //Function  
        //1.save to bin
        //2.get system parameter
        public int GetSwipBin(byte[] bSwipData)
        {

            try
            {
                AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "check device online.");

                AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "get version.");
                int iret1 = GetDeviceVersion(ref sDeviceVersion);
                if (GetDeviceSysParam() == 0)
                {
                    AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "start dump...");
                    int iTotallen = 0;
                    //-->int tempRes = CHTA_DLL.DumpRecord(_address, iPort, _nodeid, bSwipData, ref iTotallen);
                    int tempRes = 0;
                    if (tempRes == 0)
                    {
                        AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "Dump OK,Length:" + iTotallen.ToString());
                        iDumpMemoryLen = iTotallen;
                        //存BIN檔
                        string sBinFile = "";

                        sBinFile = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "dumpdata.bin";

                        AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "save Dump BIN to:" + sBinFile);
                        FileStream fs = new FileStream(sBinFile, FileMode.Create, FileAccess.Write, FileShare.Write);
                        BinaryWriter bw = new BinaryWriter(fs, System.Text.Encoding.UTF8);
                        bw.Write(bSwipData);
                        bw.Close();
                        fs.Close();

                    }
                    else
                    {
                        AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "Dump fail,return code:" + tempRes.ToString());
                    }
                    return tempRes;
                }
                else
                {
                    AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "GetDeviceSysParam fail.");
                    return -3;
                }

            }
            catch (Exception ex)
            {
                AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "exception:" + ex.Message);
                return -1;
            }
        }

        public int BeginDump()
        {
            int resultCode = -1;
            int timeout = 2000;

            byte[] bsend = new byte[255];
            byte[] breceive = new byte[255];
            int ireceivelen = 0;
            bsend[0] = 02;

            resultCode = TRAC2000DLL.hacHWRWCommandCCH(0x14, RAC2000Id, 0x01, bsend, 1, breceive, ref ireceivelen, RAC2000Handle, timeout);
            return resultCode;
        }

        public int KeepDump(ref byte[] b, ref byte[] r, Int32 iOffset, ref int iReceiveLength)
        {
            int resultCode = -1;
            int timeout = 2000;

            byte[] bsend = new byte[255];
            byte[] breceive = new byte[1024];
            int ireceivelen = 0;
            int iprotocollen = 0xf + 0x3;
            bsend[0] = 02;

            resultCode = TRAC2000DLL.hacHWRWCommandCCH(0x14, RAC2000Id, 0x02, bsend, 0, breceive, ref ireceivelen, RAC2000Handle, timeout);
            iReceiveLength = ireceivelen - iprotocollen;
            if (iReceiveLength <= 0)
            {
                resultCode = 0xff;
                return resultCode;
            }
            if (iOffset + ireceivelen < b.Length)
                Array.Copy(breceive, 0x0f, b, iOffset, ireceivelen - iprotocollen);
            Array.Resize(ref r, ireceivelen - iprotocollen);
            Array.Copy(breceive, 0x0f, r, 0, r.Length);
            bool bcheckallff = true;
            for (int i = 0xf; i < 20; i++)
            {
                if (breceive[i] != 0xff)
                {
                    bcheckallff = false;
                    break;
                }
            }
            if (bcheckallff)
                resultCode = 0xff;
            return resultCode;
        }

        public int DecodeSwip(byte[] bSwipData)
        {

            int iCardLen = 16;
            int iRecordSize = 20;
            int iRecords = bSwipData.Length / iRecordSize;
            int iOffset = 0;
            //解碼方法
            byte[] bRecord = new byte[iRecordSize];

            ArrayList alData = new ArrayList();

            string sstartdate = "";
            string senddate = "";

            char ctab = Convert.ToChar(9);

            for (int ir = 0; ir < iRecords; ir++)
            {
                try
                {
                    Array.Copy(bSwipData, iOffset, bRecord, 0, iRecordSize);
                    //bypass FFFFFFFF....
                    string sFF = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF";
                    string s00 = "000000000000000000000000000000000000000000000000000000000000";
                    if (clExchangeByteString.Bytes2HexStr(bRecord, false) == sFF.Substring(0, iRecordSize * 2) || clExchangeByteString.Bytes2HexStr(bRecord, false) == s00.Substring(0, iRecordSize * 2))
                    {
                        iOffset = iOffset + iRecordSize;
                        continue;
                    }
                    int icheck = 0;
                    for (int i = 0; i < iRecordSize; i++)
                    {
                        icheck = icheck + bRecord[i];
                    }
                    if (icheck == 0)
                    {
                        iOffset = iOffset + iRecordSize;
                        continue;
                    }
                    iCardLen = Convert.ToInt16((bRecord[0] & 0xF0) >> 4);
                    if (iCardLen > 13)
                        iCardLen = 13;
                    byte[] bCardNo = new byte[iCardLen];
                    DateTime uiTime = GetTimeFrom4Bytes(bRecord[2], bRecord[3], bRecord[4], bRecord[5]);
                    Array.Copy(bRecord, 7, bCardNo, 0, iCardLen);
                    string sCardNo = clExchangeByteString.ByteToString(bCardNo);
                    if (bRecord[7] >= 0x80)
                    {
                        int icard7len = bRecord[7] & 0x0F;
                        sCardNo = clExchangeByteString.Bytes2HexStr(bCardNo, false).Substring(2, icard7len);
                    }
                    //string sInOut = Convert.ToString((bRecord[0] & 6) >> 1);
                    string sInOut = Convert.ToString((bRecord[0] & 2) >> 1); //rac940 的 bit2 為0:一般卡,1:壓縮卡
                    string sShift = Convert.ToString(bRecord[6]);
                    string sInput = Convert.ToString(bRecord[0] & 1);
                    string sEventCode = "00" + Convert.ToInt16(bRecord[1]).ToString("X2");
                    string sEventType = Convert.ToString((bRecord[0] & 6) >> 1);
                    string sReserve = "0";
                    string soutput = "";

                    if (CallbackPubEvent != null)
                    {
                        clPubevent myEvent = new clPubevent();
                        myEvent.EventCard = sCardNo;
                        myEvent.EventCode = sEventCode;
                        myEvent.EventDate = uiTime.ToString("yyyy-MM-dd");
                        myEvent.EventTime = uiTime.ToString("HH:mm:ss");
                        myEvent.EventType = sEventType;
                        CallbackPubEvent(myEvent);
                    }
                    iOffset = iOffset + iRecordSize;

                }
                catch { }

            }

            return 0;
        }




        private int GetDeviceSysParam()
        {
            try
            {

                int iretlen = 0;
                int resultCode = -1;
                int timeout = 2000;
                byte[] dataBuffer = new byte[255];
                int dataLen = 0;
                resultCode = TRAC2000DLL.GetSysParaData(RAC2000Id, dataBuffer, ref dataLen, RAC2000Handle, timeout);
                AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "-->Modul(Dump96)System Parameter,Device:" + RAC2000Id + " Retrun Code:" + resultCode);
                return resultCode;
            }
            catch (Exception ex)
            {
                AppendLog(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + "-->Modul(Dump96),Device:" + RAC2000Id + " Exception:" + ex.StackTrace);
                return -1;
            }
        }

        private int GetDeviceVersion(ref string sVersion)
        {
            try
            {
                int timeout = 2000;
                byte[] dataBuffer = new byte[255];
                int iret = TRAC2000DLL.GetVersion(RAC2000Id, dataBuffer, RAC2000Handle, timeout);
                return iret;
            }
            catch (Exception ex)
            {
                return -1;
            }
        }




        private void AppendLog(string s)
        {
            if (Callbackmsg != null)
                Callbackmsg(s);
        }

        public DateTime GetTimeFrom4Bytes(byte inByte1, byte inByte2, byte inByte3, byte inByte4)
        {
            int sec = (inByte1 & 0x3F);
            int min = (((inByte2 & 0x0F) << 2) | ((inByte1 & 0xC0) >> 6));
            int hour = (((inByte3 & 0x01) << 4) | ((inByte2 & 0xF0) >> 4));
            int day = ((inByte3 & 0x3E) >> 1);
            int mon = (((inByte4 & 0x03) << 2) | ((inByte3 & 0xC0) >> 6));
            int year = (2000 + ((inByte4 & 0xFC) >> 2));

            DateTime resTime;
            try { resTime = new DateTime(year, mon, day, hour, min, sec); }
            catch { resTime = new DateTime(); }
            return resTime;
        }
    }
}

