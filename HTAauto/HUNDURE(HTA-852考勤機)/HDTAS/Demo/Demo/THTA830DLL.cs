using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace Demo
{
    //HTA830收取的刷卡资料格式


    public struct EventFormat
    {
        public string ClassCode;		//Class Code
        public string IllegalCode;	//Illegal Code
        public string sDateTime;	//Date Time[20]
        public string sCard;		//Card Number[20]
        public string sDeviceID;	//Device ID[10]
    };

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    //[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
    public struct struct_FingerPrinterFormat2
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
        public byte[] Card;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public byte[] Stay;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
        public byte[] DisplayMsg;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 386)]
        public byte[] FingerPrinter1;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 386)]
        public byte[] FingerPrinter2;
        //char Card[16];
        //char Stay[2];
        //char DisplayMsg[16];
        //unsigned char FingerPrinter1[386];
        //unsigned char FingerPrinter2[386];

    }
    class THTA830DLL
    {
        #region DLL函数引用
        //打开Socket

        [DllImport("HDTAS.dll", EntryPoint = "htaOpenChannel")]
        public static extern int HUNHTAOpenSocket(ref int hComm, string ip, int port);
        //关闭Socket
        [DllImport("HDTAS.dll", EntryPoint = "htaCloseChannel")]
        public static extern int HUNHTACloseSocket(int hComm);
        //获取时间
        [DllImport("HDTAS.dll", EntryPoint = "htaGetDateTime")]
        public static extern int HUNReadHTADateTime(int hComm, int inodeid, byte[] date, byte[] time, uint timeout);
        //写HTA8时间
        [DllImport("HDTAS.dll", EntryPoint = "htaSetDateTime")]
        public static extern int HUNWriteHTADateTime(int hComm, int inodeid, string date, string time, uint timeout);
        //读取HTA版本
        [DllImport("HDTAS.dll", EntryPoint = "htaGetVersion")]
        public static extern int HUNReadHTAVersion(int hComm, int inodeid, byte[] data, uint timeout);

        //加入合法卡

        [DllImport("HDTAS.dll", EntryPoint = "htaAddCard")]
        public static extern int HUNAddHTACard(int hComm, int inodeid, string cardNo, int cardlen, uint timeout);

        //单笔删除合法卡

        [DllImport("HDTAS.dll", EntryPoint = "htaDelCard")]
        public static extern int HUNDelHTACard(int hComm, int inodeid, string cardNo, int cardlen, uint timeout);
        //加入压缩卡


        [DllImport("HDTAS.dll", EntryPoint = "htaAddZCard")]
        public static extern int HUNAddHTAZCard(int hComm, int inodeid, string cardNo, int cardlen, uint timeout);
        //删除压缩卡


        [DllImport("HDTAS.dll", EntryPoint = "htaDelZCard")]
        public static extern int HUNDelHTAZCard(int hComm, int inodeid, string cardNo, int cardlen, uint timeout);
        //删除所有合法卡
        [DllImport("HDTAS.dll", EntryPoint = "htaDeleteAllCard")]
        public static extern int HUNDelHTAAllCard(int hComm, int inodeid, uint timeout);
        //获取装置Memory值


        [DllImport("HDTAS.dll", EntryPoint = "htaGetMemoryData")]
        public static extern int HUNGetHTAMemoryData(int hComm, int inodeid, byte[] Memdata, ref int iReceivelen, int iMemaddr, int iMEMLen, uint timeout);
        //写装置Memory值


        [DllImport("HDTAS.dll", EntryPoint = "htaSetMemoryData")]
        public static extern int HUNSetHTAMemoryData(int hComm, int inodeid, byte[] Memdata, int iMemaddr, int iMEMLen, uint timeout);
        //获取HTA-Log数据
        [DllImport("HDTAS.dll", EntryPoint = "htaGetLogData")]
        public static extern int HUNGetHTALogData(int hComm, int inodeid, byte[] logdata, ref int iLoglen, int iBank, int icompress, uint timeout);
        //获取HTA-Card数据
        [DllImport("HDTAS.dll", EntryPoint = "htaGetCardData")]
        public static extern int HUNGetHTACardData(int hComm, int inodeid, byte[] logdata, ref int iReceivelen, int iBank, int icompress, uint timeout);
        //清除flash-初始化HTA
        [DllImport("HDTAS.dll", EntryPoint = "htaEraseFlash")]
        public static extern int HUNEraseHTAFlash(int hComm, int inodeid, uint timeout);
        //删除HTA所有的Log
        [DllImport("HDTAS.dll", EntryPoint = "htaDeleteAllLog")]
        public static extern int HUNDeleteHTAAllLog(int hComm, int inodeid, uint timeout);

        //重启HTA
        [DllImport("HDTAS.dll", EntryPoint = "htaRestart")]
        public static extern int HUNRestartHTA(int hComm, int inodeid, uint timeout);
        //Declare Function htaGetLogRecord Lib "HTA8.dll" (ByVal hComm As Integer, ByVal inodeid As Integer, ByVal iBank As Integer, ByRef bAa As Byte, _
        //  ByRef irecord As Integer, ByVal icardtype As Integer, ByVal itimeout As Integer) As Integer
        [DllImport("HDTAS.dll", EntryPoint = "htaGetLogRecord")]
        public static extern int HUNGetHTALogRecord(int hComm, int inodeid, int iBank, ref EventFormat[] stRecord, ref int iRecord, int icardType, uint timeout);
        //polling读取刷卡资料
        [DllImport("HDTAS.dll", EntryPoint = "htaPolling")]
        public static extern int HUNHTAPolling(int hComm, int inodeid, int iprevRecord, byte[] stRecord, ref int iRecord, int icardType, uint timeout);
        #endregion

        // 2010/06/22 Netpool Add --- Start ---
        // 加入 HTA-860 的 SDK Demo Code 所必需新增加的 Function
        //int __stdcall hsSetGcuID(int aGcuid)
        [DllImport("HDTAS.dll", EntryPoint = "hsSetGcuID")]
        public static extern int SetGcuID(int aGcuid);

        // int __stdcall hsHTA850GetFPInfo(HANDLE hComm,unsigned char *cInfoData,int *iInfoLen,int *iReturnCode,unsigned int iTimeOut)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850GetFPInfo")]
        public static extern int HTA850GetFPInfo(int hComm, byte[] cInfoData, ref int iInfoLen, ref int iReturnCode, uint iTimeOut);

        // int __stdcall hsHTA850QueryMasterFP(HANDLE hComm, LPBYTE cFingerPrinterData1,unsigned char *cFingerPrinterData2,int *iReturnCode,unsigned int iTimeOut)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850QueryMasterFP")]
        public static extern int HTA850QueryMasterFP(int hComm, byte[] cFingerPrinterData1, byte[] cFingerPrinterData2, ref int iReturnCode, uint iTimeOut);

        //int __stdcall hsHTA850UpdateMasterFP(HANDLE hComm, LPBYTE cFingerPrinterData1,unsigned char *cFingerPrinterData2,int *iReturnCode,unsigned int iTimeOut)
        //int __stdcall hsHTA850UpdateMasterFP(HANDLE hComm,unsigned char *cFingerPrinterData1,unsigned char *cFingerPrinterData2,int *iReturnCode,unsigned int iTimeOut)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850UpdateMasterFP")]
        public static extern int HTA850UpdateMasterFP(int hComm, byte[] cFingerPrinterData1, byte[] cFingerPrinterData2, ref int iReturnCode, uint iTimeOut);

        // int __stdcall hsHTA850ReadTableEx(HANDLE hComm,int iTable,unsigned char *cTableData,int *iTableLen,int * iReturnCode,unsigned int iTimeout)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850ReadTableEx")]
        public static extern int HTA850ReadTableEx(int hComm, int iTable, byte[] cTableData, ref int iTableLen, ref int iReturnCode, uint iTimeout);

        // int __stdcall hsHTA850WriteTableEx(HANDLE hComm,int iTable,unsigned char *cTableData,int iTableLen,int * iReturnCode,unsigned int iTimeout)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850WriteTableEx")]
        public static extern int HTA850WriteTableEx(int hComm, int iTable, byte[] cTableData, int iTableLen, ref int iReturnCode, uint iTimeout);

        // int __stdcall hsHTA850WriteSRAM(HANDLE hComm,int iaddress,unsigned char *cTableData,int iTableLen,int * iReturnCode,unsigned int iTimeout)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850WriteSRAM")]
        public static extern int HTA850WriteSRAM(int hComm, int iaddress, byte[] cTableData, int iTableLen, ref int iReturnCode, uint iTimeout);

        // int __stdcall hsHTA850ReadSRAM(HANDLE hComm,int iAddress,unsigned char *cTableData,int *iTableLen,int * iReturnCode,unsigned int iTimeout)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850ReadSRAM")]
        public static extern int HTA850ReadSRAM(int hComm, int iAddress, byte[] cTableData, ref int iTableLen, ref int iReturnCode, uint iTimeout);
        // 2010/06/22 Netpool Add --- End ---

        // 2011/05/03 Netpool Add --- Start ---
        // 860 Add Card
        // int __stdcall hsHTA850InsertMultiUserRecord(HANDLE hComm,int CardLen,int MsgLen,int iRecord,struct_CardFormat *stRecord,int *iReturnCode,unsigned int iTimeOut);
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850InsertMultiUserRecord")]
        public static extern int HTA850InsertMultiUserRecord(int hComm, int CardLen, int MsgLen, int iRecord, byte[] stRecord, ref int iReturnCode, uint iTimeOut);

        // 860 Delete Card
        // int __stdcall hsHTA850DeleteUserRecord(HANDLE hComm,int CardLen,char *cCardNo,int *iReturnCode,unsigned int iTimeOut)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850DeleteUserRecord")]
        public static extern int HTA850DeleteUserRecord(int hComm, int CardLen, byte[] cCardNo, ref int iReturnCode, uint iTimeOut);

        // 860 Polling Data
        // int __stdcall hsHTA850PollingData(HANDLE hComm,int iPrevRecord,stPollRecord *stRecord,int *iRecord,int * iReturnCode,unsigned int iTimeout)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850PollingData")]
        public static extern int HTA850PollingData(int hComm, int iPrevRecord, byte[] stRecord, ref int iRecord, ref int iReturnCode, uint iTimeOut);
        // 2011/05/03 Netpool Add --- End ---

        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850RemoteOpen")]
        public static extern int hsHTA850RemoteOpen(int hComm, int openSec, ref int iReturnCode, uint iTimeout);

        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850QueryUserFingerPrinter2")]
        public static extern int hsHTA850QueryUserFingerPrinter2(int hComm, int CardLen, string cCardNo, byte[] cFingerPrinterData1, byte[] cFingerPrinterData2, ref int iCardFormatLen, ref int iReturnCode, uint iTimeOut);

        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850InsertMultiUserFingerPrinter2")]
        public static extern int hsHTA850InsertMultiUserFingerPrinter2(int hComm, int CardLen, int MsgLen, int iRecord, byte[] stRecord, ref int iReturnCode, uint iTimeOut);

        //[DllImport("HDTAS.dll", EntryPoint = "hsHTA850InsertMultiUserFingerPrinter2")]
        //public static extern int hsHTA850InsertMultiUserFingerPrinter2Struc(int hComm, int CardLen, int MsgLen, int iRecord, ref  struct_FingerPrinterFormat2 stRecord, ref int iReturnCode, int iTimeOut);
        //int __stdcall hsHTA850ReadEEPROM (HANDLE hComm,unsigned char *cEESendData,int iEESendLen,unsigned char *cEEReceiveData,int *iEEReceiveLen,int *iReturnCode,unsigned int iTimeout)
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850ReadEEPROM")]
        public static extern int hsHTA850ReadEEPROM(int hComm, byte[] cEESendData, int iEESendLen, byte[] cEEReceiveData, ref int iEEReceiveLen, ref  int iReturnCode, uint iTimeOut);
        
        //int __stdcall hsHTA850Set(HANDLE hComm,int iFunction,unsigned char *cSendData,int iSendLen,unsigned char *cReceiveData,int *iReceiveLen,int *iReturnCode,unsigned int iTimeOut)
        //iReturn=hsHTA850Set(hComm,10,cOutBuff,4,cReceiveData,&iReceiveLen,iReturnCode,iTimeout);
        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850Set")]
        public static extern int hsHTA850Set(int hComm, int iFunction, byte[] cSendData, int iSendLen, byte[] cReceiveData, ref int iReceiveLen, ref  int iReturnCode, uint iTimeOut);

        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850Initial")]
        public static extern int hsHTA850Initial(int hComm,  byte cInitFlag, ref  int iReturnCode, uint iTimeOut);
      
    }
}
