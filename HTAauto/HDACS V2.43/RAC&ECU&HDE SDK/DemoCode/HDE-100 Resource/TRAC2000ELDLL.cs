using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace DEMO
{
    class TRAC2000ELDLL
    {
        //Open RAC2000 connect
        [DllImport("HDACS.dll", EntryPoint = "hsELOpenChannel")]
        public static extern int OpenChannel(ref UInt32 hComm, string sCommiPort, int cardinal);
        //[DllImport("HDACS.dll", EntryPoint = "hacOpenChannel")]
        //public static extern int OpenChannel(string sCommiPort, int cardinal, ref UInt32 hComm);
        // Close the communication channel
        [DllImport("HDACS.dll", EntryPoint = "hsELCloseChannel")]
        public static extern int CloseChannel(UInt32 hComm);

        [DllImport("HDACS.dll", EntryPoint = "hsELWriteTable")]
        public static extern int WriteTable(UInt32 hComm, int iELID, byte[] cReceiveData, int iReceiveLen, ref int iReturnCode, int iTimeOut);

        [DllImport("HDACS.dll", EntryPoint = "hsELReadTable")]
        public static extern int ReadTable(UInt32 THandle, int iELID, byte[] cSendData,ref int iSendLen, ref int iRecutrnCode, int iTimeOut);

        [DllImport("HDACS.dll", EntryPoint = "hsELPolling")]
        public static extern int Polling(UInt32 hComm, int iELID, int iPrevRecord, byte[] stRecord, ref int iRecord,ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELReadParameter")]
        public static extern int hsELReadParameter(UInt32 hComm, int iELID, byte[] cParaData, ref int iParaLen, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELWriteParameter")]
        public static extern int hsELWriteParameter(UInt32 hComm, int iELID, byte[] cTableData, int iParaLen, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELInitialize")]
        public static extern int Initialize(UInt32 hComm, int iELID, string cInitFlag, ref int iReturnCode, int iTimeOut);

        [DllImport("HDACS.dll", EntryPoint = "hsELGetInfo")]
        public static extern int GetELInfo(UInt32 hComm, int iELID, byte[] cInfoData, ref int iInfoLen, ref int iReturnCode, int iTimeOut);

        [DllImport("HDACS.dll", EntryPoint = "hsELAddAuthorization")]
        public static extern int hsELAddAuthorization(UInt32 hComm, int iELID, int iCardCount, byte[] cCardNo, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELDeleteAuthorization")]
        public static extern int hsELDeleteAuthorization(UInt32 hComm, int iELID, string cCardNo, ref int iReturnCode, int iTimeout);
        //查询单笔卡号权限
        [DllImport("HDACS.dll", EntryPoint = "hsELQueryAuthorization")]
        public static extern int hsELQueryAuthorization(UInt32 hComm, int iELID, string cCardNo, byte[] stRecord, ref int iReadLen, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELDeleteAllAuthorization")]
        public static extern int hsELDeleteAllAuthorization(UInt32 hComm, int iELID, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELReadDeviceInfo")]
        public static extern int hsELReadDeviceInfo(UInt32 hComm, int iELID, byte[] stRecord, ref int iInfoLen, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELSetTime")]
        public static extern int SetDateTime(UInt32 hComm, int iELID, string cDate, string cTime, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELGetTime")]
        public static extern int GetDateTime(UInt32 hComm, int iELID, byte[] cDate, byte[] cTime, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELReleaseAlarm")]
        public static extern int hsELReleaseAlarm(UInt32 hComm, int iELID, string cAlarmFlag, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hsELPublicFloor")]
        public static extern int hsELPublicFloor(UInt32 hComm, int iELID, int iTotal, byte[] cFloorData, ref int iReadLen, ref int iReturnCode, int iTimeout);


        /*

        // Delete the Card NO.(Not Compressed Card)
        [DllImport("HDACS.dll", EntryPoint = "hacDelCard")]
        public static extern int DelCard(int iNodeID, string cCardNo, int iCardLen, UInt32 hComm, int iTimeout);
        // Add the compressed Card NO.
        [DllImport("HDACS.dll", EntryPoint = "hacAddZCard")]
        public static extern int AddZCard(int iNodeID, string cCardNo, int iCardLen, string cPassword, int iPassLen, char cStatus, UInt32 hComm, int iTimeout);
        // Delete the compressed Card NO.
        [DllImport("HDACS.dll", EntryPoint = "hacDelZCard")]
        public static extern int DelZCard(int iNodeID, string cCardNo, int iCardLen, UInt32 hComm, int iTimeout);
        // Make Relay Action
        [DllImport("HDACS.dll", EntryPoint = "hacRelayAction")]
        public static extern int RelayAction(int iNodeID, char cAction, char cMask, UInt32 hComm, int iTimeout);
        // Read datetime from equipment
        [DllImport("HDACS.dll", EntryPoint = "hacGetDateTime")]
        public static extern int GetDateTime(int iNodeID, byte[] cDate, byte[] cTime, UInt32 hComm, int iTimeout);
        // Write datetime in equipment
        [DllImport("HDACS.dll", EntryPoint = "hacSetDateTime")]
        public static extern int SetDateTime(int iNodeID, string cDate, string cTime, UInt32 hComm, int iTimeout);
        // Read data from EEPROM
        [DllImport("HDACS.dll", EntryPoint = "hacGetEEData")]
        public static extern int GetEEData(int iNodeID, byte[] cEEData, ref int iReceiveDataLen, int iEEAddr, int iEELen, UInt32 hComm, int iTimeout);
        // Write data in EEPROM
        [DllImport("HDACS.dll", EntryPoint = "hacSetEEData")]
        public static extern int SetEEData(int iNodeID, byte[] cEEData, int iEEAddr, int iEELen, UInt32 hComm, int iTimeout);
        // Read the RAM data
        [DllImport("HDACS.dll", EntryPoint = "hacGetRAMData")]
        public static extern int GetRAMData(int iNodeID, byte[] cRAMData, ref int iReceiveDataLen, int iRAMAddr, int iRAMLen, UInt32 hComm, int iTimeout);
        // Write data in RAM
        [DllImport("HDACS.dll", EntryPoint = "hacSetRAMData")]
        public static extern int SetRAMData(int iNodeID, byte[] cRAMData, int iRAMAddr, int iRAMLen, UInt32 hComm, int iTimeout);
        // Clear the data in buffer
        [DllImport("HDACS.dll", EntryPoint = "hacClearBuffer")]
        public static extern int ClearBuffer(UInt32 hComm);
        // Read the data from the buffer
        [DllImport("HDACS.dll", EntryPoint = "hacReadData")]
        public static extern int ReadData(byte[] cBuffer, ref int iDataLen, UInt32 hComm, int iTimeout);
        // Write data in buffer
        [DllImport("HDACS.dll", EntryPoint = "hacWriteData")]
        public static extern int WriteData(byte[] cBuffer, int iDataLen, ref int iWrittenlen, UInt32 hComm);
        // Get the device Version
        [DllImport("HDACS.dll", EntryPoint = "hacGetVersion")]
        public static extern int GetVersion(int iNodeID, byte[] cData, UInt32 hComm, int iTimeout);
        // Send out the order of Polling to the device
        [DllImport("HDACS.dll", EntryPoint = "hacPolling")]
        public static extern int Polling(int iNodeId, int iPrevRecord, byte[] stRecord, ref int iRecord, UInt32 hComm, int iTimeout, int iCardType);
        // Get the device Sensor
        [DllImport("HDACS.dll", EntryPoint = "hacGetSensor")]
        public static extern int GetSensor(int iNodeID, ref int iSensor, UInt32 hComm, int iTimeout);
        // Add the Card NO.(Not Compressed Card)
        [DllImport("HDACS.dll", EntryPoint = "hacAddCard")]
        public static extern int AddCard(int iNodeID, string cCardNo, int iCardLen, string ipassword, int passlen, int iTimeZone, char cStatus, UInt32 hComm, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hacAddVisitorCard")]
        public static extern int AddVisitorCard(int iNodeID, string cCardNo, int iCardStart, int iCardLen, string cStartDate, string cStartTime, string cEndDate, string cEndTime, int iWeek, int iTimes, int iSerial, int cStatus, UInt32 hComm, int iTimeout);
        [DllImport("HDACS.dll", EntryPoint = "hacDelVisitorCard")]
        public static extern int DelVisitorCard(int iNodeID, int iSerial, UInt32 hComm, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hacSetLanMode")]
        public static extern int hacSetLanMode(int iNodeId, byte cMode, UInt32 hComm, int iTimeout);
         */ 
    }
}
