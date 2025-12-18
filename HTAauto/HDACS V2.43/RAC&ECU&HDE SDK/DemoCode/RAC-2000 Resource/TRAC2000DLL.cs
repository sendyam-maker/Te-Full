using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace DEMO
{
    class TRAC2000DLL
    {
        //Open RAC2000 connect
        [DllImport("HDACS.dll", EntryPoint = "hacOpenChannel")]
        public static extern int OpenChannel(string sCommiPort, int cardinal, ref UInt32 hComm);
        // Close the communication channel
        [DllImport("HDACS.dll", EntryPoint = "hacCloseChannel")]
        public static extern int CloseChannel(UInt32 hComm);

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



        // Read data from Flash
        [DllImport("HDACS.dll", EntryPoint = "hacGetFlashData")]
        public static extern int GetFlashData(int iNodeID, byte[] cFlashData, ref int iReceiveDataLen, int iFlashAddr, int iFlashLen, UInt32 hComm, int iTimeout);
        // Write data in Flash
        [DllImport("HDACS.dll", EntryPoint = "hacSetFlashData")]
        public static extern int SetFlashData(int iNodeID, byte[] cFlashData, int iFlashAddr, int iFlashLen, UInt32 hComm, int iTimeout);
        // Read data from Para
        [DllImport("HDACS.dll", EntryPoint = "hacGetParaData")]
        public static extern int GetParaData(int iNodeID, byte[] cFlashData, ref int iReceiveDataLen, UInt32 hComm, int iTimeout);
        // Read data from system Para
        [DllImport("HDACS.dll", EntryPoint = "hacGetSysParaData")]
        public static extern int GetSysParaData(int iNodeID, byte[] cFlashData, ref int iReceiveDataLen, UInt32 hComm, int iTimeout);
        // Write data in system Para
        [DllImport("HDACS.dll", EntryPoint = "hacSetSysParaData")]
        public static extern int SetSysParaData(int iNodeID, byte[] cFlashData, int iFlashLen, UInt32 hComm, int iTimeout);
        // Read Card from Mifare
        [DllImport("HDACS.dll", EntryPoint = "hacGetMifare")]
        public static extern int GetMifare(int iNodeID, ref  int iKeyType, ref int iBlock, ref int iStartDigit, ref int iDigitLength, ref int iCompact, UInt32 hComm, int iTimeout);
        // Read Card from Mifare
        [DllImport("HDACS.dll", EntryPoint = "hacSetMifare")]
        public static extern int SetMifare(int iNodeID, int iKeyType, int iBlock, int iStartDigit, int iDigitLength, int iCompact, byte[] cKeyValue, UInt32 hComm, int iTimeout);
        // Add the compressed Card NO.
        [DllImport("HDACS.dll", EntryPoint = "hacAddCardEX")]
        public static extern int AddCardEX(int iNodeID, string cCardNo, int iCardLen, string cPassword, int iPassLen, string cName, int iNameLen, int iTimeZone, char cStatus, UInt32 hComm, int iTimeout);



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

        [DllImport("HDACS.dll", EntryPoint = "hacAddCardFingerPrint")]
        public static extern int hacAddCardFingerPrint(int iNodeID, string cCardNo, int iCardLen, string cPassword, int iPassLen, int iTimeZone, int cStatus, byte[] cFingerPrinterData1, byte[] cFingerPrinterData2, uint hComm, int iTimeout);
        [DllImport("HDACS.dll", EntryPoint = "hacAddCardFingerPrintEx")]
        public static extern int hacAddCardFingerPrintEx(int iNodeID, string cCardNo, int iCardLen, string cPassword, int iPassLen, string cName, int iNameLen, int iTimeZone, int cStatus, byte[] cFingerPrinterData1, byte[] cFingerPrinterData2, uint hComm, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hacFingerPrinterQueryUser")]
        public static extern int hacFingerPrinterQueryUser(int iNodeID, UInt32 hComm, int CardLen, string cCardNo, byte[] cFingerPrinterData1, byte[] cFingerPrinterData2, ref int iCardFormatLen, ref int iReturnCode, int iTimeOut);

        [DllImport("HDACS.dll", EntryPoint = "hsECUReadIO")]
        public static extern int hsECUReadIO(UInt32 hComm, int iECUID, ref int cSensor, ref int cRelay, ref int iReturnCode, int iTimeout);
        [DllImport("HDACS.dll", EntryPoint = "hsECUReadParamenter")]
        public static extern int hsECUReadParamenter(UInt32 hComm, int iECUID, byte[] cParaData, ref int iParaLen, ref int iReturnCode, int iTimeout);
        [DllImport("HDACS.dll", EntryPoint = "hsECUReadPower")]
        public static extern int hsECUReadPower(UInt32 hComm, int iECUID, ref byte cPower, byte[] cCardNo, ref int iReturnCode, int iTimeout);
        [DllImport("HDACS.dll", EntryPoint = "hsECUAddCard")]
        public static extern int hsECUAddCard(UInt32 hComm, int iECUID, string cCardNo, int iCardLen, int iTimeZone, int Validity, int ValidityYear, int ValidityMouth, int ValidityyDay, int ValidityHour, int ValidityMinute, int iTimeout);
        [DllImport("HDACS.dll", EntryPoint = "hsECUPolling")]
        public static extern int hsECUPolling(UInt32 hComm, int iELID, int iPrevRecord, byte[] stRecord, ref int iRecord, ref int iReturnCode, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "hacDumpLegalCard")]
        public static extern int hacDumpLegalCard(int iNodeID, UInt32 hComm, string cBinFilename, ref int iReturnCode, int iTimeOut);

        [DllImport("HDTAS.dll", EntryPoint = "hsHTA850DumpFile")]
        public static extern int hsHTA850DumpFile(UInt32 hComm, string cBinFilename, ref int iReturnCode, int iTimeOut);
        //2014.2.17 Add
        //RAC2000_API int __stdcall hacHWReadCommand(int iNodeID,int iIndex,unsigned char *cSendData,int iSendDataLen,
        //                                   unsigned char *cReceiveData,int *iReceiveLen,HANDLE hComm,unsigned int iTimeout)
        //iCommandType=0 Read Command, 1:Write Command, 2:ReadWrite Command
        [DllImport("HDACS.dll", EntryPoint = "hacHWRWCommandCCH")]
        public static extern int hacHWRWCommandCCH(int iCommandType, int iNodeID, int iIndex, byte[] cSendData, int iSendDataLen, byte[] cReceiveData, ref int iReceiveLen, UInt32 hComm, int iTimeOut);
        [DllImport("HDACS.dll", EntryPoint = "hacGetDesFire")]
        public static extern int hacHWRWCommandCCH(int iNodeID, ref int iKeyType, ref int ReadOffset, ref int ReadLength, byte[] ApplicationID, ref int FileID, ref int KeyNo, UInt32 hComm, int iTimeOut);
        [DllImport("HDACS.dll", EntryPoint = "hacSetDesFire")]
        public static extern int hacHWRWCommandCCH(int iNodeID, ref int iKeyType, ref int ReadOffset, ref int ReadLength, byte[] ApplicationID, ref int FileID, ref int KeyNo, byte[] cKeyDate, UInt32 hComm, int iTimeOut);
        #region RAC340/520
        // Send out the order of Polling to the device
        [DllImport("HDACS.dll", EntryPoint = "hac34Polling")]
        public static extern int hac34Polling(int iNodeId, int iPrevRecord, byte[] stRecord, ref int iRecord, UInt32 hComm, int iTimeout, int iCardType);
        // Read datetime from equipment
        [DllImport("HDACS.dll", EntryPoint = "hac34GetDateTime")]
        public static extern int hac34GetDateTime(int iNodeID, byte[] cDate, byte[] cTime, UInt32 hComm, int iTimeout);
        // Write datetime in equipment
        [DllImport("HDACS.dll", EntryPoint = "hac34SetDateTime")]
        public static extern int hac34SetDateTime(int iNodeID, string cDate, string cTime, UInt32 hComm, int iTimeout);
        // Read data from EEPROM
        [DllImport("HDACS.dll", EntryPoint = "hac34GetEEData")]
        public static extern int hac34GetEEData(int iNodeID, byte[] cEEData, ref int iReceiveDataLen, int iEEArea, int iEEAddr, int iEELen, UInt32 hComm, int iTimeout);
        [DllImport("HDACS.dll", EntryPoint = "hac34SetEEData")]
        public static extern int hac34SetEEData(int iNodeID, byte[] cEEData, int iEEArea, int iEEAddr, int iEELen, UInt32 hComm, int iTimeout);
        // Get the device Version
        [DllImport("HDACS.dll", EntryPoint = "hac34GetVersion")]
        public static extern int hac34GetVersion(int iNodeID, byte[] cData, UInt32 hComm, int iTimeout);
        // Add the Card NO.(Not Compressed Card)
        [DllImport("HDACS.dll", EntryPoint = "hac34AddCard")]
        public static extern int hac34AddCard(int iNodeID, string cCardNo, int iCardLen, string ipassword, int passlen, int iTimeZone, char cStatus, UInt32 hComm, int iTimeout);
        // Delete the Card NO.(Not Compressed Card)
        [DllImport("HDACS.dll", EntryPoint = "hac34DelCard")]
        public static extern int hac34DelCard(int iNodeID, string cCardNo, int iCardLen, UInt32 hComm, int iTimeout);
        // Delete All Card
        [DllImport("HDACS.dll", EntryPoint = "hac34DelAllCard")]
        public static extern int hac34DelAllCard(int iNodeID, UInt32 hComm, int iTimeout);
        //Read Card Parameter
        [DllImport("HDACS.dll", EntryPoint = "hac34GetReadCardParameter")]
        public static extern int hac34GetReadCardParameter(int iNodeID, ref int iKeyType, ref int iBlock, ref int iStartDigit,
            ref int iDigitLength, ref int iCompact, UInt32 hComm, int iTimeout);
        //Set Card Parameter
        [DllImport("HDACS.dll", EntryPoint = "hac34SetReadCardParameter")]
        public static extern int hac34SetReadCardParameter(int iNodeID, int iKeyType, int iBlock, int iStartDigit,
            int iDigitLength, int iCompact, byte[] cKeyValue, UInt32 hComm, int iTimeout);

        [DllImport("HDACS.dll", EntryPoint = "ret_time")]
        public static extern int ret_time(uint utime, ref  RTC_STRUCT_ret_time str);


        [DllImport("HDACS.dll", EntryPoint = "make7byteCardNum")]
        public static extern int make7byteCardNum(string cCardNo, int iCardLen, byte[] c7bytesCardNo);

        #endregion
    }
}
