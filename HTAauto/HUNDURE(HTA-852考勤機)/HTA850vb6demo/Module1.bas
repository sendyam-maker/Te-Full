Attribute VB_Name = "Module1"
    
    Public Type TEvent
         sDate As String
         sTime As String
         Reader As Byte
         InputType As Byte
         ASection As Byte
         AClass As Byte
         EventCode As Byte
         Card As String
    End Type
    'Connect To HTA850
    Declare Function hsOpenChannel Lib "HDTAS.dll" (ByRef hComm As Integer, ByVal sComm As String, ByVal iport As Integer) As Integer

    'Close HTA850 Connction
    Declare Function hsCloseChannel Lib "HDTAS.dll" (ByVal hComm As Integer) As Integer
    'Initial HTA850
    'int __stdcall hsHTA850Initial(HANDLE hComm,char cIniFlag,int *iReturnCode,unsigned int iTimeOut)
    Declare Function hsHTA850Initial Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal cInitFlag As Byte, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
    'Get Basic HTA850 Information
    'int __stdcall hsHTA850GetInfo(HANDLE hComm,unsigned char *cInfoData,int *iInfoLen,int *iReturnCode,unsigned int iTimeOut)
    Declare Function hsHTA850GetInfo Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cInfodata As Byte, ByRef iinfolen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850ReadParameter(HANDLE hComm,unsigned char *cParaData,int *iParaLen,int *iReturnCode,unsigned int iTimeOut)
    'Read HTA850 Params
    Declare Function hsHTA850ReadParameter Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cParaData As Byte, ByRef iparalen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
    'int __stdcall hsHTA850WriteParameter(HANDLE hComm,unsigned char *cParaData,int iParaLen,int *iReturnCode,unsigned int iTimeOut)
    'Write HTA850 Params
    Declare Function hsHTA850WriteParameter Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cParaData As Byte, ByVal iparalen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850WriteTime(HANDLE hComm,char * cDate,char *cTime,int *iReturnCode,unsigned int iTimeOut)
    ' Set the host time
    Declare Function hsHTA850WriteTime Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal sDate As String, ByVal sTime As String, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
    'int __stdcall hsHTA850ReadTime(HANDLE hComm,int iGCUID,char *cDate,char *cTime,int * iReturnCode,unsigned int iTimeout)
    Declare Function hsHTA850ReadTime Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal sDate As String, ByVal sTime As String, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850WriteTable(HANDLE hComm,int iTable,unsigned char *cTableData,int iTableLen,int * iReturnCode,unsigned int iTimeout)
    Declare Function hsHTA850WriteTable Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal itable As Integer, ByRef cTableData As Byte, ByVal itablelen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850ReadTable(HANDLE hComm,int iTable,unsigned char *cTableData,int *iTableLen,int * iReturnCode,unsigned int iTimeout)
    Declare Function hsHTA850ReadTable Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal itable As Integer, ByRef cTableData As Byte, ByRef itablelen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850InsertMultiUserRecord(HANDLE hComm,int CardLen,int MsgLen,int iRecord,struct_CardFormat *stRecord,int *iReturnCode,unsigned int iTimeOut)
    Declare Function hsHTA850InsertMultiUserRecord Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByVal MsgLen As Integer, ByVal iRecord As Integer, ByRef stRecord As Byte, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850DeleteUserRecord(HANDLE hComm,int CardLen,char *cCardNo,int *iReturnCode,unsigned int iTimeOut)

    Declare Function hsHTA850DeleteUserRecord Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByRef cCardNo As Byte, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850QueryUserRecord(HANDLE hComm,int CardLen,char *cCardNo,unsigned char *cCardFormatData,int *iCardFormatLen,int *iReturnCode,unsigned int iTimeOut)
    Declare Function hsHTA850QueryUserRecord Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByRef cCardNo As Byte, ByRef cCardFormatData As Byte, ByRef iCardFormatLen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850DeleteAllUserRecord(HANDLE hComm,int *iReturnCode,unsigned int iTimeOut)
    Declare Function hsHTA850DeleteAllUserRecord Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850ReadEEPROM(HANDLE hComm,unsigned char *cEESendData,int iEESendLen,unsigned char * cEEReceiveData,int *iEEReceiveLen,int *iReturnCode,unsigned int iTimeOut)
    Declare Function hsHTA850ReadEEPROM Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cEESendData As Byte, ByVal iEESendLen As Integer, ByRef cEEReceiveData As Byte, ByRef iEEReceiveLen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850SetEEPROM(HANDLE hComm,unsigned char * cEEData,int iEELen,int *iReturnCode,unsigned int iTimeOut)
    Declare Function hsHTA850SetEEPROM Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cEEData As Byte, ByVal iEELen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850SetMifareReader(HANDLE hComm,unsigned char *cData,int iLen,int *iReturnCode,unsigned int iTimeOut)
    Declare Function hsHTA850SetMifareReader Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cdata As Byte, ByVal iLen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer

    'int __stdcall hsHTA850PollingData(HANDLE hComm,int iPrevRecord,stPollRecord *stRecord,int *iRecord,int * iReturnCode,unsigned int iTimeout)
    Declare Function hsHTA850PollingData Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal iPrevRecord As Integer, ByRef stRecord As Byte, ByRef iRecord As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
    
    'int __stdcall hsHTA850InsertMultiUserFingerPrinter(HANDLE hComm,int CardLen,int MsgLen,int iRecord,struct_FingerPrinterFormat *stRecord,int *iReturnCode,unsigned int iTimeOut);
    Declare Function hsHTA850InsertMultiUserFingerPrinter Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByVal MsgLen As Integer, ByVal iRecord As Integer, ByRef stRecord As Byte, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
    Declare Function hsHTA850InsertMultiUserFingerPrinter2 Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByVal MsgLen As Integer, ByVal iRecord As Integer, ByRef stRecord As Byte, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
    
    'int __stdcall hsHTA850QueryUserFingerPrinter(HANDLE hComm,int CardLen,char *cCardNo,unsigned char *cFingerPrinterData1,unsigned char *cFingerPrinterData2,int *iCardFormatLen,int *iReturnCode,unsigned int iTimeOut);
    Declare Function hsHTA850QueryUserFingerPrinter Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByRef cCardNo As Byte, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByRef iCardFormatLen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
    Declare Function hsHTA850QueryUserFingerPrinter2 Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByRef cCardNo As Byte, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByRef iCardFormatLen As Integer, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
   
    Declare Function hsHTA850UpdateMasterFP Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
    Declare Function hsHTA850QueryMasterFP Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByRef ireturncode As Integer, ByVal iTimeout As Integer) As Integer
    
    Public ghComm As Integer
    
