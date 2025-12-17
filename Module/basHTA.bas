Attribute VB_Name = "basHTA"
Option Explicit
    
Public HTAip As String
Public HTAips As String
Public HTAips2 As String 'Added by Morgan 2020/9/3 門禁機IP
Private Const HTAport As Integer = 4660

'Added by Morgan 2020/8/25
Public pubIsNewDevice As Boolean '是否新考勤機(門禁機)
Private Const cNodeID As Integer = 1
'end 2020/8/25

Public Type tEvent
     Sdate As String
     sTime As String
     Reader As Byte
     InputType As Byte
     ASection As Byte
     AClass As Byte
     EventCode As Byte
     Card As String
End Type
'Connect To HTA850
Declare Function hsOpenChannel Lib "HDTAS.dll" (ByRef hComm As Integer, ByVal sComm As String, ByVal iPort As Integer) As Integer

'Close HTA850 Connction
Declare Function hsCloseChannel Lib "HDTAS.dll" (ByVal hComm As Integer) As Integer
'Initial HTA850
'int __stdcall hsHTA850Initial(HANDLE hComm,char cIniFlag,int *iReturnCode,unsigned int iTimeOut)
'cInitFlag:轉2進位 bit0:1 刪除所有合法卡 ;bit1:1 初始化所有表格;bit2:1 刪除刷卡記錄;bit3:1 其他參數初始化,但不包含系統參數;bit4~7 未使用
Declare Function hsHTA850Initial Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal cInitFlag As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer
'Get Basic HTA850 Information
'int __stdcall hsHTA850GetInfo(HANDLE hComm,unsigned char *cInfoData,int *iInfoLen,int *iReturnCode,unsigned int iTimeOut)
Declare Function hsHTA850GetInfo Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cInfodata As Byte, ByRef iinfolen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850ReadParameter(HANDLE hComm,unsigned char *cParaData,int *iParaLen,int *iReturnCode,unsigned int iTimeOut)
'Read HTA850 Params
Declare Function hsHTA850ReadParameter Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cParaData As Byte, ByRef iparalen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer
'int __stdcall hsHTA850WriteParameter(HANDLE hComm,unsigned char *cParaData,int iParaLen,int *iReturnCode,unsigned int iTimeOut)
'Write HTA850 Params
Declare Function hsHTA850WriteParameter Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cParaData As Byte, ByVal iparalen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850WriteTime(HANDLE hComm,char * cDate,char *cTime,int *iReturnCode,unsigned int iTimeOut)
' Set the host time
Declare Function hsHTA850WriteTime Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal Sdate As String, ByVal sTime As String, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer
'int __stdcall hsHTA850ReadTime(HANDLE hComm,int iGCUID,char *cDate,char *cTime,int * iReturnCode,unsigned int iTimeout)
Declare Function hsHTA850ReadTime Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef Sdate As Byte, ByRef sTime As Byte, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850WriteTable(HANDLE hComm,int iTable,unsigned char *cTableData,int iTableLen,int * iReturnCode,unsigned int iTimeout)
Declare Function hsHTA850WriteTable Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal itable As Integer, ByRef cTableData As Byte, ByVal itablelen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850ReadTable(HANDLE hComm,int iTable,unsigned char *cTableData,int *iTableLen,int * iReturnCode,unsigned int iTimeout)
Declare Function hsHTA850ReadTable Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal itable As Integer, ByRef cTableData As Byte, ByRef itablelen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850InsertMultiUserRecord(HANDLE hComm,int CardLen,int MsgLen,int iRecord,struct_CardFormat *stRecord,int *iReturnCode,unsigned int iTimeOut)
Declare Function hsHTA850InsertMultiUserRecord Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByVal MsgLen As Integer, ByVal iRecord As Integer, ByRef stRecord As Byte, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850DeleteUserRecord(HANDLE hComm,int CardLen,char *cCardNo,int *iReturnCode,unsigned int iTimeOut)

Declare Function hsHTA850DeleteUserRecord Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByRef cCardNo As Byte, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850QueryUserRecord(HANDLE hComm,int CardLen,char *cCardNo,unsigned char *cCardFormatData,int *iCardFormatLen,int *iReturnCode,unsigned int iTimeOut)
Declare Function hsHTA850QueryUserRecord Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByRef cCardNo As Byte, ByRef cCardFormatData As Byte, ByRef iCardFormatLen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850DeleteAllUserRecord(HANDLE hComm,int *iReturnCode,unsigned int iTimeOut)
Declare Function hsHTA850DeleteAllUserRecord Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850ReadEEPROM(HANDLE hComm,unsigned char *cEESendData,int iEESendLen,unsigned char * cEEReceiveData,int *iEEReceiveLen,int *iReturnCode,unsigned int iTimeOut)
Declare Function hsHTA850ReadEEPROM Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cEESendData As Byte, ByVal iEESendLen As Integer, ByRef cEEReceiveData As Byte, ByRef iEEReceiveLen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850SetEEPROM(HANDLE hComm,unsigned char * cEEData,int iEELen,int *iReturnCode,unsigned int iTimeOut)
Declare Function hsHTA850SetEEPROM Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cEEData As Byte, ByVal iEELen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850SetMifareReader(HANDLE hComm,unsigned char *cData,int iLen,int *iReturnCode,unsigned int iTimeOut)
Declare Function hsHTA850SetMifareReader Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cdata As Byte, ByVal iLen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850PollingData(HANDLE hComm,int iPrevRecord,stPollRecord *stRecord,int *iRecord,int * iReturnCode,unsigned int iTimeout)
Declare Function hsHTA850PollingData Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal iPrevRecord As Integer, ByRef stRecord As Byte, ByRef iRecord As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850InsertMultiUserFingerPrinter(HANDLE hComm,int CardLen,int MsgLen,int iRecord,struct_FingerPrinterFormat *stRecord,int *iReturnCode,unsigned int iTimeOut);
Declare Function hsHTA850InsertMultiUserFingerPrinter Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByVal MsgLen As Integer, ByVal iRecord As Integer, ByRef stRecord As Byte, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer
Declare Function hsHTA850InsertMultiUserFingerPrinter2 Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByVal MsgLen As Integer, ByVal iRecord As Integer, ByRef stRecord As Byte, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'int __stdcall hsHTA850QueryUserFingerPrinter(HANDLE hComm,int CardLen,char *cCardNo,unsigned char *cFingerPrinterData1,unsigned char *cFingerPrinterData2,int *iCardFormatLen,int *iReturnCode,unsigned int iTimeOut);
Declare Function hsHTA850QueryUserFingerPrinter Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByRef cCardNo As Byte, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByRef iCardFormatLen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer
Declare Function hsHTA850QueryUserFingerPrinter2 Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal CardLen As Integer, ByRef cCardNo As Byte, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByRef iCardFormatLen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

Declare Function hsHTA850UpdateMasterFP Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer
Declare Function hsHTA850QueryMasterFP Lib "HDTAS.dll" (ByVal hComm As Integer, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

'iFunction:9 Write,10 Read
Declare Function hsHTA850Set Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal iFunction As Integer, ByRef cSendData As Byte, ByVal iSendLen As Integer, ByRef cReceiveData As Byte, ByRef iReceiveLen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer

Public ghComm As Integer

'Added by Morgan 2020/8/25 RAC960PMF門禁機 API
Declare Function hacOpenChannelEX Lib "HDACS.dll" (ByVal sComm As String, ByVal iPort As Integer, ByVal iCheckStatus As Integer, ByRef hComm As Integer, ByVal iTimeout As Integer) As Integer
Declare Function hacCloseChannel Lib "HDACS.dll" (ByVal hComm As Integer) As Integer
Declare Function hacGetDateTime Lib "HDACS.dll" (ByVal iNodeID As Integer, ByRef Sdate As Byte, ByRef sTime As Byte, ByVal hComm As Integer, ByVal iTimeout As Integer) As Integer
Declare Function hacSetDateTime Lib "HDACS.dll" (ByVal iNodeID As Integer, ByVal Sdate As String, ByVal sTime As String, ByVal hComm As Integer, ByVal iTimeout As Integer) As Integer
'Adds a card & display name
'iStatus:轉2進位 Bit0:0=正常卡,1=黑名單;Bit1:0=卡片,1=指紋;Bit2:0=假日不檢查,1=假日檢查;Bit3:0=時段不檢查,1=時段檢查
Declare Function hacAddCardEX Lib "HDACS.dll" (ByVal iNodeID As Integer, ByVal cCardNo As String, ByVal iCardLen As Integer, ByVal cPassWord As String, ByVal iPassLen As Integer, ByVal cName As String, ByVal iNameLen As Integer _
, ByVal iTimeZone As Integer, ByVal cStatus As Integer, ByVal hComm As Integer, ByVal iTimeout As Integer) As Integer
Declare Function hacDelCard Lib "HDACS.dll" (ByVal iNodeID As Integer, ByVal cCardNo As String, ByVal iCardLen As Integer, ByVal hComm As Integer, ByVal iTimeout As Integer) As Integer
'Retrieve Finger Pattern
Declare Function hacFingerPrinterQueryUser Lib "HDACS.dll" (ByVal iNodeID As Integer, ByVal hComm As Integer, ByVal iCardLen As Integer, ByRef cCardNo As Byte _
, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByRef iCardFormatLen As Integer, ByRef iReturnCode As Integer, ByVal iTimeout As Integer) As Integer
'Insert Finger Pattern
'iStatus:轉2進位 Bit0:0=正常卡,1=黑名單;Bit1:0=卡片,1=指紋;Bit2:0=假日不檢查,1=假日檢查;Bit3:0=時段不檢查,1=時段檢查
Declare Function hacAddCardFingerPrintEx Lib "HDACS.dll" (ByVal iNodeID As Integer, ByVal cCardNo As String, ByVal iCardLen As Integer, ByVal cPassWord As String, ByVal iPassLen As Integer, ByVal cName As String, ByVal iNameLen As Integer _
, ByVal iTimeZone As Integer, ByVal cStatus As Integer, ByRef cFingerPrinterData1 As Byte, ByRef cFingerPrinterData2 As Byte, ByVal hComm As Integer, ByVal iTimeout As Integer) As Integer
Declare Function hacPolling Lib "HDACS.dll" (ByVal iNodeID As Integer, ByVal iPrevRecord As Integer, ByRef stRecord As Byte, ByRef iRecord As Integer, ByVal hComm As Integer, ByVal iTimeout As Integer, ByRef iCardType As Integer) As Integer
Declare Function hacGetFlashData Lib "HDACS.dll" (ByVal iNodeID As Integer, ByRef cFlashData As Byte, ByRef iReceiveDataLen As Integer, ByVal iFlashAddr As Integer, ByVal iFlashLen As Integer, ByVal hComm As Integer, ByVal iTimeout As Integer) As Integer
Declare Function hacSetFlashData Lib "HDACS.dll" (ByVal iNodeID As Integer, ByRef cFlashData As Byte, ByVal iFlashAddr As Integer, ByVal iFlashLen As Integer, ByVal hComm As Integer, ByVal iTimeout As Integer) As Integer
'end 2020/8/25

Declare Function htaGetCardData Lib "HDTAS.dll" (ByVal hComm As Integer, ByVal iNodeID As Integer, ByRef cCardData As Byte, ByRef iReceiveDataLen As Integer, ByVal iBank As Integer, ByVal iCompress As Integer, ByVal iTimeout As Integer) As Integer

'Added by Morgan 2023/7/26
Private Const cHTATimOut1 As Integer = 3000
Private Const cHTATimOut2 As Integer = 6000
Private Const cHTATimOut3 As Integer = 10000
Private Const cHTATimOut4 As Integer = 30000

'建立考勤機連線
'Modified by Morgan 2020/8/25 +門禁機
Public Function HTAconnect(Optional NoErrMsg As Boolean, Optional iReturn As Integer, Optional pRetry As Integer = 3) As Boolean
   Dim iErrCount As Integer
      
On Error GoTo ErrHnd
      
   pubIsNewDevice = CheckLevel(HTAip & ";", "門禁機IP")
   
   iErrCount = 0
   Do While iErrCount < pRetry
      If pubIsNewDevice Then
         'Modified by Morgan 2023/6/12 +timeout延長為3000
         'Modified by Morgan 2023/7/27 Timeout改用常數
         iReturn = hacOpenChannelEX(HTAip, HTAport, 1, ghComm, cHTATimOut2)
      Else
         iReturn = hsOpenChannel(ghComm, HTAip, HTAport)
      End If
      If iReturn = 0 And ghComm > 0 Then
         HTAconnect = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
      
   If iReturn <> 0 Then
      Pub_WriteSysLog "建立考勤機連線失敗！ (" & HTAip & ")"
      PUB_SendMail strUserNum, GetDeptMan("M51") & ";92012", "", "建立考勤機連線失敗！(" & HTAip & ",Return No:" & iReturn & ")", "如旨"
      If NoErrMsg = False Then
         MsgBox "建立考勤機連線失敗！" & " (" & HTAip & ")", vbCritical
      End If
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
End Function
'關閉考勤機連線
Public Function HTAclose(Optional NoErrMsg As Boolean, Optional iReturn As Integer, Optional pRetry As Integer = 3) As Boolean
   Dim iErrCount As Integer
   
On Error GoTo ErrHnd

   If ghComm = 0 Then
      HTAclose = True
   Else
      iErrCount = 0
      Do While iErrCount < pRetry
         'Modified by Morgan 2020/8/25 +門禁機
         If pubIsNewDevice Then
            iReturn = hacCloseChannel(ghComm)
         Else
            iReturn = hsCloseChannel(ghComm)
         End If
         'end 2020/8/25
         If iReturn = 0 Then
            HTAclose = True
            ghComm = 0
            Exit Do
         End If
         iErrCount = iErrCount + 1
      Loop
      
      If iReturn <> 0 Then
         Pub_WriteSysLog "關閉考勤機連線失敗！( Return Code: " & iReturn & ")"
         
         If NoErrMsg = False Then
            MsgBox "關閉考勤機連線失敗！( Return Code: " & iReturn & ")", vbCritical
         End If
      End If
   End If
   Exit Function
   
ErrHnd:
   Pub_WriteSysLog Err.Description
   
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

'刪除考勤機所有卡號資料
Public Function HTAdeleteAllCard(Optional NoErrMsg As Boolean, Optional iReturn As Integer, Optional iReturnCode As Integer, Optional pRetry As Integer = 3) As Boolean
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   
   If pubIsNewDevice Then HTAdeleteAllCard = True: Exit Function 'Added by Morgan 2020/8/25 門禁機無此功能
   
On Error GoTo ErrHnd

   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   
   iErrCount = 0
   Do While iErrCount < pRetry
      'Modified by Morgan 2023/7/27 Timeout改用常數
      iReturn = hsHTA850DeleteAllUserRecord(ghComm, iReturnCode, cHTATimOut1)
      If iReturn = 0 Then
         HTAdeleteAllCard = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And NoErrMsg = False Then
      MsgBox "刪除所有卡號失敗！( Error Code: " & iReturn & "," & iReturnCode & ")", vbCritical
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

'刪除考勤機單筆卡號資料
Public Function HTAdeleteCard(CardID As String, Optional NoErrMsg As Boolean, Optional iReturn As Integer, Optional iReturnCode As Integer, Optional pRetry As Integer = 3) As Boolean
   Dim bolNewComm As Boolean
   Dim i As Integer
   Dim stRecord(15) As Byte
   Dim iErrCount As Integer
   
On Error GoTo ErrHnd

   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   
   For i = 0 To Len(CardID) - 1 '.Length - 1
       stRecord(i) = Asc(Mid(CardID, i + 1, 1))
   Next
   For i = Len(CardID) To 15
      stRecord(i) = 0
   Next

   iErrCount = 0
   Do While iErrCount < pRetry
      'Modified by Morgan 2020/8/25 +門禁機
      'Modified by Morgan 2023/7/27 +Timeout改用常數
      If pubIsNewDevice Then
         iReturn = hacDelCard(cNodeID, CardID, Len(CardID), ghComm, cHTATimOut2)
      Else
         iReturn = hsHTA850DeleteUserRecord(ghComm, 16, stRecord(0), iReturnCode, cHTATimOut1)
      End If
      'end 2020/8/25
      
      If iReturn = 0 Then
         HTAdeleteCard = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And NoErrMsg = False Then
      MsgBox "刪除卡號 " & CardID & " 失敗！( Error Code: " & iReturn & "," & iReturnCode & ")", vbCritical
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

'新增考勤機卡號資料(卡)
Public Function HTAaddCard(CardNo As String, DisplayName As String, Optional NoErrMsg As Boolean, Optional iReturn As Integer, Optional iReturnCode As Integer, Optional pRetry As Integer = 3, Optional pTimeZone As Integer = 1, Optional pNoCheck As Boolean = False) As Boolean
   Dim i As Integer, k As Integer, iASC As Integer
   Dim stRecord(34) As Byte
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   Dim iStatus As Integer 'Added by Morgan 2020/8/28
   
On Error GoTo ErrHnd

   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   
   For i = 0 To UBound(stRecord)
       stRecord(i) = 0
   Next
        
   k = 0
   For i = 1 To Len(CardNo)
       stRecord(k) = Asc(Mid(CardNo, i, 1))
       k = k + 1
   Next
   
   k = 18
   '中文ASCII碼為負數要拆成兩組16進位碼
   For i = 1 To Len(DisplayName)
      iASC = Asc(Mid(DisplayName, i, 1))
      If iASC > 0 Then
          stRecord(k) = iASC
      Else
         stRecord(k) = Val("&H" & Left(Hex(iASC), 2))
         k = k + 1
          stRecord(k) = Val("&H" & Right(Hex(iASC), 2))
      End If
      k = k + 1
   Next
   
   iErrCount = 0
   Do While iErrCount < pRetry
      'Modified by Morgan 2020/8/25 +門禁機
      'Modified by Morgan 2023/7/27 Timeout改用常數
      If pubIsNewDevice Then
         If pNoCheck = True Then
            'Added by Morgan 2023/10/16 改只有假日不檢查
            'iStatus = 0
            iStatus = 8
            'end 2023/10/16
         Else
            iStatus = 12
         End If
         'iStatus:轉2進位 Bit0:0=正常卡,1=黑名單;Bit1:0=卡片,1=指紋;Bit2:0=假日不檢查,1=假日檢查;Bit3:0=時段不檢查,1=時段檢查
         iReturn = hacAddCardEX(cNodeID, CardNo, Len(CardNo), "", 0, DisplayName, LenB(DisplayName), pTimeZone, iStatus, ghComm, cHTATimOut2)
      Else
         iReturn = hsHTA850InsertMultiUserRecord(ghComm, 16, 16, 1, stRecord(0), iReturnCode, cHTATimOut2)
      End If
      'end 2020/8/25
      
      If iReturn = 0 Then
         HTAaddCard = True
         Exit Do
      End If
      Sleep 3000
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And NoErrMsg = False Then
      MsgBox "新增卡號 " & CardNo & " 失敗！(" & iReturn & "," & iReturnCode & ")", vbCritical
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
End Function
'新增考勤機卡號資料(指紋)
Public Function HTAaddFingerPrinter(CardNo As String, DisplayName As String, sFingerPrinter1 As String, sFingerPrinter2 As String, Optional NoErrMsg As Boolean, Optional iReturn As Integer, Optional iReturnCode As Integer, Optional pRetry As Integer = 3, Optional pTimeZone As Integer = 1, Optional pNoCheck As Boolean = False) As Boolean
   Dim sFPtData1(385) As Byte
   Dim sFPtData2(385) As Byte
   Dim newFP1, newFP2
   Dim stRecord(805) As Byte
   Dim i As Integer, k As Integer
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   Dim iStatus As Integer 'Added by Morgan 2020/8/28
    
On Error GoTo ErrHnd

   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
    
   For i = 0 To UBound(sFPtData1)
      sFPtData1(i) = 0
   Next
   For i = 0 To UBound(sFPtData2)
      sFPtData2(i) = 0
   Next
      
   
   newFP1 = Split(sFingerPrinter1, " ")
   newFP2 = Split(sFingerPrinter2, " ")
   For i = 0 To UBound(newFP1)
      If i <= UBound(sFPtData1) Then
         sFPtData1(i) = Val("&H" & newFP1(i))
      End If
   Next
   For i = 0 To UBound(newFP2)
      If i <= UBound(sFPtData2) Then
         sFPtData2(i) = Val("&H" & newFP2(i))
      End If
   Next
   
   If sFPtData1(0) <> "0" Then
      For i = 0 To UBound(stRecord)
          stRecord(i) = 0
      Next
      
      For i = 0 To Len(CardNo) - 1
          stRecord(i) = Asc(Mid(CardNo, i + 1, 1))
      Next
      
      stRecord(16) = 0
      stRecord(17) = 1
        
      k = 0
      For i = 18 To 18 + Len(DisplayName) - 1
         If Asc(Mid(DisplayName, (i - 17), 1)) > 0 Then
            stRecord(i + k) = Asc(Mid(DisplayName, (i - 17), 1))
         Else
            stRecord(i + k) = Val("&H" & Left(Hex(Asc(Mid(DisplayName, (i - 17), 1))), 2))
            k = k + 1
            stRecord(i + k) = Val("&H" & Right(Hex(Asc(Mid(DisplayName, (i - 17), 1))), 2))
         End If
      Next
      
      For i = 0 To 385
          stRecord(34 + i) = sFPtData1(i)
      Next
      
      For i = 0 To 385
          stRecord(i + 420) = sFPtData2(i)
      Next
      
      iErrCount = 0
      Do While iErrCount < pRetry
         'Modified by Morgan 2020/8/25 +門禁機
         'Modified by Morgan 2023/7/27 Timeout改用常數
         If pubIsNewDevice Then
            If pNoCheck = True Then
               'Added by Morgan 2023/10/16 改只有假日不檢查
               'iStatus = 2
               iStatus = 10
               'end 2023/10/16
            Else
               iStatus = 14
            End If
            'iStatus:轉2進位 Bit0:0=正常卡,1=黑名單;Bit1:0=卡片,1=指紋;Bit2:0=假日不檢查,1=假日檢查;Bit3:0=時段不檢查,1=時段檢查
            iReturn = hacAddCardFingerPrintEx(cNodeID, CardNo, Len(CardNo), "", 0, DisplayName, 16, pTimeZone, iStatus, sFPtData1(0), sFPtData2(0), ghComm, cHTATimOut4)
         Else
            iReturn = hsHTA850InsertMultiUserFingerPrinter2(ghComm, 16, 16, 1, stRecord(0), iReturnCode, cHTATimOut4)
         End If
         
         If iReturn = 0 Then
            HTAaddFingerPrinter = True
            Exit Do
         End If
         iErrCount = iErrCount + 1
      Loop
      
      If iReturn <> 0 And NoErrMsg = False Then
         MsgBox "新增卡號 " & CardNo & " 失敗！(" & iReturn & "," & iReturnCode & ")", vbCritical
      End If
   End If

   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function
'查詢考勤機卡號資料
Public Function HTAqueryCard(CardNo As String, Optional DisplayName As String, Optional NoErrMsg As Boolean, Optional iReturn As Integer, Optional iReturnCode As Integer, Optional pRetry As Integer = 3) As Boolean
   Dim i As Integer, k As Integer
   Dim stRecord(15) As Byte, sFPtData1(385) As Byte, sFPtData2(385) As Byte
   Dim sCardFormatData(255) As Byte
   Dim iCardFormatLen As Integer
   Dim gStr As String
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   
On Error GoTo ErrHnd

   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
        
   k = 0
   For i = 1 To Len(CardNo)
      stRecord(k) = Asc(Mid(CardNo, i, 1))
      k = k + 1
   Next
   For i = k To UBound(stRecord)
      stRecord(i) = 0
   Next
   
   iErrCount = 0
   Do While iErrCount < pRetry
      iReturnCode = 0
      iReturn = 0
      'Modified by Morgan 2020/8/25 +門禁機
      'Modified by Morgan 2023/7/27 Timeout改用常數
      If pubIsNewDevice Then
         iReturn = hacFingerPrinterQueryUser(cNodeID, ghComm, Len(CardNo), stRecord(0), sFPtData1(0), sFPtData2(0), iCardFormatLen, iReturnCode, cHTATimOut2)
         If iReturn = 1001 And iReturnCode = 0 Then
            iReturnCode = 6
         End If
      Else
         iReturn = hsHTA850QueryUserRecord(ghComm, 16, stRecord(0), sCardFormatData(0), iCardFormatLen, iReturnCode, cHTATimOut2)
      End If
      'end 2020/8/25
      
      If iReturn = 0 Then
         gStr = ""
         If Not pubIsNewDevice Then
            For i = 0 To 15
              If sCardFormatData(i) > 0 Then
                 gStr = gStr & Chr(sCardFormatData(i))
              End If
            Next
            'CardNo = RTrim(gStr)
            
            gStr = ""
            For i = 18 To iCardFormatLen - 1
              If sCardFormatData(i) > 0 Then
                 If Val(sCardFormatData(i)) < 128 Then
                    gStr = gStr & Chr(sCardFormatData(i))
                 Else
                    gStr = gStr & Chr(Val("&H" & Hex(sCardFormatData(i)) & Hex(sCardFormatData(i + 1))))
                    i = i + 1
                 End If
              End If
            Next
            DisplayName = RTrim(gStr)
         End If
         HTAqueryCard = True
         Exit Do
         
      ElseIf iReturnCode = 6 Then
         If NoErrMsg = False Then
            MsgBox "卡號不存在！", vbExclamation
         End If
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And iReturnCode <> 6 And NoErrMsg = False Then
      MsgBox "卡號資料讀取失敗!(" & iReturn & "," & iReturnCode & ")", vbCritical
   End If
             
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function
'查詢考勤機指紋資料
Public Function HTAqueryFingerPrinter(CardNo As String, Optional sFingerPrinter1 As String, Optional sFingerPrinter2 As String, Optional NoErrMsg As Boolean, Optional iReturn As Integer, Optional iReturnCode As Integer, Optional pRetry As Integer = 3) As Boolean
   Dim i As Integer, k As Integer
   Dim stRecord(15) As Byte
   Dim sFPtData1(386) As Byte
   Dim sFPtData2(386) As Byte
   Dim iCardFormatLen As Integer
   Dim gStr As String
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   
On Error GoTo ErrHnd

   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   
   
   
   k = 0
   For i = 1 To Len(CardNo)
      stRecord(k) = Asc(Mid(CardNo, i, 1))
      k = k + 1
   Next
   For i = k To UBound(stRecord)
      stRecord(i) = 0
   Next
   
   For i = 0 To 385
       sFPtData1(i) = 0
   Next
   For i = 0 To 385
       sFPtData2(i) = 0
   Next
   
   iErrCount = 0
   Do While iErrCount < pRetry
      'Modified by Morgan 2020/8/25 +門禁機
      'Modified by Morgan 2023/7/27 Timeout改用常數
      If pubIsNewDevice Then
         iReturn = hacFingerPrinterQueryUser(cNodeID, ghComm, Len(CardNo), stRecord(0), sFPtData1(0), sFPtData2(0), iCardFormatLen, iReturnCode, cHTATimOut2)
      Else
         DoEvents 'Added by Morgan 2023/12/5 win10要加才不會失敗
         iReturn = hsHTA850QueryUserFingerPrinter2(ghComm, 16, stRecord(0), sFPtData1(0), sFPtData2(0), iCardFormatLen, iReturnCode, cHTATimOut2)
      End If
      'end 2020/8/25
      If iReturn = 0 Then
         gStr = ""
         For i = 0 To 385
             gStr = gStr & Right("0" & Hex(sFPtData1(i)), 2) & " "
         Next
         sFingerPrinter1 = gStr
         
         gStr = ""
         For i = 0 To 385
             gStr = gStr & Right("0" & Hex(sFPtData2(i)), 2) & " "
         Next
         sFingerPrinter2 = gStr
         HTAqueryFingerPrinter = True
         Exit Do
      ElseIf iReturnCode = 6 Then
         If NoErrMsg = False Then
            MsgBox "指紋資料不存在！", vbExclamation
         End If
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And iReturnCode <> 6 And NoErrMsg = False Then
      MsgBox "指紋資料讀取失敗!(" & iReturn & "," & iReturnCode & ")", vbCritical
   End If
          
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'下載刷卡紀錄
Public Function HTAPolling(Optional iRecordCount As Integer, Optional NoErrMsg As Boolean, Optional iReturn As Integer, Optional iReturnCode As Integer, Optional iDBreturn As Integer) As Boolean
   Dim rByte(1024) As Byte
   Dim iRecCount As Integer '前次紀錄位置
   Dim iRecord As Integer
   Dim PR(9) As String
   Dim stTemp As String
   Dim i As Integer, j As Integer, k As Integer
   Dim bolNewComm As Boolean
        
On Error GoTo ErrHnd

   'Added by Morgan 2017/6/28 改若有連線時一律先斷線再重連
   If ghComm <> 0 Then
      Pub_WriteSysLog "關閉舊連線..."
      If HTAclose(NoErrMsg) = False Then Exit Function
   End If
   'end 2017/6/28
   
   Pub_WriteSysLog "建立連線..."
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   iRecordCount = 0
   
   Do
      
      Pub_WriteSysLog "下載刷卡紀錄..."
      
      Erase rByte
      
      'Modified by Morgan 2020/8/25 +門禁機
      'Modified by Morgan 2023/7/27 Timeout改用常數
      If pubIsNewDevice Then
         iReturn = hacPolling(cNodeID, iRecCount, rByte(0), iRecord, ghComm, cHTATimOut3, 3)
         If iReturn = "1010" Then
            iRecord = 0
            iReturn = 0
         End If
      Else
         iReturn = hsHTA850PollingData(ghComm, iRecCount, rByte(0), iRecord, iReturnCode, cHTATimOut3)
         'iRecordCount = iRecordCount + iRecord 'Removed by Morgan 2020/11/5 移到迴圈內
      End If
      
      If iReturn = 0 Then
         k = 0
         For i = 0 To iRecord - 1
            iRecordCount = iRecordCount + 1 'Added by Morgan 2020/11/5
            If pubIsNewDevice Then
               'cEventCode[5] 事件代碼
               '0-1 班別
               PR(7) = ""
               For j = 1 To 2
                   If (rByte(k) <> 0) Then
                       PR(7) = PR(7) & Chr(rByte(k))
                   End If
                   k = k + 1
               Next j
               
               '2-3 狀態
               PR(8) = ""
               For j = 1 To 2
                   If (rByte(k) <> 0) Then
                       PR(8) = PR(8) & Chr(rByte(k))
                   End If
                   k = k + 1
               Next j
               
               '4 無
               
               k = k + 1
               'cDateTime[20] 日期時間 YYYY/MM/DD HH:MI:SS
               '5-14 Date
               PR(1) = ""
               For j = 1 To 10
                   If (rByte(k) <> 0) Then
                       PR(1) = PR(1) & Chr(rByte(k))
                   End If
                   k = k + 1
               Next j
               PR(1) = Replace(PR(1), "/", "")
               
               '15-24 Time
               PR(2) = ""
               For j = 1 To 10
                   If (rByte(k) <> 0) Then
                       PR(2) = PR(2) & Chr(rByte(k))
                   End If
                   k = k + 1
               Next j
               PR(2) = Replace(Trim(PR(2)), ":", "")
               
               'cCard[20] 卡號
               PR(3) = ""
               For j = 1 To 20
                   If (rByte(k) <> 0) Then
                       PR(3) = PR(3) & Chr(rByte(k))
                   End If
                   k = k + 1
               Next j
               
               'cDeviceID[10] 設備ID
               PR(4) = ""
               For j = 1 To 10
                   If (rByte(k) <> 0) Then
                       PR(4) = PR(4) & Chr(rByte(k))
                   End If
                   k = k + 1
               Next j
               
               'cReaderID[10] 讀頭ID
               PR(5) = ""
               For j = 1 To 10
                   If (rByte(k) <> 0) Then
                       PR(5) = PR(5) & Chr(rByte(k))
                   End If
                   k = k + 1
               Next j
            
               PR(6) = "0"
            Else
               '0-9 Date
               PR(1) = ""
               For j = 1 To 10
                   If (rByte(k) <> 0) Then
                       PR(1) = PR(1) & Chr(rByte(k))
                   End If
                   k = k + 1
               Next j
               '10-19 Time
               PR(2) = ""
               For j = 1 To 10
                   If (rByte(k) <> 0) Then
                       PR(2) = PR(2) & Chr(rByte(k))
                   End If
                   k = k + 1
               Next j
               '20 Reader NO.
               PR(4) = rByte(k)
               '21 Input Type
               k = k + 1
               PR(5) = rByte(k)
               '22 Section
               k = k + 1
               PR(6) = rByte(k)
               '23 Class
               k = k + 1
               PR(7) = rByte(k)
               '24 Event Code
               k = k + 1
               PR(8) = rByte(k)
               '25-40
               PR(3) = ""
               k = k + 1
               For j = 1 To 16
                  If (rByte(k) <> 0) Then
                      PR(3) = PR(3) & Chr(rByte(k))
                  End If
                  k = k + 1
               Next j
            End If
            
            PR(9) = HTAip
            'If PR(8) = "0" Or PR(8) = "9" Then
               'Debug.Print PR(1) & ":" & PR(2) & ":" & PR(3)
               If PR(3) <> "" Then 'Added by Morgan 2020/11/5 發現有無ID的異常資料
                  If SaveRecord(PR, NoErrMsg) = False Then
                     iDBreturn = 1
                     Exit Function
                  End If
               End If
            'End If
         Next i
         iRecCount = iRecord
      Else
         Pub_WriteSysLog "刷卡紀錄讀取失敗！"
         'Added by Morgan 2015/8/21
         'Modify By Sindy 2023/7/25 經理改抓 GetDeptMan("M51")
         PUB_SendMail strUserNum, GetDeptMan("M51") & ";92012", "", "刷卡紀錄讀取失敗！(" & HTAip & ",Return Code:" & iReturnCode & ",Return No:" & iReturn & ")", "如旨"
         If NoErrMsg = False Then
            MsgBox "刷卡紀錄讀取失敗！"
         End If
         Exit Do
      End If
      
   Loop While (iRecCount > 0)
   
   If iReturn = 0 Then
      HTAPolling = True
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   Pub_WriteSysLog Err.Description
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If

   'Added by Morgan 2017/5/8 發生錯誤時連線也要中斷後面的連線才會正確
   If bolNewComm = True Then
      HTAclose True
   End If
   'end 2017/5/8
End Function

'Added by Morgan 2016/2/23
Public Function HTAReadTime(ByRef pDate As String, ByRef pTime As String, Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturnCode, iReturn
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   Dim btDate(8) As Byte
   Dim btTime(5) As Byte
   Dim ii As Integer
   
On Error GoTo ErrHnd

   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function

   iErrCount = 0
   Do While iErrCount < pRetry
      iReturn = 0
      iReturnCode = 0
      'Modified by Morgan 2020/8/25 +門禁機
      'Modified by Morgan 2023/7/27 Timeout改用常數
      If pubIsNewDevice Then
         iReturn = hacGetDateTime(cNodeID, btDate(0), btTime(0), ghComm, cHTATimOut2)
      Else
         iReturn = hsHTA850ReadTime(ghComm, btDate(0), btTime(0), iReturnCode, cHTATimOut1)
      End If
      'end 2020/8/25
      If iReturn = 0 Then
         HTAReadTime = True
         pDate = ""
         For ii = 0 To UBound(btDate)
            pDate = pDate & Chr(btDate(ii))
         Next
         pTime = ""
         For ii = 0 To UBound(btTime)
            pTime = pTime & Chr(btTime(ii))
         Next
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And NoErrMsg = False Then
      MsgBox "指紋機時間讀取失敗！", vbExclamation
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
      
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function


Public Function HTAWriteTime(Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturnCode, iReturn, iELID As Integer
   Dim iweek As Integer
   Dim Sdate As String
   Dim sTime As String, sTime1 As String
   Dim dtNow As Date
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   
On Error GoTo ErrHnd

   'Added by Morgan 2021/9/2 先同步本機與DB時間
   Date = Format(ServerDate, "####/##/##") 'Added by Morgan 2025/3/11 日期也要同步
   sTime1 = ServerTime
   sTime = sTime1
   Do While (sTime1 = sTime)
      Sleep 50
      sTime = ServerTime
   Loop
   Time = Format(ServerTime, "00:00:00")
   'end 2021/9/2
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function

   iErrCount = 0
   Do While iErrCount < pRetry
      dtNow = Now
      sTime1 = Format(dtNow, "HHmmss")
      sTime = sTime1
      Do While (sTime1 = sTime)
         Sleep 50
         dtNow = Now
         sTime = Format(dtNow, "HHmmss")
      Loop
   
      iweek = Weekday(dtNow) - 1
      If iweek = 0 Then iweek = 7
      Sdate = Format(dtNow, "yyyymmdd")
      Sdate = Sdate & Trim(str(iweek))
      
      iReturn = 0
      iReturnCode = 0
      
      'Modified by Morgan 2020/8/25 +門禁機
      'Modified by Morgan 2023/7/27 Timeout改用常數
      If pubIsNewDevice Then
         iReturn = hacSetDateTime(cNodeID, Sdate, sTime, ghComm, cHTATimOut2)
      Else
         iReturn = hsHTA850WriteTime(ghComm, Sdate, sTime, iReturnCode, cHTATimOut1)
      End If
      
      If iReturn = 0 Then
         HTAWriteTime = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And NoErrMsg = False Then
      MsgBox "更新指紋機時間失敗！", vbExclamation
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
      
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Function SaveRecord(pPR() As String, Optional NoErrMsg As Boolean) As Boolean
   Dim stSQL As String, iRtn As Integer
   
   cnnConnection.BeginTrans
On Error GoTo ErrHnd

   stSQL = "delete PollRecord where PR01=" & pPR(1) & " and pr02=" & pPR(2) & " and pr03='" & pPR(3) & "'"
   cnnConnection.Execute stSQL, iRtn
   stSQL = "insert into PollRecord(pr01,pr02,pr03,pr04,pr05,pr06,pr07,pr08,pr09) values (" & pPR(1) & "," & pPR(2) & ",'" & pPR(3) & "'," & pPR(4) & "," & pPR(5) & "," & pPR(6) & "," & pPR(7) & "," & pPR(8) & ",'" & pPR(9) & "')"
   cnnConnection.Execute stSQL, iRtn
   cnnConnection.CommitTrans
   SaveRecord = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
End Function
'Modified by Morgan 2020/8/25 +門禁機
Public Function GetHtaIP(Optional pType As Integer = 0) As String
   Dim stIPList As String, stIPList2 As String
   If pType = 0 Or pType = 1 Then
      stIPList = Pub_GetSpecMan("考勤機IP")
   End If
   If pType = 0 Or pType = 2 Then
      stIPList2 = Pub_GetSpecMan("門禁機IP")
   End If
   
   GetHtaIP = stIPList & IIf(stIPList <> "", ";", "") & stIPList2
End Function
'Added by Morgan 2020/9/3
'回寫門禁機時間表
Public Function HTAWriteTimeSheet(pSNo As String, pFromTime As String, pToTime As String, Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturn
   Dim iErrCount  As Integer
   Dim bolNewComm As Boolean
   Dim rByte(256) As Byte
   Dim sAddr As String
   Dim iSNo As Integer
      
On Error GoTo ErrHnd
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   
   iSNo = Val(pSNo)
   sAddr = &H480 + 4 * iSNo
   rByte(0) = Val("&H" & (pFromTime \ 100))
   rByte(1) = Val("&H" & Right(pFromTime, 2))
   rByte(2) = Val("&H" & (pToTime \ 100))
   rByte(3) = Val("&H" & Right(pToTime, 2))
   
   iErrCount = 0
   Do While iErrCount < pRetry
      'Modified by Morgan 2023/7/27 Timeout改用常數
      iReturn = hacSetFlashData(cNodeID, rByte(0), sAddr, 4, ghComm, cHTATimOut2)
      If iReturn = 0 Then
         HTAWriteTimeSheet = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
End Function
'Added by Morgan 2020/9/3
'清除門禁機時間表 128組*4Byte
Public Function HTAClearTimeSheet(Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturn
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   Dim rByte(256) As Byte
   Dim sAddr As String
   Dim iDataLen As Byte
   Dim iSNo As Integer
   Dim ii As Integer
   
On Error GoTo ErrHnd
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   
   sAddr = &H480 '128*4
   '先鎖定50筆以內
   For ii = 0 To 9
      rByte(0 + ii * 4) = &HFF
      rByte(1 + ii * 4) = &HFF
      rByte(2 + ii * 4) = &HFF
      rByte(3 + ii * 4) = &HFF
   Next
   iDataLen = 4 * 50
   iErrCount = 0
   Do While iErrCount < pRetry
      'Modified by Morgan 2023/7/27 Timeout改用常數
      iReturn = hacSetFlashData(cNodeID, rByte(0), sAddr, iDataLen, ghComm, cHTATimOut2)
      If iReturn = 0 Then
         HTAClearTimeSheet = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
End Function
'Added by Morgan 2020/9/3
'回寫門禁機時段表
Public Function HTAWriteTimeZone(pSNo As String, pDay1 As String, pDay2 As String, pDay3 As String, pDay4 As String, pDay5 As String, pDay6 As String, pDay7 As String, Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturn
   Dim iErrCount  As Integer
   Dim bolNewComm As Boolean
   Dim rByte(256) As Byte
   Dim sAddr As String
   Dim iSNo As Integer
      
On Error GoTo ErrHnd
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   
   iSNo = Val(pSNo)
   sAddr = &H100 + 7 * iSNo
   rByte(0) = Val("&H" & pDay1)
   rByte(1) = Val("&H" & pDay2)
   rByte(2) = Val("&H" & pDay3)
   rByte(3) = Val("&H" & pDay4)
   rByte(4) = Val("&H" & pDay5)
   rByte(5) = Val("&H" & pDay6)
   rByte(6) = Val("&H" & pDay7)
   
   iErrCount = 0
   Do While iErrCount < pRetry
      'Modified by Morgan 2023/7/27 Timeout改用常數
      iReturn = hacSetFlashData(cNodeID, rByte(0), sAddr, 7, ghComm, cHTATimOut2)
      If iReturn = 0 Then
         HTAWriteTimeZone = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
End Function
'Added by Morgan 2020/9/3
'清除門禁機時段表 128組*4Byte
Public Function HTAClearTimeZone(Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturn
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   Dim rByte(256) As Byte
   Dim sAddr As String
   Dim iDataLen As Byte
   Dim iSNo As Integer
   Dim ii As Integer
   
On Error GoTo ErrHnd
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   
   sAddr = &H100 '128*7
   '先鎖定10筆以內
   For ii = 0 To 9
      rByte(0 + ii * 7) = &HFF
      rByte(1 + ii * 7) = &HFF
      rByte(2 + ii * 7) = &HFF
      rByte(3 + ii * 7) = &HFF
      rByte(4 + ii * 7) = &HFF
      rByte(5 + ii * 7) = &HFF
      rByte(6 + ii * 7) = &HFF
   Next
   iDataLen = 7 * 10
   iErrCount = 0
   Do While iErrCount < pRetry
      'Modified by Morgan 2023/7/27 Timeout改用常數
      iReturn = hacSetFlashData(cNodeID, rByte(0), sAddr, iDataLen, ghComm, cHTATimOut2)
      If iReturn = 0 Then
         HTAClearTimeZone = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'Added by Morgan 2020/11/16
'讀取時間表
Public Function HTAReadTimeSheet(ByRef rDate() As String, Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturnCode, iReturn
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   Dim rByte(256) As Byte
   Dim iGetLen As Integer
   Dim sAddr As String
   Dim iDataLen As Integer
   Dim ii As Integer, jj As Integer
   
On Error GoTo ErrHnd
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function

   iErrCount = 0
   Do While iErrCount < pRetry
      iGetLen = 40
      sAddr = Val("&H" & "480")
      'Modified by Morgan 2023/7/27 Timeout改用常數
      iReturn = hacGetFlashData(cNodeID, rByte(0), iDataLen, sAddr, iGetLen, ghComm, cHTATimOut2)
      If iReturn = 0 And iDataLen > 0 Then
         For ii = 1 To iDataLen
            rDate(ii) = Right("0" & Hex(rByte(ii - 1)), 2)
         Next
         HTAReadTimeSheet = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And NoErrMsg = False Then
      MsgBox "時間表讀取失敗！", vbExclamation
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
      
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

'Added by Morgan 2020/11/16
'讀段時間表
Public Function HTAReadTimeZone(ByRef rDate() As String, Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturnCode, iReturn
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   Dim rByte(256) As Byte
   Dim iGetLen As Integer
   Dim sAddr As String
   Dim iDataLen As Integer
   Dim ii As Integer, jj As Integer
   
On Error GoTo ErrHnd
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function

   iErrCount = 0
   Do While iErrCount < pRetry
      iGetLen = 70
      sAddr = Val("&H" & "100")
      'Modified by Morgan 2023/7/27 Timeout改用常數
      iReturn = hacGetFlashData(cNodeID, rByte(0), iDataLen, sAddr, iGetLen, ghComm, cHTATimOut2)
      If iReturn = 0 And iDataLen > 0 Then
         For ii = 1 To iDataLen
            rDate(ii) = Right("0" & Hex(rByte(ii - 1)), 2)
         Next
         HTAReadTimeZone = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And NoErrMsg = False Then
      MsgBox "時段表讀取失敗！", vbExclamation
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
      
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function
'Added by Morgan 2020/9/8
'Modified by Morgan 2024/12/2 加讀取次年
Public Function HTAReadHoliday(ByRef rDate() As String, Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturnCode, iReturn
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   Dim rByte(600) As Byte
   Dim iGetLen As Integer, iGetLen1 As Integer
   Dim sAddr As String
   Dim iDataLen As Integer
   Dim ii As Integer, jj As Integer
   
On Error GoTo ErrHnd
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function

   iErrCount = 0
   Do While iErrCount < pRetry
      iGetLen = 3 * 50 '無法一次讀100筆，改先讀50筆，再往後
      iGetLen1 = iGetLen
      sAddr = Val("&H" & "680")
      'Modified by Morgan 2023/7/27 Timeout改用常數
      iReturn = hacGetFlashData(cNodeID, rByte(0), iDataLen, sAddr, iGetLen1, ghComm, cHTATimOut2)
      If iReturn = 0 And iDataLen > 0 Then
         If iDataLen < 600 Then
            iGetLen = iDataLen
            Do While (iGetLen < 600)
               sAddr = Val("&H" & "680") + iGetLen
               iDataLen = 0
               'Modified by Morgan 2023/7/27 Timeout改用常數
               iReturn = hacGetFlashData(cNodeID, rByte(iGetLen), iDataLen, sAddr, iGetLen1, ghComm, cHTATimOut2)
               If iReturn = 0 And iDataLen > 0 Then
                  iGetLen = iGetLen + iDataLen
               Else
                  Exit Do
               End If
            Loop
         End If
         
         If iReturn = 0 Then
            For ii = 0 To iGetLen - 1
               jj = 1 + ii \ 3
               rDate(jj) = rDate(jj) & Right("0" & Hex(rByte(ii)), 2)
            Next
         End If
         
         HTAReadHoliday = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And NoErrMsg = False Then
      MsgBox "假日表讀取失敗！", vbExclamation
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
      
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

'Added by Morgan 2020/9/8
'Modified by Morgan 2024/12/2 次年改放在後面的區塊(100-199)
Public Function HTAWriteHoliday(ByRef rDate() As String, Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturn
   Dim iErrCount  As Integer
   Dim bolNewComm As Boolean
   Dim rByte(600) As Byte 'Modified by Morgan 2024/12/2 300->600
   Dim sAddr As String
   Dim iDataLen As Byte
   Dim ii As Integer, idx As Integer
      
On Error GoTo ErrHnd
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function
   
   iDataLen = 0
   
   'Modified by Morgan 2024/12/2
   'For ii = 0 To 99
   For ii = 0 To 199
   'end 2024/12/2
      If ii >= UBound(rDate) Then
         rByte(0 + ii * 3) = Val("&HFF")
         rByte(1 + ii * 3) = Val("&HFF")
         rByte(2 + ii * 3) = Val("&HFF")
         
      ElseIf rDate(ii + 1) <> "" Then
         rByte(0 + ii * 3) = Val("&H" & (Left(rDate(ii + 1), 2)))
         rByte(1 + ii * 3) = Val("&H" & (Mid(rDate(ii + 1), 3, 2)))
         rByte(2 + ii * 3) = Val("&H1")
         'rByte(2 + ii * 3) = Val("&H4")
         
      Else
         rByte(0 + ii * 3) = Val("&HFF")
         rByte(1 + ii * 3) = Val("&HFF")
         rByte(2 + ii * 3) = Val("&HFF")
      End If
   Next
      
'Modified by Morgan 2024/12/2
'   iErrCount = 0
'   Do While iErrCount < pRetry
'      iDataLen = 3 * 50 '一次100筆寫會當，分2次，每次50筆
'      sAddr = Val("&H" & "680")
'      'Modified by Morgan 2023/7/27 Timeout改用常數
'      iReturn = hacSetFlashData(cNodeID, rByte(0), sAddr, iDataLen, ghComm, cHTATimOut2)
'      If iReturn = 0 Then
'         sAddr = Val("&H" & "680") + iDataLen
'         iReturn = hacSetFlashData(cNodeID, rByte(iDataLen), sAddr, iDataLen, ghComm, cHTATimOut2)
'         If iReturn = 0 Then
'            HTAWriteHoliday = True
'            Exit Do
'         End If
'      End If
'      iErrCount = iErrCount + 1
'   Loop
   
   iDataLen = 3 * 50 '一次100筆寫會當，分2次，每次50筆
   For ii = 0 To 3
      idx = ii * iDataLen
      sAddr = Val("&H" & "680") + idx
      iErrCount = 0
      Do While iErrCount < pRetry
         iReturn = hacSetFlashData(cNodeID, rByte(idx), sAddr, iDataLen, ghComm, cHTATimOut2)
         If iReturn = 0 Then
            Exit Do
         Else
            iErrCount = iErrCount + 1
         End If
      Loop
   Next
   If ii > 3 Then
      HTAWriteHoliday = True
   End If
'end 2024/12/2
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
   
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'Added by Morgan 2020/9/8
'讀取系統日往後一年的假日(排除星期日)
Public Function PUB_GetHoliday(ByRef pRst As ADODB.Recordset, Optional bDesc As Boolean = False) As Boolean
   Dim stSQL As String, intQ As Integer
   'Modified by Morgan 2023/4/26 排除週六的5/1(智慧局有上班，公司也配合有人需上班)
   stSQL = "select to_char(dt,'yyyy/mm/dd') td,to_char(dt,'D')-1 wd" & _
      " from ( select sysdate+rownum-1 dt from workday where rownum<366 ) X" & _
      " where to_char(dt,'D')>1 and not (to_char(dt,'MMDD')='0501' and to_char(dt,'D')<7)" & _
      " and not exists(select * from workday where wd01=to_char(dt,'yyyymmdd'))" & _
      " order by 1 " & IIf(bDesc, "desc", "asc")
   intQ = 1
   Set pRst = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      PUB_GetHoliday = True
   End If
End Function

'Added by Morgan 2020/9/8
'更新門禁機假日表
Public Function PUB_WriteHoliday(Optional pNoMessage As Boolean = False, Optional pErrMsg As String) As Boolean
   Dim rsQuery As ADODB.Recordset
   Dim arrDate() As String
   Dim iCount As Integer
   Dim arrIP() As String
   Dim ii As Integer
   Dim bErr As Boolean
   
   If PUB_GetHoliday(rsQuery) = True Then
      With rsQuery
      ReDim Preserve arrDate(.RecordCount) As String
      Do While Not .EOF
         arrDate(.AbsolutePosition) = Mid(Replace(.Fields(0), "/", ""), 5, 4)
         .MoveNext
      Loop
      End With
      
      HTAips2 = GetHtaIP(2)
      arrIP = Split(HTAips2, ";")
      For ii = 0 To UBound(arrIP)
         HTAip = arrIP(ii)
         If HTAip <> "" Then
            If HTAWriteHoliday(arrDate, pNoMessage) = False Then
               bErr = True
               pErrMsg = "門禁機 ( " & HTAip & " ) 假日表更新失敗!!"
               If pNoMessage = False Then MsgBox pErrMsg
               Exit For
            End If
         End If
      Next
      If bErr = False Then
         PUB_WriteHoliday = True
      End If
   Else
      pErrMsg = "無法讀取假日資料，門禁機假日表更新失敗！"
      If pNoMessage = False Then MsgBox pErrMsg, vbCritical
   End If
End Function

'Added by Morgan 2024/3/5
'Modified by Morgan 2024/9/19
'考勤機初始化
Public Function HTAInitial() As Boolean
   Dim iInitFlag As Integer
   Dim iReturnCode As Integer, iReturn As Integer
   
On Error GoTo ErrHnd

   If ghComm = 0 Then HTAconnect
   If ghComm = 0 Then Exit Function
   
   'cInitFlag:轉2進位
   'bit0:1 刪除所有合法卡
   'bit1:1 初始化所有表格
   'bit2:1 刪除刷卡記錄
   'bit3:1 其他參數初始化,但不包含系統參數
   'bit4~7 未使用
   iInitFlag = 1 + 2 + 4 + 8 '
   iReturn = hsHTA850Initial(ghComm, iInitFlag, iReturnCode, cHTATimOut2)
   If iReturn = 0 Then
      HTAInitial = True
   End If
   
   If ghComm > 0 Then HTAclose
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

'Added by Morgan 2013/8/2
'Modified by Morgan 2013/9/2 只要清除指紋內容(卡片不刪除,否則刷卡紀錄會對應不到員工號)
'Modified by Morgan 2020/4/27 改為更新打卡紀錄後連卡片資料一併清除
'Modified by Morgan 2024/5/23 +bTrans
'Modified by Morgan 2024/5/27 +卡片也要刪除(避免門禁機有私人卡片)
Public Function PUB_ClearCardData(pUserNo As String, Optional ByRef bTrans As Boolean, Optional ByRef oListBox As ListBox) As Boolean
   Dim arrIpList
   Dim ii As Integer, bolDelete As Boolean, iRCode As Integer
   Dim strQ As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   HTAips = GetHtaIP()
   arrIpList = Split(HTAips, ";")
   strQ = "select scd01,scd02,scd03 from staffcarddata where scd01='" & pUserNo & "' and (scd03 is not null or scd01<>scd02)"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      Do While Not rsQuery.EOF
      '刪除考勤機卡號資料
      For ii = LBound(arrIpList) To UBound(arrIpList)
         HTAip = arrIpList(ii)
         bolDelete = False
         If HTAqueryCard(rsQuery("scd02"), , True, , iRCode) = True Then
            'Added by Morgan 2024/9/9
            If TypeName(oListBox) = "ListBox" Then
               oListBox.AddItem HTAip & " -> 檢查[" & pUserNo & "]卡號資料[" & rsQuery("scd02") & "]...成功", 0
               DoEvents
            End If
            'end 2024/9/9
            bolDelete = True
         ElseIf iRCode <> 6 Then
            cnnConnection.RollbackTrans
            bTrans = False
            'Added by Morgan 2024/9/9
            If TypeName(oListBox) = "ListBox" Then
               oListBox.AddItem HTAip & " -> 檢查[" & pUserNo & "]卡號資料[" & rsQuery("scd02") & "]...失敗", 0
               DoEvents
            End If
            'end 2024/9/9
            MsgBox "考勤機(" & HTAip & ")指紋/卡號(" & rsQuery("scd02") & ")讀取失敗，作業取消!!"
            Exit Function
         
         'Added by Morgan 2024/9/9
         Else
            If TypeName(oListBox) = "ListBox" Then
               oListBox.AddItem HTAip & " -> 檢查[" & pUserNo & "]卡號資料[" & rsQuery("scd02") & "]...無此卡號", 0
               DoEvents
            End If
         'end 2024/9/9
         End If
         If bolDelete = True Then
            If HTAdeleteCard(rsQuery("scd02")) = False Then
               cnnConnection.RollbackTrans
               bTrans = False
               'Added by Morgan 2024/9/9
               If TypeName(oListBox) = "ListBox" Then
                  oListBox.AddItem HTAip & " -> 刪除查[" & pUserNo & "]卡號資料[" & rsQuery("scd02") & "]...失敗", 0
                  DoEvents
               End If
               'end 2024/9/9
               MsgBox "考勤機(" & HTAip & ")指紋/卡號(" & rsQuery("scd02") & ")清除失敗，作業取消!!"
               Exit Function
            'Added by Morgan 2024/9/9
            Else
               If TypeName(oListBox) = "ListBox" Then
                  oListBox.AddItem HTAip & " -> 刪除查[" & pUserNo & "]卡號資料[" & rsQuery("scd02") & "]...成功", 0
                  DoEvents
               End If
            'end 2024/9/9
            End If
         End If
      Next
      rsQuery.MoveNext
      Loop
   End If
   
   PUB_DelRepRec pUserNo 'Added by Morgan 2022/7/26 刪除重複的打卡紀錄(保留卡號最大那一筆)
   
   'Modified by Morgan 2020/4/27 更新打卡紀錄後刪除指紋及卡片資料
   '打卡紀錄卡號改員工號
   'strSql = "update staffcarddata set scd03=null,scd04=null where scd01='" & pUserNo & "'"
   strSql = "update pollrecord set pr03='" & pUserNo & "' where pr03 in (select scd02 from staffcarddata where scd01='" & pUserNo & "' and scd02<>scd01)"
   cnnConnection.Execute strSql, intQ
   
   'Added by Morgan 2024/9/9
   If TypeName(oListBox) = "ListBox" Then
      oListBox.AddItem "更新[" & pUserNo & "]打卡紀錄..." & intQ & "筆", 0
      DoEvents
   End If
   'end 2024/9/9
   
   '清除指紋及卡片資料
   strSql = "delete staffcarddata where scd01='" & pUserNo & "'"
   cnnConnection.Execute strSql, intQ
   
   'Added by Morgan 2024/9/9
   If TypeName(oListBox) = "ListBox" Then
      oListBox.AddItem "刪除[" & pUserNo & "]指紋及卡片資料..." & intQ & "筆", 0
      DoEvents
   End If
   'end 2024/9/9
   
   'Added by Morgan 2020/7/13 新增一筆空的指紋紀錄，否則打卡紀錄無法正常顯示
   strSql = "insert into staffcarddata(scd01,scd02,scd05) values ('" & pUserNo & "','" & pUserNo & "','空指紋.查詢記錄用.勿刪')"
   cnnConnection.Execute strSql, intQ
   'end 2020/7/13
   'end 2020/4/27
   
   'Added by Morgan 2024/9/9
   If TypeName(oListBox) = "ListBox" Then
      oListBox.AddItem "已新增[" & pUserNo & "]空指紋", 0
      DoEvents
   End If
   'end 2024/9/9
   
   PUB_ClearCardData = True
   
   Set rsQuery = Nothing
End Function

'Added by Morgan 2022/7/26
'刪除重複的打卡紀錄(保留卡號最大那一筆)
Public Sub PUB_DelRepRec(pUserNo As String)
   strSql = "delete pollrecord a where pr03 in (select scd02 from staffcarddata where scd01='" & pUserNo & "')" & _
      " and exists(select * from pollrecord b,staffcarddata where pr01=a.pr01 and pr02=a.pr02 and pr03>a.pr03 and scd02(+)=pr03 and scd01='" & pUserNo & "')"
   cnnConnection.Execute strSql, intI
End Sub


