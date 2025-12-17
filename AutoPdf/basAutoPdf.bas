Attribute VB_Name = "basAutoPDF"
Option Explicit

'右下角圖示用
Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4

'Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208

Private mlngID As Long
Private mcolNID As Collection
Private Declare Function Shell_NotifyIconA Lib "SHELL32.DLL" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'連線
Public Function fConnect(ByRef oForm As Form) As Boolean
   Dim sChoice As String, sDS As String
   
On Error GoTo ErrHand
   
   'Modify By Sindy 2019/10/25
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      sChoice = Trim(InputBox("1.【 正式 】資料庫" & vbCrLf & vbCrLf & "2.【 測試 】資料庫", "請選擇連線", "2"))
   Else
      sChoice = "1"
   End If
   If sChoice = Empty Then
      End
   End If
   Select Case sChoice
      Case 1
         sDS = "Live"
      Case 2
         sDS = "Test"
      Case Else
         sDS = "Test"
   End Select
   
   'Removed by Morgan 2025/5/7
   'If UPdateTNSName(sDS) = True Then
   '   sDS = "M51CON"
   'End If
   'end 2025/5/7
   
   Set cnnConnection = Nothing
   If cnnConnection Is Nothing Then
      Set cnnConnection = New ADODB.Connection
   Else
      If cnnConnection.State = adStateOpen Then
         cnnConnection.Close
      End If
   End If
   cnnConnection.ConnectionTimeout = 60
   cnnConnection.Provider = IIf(strProvider <> "", strProvider, cProvider)
   cnnConnection.Properties("Data Source").Value = sDS
   cnnConnection.Properties("User ID").Value = UserName
   cnnConnection.Properties("Password").Value = Password
   cnnConnection.Open
   strServerName = sDS 'Added by Morgan 2019/5/22
   
'   If cnnConnection.State = adStateClosed Then
'      oForm.StatusBar1.Panels(1).Text = "連線中..."
'      cnnConnection.ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=m51con;Persist Security Info=True"
'      cnnConnection.Open
'   End If
   '2019/10/25 END
   
   pub_HostName = PUB_ReadHostName '要記錄電腦名稱否則寄信會失敗
   oForm.Caption = oForm.Caption & PUB_GetDbTerminal
   
   oForm.StatusBar1.Panels(1).Text = "已連線..."
   
   strSrvDate(1) = Format(ServerDate)
   strSrvDate(2) = Format(Val(strSrvDate(1)) - 19110000)
   
   'Modified by Morgan 2017/1/16 改用QPGMR(原來抓windows登入為74001)
   'If ClsPDSetUserData(strUserNum, strUserName, strGroup) = False Then
   strUserNum = "QPGMR"
   If SetUserData_1() = False Then
       End
   'Added by Morgan 2017/9/11
   ElseIf PUB_SetSystemVar = False Then
       End
   'end 2017/9/11
   End If
   'end 2017/1/16
   Set adoTaie = cnnConnection 'Added by Morgan 2020/4/8
   fConnect = True
   
ErrHand:

   If Err.Number <> 0 Then
      oForm.StatusBar1.Panels(1).Text = "連線失敗..."
      MsgBox Err.Description
   End If

End Function

Private Function SetUserData_1() As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset

On Error GoTo ErrHand
    SetUserData_1 = False
    strSql = "select st04,st02,st11 from staff where upper(st01)=" + CNULL(strUserNum)
    rsRecordset.CursorLocation = adUseClient
    rsRecordset.Open strSql, cnnConnection
    If rsRecordset.RecordCount > 0 Then
        If rsRecordset.Fields(0) = "1" Then
            strSql = "begin " + _
                "select st02,st03,st05,st11 into user_data.user_name,user_data.user_department," + _
                "user_data.user_level,user_data.user_group from staff where upper(st01)=" + CNULL(strUserNum) + ";" + _
                "user_data.user_num:=" + CNULL(strUserNum) + ";" + _
                "end;"
            cnnConnection.Execute strSql
            strUserName = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
            strGroup = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
            SetUserData_1 = True
        Else
            ShowMsg MsgText(9165)
        End If
    Else
        ShowMsg MsgText(9166)
    End If
    rsRecordset.Close
Exit Function
ErrHand:
    MsgBox Err.Description
End Function

Public Sub PUB_SetStaffVar()
   If strUserNum <> "" Then
      strExc(0) = "Select ST06,ST03,ST17,ST15 From Staff Where ST01='" & strUserNum & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
        pub_strUserOffice = "" & RsTemp.Fields("ST06")
        Pub_StrUserSt03 = "" & RsTemp.Fields("ST03")
        Pub_StrUserSt17 = "" & RsTemp.Fields("ST17")
        Pub_StrUserSt15 = "" & RsTemp.Fields("ST15")
      End If
   End If
End Sub

'右下角圖示用
Public Function AddToSystemTray(ByVal hWnd As Long, _
                                ByVal vlngCallbackMessage As Long, _
                                ByVal vipdIcon As IPictureDisp, _
                                ByVal vstrTip As String) As Long

    mlngID = mlngID + 1
   
    Dim nidTemp As NOTIFYICONDATA
   
    With nidTemp
        .cbSize = Len(nidTemp)
        .hWnd = hWnd
        .uID = mlngID
        .uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
        .uCallbackMessage = vlngCallbackMessage
        .hIcon = CLng(vipdIcon)
        .szTip = vstrTip & vbNullChar
    End With

    If mcolNID Is Nothing Then Set mcolNID = New Collection

    mcolNID.add hWnd, CStr(mlngID)

    Shell_NotifyIconA NIM_ADD, nidTemp
   
    AddToSystemTray = mlngID
End Function

Public Sub DeleteFromSystemTray(ByVal vlngID As Long)

Dim nidTemp As NOTIFYICONDATA

With nidTemp
.cbSize = Len(nidTemp)
.hWnd = mcolNID(CStr(vlngID))
.uID = vlngID
.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
End With

Shell_NotifyIconA NIM_DELETE, nidTemp

End Sub

'Removed by Morgan 2018/7/3 沒用
''Added by Morgan 2014/3/6
''新增定稿至卷宗區
''pDocNo:收文號,pFileName:pdf檔名
'Public Sub PUB_ConvLetter2Case()
'   Dim stSQL As String, intR As Integer
'   Dim rsQuery As New ADODB.Recordset
'   Dim rsCheck As ADODB.Recordset
'   Dim strFullFileName As String, strFileName As String
'   Dim fs, f
'
'   stSQL = "select * from standbyletter where sl06 is null order by sl03,sl04,sl02"
'   With rsQuery
'   .CursorLocation = adUseClient
'   .MaxRecords = 1
'   .Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
'   If Not (.EOF And .BOF) Then
'      stSQL = "update standbyletter set sl06=sysdate where sl01='" & .Fields("sl01") & "' and sl02='" & .Fields("sl02") & "' and sl03=" & .Fields("sl03") & " and sl04=" & .Fields("sl04")
'      cnnConnection.Execute stSQL, intR
'      If intR = 1 Then
'         If .Fields("sl05") = "1" Then
'            strFileName = "Letter"
'         ElseIf rsQuery("sl05") = "2" Then
'            strFileName = "Order"
'         ElseIf rsQuery("sl05") = "3" Then
'            strFileName = "App"
'         Else
'            strFileName = "Misc"
'         End If
'
'         strFileName = .Fields("sl01") & "." & strFileName & ".pdf"
'
'         If PUB_ConvLetter2PDF(rsQuery("sl02"), rsQuery("sl03"), rsQuery("sl04"), strFullFileName) = True Then
'            Set fs = CreateObject("Scripting.FileSystemObject")
'            Set f = fs.GetFile(strFullFileName)
'            'Modify By Sindy 2015/5/18
'            'If SaveAttFile_PDF(.Fields("sl01"), strFullFileName, strFileName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False, "4") = True Then
'            If SaveAttFile_PDF(.Fields("sl01"), strFullFileName, strFileName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = True Then
'            '2015/5/18 END
'               stSQL = "delete standbyletter where sl01='" & .Fields("sl01") & "' and sl02='" & .Fields("sl02") & "' and sl03=" & .Fields("sl03") & " and sl04=" & .Fields("sl04")
'               cnnConnection.Execute stSQL, intR
'            Else
'               stSQL = "update standbyletter set sl06=null where sl01='" & .Fields("sl01") & "' and sl02='" & .Fields("sl02") & "' and sl03=" & .Fields("sl03") & " and sl04=" & .Fields("sl04")
'               cnnConnection.Execute stSQL, intR
'            End If
'            Kill strFullFileName
'         Else
'            stSQL = "update standbyletter set sl06=null where sl01='" & .Fields("sl01") & "' and sl02='" & .Fields("sl02") & "' and sl03=" & .Fields("sl03") & " and sl04=" & .Fields("sl04")
'            cnnConnection.Execute stSQL, intR
'         End If
'      End If
'   End If
'   End With
'   Set rsQuery = Nothing
'End Sub

'Added by Morgan 2014/4/9
'檢查檔案數量是否符合
Public Function PUB_CheckPDF(pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String, iFileQty As Integer, Optional pDocNo As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stCPP01 As String
   Dim oFileSys As New FileSystemObject
   Dim oFolder As Folder
   Dim oFiles As files
   Dim oFile As File
   Dim iQty As Integer
   Dim iPos As Integer
   Dim stFileName As String
   Dim strCP02 As String, strCP03 As String, strCP04 As String
   
   '先檢查資料庫
   If pDocNo <> "" Then
      stCPP01 = pDocNo
   Else
      stCPP01 = pCP01 & pCP02 & pCP03 & pCP04
   End If
   stSQL = "select nvl(count(*),0) from casepaperpdf where cpp01='" & stCPP01 & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      iQty = rsQuery(0)
   End If
   
   '若數量不足時再檢查匯入區
   If iQty < iFileQty Then
      If Dir("C:\temp\OA\*.*") <> "" Then
         Set oFolder = oFileSys.GetFolder("C:\temp\OA")
         Set oFiles = oFolder.files
         For Each oFile In oFiles
            If UCase(Right(oFile.Name, 4)) = ".PDF" Then
               stFileName = oFile.Name
               '去除路徑
               iPos = InStrRev(stFileName, "\")
               If iPos > 0 Then
                  stFileName = Mid(stFileName, iPos + 1)
               End If
               '抓主檔名
               iPos = InStr(stFileName, ".")
               If iPos > 0 Then
                  stFileName = Left(stFileName, iPos - 1)
               End If
               '轉本所案號格式(追加或多國碼用-分隔)
               If InStr(stFileName, pCP01) = 1 Then
                  stFileName = Mid(stFileName, Len(pCP01) + 1)
                  iPos = InStr(stFileName, "-")
                  If iPos = 0 Then
                     strCP02 = Format(stFileName, "000000")
                     strCP03 = "0"
                     strCP04 = "00"
                  Else
                     strCP02 = Format(Left(stFileName, iPos - 1), "000000")
                     strCP03 = Mid(stFileName, iPos + 1, 1)
                     iPos = InStr(iPos + 1, stFileName, "-")
                     If iPos > 0 Then
                        strCP04 = Format(Mid(stFileName, iPos + 1), "00")
                     Else
                        strCP04 = "00"
                     End If
                  End If
                  If strCP02 = pCP02 And strCP03 = pCP03 And strCP04 = pCP04 Then
                     iQty = iQty + 1
                  End If
               End If
            End If
         Next
      End If
   End If
   PUB_CheckPDF = True
   
flgOK:
   Set rsQuery = Nothing
End Function

'Add By Sindy 2022/10/4 因為Account有引用 basFlow 會連帶需要引用到 Service1
'但因接洽單電子收文就會呼叫到一些案件系統函數, 所以才建此虛函數
Public Function PUB_AutoRecvCRLMain(strSys As String, strCRL01 As String) As Boolean
End Function
