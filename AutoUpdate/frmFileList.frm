VERSION 5.00
Begin VB.Form frmFileList 
   BorderStyle     =   0  '沒有框線
   Caption         =   "做清單"
   ClientHeight    =   1788
   ClientLeft      =   0
   ClientTop       =   -36
   ClientWidth     =   4272
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1788
   ScaleWidth      =   4272
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1020
      Top             =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "清單建立中..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   720
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   4095
   End
End
Attribute VB_Name = "frmFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/2/22 改成Form2.0 (無)
'Memo By Morgan 2012/12/10 智權人員欄已修改
Option Explicit
Dim IsOK As Boolean

'Add by Morgan 2011/8/9
'檢查程式是否正在執行中,本程式則檢查是否有2個以上相同程式
Private Function CheckIsRunning(pProcessName As String) As Boolean
   Dim Processes, Process
   
   Set Processes = Interaction.GetObject("winmgmts:").ExecQuery("select * from Win32_Process where name='" & pProcessName & "'")
   If UCase(pProcessName) = UCase(App.EXEName & ".exe") Then
      If Processes.Count > 1 Then CheckIsRunning = True
   Else
      If Processes.Count > 0 Then CheckIsRunning = True
   End If
End Function

Private Sub Form_Load()
If CheckIsRunning(App.EXEName & ".exe") Then
   MsgBox "程式正在執行中，請耐心等候...", vbExclamation
   End
End If
IsOK = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFileList = Nothing
End Sub

Private Sub Timer1_Timer()
Dim MyPath As String
Dim MyName As String
Dim MyDateTime As String
Dim FileHandle As Long
Dim lpReOpenBuff As OFSTRUCT
Dim ft As SYSTEMTIME
Dim FileInfo As BY_HANDLE_FILE_INFORMATION
Dim tZone As TIME_ZONE_INFORMATION
Dim bias As Long
Dim tmpDate As Date
Dim F1 As Integer
Dim tmpFileList() As String
Dim CntFileList As Integer
Dim StrFileList As String
Dim IsHaveStar As Boolean
Dim tmpI As Integer
Dim ArrFileStr As Variant
Dim stIpZone As String

If IsOK = True Then
    IsOK = False
    MyPath = App.Path & "\"
    '先暫存原先有 * 的檔案
    CntFileList = 0
    ReDim Preserve tmpFileList(CntFileList) As String
    If Dir(MyPath & "filelistByTaipei.lst") <> "" Then
         F1 = FreeFile
         Open MyPath & "filelistByTaipei.lst" For Input As F1
         Do While Not EOF(F1)
              Input #F1, StrFileList
              If Mid(StrFileList, 1, 1) = "*" Then
                 CntFileList = CntFileList + 1
                 ReDim Preserve tmpFileList(CntFileList) As String
                 tmpFileList(CntFileList) = StrFileList
              End If
         Loop
         Close #F1
    End If
    If Dir(MyPath & "filelist.lst") <> "" Then
            F1 = FreeFile
            Open MyPath & "filelist.lst" For Input As F1
            Do While Not EOF(F1)
                 Input #F1, StrFileList
                 If Mid(StrFileList, 1, 1) = "*" Then
                    CntFileList = CntFileList + 1
                    ReDim Preserve tmpFileList(CntFileList) As String
                    tmpFileList(CntFileList) = StrFileList
                 End If
            Loop
            Close #F1
      End If
        MyName = Dir(MyPath)
    F1 = FreeFile
    Open MyPath & "filelist.lst" For Output As F1
    Do While MyName <> ""   ' 執行迴圈。
       If MyName <> "." And MyName <> ".." Then
          If (GetAttr(MyPath & MyName) And vbDirectory) <> vbDirectory And UCase(MyName) <> UCase("filelist.lst") And UCase(MyName) <> UCase("filelistByTaipei.lst") Then
            FileHandle = OpenFile(MyPath & MyName, lpReOpenBuff, OF_READ)
            GetFileInformationByHandle FileHandle, FileInfo
            CloseHandle FileHandle
            Call GetTimeZoneInformation(tZone)
            bias = tZone.bias
            FileTimeToSystemTime FileInfo.ftLastWriteTime, ft
            tmpDate = CDate(ft.wYear & "/" & ft.wMonth & "/" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
            IsHaveStar = False
            For tmpI = 0 To UBound(tmpFileList) - 1
               ArrFileStr = Split(tmpFileList(tmpI + 1), "||")
               If UCase(MyName) = UCase(Replace(ArrFileStr(0), "*", "")) Then
                  IsHaveStar = True
                  Print #F1, ArrFileStr(0) & "||" & FileLen(MyPath & MyName) & "||" & Format(tmpDate, "YYYYMMDDHHmmss") & "||" & ArrFileStr(3)
                  Exit For
               End If
            Next tmpI
            If IsHaveStar = False Then
               Print #F1, MyName & "||" & FileLen(MyPath & MyName) & "||" & Format(tmpDate, "YYYYMMDDHHmmss")
            End If
          End If
       End If
       MyName = Dir   ' 尋找下一個目錄。
    Loop
    Close F1
   
   'Add by Morgan 2010/9/10
   'Modified by Morgan 2017/6/1
   'Modified by Morgan 2024/7/23
   'If Left(MyPath, 2) = "\\192.168." Then
   '   stIpZone = Mid(MyPath, 3, 10)
   'Else
   '   stIpZone = Left(GetLocalIP, 10)
   'End If
   stIpZone = Left(GetLocalIP, 10)
   'end 2024/7/23
   
   'Modified by Morgan 2014/5/21 北所也有用 .0 的網段
   'Modified by Morgan 2025/11/10 北所也有用 .5 的網段
   If stIpZone = "192.168.1." Or stIpZone = "192.168.0." Or stIpZone = "192.168.5." Then
      If UpdateDB(MyPath & "filelist.lst") = True Then
         'Added by Morgan 2016/10/
         frmAutoUpdBranch.Called = True
         frmAutoUpdBranch.Show
      End If
      Unload Me
      'End
   Else
      End
   End If
End If
End Sub
'Add by Morgan 2010/9/10
'更新資料庫清單
Private Function UpdateDB(p_fileName As String) As Boolean
   Dim F1 As Integer
   Dim strInput As String
   Dim ArrFileStr
   
   If PUB_Connect2DB Then
      If ChkDB() = False Then Exit Function 'Added by Morgan 2018/3/20
      cnnConnection.BeginTrans
      cnnConnection.Execute "delete filelist"
      F1 = FreeFile
      Open p_fileName For Input As F1
      Do Until EOF(F1)
         Line Input #F1, strInput
         ArrFileStr = Split(strInput, "||")
         If UBound(ArrFileStr) = 2 Then
            cnnConnection.Execute "insert into filelist(fl01,fl02,fl03) values ('" & ArrFileStr(0) & "'," & Val(ArrFileStr(1)) & "," & Val(ArrFileStr(2)) & ")"
         End If
      Loop
      cnnConnection.CommitTrans
      UpdateDB = True
   End If
End Function

Private Function ChkDB() As Boolean
   Dim stSQL As String, iQ As Integer
   Dim rsQuery As New ADODB.Recordset
   Dim stLiveDB As String, stNowDB As String
   
   stSQL = "select TERMINAL FROM V$SESSION where SID=1"
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      stNowDB = "" & rsQuery(0)
   Else
      MsgBox "目前連線資料庫電腦名稱讀取失敗！", vbCritical
      GoTo ExitPoint
   End If
   
   stSQL = "select oMan from SetSpecMan  where oCode='正式資料庫電腦名稱'"
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      stLiveDB = "" & rsQuery(0)
   Else
      MsgBox "正式資料庫電腦名稱讀取失敗！", vbCritical
      GoTo ExitPoint
   End If
   
   If UCase(stNowDB) = UCase(stLiveDB) Then
      ChkDB = True
   ElseIf MsgBox("目前連線資料庫(" & stNowDB & ")並非正式資料庫(" & stLiveDB & ")，是否確定要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
      ChkDB = True
   End If
   
ExitPoint:
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
End Function
