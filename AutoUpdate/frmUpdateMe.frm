VERSION 5.00
Begin VB.Form frmUpdateMe 
   BorderStyle     =   1  '單線固定
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   945
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   300
   End
End
Attribute VB_Name = "frmUpdateMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/2/22 改成Form2.0 (無)
'Memo By Morgan 2012/12/10 智權人員欄已修改
Option Explicit

Private Sub Form_Load()

On Error GoTo GetErr
IsGo = True
    'UpdateCommand = Command()
    UpdateCommand = Clipboard.GetText
    Clipboard.SetText ""
Me.Hide
Exit Sub
GetErr:
    MsgBox Err.Description, vbOKOnly, "警告"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUpdateMe = Nothing
End Sub

Private Sub Timer1_Timer()
Dim ShFileOP As SHFILEOPSTRUCT
Dim len5 As Long
Dim tZone As TIME_ZONE_INFORMATION
Dim bias As Long
Dim writedate As Date
Dim ft As SYSTEMTIME
Dim ofs As OFSTRUCT
Dim hFile As Long
Dim lpct As FILETIME, lplac As FILETIME, lplwr As FILETIME

If IsGo = True Then
    '因為只要觸發一次就好 , 所以再鎖起來
    IsGo = False
    If UpdateCommand = "" Then
        Unload Me
        Exit Sub
    End If
    RecCommand UpdateCommand
    Do While CheckRunIs(UpdateExe)
        DoEvents
    Loop
'    ShFileOP.wFunc = FO_DELETE
'    ShFileOP.pFrom = Trim(UpdatePath) & Trim(UpdateExe) & Chr(0)
'    ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
'    SHFileOperation ShFileOP
'    '將新檔移到剛剛的位置
'    tempPath = String(255, 0)
'    len5 = GetTempPath(256, tempPath)
'    tempPath = Left(tempPath, len5)
    Do While Dir(UpdatePath & UpdateExe) <> ""
        ShFileOP.wFunc = FO_MOVE
        ShFileOP.pFrom = UpdatePath & UpdateExe + Chr(0)
        ShFileOP.pTo = Left(UpdatePath, Len(UpdatePath) - 1) & "\" & UpdateExe & ".f2"
        ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
        SHFileOperation ShFileOP
    Loop
    Do While Dir(UpdatePath & UpdateExe & ".f") <> ""
        ShFileOP.wFunc = FO_MOVE
        ShFileOP.pFrom = Trim(UpdatePath) & Trim(UpdateExe) & ".f" + Chr(0)
        ShFileOP.pTo = Left(Trim(UpdatePath), Len(Trim(UpdatePath)) - 1) & "/" & UpdateExe
        ShFileOP.fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
        SHFileOperation ShFileOP
    Loop
    Do While Dir(UpdatePath & UpdateExe & ".f2") <> ""
        ShFileOP.wFunc = FO_DELETE
        ShFileOP.pFrom = Trim(UpdatePath) & Trim(UpdateExe) & ".f2" & Chr(0)
        ShFileOP.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
        SHFileOperation ShFileOP
    Loop
    '更改日期
    Call GetTimeZoneInformation(tZone)
    bias = tZone.bias
    writedate = CDate(Mid(UpdateDate, 1, 4) & "/" & Mid(UpdateDate, 5, 2) & "/" & Mid(UpdateDate, 7, 2) & " " & Mid(UpdateDate, 9, 2) & ":" & Mid(UpdateDate, 11, 2) & ":" & Mid(UpdateDate, 13, 2)) + TimeSerial(0, bias, 0)
    ft.wYear = Year(writedate)
    ft.wMonth = Month(writedate)
    ft.wDay = Day(writedate)
    ft.wHour = Hour(writedate)
    ft.wMinute = Minute(writedate)
    ft.wSecond = Second(writedate)
    ft.wDayOfWeek = Weekday(writedate)
    ft.wMilliseconds = Mid(UpdateDate, 15)
    hFile = OpenFile(UpdatePath & UpdateExe, ofs, OF_READWRITE)
    'UpdateTaieProc(oIjk).RemoteDate.wMinute = UpdateTaieProc(oIjk).RemoteDate.wMinute + bias
    Call SystemTimeToFileTime(ft, lplwr)
    '更動hFile的時間，第2個參數改Create DateTime
    '第3個參數改Last Access DateTime
    '第四個參數改Last Modify DateTime
    Call SetFileTime(hFile, ByVal 0, ByVal 0, lplwr) '只更改第三個(最後寫入)時間
    Call CloseHandle(hFile) '關閉檔案
    StartAutoUpdteExe
    End
End If
End Sub

Sub StartAutoUpdteExe()
'再開另一個trared  執行程式
    Dim strExc As String
    Dim lngRCode As Long
    Dim udtStartupInfo As STARTUPINFO
    Dim udtProcessInfo As PROCESS_INFORMATION
    
    udtStartupInfo.cb = Len(udtStartupInfo)
    udtStartupInfo.dwFlags = STARTF_USESHOWWINDOW
    udtStartupInfo.wShowWindow = SW_SHOWDEFAULT

    udtProcessInfo.dwProcessId = 0&
    udtProcessInfo.dwThreadId = 0&
    udtProcessInfo.hProcess = 0&
    udtProcessInfo.hThread = 0&
    strExc = Trim(UpdatePath) & Trim(UpdateExe)
    lngRCode = CreateProcess(Trim(strExc), _
                             "", _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             udtStartupInfo, _
                             udtProcessInfo)
    If lngRCode = 0& Then
        MsgBox "Failed to call CreateProcess() function."
    Else
        CloseHandle udtProcessInfo.hProcess
        udtProcessInfo.hProcess = 0&
        CloseHandle udtProcessInfo.hThread
        udtProcessInfo.hThread = 0&
    End If
'結束本程式
'Shell strExc, vbMaximizedFocus
Unload Me
End Sub
