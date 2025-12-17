VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAutoUpdate 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  '沒有框線
   Caption         =   "台一國際專利商標-----自動更新"
   ClientHeight    =   5652
   ClientLeft      =   4116
   ClientTop       =   1476
   ClientWidth     =   9300
   ControlBox      =   0   'False
   Icon            =   "frmAutoUpdate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5652
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3090
      TabIndex        =   29
      Top             =   4740
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Debug"
      Height          =   495
      Left            =   1800
      TabIndex        =   28
      Top             =   4740
      Width           =   1125
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   2
      Left            =   1710
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   1110
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   14
      Left            =   3420
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   2790
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   13
      Left            =   2565
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   2790
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   12
      Left            =   1710
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   2790
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   11
      Left            =   855
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   2790
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   10
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   2790
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "結束"
      Enabled         =   0   'False
      Height          =   405
      Left            =   30
      TabIndex        =   20
      Top             =   630
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Left            =   4950
      ScaleHeight     =   864
      ScaleWidth      =   960
      TabIndex        =   19
      Top             =   6180
      Width           =   1005
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   15
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   3630
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   16
      Left            =   855
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   3630
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   17
      Left            =   1710
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   3630
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   18
      Left            =   2565
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   3630
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   19
      Left            =   3420
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   3630
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   9
      Left            =   3420
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   1950
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   8
      Left            =   2565
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   1950
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   7
      Left            =   1710
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   1950
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   6
      Left            =   855
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   1950
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   5
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   1950
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   4
      Left            =   3420
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   1110
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   3
      Left            =   2565
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   1110
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFC0&
      Height          =   840
      Index           =   1
      Left            =   855
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1110
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6975
      Top             =   2430
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   4335
      _ExtentX        =   7641
      _ExtentY        =   360
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   840
      Index           =   0
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1110
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox DebugList 
      Height          =   5448
      Left            =   4470
      TabIndex        =   27
      Top             =   30
      Width           =   4695
   End
   Begin VB.Label lblState2 
      BackStyle       =   0  '透明
      Caption         =   "印表機檢查"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   1650
      TabIndex        =   30
      Top             =   540
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   180
      Left            =   1080
      TabIndex        =   26
      Top             =   840
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6015
      Top             =   3060
      Width           =   555
   End
   Begin VB.Label lblSpeed 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   180
      Left            =   4230
      TabIndex        =   3
      Top             =   570
      Width           =   45
   End
   Begin VB.Label lbldownState 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   180
      Left            =   4230
      TabIndex        =   2
      Top             =   840
      Width           =   45
   End
   Begin VB.Label lblState 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0080C0FF&
      Caption         =   "狀態顯示區....                    "
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   4335
   End
End
Attribute VB_Name = "frmAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/2 改成Form2.0 (無)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit
Public pOld As Boolean
Dim IsTimeOut As Boolean
Dim SeekTimer
'add by nickc 2008/04/08 debug 用
Dim comString As String
Const cO12Flag As String = "O12.flg"

Private Sub cmd_Click(Index As Integer)

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
    strExc = AllTaieProc(Index + 1).LocalPath & AllTaieProc(Index + 1).ProcName
    'Modified by Morgan 2017/9/8 +改第2參數(lpCommandLine): "" => " 1" (前面空白不可省)
    lngRCode = CreateProcess(strExc, _
                             " 1", _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             udtStartupInfo, _
                             udtProcessInfo)
    If lngRCode = 0& Then
        'MsgBox "Failed to call CreateProcess() function."
        lblState.Caption = "呼叫程式失敗"
        CmdEnd.Visible = True
    Else
        CloseHandle udtProcessInfo.hProcess
        udtProcessInfo.hProcess = 0&
        CloseHandle udtProcessInfo.hThread
        udtProcessInfo.hThread = 0&
    End If
'改以 shell 方式呼叫
'Shell strExc, vbNormalFocus
'解 Ctrl + Alt + Del
'SystemParametersInfo SPI_SCREENSAVERRUNNING, False, pOld, 0
'結束本程式
'Unload Me
End
End Sub

Private Sub CmdEnd_Click()
'SystemParametersInfo SPI_SCREENSAVERRUNNING, False, pOld, 0
'Unload Me
End
End Sub

Private Sub Command1_Click()
'Me.DebugList.Clear
'debugListItem = 0
IsGo = True
Timer2.Enabled = True 'Added by Morgan 2017/9/14
End Sub

Private Sub Command2_Click()
'Unload Me
End
End Sub

Private Sub Form_Load()

'Add by Morgan 2010/7/13
'Dim strLocalIP As String 'Removed by Morgan 2020/8/27 改pub_LocalIP以便共用
Dim strIP4th As String
Dim strUser As String * 10
Dim lngRt As Long

debugListItem = 0
DBMode = False
Me.Width = 4275
Me.Height = 1095
MoveFormToCenter
Me.Show 'Added by Morgan 2024/12/13 先顯示畫面已避免使用者誤以為沒點到

comString = Command '讀取啟動參數

'Added by Morgan 2017/9/26
pub_HostName = PUB_ReadHostName()
'Added by Morgan 2023/3/24
'非多人使用的電腦只允許執行1隻
If Not (Left(LCase(pub_HostName), 4) = "taie" Or LCase(pub_HostName) = "client-win7") Then
   If CheckIsRunning(App.EXEName & ".exe") Then End
End If
'end 2023/3/24
lngRt = WNetGetUser("", strUser, 10)
If lngRt = 0 Then pub_LoginUser = Replace(strUser, Chr(0), "")
'end 2017/9/26

'北所FTP
'Modified by Morgan 2024/8/12
'Taipei_Ftp_ip = "192.168.1.253" 'Added by Morgan 2014/4/10
Taipei_Ftp_ip = GetTpFtpIP
If Taipei_Ftp_ip = "" Then End
'end 2024/7/29

'取得ip
pub_LocalIP = GetLocalIP

'Modified by Morgan 2020/3/17 除分所IP(192.168.2-.5)外,其他都抓北所,目前有 .0,.1,.6(VPN)
'分所FTP
If pub_LocalIP > "192.168.2." And pub_LocalIP < "192.168.5." Then
   Local_Ftp_ip = Left(pub_LocalIP, InStrRev(pub_LocalIP, ".")) & "253"
Else
   Local_Ftp_ip = Taipei_Ftp_ip
End If

'Modified by Morgan 2020/8/27 若分所失敗時改抓北所
'External_Ftp_ip = Local_Ftp_ip
External_Ftp_ip = Taipei_Ftp_ip
'end 2020/8/27
'end 2010/7/13

'O12OnlineCheck 'Added by Morgan 2018/6/7 檢查O12是否上線,要放在連線資料庫前

UpdateProgramData 'Added by Morgan 2017/9/26

'Added by Morgan 2014/6/10
'pdf Creater 設定永不更新
SaveString HKEY_CURRENT_USER, "Software\PDFCreator\Program", "UpdateInterval", "0"
'end 2014/6/10

'Added by Morgan 2023/4/17
'設定造字檔路徑
SaveString HKEY_CURRENT_USER, "EUDC\950", "SystemDefaultEUDCFont", App.Path & "\EUDC.TTE"
'end 2023/4/17


If UCase(comString) = "X" Or InStr(UCase(p_GetModuleFileName), "VB6.EXE") <> 0 Then
    Me.Width = 9270
    Me.Height = 5565
    debugListItem = 0
    DBMode = True
    Put2DBList "啟用偵錯模式！"
End If

Dim ECHO As ICMP_ECHO_REPLY
Dim oNowState As String
Dim oIjk As Integer
'frmMemory.Show
On Error GoTo GetErr
lblState.Caption = ""
IsNetErr = False
'IsStartListen = False
HaveProc = True
Put2DBList "設定隱藏"
'不顯示在工作管理
App.TaskVisible = False
lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision

'取得 hostname
'MsgBox GetIPHostName
'將觸發 timer 的 key 鎖起來
IsGo = False
'add by nickc 2007/12/17
IsHaveEudc = False

For oIjk = 0 To 19
    cmd(oIjk).Visible = False
    
Next oIjk
Put2DBList "程式開始！"

'最上層
If DBMode = False Then
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If


'Me.Top = 1500
Put2DBList "鎖 C+A+D"
'鎖控制按鈕
'鎖 Ctrl+Alt + Del
SystemParametersInfo SPI_SCREENSAVERRUNNING, True, pOld, 0

'檢查有哪些程式準備要檢查
lblState.Caption = "檢查程式...."
DoEvents
Put2DBList "檢查程式"

'Add by Morgan 2011/3/23 改存變數以便共用
'Modified by Morgan 2019/4/15 外專工程師造字已同步,不用再檢查例外--經理
'If Dir(App.Path & "\nochgeudc.txt") = "" Then
   bolNoEUDC = False
'Else
'   bolNoEUDC = True
'End If
'end 2019/4/15

ScanAllTaieProc
'若是沒有程式，一支都沒有，則秀訊息，結束程式
If MaxTaieProc = 0 Then
    HaveProc = False
End If
''檢查分所還是北所
lblState.Caption = "檢查網路...."
DoEvents
Put2DBList "檢查網路"
'lblState.Caption = "檢查網路...."
IsTaipei = True
If GetStatusCode(Ping(Local_Ftp_ip, "test", ECHO)) = "11010" Then
    'Modified by Morgan 2020/8/27
    'If GetStatusCode(Ping(External_Ftp_ip, "test", ECHO)) = "11010" Then
    If External_Ftp_ip = Local_Ftp_ip Then
        lblState.Caption = "請檢查網路狀態！"
        Put2DBList "網路回應錯誤"
        IsNetErr = True
    ElseIf GetStatusCode(Ping(External_Ftp_ip, "test", ECHO)) = "11010" Then
    'end 2020/8/27
    
        'MsgBox "請檢查網路狀態！", , "警告！"
        lblState.Caption = "請檢查網路狀態！"
        Put2DBList "網路回應錯誤"
        IsNetErr = True
    Else
        IsTaipei = False
        IsTimeOut = False
'Modified by Morgan 2024/8/12 取消Winsock檢查--經理
'        oNowState = "checkTrue"
'        If Winsock1.State <> 0 Then Winsock1.Close
'        Winsock1.Connect External_Ftp_ip, External_Ftp_port
'        SeekTimer = Timer
'        Do While Winsock1.State = 6 And IsTimeOut = False
'            DoEvents
'            If Timer - SeekTimer > 10 Then
'                IsTimeOut = True
'            End If
'        Loop
'        '若是超過時間，已舊程式執行
'        If IsTimeOut = False Then
'            If Winsock1.State = 7 Then
'                Winsock1.Close
'                Local_Ftp_ip = External_Ftp_ip 'Added by Morgan 2020/8/27
'            Else
'                'MsgBox "請檢查網路狀態！", , "警告！"
'                Put2DBList "網路錯誤"
'                lblState.Caption = "請檢查網路狀態！"
'                IsNetErr = True
'            End If
'        End If
'        'Winsock1.Close 'Removed by Morgan 2020/8/27
      Local_Ftp_ip = External_Ftp_ip
'end 2024/8/12
    End If
Else
    IsTimeOut = False
'Removed by Morgan 2024/8/12 取消Winsock檢查--經理
'    oNowState = "checkFalse"
'    If Winsock1.State <> 0 Then Winsock1.Close
'    Winsock1.Connect Local_Ftp_ip, Local_Ftp_port
'    SeekTimer = Timer
'    'Modified by Morgan 2024/8/2 改判斷非已連線都要等待，因用名稱連線時會先State=4(識別主機)
'    'Do While Winsock1.State = 6 And IsTimeOut = False
'    Do While Winsock1.State <> 7 And IsTimeOut = False
'    'end 2024/8/2
'        DoEvents
'        If Timer - SeekTimer > 10 Then
'            IsTimeOut = True
'        End If
'    Loop
'    If IsTimeOut = False Then
'        If Winsock1.State = 7 Then
'            Winsock1.Close
'            Local_Ftp_ip = Winsock1.RemoteHostIP 'Added by Morgan 2024/8/2 後面再連線要用IP,否則可能會連不上
'        Else
'            'MsgBox "請檢查網路狀態！", , "警告！"
'            Put2DBList "網路錯誤"
'            lblState.Caption = "請檢查網路狀態！"
'            IsNetErr = True
'        End If
'    End If
'end 2024/8/12
End If
'將觸發 timer 的 key 打開
IsGo = True
Exit Sub
GetErr:
        If Err.Number = "40006" Then
            If oNowState = "checkFalse" Then
                If GetStatusCode(Ping(External_Ftp_ip, "test", ECHO)) = "11010" Then
                    'MsgBox "請檢查網路狀態！", , "警告！"
                    Put2DBList "網路錯誤"
                    lblState.Caption = "請檢查網路狀態！"
                    IsNetErr = True
                Else
                    IsTaipei = False
                    
'Modified by Morgan 2024/8/12 取消Winsock檢查--經理
'                    oNowState = "checkTrue"
'                    If Winsock1.State <> 0 Then Winsock1.Close
'                    Winsock1.Connect External_Ftp_ip, External_Ftp_port
'                    IsTimeOut = False
'                    SeekTimer = Timer
'                    Do While Winsock1.State = 6 And IsTimeOut = False
'                        DoEvents
'                        If Timer - SeekTimer > 10 Then
'                            IsTimeOut = True
'                        End If
'                    Loop
'                    If IsTimeOut = False Then
'                        If Winsock1.State = 7 Then
'                            Winsock1.Close
'                        Else
'                            'MsgBox "請檢查網路狀態！", , "警告！"
'                            Put2DBList "網路錯誤"
'                            lblState.Caption = "請檢查網路狀態！"
'                            IsNetErr = True
'                        End If
'                    End If
'                    Winsock1.Close
                  Local_Ftp_ip = External_Ftp_ip
'end 2024/8/12

                End If
            Else
                'MsgBox "請檢查網路狀態！", , "警告！"
                Put2DBList "網路錯誤"
                lblState.Caption = "請檢查網路狀態！"
                IsNetErr = True
            End If
        End If
        'MsgBox Err.Number & "==>" & Err.Description
        lblState.Caption = Err.Number & "==>" & Err.Description
        Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
If hConnection <> 0 Then InternetCloseHandle hConnection
If hOpen <> 0 Then InternetCloseHandle hOpen
If DBMode = False Then
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
SystemParametersInfo SPI_SCREENSAVERRUNNING, False, pOld, 0
Set frmAutoUpdate = Nothing
End Sub

'Added by Morgan 2020/4/1
Private Sub Timer1_Timer()
   End
End Sub

'第二版用
Private Sub Timer2_Timer()
Dim oIjk As Integer
Dim bolAnswer As String 'Added by Morgan 2020/4/1
Dim stRunners As String 'Added by Morgan 2020/9/30

If IsGo = True Then
    Timer2.Enabled = False 'Added by Morgan 2017/9/14
    '因為只要觸發一次就好，所以再鎖起來
    IsGo = False
    
   'Added by Morgan 2021/6/25 同步系統與控制台印表機
   lblState.Caption = "系統印表機同步中...."
   DoEvents
   lblState2.Caption = ""
   lblState2.Visible = True
   SyncPrinters
   lblState2.Visible = False
   'end 2021/6/25

    'edit by nickc 若是只判斷網路會有問題
    'If IsNetErr = True Then
    If IsNetErr = True Or IsTimeOut = True Then
        If DBMode = False Then
            SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            If MsgBox("目前的程式可能非最新版本，是否執行舊版本？" & vbCrLf & vbCrLf & "(若有疑問請洽程式管理員！)", vbYesNo + vbExclamation, "暫時無法與更新伺服器聯繫！") = vbNo Then
                'Unload Me
                End
                Exit Sub
            End If
            SetWindowPos frmAutoUpdate.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        Else
            Put2DBList "執行舊版程式"
        End If
        IsNetErr = False
    End If
    If HaveProc = False Then
        Put2DBList "搜尋不到任何台一程式"
        lblState.Caption = "找不到台一的任何程式，請重新安裝或通知電腦中心人員！"
        CmdEnd.Visible = True
        CmdEnd.Enabled = True
        Exit Sub
    End If
    
    'If App.PrevInstance Then
    'Modified by Morgan 2020/9/30
    'If CheckIsRun(App.EXEName & ".exe") Then
    If CheckIsRunning(App.EXEName & ".exe", stRunners) Then
    'end 2020/9/30
        If DBMode = False Then
            UpdateProgramData , "-" & App.EXEName & ".exe(" & Left(stRunners, 30) & ")" 'Added by Morgan 2020/9/10
            'Added by Morgan 2020/4/13
            '檢查檔案是否已是最新(多人登入)
            'Modified by Morgan 2020/9/11
            'If ChkFileReady = True Then
            If CheckAllUpdateFile2() = True Then
            'end 2020/9/11
            
               GoTo IsReady
            End If
            'end 2020/4/13
         
            SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            'Modified by Morgan 2020/4/1
            'MsgBox "台一案件管理系統目前正在更新系統，請稍後再執行...", vbExclamation, "無法重覆更新"
            MsgBox "台一案件管理系統目前正在更新系統，請稍後再執行..." & vbCrLf & vbCrLf & "(建議等候 60 秒以上)", vbExclamation, "無法重覆更新"
            SystemParametersInfo SPI_SCREENSAVERRUNNING, True, pOld, 0
            'Unload Me
            End
            Exit Sub
        Else
            Put2DBList "重覆執行"
        End If
    End If
    
    'If IsTimeOut = False And IsNetErr = False Then
    If IsNetErr = False Then
         'FTP 連線
         lblState.Caption = "連線中...."
         Put2DBList "與 FTP 連線中......"
         'Modified by Morgan 2024/8/2
         'FTP_Conn
         If FTP_Conn = False Then
            If MsgBox("目前的程式可能非最新版本，是否執行舊版本？" & vbCrLf & vbCrLf & "(若有疑問請洽程式管理員！)", vbYesNo + vbExclamation, "無法與 FTP 連線") = vbNo Then
                  SystemParametersInfo SPI_SCREENSAVERRUNNING, True, pOld, 0
               End
               Exit Sub
            End If
                  
            GoTo IsReady
         End If
         'end 2024/8/2
         
         'Removed by Morgan 2024/8/12 沒用了
         'KillInstantclient_taie 'Added by Morgan 2019/7/19
         'O12ClientUpdate 'Added by Morgan 2019/6/27 O12Client更新
         'end 2024/8/12
         'SetO12Client 'Added by Morgan 2017/9/27
         
         'add by nickc 2005/04/06 北所及分所收文
'         If IsTaipei = True Or (IsTaipei = False And IsStartListen = True) Then
               '下載清單
               Put2DBList "下載清單"
               DownFileList
               '清單存在就繼續
               'Add by Morgan 2011/4/21 沒有清單時要顯示訊息
               If Dir(App.Path & "\" & "filelist.lst") = "" Then
                  If MsgBox("目前的程式可能非最新版本，是否執行舊版本？" & vbCrLf & vbCrLf & "(若有疑問請洽程式管理員！)", vbYesNo + vbExclamation, "伺服器正在更新中，暫時無法下載程式清單！") = vbNo Then
                        SystemParametersInfo SPI_SCREENSAVERRUNNING, True, pOld, 0
                     End
                     Exit Sub
                  End If
               Else
                       '檢查必須下載之新檔案
                       lblState.Caption = "檢查程式更新...."
                       Put2DBList "檢查新程式"
                       If CheckUpdateNewFile = True Then
                           lblState.Caption = "下載新程式...."
                           TaieAllFileSize_OK = 0
                           Put2DBList "下載新程式"
                           DownLoadAllNewFileToTemp
                           lblState.Caption = "檢查下載之新程式...."
                           Put2DBList "檢查下載的程式檔案大小"
                           CheckNewFileSize
                           Put2DBList "準備替換程式"
                           MoveNewFileToReady
                       End If
                       '檢查自己 ************************************
                       lblState.Caption = "檢查程式更新...."
                       Put2DBList "檢查本程式的更新"
                       If CheckUpdateThisFile2 = True Then
                           
                           lblState.Caption = "下載新程式...."
                           TaieAllFileSize_OK = 0
                           Put2DBList "下載新的本程式"
                           Call DownLoadFileToTemp(UpdateThisProc)
                           lblState.Caption = "檢查下載之新程式...."
                           Put2DBList "檢查本程式的檔案大小"
                           CheckMeSize
                           lblState.Caption = "更新本機程式...."
                           Put2DBList "更新本程式"
                           UpdateMe
                           lblState.Caption = "更新完成，重新啟動更新...."
                           Put2DBList "關閉本支程式"
                           If DBMode = False Then
                                'Unload Me
                                End
                           End If
                           Exit Sub
                            
                       End If
                       
                       'Modify by Morgan 2011/5/31 改先更新程式(若造字失敗仍有新的程式可用)
                       '正常更新 ************************************
                       '檢查更新的檔案數量
                       Put2DBList "檢查需要更新的台一程式"
                       lblState.Caption = "檢查程式更新...."
                       CheckAllUpdateFile
                       '下載到暫存目錄
                       Put2DBList "下載需要更新的台一程式"
                       lblState.Caption = "下載新程式...."
                       DownLoadAllFileToTemp
                       '下載成功，檢查檔案大小
                       Put2DBList "檢查下載的台一程式大小"
                       lblState.Caption = "檢查下載之新程式...."
                       CheckSize
                       
                        '搬檔，若有 dll 則要註冊
                        Put2DBList "開始更新本機台一程式"
                        lblState.Caption = "更新本機程式...."
                        If MoveFileToReady = False Then
                           Put2DBList "更新失敗，提示重新開機再試"
               '            If IsHaveEudc And IsEudcUsing Then
               '                MsgBox "請重新開機！更新造字失敗！", vbExclamation, "使用中..."
               '            Else
                            If DBMode = False Then
                               SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                               MsgBox "請關閉所有新系統！重新執行一次！", vbExclamation, "使用中..."
               '            End If
                               'Unload Me
                               End
                           End If
                           Exit Sub
                        End If
                       
                       '檢查造字 ************************************
                       If bolNoEUDC = False Then
                            lblState.Caption = "檢查造字更新...."
                            Put2DBList "檢查造字更新"
                            If CheckUpdateEUDCFile = True Then
                                lblState.Caption = "下載新造字...."
                                TaieEUDCFileSize_OK = 0
                                Put2DBList "下載新造字"
                                DownLoadEudcNewFileToTemp
                                If IsCntChg = True Then
                                    If DBMode = False Then
                                        SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                                        MsgBox "請通知電腦中心開啟 Fonts 權限，目前將使用舊造字！", vbExclamation, "沒有權限..."
                                    Else
                                        Put2DBList "沒有更新  fonts 權限"
                                    End If
                                 Else
                                     lblState.Caption = "檢查下載之新造字...."
                                     Put2DBList "檢查造字的檔案大小"
                                     CheckEudcFileSize
                                     lblState.Caption = "更新本機造字...."
                                     Put2DBList "更新造字"
                                     bolReboot = False 'Add by Morgan 2008/7/29
                                     If MoveEudcFileToReady Then
                                       If bolReboot = True Then 'Add by Morgan 2008/7/25
                                         SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                                         MsgBox "請重新開機，完成造字更新！", vbExclamation, "使用中..."
                                         End
                                       End If
                                        
                                     Else
                                       If IsEudcUsing Then
                                           SetWindowPos frmAutoUpdate.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                                           MsgBox "請重新開機！造字使用中！", vbExclamation, "使用中..."
                                       End If
                                       End
                                     End If
                                End If
                            End If
                        End If

               End If
               '斷線
               Put2DBList "與 FTP 斷線"
               FTP_Disc
'         Else   'add by nickc 2005/04/06  分所非收文
'               '搜尋 WinSock Server
'               '針對本機，對 Server 發出請求，兼下載程式
'
'         End If
    End If
    '秀畫面，將搜尋到的程式，用按鈕方式秀出
    '先檢查是否只有一個系統
    Put2DBList "更新完成，可以使用程式"
    
IsReady:
    CmdEnd.Enabled = True

    If MaxUpdateTaieProc = 0 Then
        lblState.Caption = "請選擇要執行的功能...."
    Else
        lblState.Caption = "更新完成，請選擇要執行的功能...."
    End If
    '啟動監聽器
'    If IsStartListen = True Then
'        StartListen
'    End If
    If MaxAllRunProc = 1 Then
        cmd(0).Visible = True
        cmd_Click (0)
    Else
        If DBMode = False Then
            Me.Width = 4275
            Me.Height = 1095 + (((MaxAllRunProc \ 5) + IIf(MaxAllRunProc Mod 5 <> 0, 1, 0)) * 850)
        Else
            
        End If
        For oIjk = 0 To MaxAllRunProc - 1
            cmd(oIjk).Visible = True
            '再按鈕上畫圖
            Pic_to_Cmd cmd(oIjk), AllRunProc(oIjk + 1)
        Next oIjk
        MoveFormToCenter
    End If
    If DBMode = True Then
        Put2DBList "========================================"
    End If
    Timer1.Enabled = True
End If
End Sub

Sub StartListen()
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
    strExc = App.Path & "\prjLis.exe"
    lngRCode = CreateProcess(strExc, _
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
        'MsgBox "Failed to call CreateProcess() function."
        lblState.Caption = "呼叫程式失敗"
        CmdEnd.Visible = True
    Else
        CloseHandle udtProcessInfo.hProcess
        udtProcessInfo.hProcess = 0&
        CloseHandle udtProcessInfo.hThread
        udtProcessInfo.hThread = 0&
    End If
End Sub

'Removed by Morgan 2024/12/24 沒用了
'Private Sub SetO12Client()
'   Dim strPath As String, iPos1 As Integer, iPos2 As Integer
'
'On Error GoTo ErrHnd
'
'   'win7以上系統自動安裝O12 Client
'   lblState.Caption = "O12用戶端程式檢查中...."
'   DoEvents
'   If getVersionNo >= 6 Then
'      'strPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%PATH%")
'      strPath = Environ("Path")
'      iPos1 = InStr(UCase(strPath), UCase("C:\instantclient"))
'      '檢查是否已安裝
'      If iPos1 = 0 And InStr(UCase(strPath), UCase("C:\app")) = 0 Then
'         '檢查是否有安裝記錄
'         If UpdateProgramData(3) = False Then
'            UpdateProgramData 1 '紀錄開始安裝時間
'            '複製安裝程式到本機
'            lblState.Caption = "O12用戶端程式複製中...."
'            DoEvents
'            If CopyInstantClient(Local_Ftp_ip) = True Then
'               ChDir "C:\instantclient_taie"
'               lblState.Caption = "O12用戶端程式安裝中...."
'               DoEvents
'               InstallO12Client
'
'               Do While CheckIsRunning("cmd.exe")
'                  Sleep 1000
'               Loop
'               ChDir App.Path
'               Name "C:\instantclient" As "C:\instantclient_"
'               lblState.Caption = "O12用戶端程式安裝完成...."
'               DoEvents
'               UpdateProgramData 2 '紀錄安裝結束時間
'               'ReStart
'            End If
'         End If
'      End If
'
'      '檢查是否啟動O12(Rename安裝目錄)
'      If Dir("C:\instantclient_", vbDirectory) <> "" And iPos1 > 0 Then
'         iPos2 = InStr(UCase(strPath), UCase("C:\orant"))
'         If iPos2 = 0 Or iPos2 > iPos1 Then
'            '沒有登入網域沒有權限Dir
'            'If Dir("\\" & Local_Ftp_ip & "\PolyCOM\Setup\NewProc\O12.flg") <> "" Then
'            If CheckO12 Then
'               Name "C:\instantclient_" As "C:\instantclient"
'               lblState.Caption = "O12用戶端程式已啟用...."
'               DoEvents
'
'            'Added by Morgan 2018/5/31
'            Else
'               O12Test
'            'end 2018/5/31
'            End If
'            UpdateProgramData 4
'         End If
'      End If
'   End If
'
'ErrHnd:
'   If Err.Number <> 0 Then
'      MsgBox Err.Description, vbCritical
'   End If
'End Sub

'Added by Morgan 2017/9/18
'複製 instantclient_taie 到本機
Private Function CopyInstantClient(pSourceIP As String) As Boolean
   Dim fs As Object
   
   On Error GoTo ErrHnd

   Set fs = CreateObject("Scripting.FileSystemObject")
   fs.CopyFolder "\\" & pSourceIP & "\PolyCOM\Setup\NewProc\instantclient_taie", "c:\instantclient_taie"
   If fs.FolderExists("c:\instantclient_") Then
      fs.DeleteFolder "c:\instantclient_"
   End If
   CopyInstantClient = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
   Set fs = Nothing
End Function

Private Sub InstallO12Client()
Dim program_name As String
Dim process_id As Long
Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError
    
    program_name = "C:\instantclient_taie\install_taie.exe"
    process_id = Shell(program_name, vbHide)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    Exit Sub

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub

Private Sub ReStart()
   SystemParametersInfo SPI_SCREENSAVERRUNNING, False, frmAutoUpdate.pOld, 0
   'Shell """" & App.Path & "\" & App.EXEName & ".exe"" R", 1
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
   strExc = App.Path & "\" & App.EXEName & ".exe"
   lngRCode = CreateProcess(strExc, _
                             " R", _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             0&, _
                             udtStartupInfo, _
                             udtProcessInfo)
   If lngRCode = 0& Then
       'MsgBox "Failed to call CreateProcess() function."
       lblState.Caption = "呼叫程式失敗"
   Else
       CloseHandle udtProcessInfo.hProcess
       udtProcessInfo.hProcess = 0&
       CloseHandle udtProcessInfo.hThread
       udtProcessInfo.hThread = 0&
       
       End
   End If
End Sub

Private Function CheckO12() As Boolean
   Dim pData As WIN32_FIND_DATA
   Dim hFind As Long
   
   rcd cRemotePath
   hFind = FtpFindFirstFile(hConnection, cO12Flag, pData, 0, 0)
   If hFind <> 0 Then
      InternetCloseHandle hFind
      CheckO12 = True
   End If
End Function

'Removed by Morgan 2024/12/24 沒用了
'Private Sub O12Test()
'   Dim strDB As String
'   Dim strExe As String
'   Dim strFolder As String, strFolder1 As String
'
'On Error GoTo ExitSub
'
'   strFolder = "C:\instantclient"
'   strFolder1 = "C:\instantclient_"
'
'   If UpdateProgramData(5) = True Then
'      GoTo ExitSub
'   End If
'
'   '紀錄測試開始時間
'   UpdateProgramData 6
'   '檢查O12目錄是否存在
'   If Dir(strFolder1, vbDirectory) = "" Then
'      'MsgBox strFolder1 & " 資料夾不存在，測試結束！", vbExclamation
'      GoTo ExitSub
'   End If
'   '檢查連線測試程式(不存在時自動從FTP下載)
'   strExe = "DbConnTest.exe"
'   If ChkExe(strExe) = False Then Exit Sub
'   '清除機碼
'   SaveSetting "TAIE", "AutoUpd", "DBTest", ""
'   'O12目錄去底線
'   Name strFolder1 As strFolder
'   '呼叫連線測試程式
'   CallExe App.Path & "\" & strExe, "Y"
'   'O12目錄加底線
'   Name strFolder As strFolder1
'   '紀錄測試結束時間
'   UpdateProgramData 7
'   '讀取測試結果
'   strDB = GetSetting("TAIE", "AutoUpd", "DBTest")
'   'O12連線測試成功
'   If strDB = "O12" Then
'      '紀錄測試成功
'      UpdateProgramData 8
'      'MsgBox strDB & " 連線成功！"
'   End If
'   '刪除機碼
'   DeleteSetting "TAIE", "AutoUpd", "DBTest"
'
'ExitSub:
'   'KillExe strExe
'End Sub

Private Function ChkExe(pExe As String) As Boolean
   
On Error GoTo ErrHnd

   If Dir(App.Path & "\" & pExe) = "" Then
      FileCopy "\\" & Local_Ftp_ip & "\PolyCOM\Setup\NewProc\" & pExe, App.Path & "\" & pExe
      
      If Dir(App.Path & "\" & pExe) = "" Then
         'MsgBox strExe & " 下載失敗！"
         Exit Function
      End If
   End If
   ChkExe = True
   Exit Function
   
ErrHnd:
   'MsgBox Err.Description, vbCritical
   
End Function

Private Sub KillExe(pExe As String)
On Error Resume Next
   Kill App.Path & "\" & pExe
End Sub

Private Function CallExe(ByVal program_name As String, parameters As String) As Boolean

   Dim process_id As Long
   Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError

    process_id = Shell("""" & program_name & """ " & parameters, vbHide)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    Exit Function

ShellError:
    'MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Function

'Added by Morgan 2018/6/7
Private Sub O12OnlineCheck()
On Error GoTo ErrHnd
   If Dir("C:\instantclient_", vbDirectory) <> "" Then
      If CheckO12x Then
         Name "C:\instantclient_" As "C:\instantclient"
      End If
   End If
ErrHnd:
End Sub

'Added by Morgan 2019/6/27
'O12用戶端程式更新
Private Sub O12ClientUpdate()
On Error GoTo ErrHnd

   lblState.Caption = "O12用戶端程式檢查中...."
   DoEvents
   
   '有安裝 app
   If Dir("C:\app", vbDirectory) <> "" Then
      '無更新記錄
      If UpdateProgramData(9) = False Then
         '測試連線
         lblState.Caption = "O12測試連線中...."
         DoEvents
         If O12OleDBTest() = True Then
            Call UpdateProgramData(10)
         End If
      End If
   '有安裝 instantclient
   ElseIf Dir("C:\instantclient", vbDirectory) <> "" Then
      '已安裝 oledb
      If Dir("C:\instantclient\oledb", vbDirectory) <> "" Then
         '無更新記錄
         If UpdateProgramData(11) = False Then
            '測試連線
            lblState.Caption = "O12測試連線中...."
            DoEvents
            If O12OleDBTest() = True Then
               Call UpdateProgramData(12)
            End If
         End If
      '未安裝 oledb
      Else
         '無更新記錄
         If UpdateProgramData(13) = False Then
            '複製安裝程式到本機
            lblState.Caption = "O12用戶端程式複製中...."
            DoEvents
            If CopyOleDBFile(Local_Ftp_ip) = True Then
               
               ChDir "C:\instantclient"
               lblState.Caption = "O12用戶端程式安裝中...."
               DoEvents
               InstallO12Oledb
               
               Do While CheckIsRunning("cmd.exe")
                  Sleep 1000
               Loop
               ChDir App.Path
               
               '記錄安裝完成時間
               Call UpdateProgramData(14)
               
               '測試連線
               lblState.Caption = "O12測試連線中...."
               DoEvents
               If O12OleDBTest() = True Then
                  Call UpdateProgramData(12)
               End If
            End If
         End If
      End If
   End If
   
   lblState.Caption = ""
   DoEvents
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Added by Morgan 2019/6/27
Private Sub InstallO12Oledb()
Dim program_name As String
Dim process_id As Long
Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError
    
    program_name = "C:\instantclient\install_oledb_taie.exe"
    process_id = Shell(program_name, vbHide)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    Exit Sub

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub
'Added by Morgan 2019/7/19
Private Sub KillInstantclient_taie()
On Error GoTo ErrHnd

   Dim strFolder As String
   Dim fs As Object
   
   lblState.Caption = "檢查 instantclient_taie ...."
   DoEvents
   
   strFolder = "C:\instantclient_taie"
   Set fs = CreateObject("Scripting.FileSystemObject")
   If fs.FolderExists(strFolder) Then
      lblState.Caption = "刪除 instantclient_taie...."
      DoEvents
      fs.DeleteFolder strFolder, True
   End If
   lblState.Caption = ""
   DoEvents
   
ErrHnd:
   'If Err.Number <> 0 Then MsgBox Err.Description
   Set fs = Nothing
   
End Sub

'Added by Morgan 2019/6/27
Private Function CopyOleDBFile(pSourceIP As String) As Boolean
   Dim fs As Object
   
   On Error GoTo ErrHnd

   Set fs = CreateObject("Scripting.FileSystemObject")
   fs.CopyFolder "\\" & pSourceIP & "\PolyCOM\Setup\NewProc\instantclient", "c:\instantclient"
   CopyOleDBFile = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
   Set fs = Nothing
End Function

'Added by Morgan 2019/6/27
'O12 OleDB 連線測試
Private Function O12OleDBTest() As Boolean
   Dim cnnO12OleDB As New ADODB.Connection
   
On Error GoTo ErrHand
   O12OleDBTest = False
   If cnnO12OleDB.State = adStateOpen Then cnnO12OleDB.Close
   cnnO12OleDB.ConnectionTimeout = 60
   cnnO12OleDB.Provider = "OraOLEDB.Oracle"
   cnnO12OleDB.Properties("Data Source").Value = "M51CON"
   cnnO12OleDB.Properties("User ID").Value = "PGMID"
   cnnO12OleDB.Properties("Password").Value = "PGMPWD"
   cnnO12OleDB.Open
   O12OleDBTest = True
   
ErrHand:
   If Err.Number <> 0 Then MsgBox Err.Description
   If cnnO12OleDB.State = adStateOpen Then cnnO12OleDB.Close
   Set cnnO12OleDB = Nothing
   
End Function

'Added by Morgan 2018/6/7 用SMB檢查O12上線旗標檔(未加入Domain的公用電腦無權限需人工處理)
Private Function CheckO12x() As Boolean
On Error GoTo ErrHnd
   'If comString = "Y" Then '上線前測試用
   '   CheckO12x = True
   'Else
      If Dir("\\" & Local_Ftp_ip & "\PolyCOM\Setup\NewProc\" & cO12Flag) <> "" Then
         CheckO12x = True
      End If
   'End If
ErrHnd:
End Function

'Added by Morgan 2020/4/13
Private Function ChkFileReady() As Boolean
   CheckAllUpdateFile
   If TaieAllFileSize = 0 Then
      ChkFileReady = True
   End If
   
End Function

'Added by Morgan 2021/6/25 同步系統與控制台印表機
Private Sub SyncPrinters()
   Dim colInstalledPrinters, objPrinter
   Dim ii As Integer, jj As Integer, kk As Integer, strPrinter As String, bFind As Boolean
   
   'Removed by Morgan 2024/11/26
   'Dim IFlags As Integer
   'Const wbemFlagReturnImmediately = 16
   'Const wbemFlagForwardOnly = 32
   'IFlags = wbemFlagReturnImmediately + wbemFlagForwardOnly
   'end 2024/11/26

On Error GoTo ErrHnd

   'Modified by Morgan 2024/11/26 有下參數時，colInstalledPrinters.Count屬性會觸發錯誤
   'Set colInstalledPrinters = Interaction.GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("Select * from Win32_Printer", , IFlags)
   Set colInstalledPrinters = Interaction.GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("Select * from Win32_Printer")
   'end 2024/11/26
   If Printers.Count > colInstalledPrinters.Count Then
      jj = 0
      kk = 0
      For ii = Printers.Count - 1 To 0 Step -1 '要從陣列最後開始,否則刪除後索引會變小
         strPrinter = Printers(ii).DeviceName
         bFind = False
         jj = jj + 1
         lblState2 = strPrinter
         Put2DBList lblState2
         For Each objPrinter In colInstalledPrinters
            If objPrinter.Name = strPrinter Then
               bFind = True
               Exit For
            End If
         Next
         If bFind = False Then
            PUB_DelPrinter strPrinter
            lblState2 = lblState2 & "...移除 (" & jj & "/" & Printers.Count & ")"
            kk = kk + 1
         Else
            lblState2 = lblState2 & "...正確 (" & jj & "/" & Printers.Count & ")"
         End If
         DoEvents
         If Printers.Count > 600 Then
            If kk >= 30 Then Exit For '600台以上,1次不要刪除超過30台
         ElseIf Printers.Count > 300 Then
            If kk >= 60 Then Exit For '300台以上,1次不要刪除超過60台
         End If
      Next
   End If
   Exit Sub
   
ErrHnd:
   Put2DBList Err.Description
   
End Sub

'Added by Morgan 2024/8/12
'以系統特殊設定取得FTP IP
Private Function GetTpFtpIP() As String
   Dim adoRst As New ADODB.Recordset
   
   PUB_OpenConn
   If adoConn.State = adStateOpen Then
      adoRst.CursorLocation = adUseClient
      adoRst.Open "select oMan from SetSpecMan  where oCode='FTP_VOL_IP_LINUX'", adoConn, adOpenStatic, adLockReadOnly
      If adoRst.RecordCount > 0 Then
         GetTpFtpIP = adoRst(0)
      End If
      PUB_CloseConn
   End If
   Set adoRst = Nothing
End Function
