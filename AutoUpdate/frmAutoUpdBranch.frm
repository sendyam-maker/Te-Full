VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAutoUpdBranch 
   Caption         =   "自動更新分所系統程式"
   ClientHeight    =   6816
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   8220
   Icon            =   "frmAutoUpdBranch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6816
   ScaleWidth      =   8220
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox Picture1 
      Height          =   345
      Left            =   6975
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   12
      Top             =   4170
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Height          =   2925
      Left            =   1260
      TabIndex        =   0
      Top             =   3960
      Width           =   5190
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   315
         TabIndex        =   1
         Top             =   2250
         Width           =   4650
      End
      Begin VB.Timer Timer1 
         Left            =   4230
         Top             =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "自動更新分所系統程式倒數計時"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   3
         Top             =   210
         Width           =   4620
      End
      Begin VB.Label lblCountDown 
         Alignment       =   2  '置中對齊
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   72
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1440
         Left            =   2025
         TabIndex        =   2
         Top             =   690
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   45
      TabIndex        =   4
      Top             =   60
      Width           =   8115
      Begin VB.CommandButton Command3 
         Caption         =   "結束"
         Height          =   525
         Left            =   6615
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Width           =   1185
      End
      Begin VB.CheckBox chkBranch 
         Caption         =   "高所"
         Height          =   225
         Index           =   2
         Left            =   3015
         TabIndex        =   10
         Top             =   2490
         Width           =   1320
      End
      Begin VB.CheckBox chkBranch 
         Caption         =   "南所"
         Height          =   225
         Index           =   1
         Left            =   1620
         TabIndex        =   9
         Top             =   2490
         Width           =   1320
      End
      Begin VB.CheckBox chkBranch 
         Caption         =   "中所"
         Height          =   225
         Index           =   0
         Left            =   225
         TabIndex        =   8
         Top             =   2490
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         Caption         =   "更新"
         Height          =   525
         Index           =   0
         Left            =   5310
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   1185
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1248
         Left            =   225
         TabIndex        =   6
         Top             =   900
         Width           =   7620
      End
      Begin VB.CommandButton Command1 
         Caption         =   "還原分所清單"
         Height          =   525
         Index           =   1
         Left            =   225
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   1320
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   2880
         Width           =   7710
         _ExtentX        =   13610
         _ExtentY        =   360
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "更新中...."
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   20.4
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   2700
         TabIndex        =   16
         Top             =   330
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "lblSpeed"
         Height          =   180
         Left            =   7275
         TabIndex        =   15
         Top             =   3150
         Width           =   600
      End
      Begin VB.Label lbldownState 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "lbldownState"
         Height          =   180
         Left            =   6975
         TabIndex        =   14
         Top             =   3390
         Width           =   900
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuDisplay 
         Caption         =   "顯示"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "結束"
      End
   End
End
Attribute VB_Name = "frmAutoUpdBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/2/22 改成Form2.0 (無)
Option Explicit

Public Called As Boolean

Const REMOTE_FTP_PORT  As String = "21"
Const MAX_PATH = 260
'右下角圖示用
Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const NIM_MODIFY = &H1
Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_MBUTTONDBLCLK = &H209
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIconA Lib "SHELL32.DLL" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private mlngID As Long
Private mcolNID As Collection
'右下角圖示用 ---

'Ping 用
Const INADDR_NONE As Long = &HFFFFFFFF
Const PING_TIMEOUT As Long = 500
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal s As String) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Dim ECHO As ICMP_ECHO_REPLY
'Ping 用 ---

Dim bolActived As Boolean
Dim stFtpIp As String, stFtpAccB As String, stFtpPwdB As String
Dim stNewProcNetPath As String, stTempFolder As String, stNewProcFtpPath As String
Dim lngAllFileSize As Long, lngUploadFileSize As Long
Dim arrList2() As String '分所清單
Dim bolError As Boolean

Private Sub cmdCancel_Click()
   Timer1.Enabled = False
   If Me.WindowState = vbNormal Then
      Me.Width = Frame2.Width + 200
      Me.Height = Frame2.Height + 400
   End If
   Frame1.Visible = False
   Frame2.Visible = True
   Frame2.BorderStyle = 0
End Sub


Private Function Process() As Boolean
   Dim hOpen As Long, hConnection As Long
   Dim stSource As String, stTarget As String
   Dim F1 As Integer, ii As Integer, jj As Integer, kk As Integer, ss As Integer
   Dim stFilesListLine As String
   Dim arrList() As String
   Dim arrList1() As String '北所清單
   Dim arrList3() As String '待更新清單
   Dim bolDelete As Boolean
   Dim stErrMsg As String
   Dim bolSaveList As Boolean
   
On Error GoTo OutPort

   
   
   '建立分所FTP連線
   hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
   If hOpen = 0 Then
       Err.Raise 999, , "網路錯誤！"
   End If
   hConnection = InternetConnect(hOpen, stFtpIp, REMOTE_FTP_PORT, stFtpAccB, stFtpPwdB, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
   If hConnection = 0 Then
      stErrMsg = "無法與FTP Server建立連線！"
      GoTo OutPort
   End If
   
   '下載分所清單
   AddListItem "下載分所清單"
   stTarget = stTempFolder & "\filelist2.lst"
   stSource = stNewProcFtpPath & "/filelist.lst"
   '有清單
   If FindFtpFile(hConnection, stSource) Then
      If DownloadFile(hConnection, stSource, stTarget) = False Then
         stErrMsg = "分所清單下載失敗！"
         GoTo OutPort
      End If
      If RenFtpFile(hConnection, stSource, stSource & ".f2") = False Then
         stErrMsg = "filelist.lst 更名為 .f2 失敗！"
         GoTo OutPort
      End If
      
   '無清單但有.f2(上次未更新完成)
   ElseIf FindFtpFile(hConnection, stSource & ".f2") Then
      If DownloadFile(hConnection, stSource & ".f2", stTarget) = False Then
         stErrMsg = "分所清單下載失敗！"
         GoTo OutPort
      End If
   End If
   
   '北所清單存到陣列
   ii = 0
   Erase arrList1
   F1 = FreeFile
   Open stTempFolder & "\filelist1.lst" For Input As F1
   Do While Not EOF(F1)
      ii = ii + 1
      Input #F1, stFilesListLine
      ReDim Preserve arrList1(4, ii)
      arrList = Split(stFilesListLine, "||")
      For ss = 0 To UBound(arrList)
         arrList1(ss, ii) = arrList(ss)
      Next
   Loop
   Close #F1
   
   '分所清單存到陣列
   ii = 0
   Erase arrList2
   Erase arrList3
   F1 = FreeFile
   Open stTempFolder & "\filelist2.lst" For Input As F1
   Do While Not EOF(F1)
      ii = ii + 1
      Input #F1, stFilesListLine
      ReDim Preserve arrList2(4, ii)
      ReDim Preserve arrList3(4, ii)
      arrList = Split(stFilesListLine, "||")
      For ss = 0 To UBound(arrList)
         arrList2(ss, ii) = arrList(ss)
         arrList3(ss, ii) = arrList(ss)
      Next
   Loop
   Close #F1
   
   bolSaveList = True '回存分所清單
   
   '比對並儲存待更新檔案清單到陣列
   kk = 0
   lngAllFileSize = 0
   For ii = 1 To UBound(arrList2, 2)
      For jj = 1 To UBound(arrList1, 2)
         If arrList2(0, ii) = arrList1(0, jj) Then
            If arrList2(2, ii) < arrList1(2, jj) Then
               kk = kk + 1
               arrList3(1, ii) = arrList1(1, jj)
               arrList3(2, ii) = arrList1(2, jj)
               arrList3(4, ii) = "*"
               lngAllFileSize = lngAllFileSize + Val(arrList1(1, jj))
            End If
            Exit For
         End If
      Next
   Next
   
   '有更新
   If kk > 0 Then
   
      ProgressBar1.Min = 0
      ProgressBar1.Max = IIf(lngAllFileSize = 0, 102400, lngAllFileSize)
      lngUploadFileSize = 0
      ShowProgressBar

      For jj = 1 To UBound(arrList3, 2)
         If arrList3(4, jj) = "*" Then
            '更新檔案
            AddListItem arrList3(0, jj) & "(" & Format(Val(arrList3(1, jj)) / 1024, "#,###") & "K)"
            
            '上傳新檔
            'If Not UploadFile(hConnection, stNewProcNetPath & "\" & arrList3(0, jj), stNewProcFtpPath & "/" & arrList3(0, jj) & ".f") Then
            If Not UploadFileByPart(hConnection, stNewProcNetPath & "\" & arrList3(0, jj), stNewProcFtpPath & "/" & arrList3(0, jj) & ".f") Then
               stErrMsg = arrList3(0, jj) & "上傳失敗！"
               GoTo OutPort
            Else
               '有舊檔才要更名
               bolDelete = False
               If FindFtpFile(hConnection, stNewProcFtpPath & "/" & arrList3(0, jj)) Then
                  '更名舊檔
                  If Not RenFtpFile(hConnection, stNewProcFtpPath & "/" & arrList3(0, jj), stNewProcFtpPath & "/" & arrList3(0, jj) & ".f2") Then
                     stErrMsg = arrList3(0, jj) & "更名為 .f2 失敗！"
                     GoTo OutPort
                  End If
                  bolDelete = True
               End If
               
               '更名新檔
               If Not RenFtpFile(hConnection, stNewProcFtpPath & "/" & arrList3(0, jj) & ".f", stNewProcFtpPath & "/" & arrList3(0, jj)) Then
                  stErrMsg = arrList3(0, jj) & ".f 更名失敗！"
                  GoTo OutPort
               Else
                  '刪除舊檔
                  If bolDelete Then
                     FtpDeleteFile hConnection, stNewProcFtpPath & "/" & arrList3(0, jj) & ".f2"
                  End If
               End If
            End If
            arrList2(1, jj) = arrList3(1, jj)
            arrList2(2, jj) = arrList3(2, jj)
         End If
      Next
      '更新分所清單
      If Not SaveBranchList(hConnection, stErrMsg) Then GoTo OutPort
      
      
   '無更新
   Else
      '還原分所清單
      If RenFtpFile(hConnection, stSource & ".f2", stSource) = False Then
         stErrMsg = "分所清單還原失敗！"
         GoTo OutPort
      End If
   End If
   
   bolSaveList = False
   Process = True
   
OutPort:
   If Err.Number <> 0 Then
      stErrMsg = Err.Description
   End If
   
   If stErrMsg <> "" Then
      bolError = True
      'MsgBox stErrMsg, vbCritical
      AddListItem stErrMsg
      If bolSaveList Then
         If Not SaveBranchList(hConnection, stErrMsg) Then
            'MsgBox stErrMsg, vbCritical
            AddListItem stErrMsg
         End If
      End If
   End If
   
   If F1 <> 0 Then Close #F1
   If hOpen <> 0 Then InternetCloseHandle hOpen
   If hConnection <> 0 Then InternetCloseHandle hConnection
End Function

Private Function DownloadFile(pConnection As Long, pSource As String, pTarget As String) As Boolean
   If FtpGetFile(pConnection, pSource, pTarget, False, FILE_ATTRIBUTE_ARCHIVE, dwInternetFlags, 0) = 1 Then
      DownloadFile = True
   End If
End Function

Private Function UploadFile(pConnection As Long, pLocalPath As String, pFtpPath As String) As Boolean
   If FindFtpFile(pConnection, pFtpPath) Then
      If FtpDeleteFile(pConnection, pFtpPath) = 0 Then
         Err.Raise 999, , pFtpPath & "檔案已存在且無法刪除!!!"
      End If
   End If
   
   If FtpPutFile(pConnection, pLocalPath, pFtpPath, dwInternetFlags, 0) = 1 Then
      UploadFile = True
   End If
   
End Function

Private Function UploadFileByPart(pConnection As Long, pLocalPath As String, pFtpPath As String) As Boolean
   
   Dim stFileName As String, stFilePath As String
   Dim hFile As Long
   Dim F1 As Integer
   Dim lngWritten As Long
   Dim lngSize As Long
   Dim Data() As Byte
   Dim jj As Integer
   Dim lngBlockSize As Long
   
On Error GoTo OutPort

   lngBlockSize = 102400
   ReDim Data(lngBlockSize - 1)
   
   '不暫存第1分所速度差不多,但第2分所就明顯比較慢
   'stFilePath = pLocalPath
   
   stFileName = Mid(pLocalPath, InStrRev(pLocalPath, "\") + 1)
   stFilePath = stTempFolder & "\" & stFileName
   If Dir(stFilePath) = "" Then
      AddListItem "複製 " & stFileName & " 到暫存區"
      FileCopy pLocalPath, stFilePath
      AddListItem "複製 " & stFileName & " 完成"
   End If
   
   hFile = FtpOpenFile(pConnection, pFtpPath, &H40000000, INTERNET_FLAG_TRANSFER_BINARY, 0)
   If hFile = 0 Then
      Exit Function
   End If
   
   F1 = FreeFile
   Open stFilePath For Binary Access Read As #F1
   
   lngSize = LOF(F1)
   For jj = 1 To lngSize \ 102400
       Get #F1, , Data
       If (InternetWriteFile(hFile, Data(0), lngBlockSize, lngWritten) = 0) Then
           Exit Function
       End If
       DoEvents
       lngUploadFileSize = lngUploadFileSize + lngBlockSize
       ShowProgressBar
   Next
   
   If lngSize Mod lngBlockSize <> 0 Then
      Get #F1, , Data
      If (InternetWriteFile(hFile, Data(0), lngSize Mod lngBlockSize, lngWritten) = 0) Then
           Exit Function
      End If
      lngUploadFileSize = lngUploadFileSize + (lngSize Mod lngBlockSize)
   End If
   ShowProgressBar
   UploadFileByPart = True
   
OutPort:
   If F1 <> 0 Then Close #F1
   If hFile <> 0 Then InternetCloseHandle (hFile)
   
End Function

Private Function RenFtpFile(pConnection As Long, pOldFileName As String, pNewFileName As String) As Boolean
   If FtpRenameFile(pConnection, pOldFileName, pNewFileName) = 1 Then
      RenFtpFile = True
   End If
End Function

Private Function FindFtpFile(pConnection As Long, pFtpPath As String) As Boolean
   Dim hFind As Long, LRet  As Long, stFileName As String
   Dim pData As WIN32_FIND_DATA
   
   stFileName = Mid(pFtpPath, InStrRev(pFtpPath, "/") + 1)
   pData.cFileName = String(MAX_PATH, 0)
   hFind = FtpFindFirstFile(pConnection, pFtpPath & "*", pData, 0, 0)
   If hFind <> 0 Then
      Do
         If InStr(pData.cFileName, stFileName & Chr(0)) = 1 Then
            FindFtpFile = True
            Exit Do
         Else
            LRet = InternetFindNextFile(hFind, pData)
         End If
      Loop While LRet <> 0
      InternetCloseHandle hFind
   End If
End Function

Private Function RestorFileList() As Boolean
   Dim hOpen As Long, hConnection As Long
   Dim stErrMsg As String
   
On Error GoTo ErrHnd

   '建立分所FTP連線
   hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
   If hOpen = 0 Then
       stErrMsg = "網路錯誤！"
       GoTo ErrHnd
   End If
   hConnection = InternetConnect(hOpen, stFtpIp, "21", stFtpAccB, stFtpPwdB, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
   If hConnection = 0 Then
      stErrMsg = "無法與FTP Server建立連線！"
      GoTo ErrHnd
   End If
   '檢查分所清單是否存在
   If Not FindFtpFile(hConnection, stNewProcFtpPath & "/filelist.lst") Then
      If FindFtpFile(hConnection, stNewProcFtpPath & "/filelist.lst.f2") Then
         If RenFtpFile(hConnection, stNewProcFtpPath & "/filelist.lst.f2", stNewProcFtpPath & "/filelist.lst") = False Then
            stErrMsg = "filelist.lst.f2 檔名還原失敗！"
            GoTo ErrHnd
      
         End If
      End If
   End If
   
ErrHnd:
   If Err.Number <> 0 Then
      stErrMsg = Err.Description
   End If
   
   If stErrMsg <> "" Then
      bolError = True
      'MsgBox stErrMsg, vbCritical
      AddListItem stErrMsg
   End If
   
   If hOpen <> 0 Then InternetCloseHandle hOpen
   If hConnection <> 0 Then InternetCloseHandle hConnection
End Function

Private Sub Command1_Click(index As Integer)
   Dim oTime
   Dim lngInterval As Long
   Dim oCheck As CheckBox
   
   If chkBranch(0).Value + chkBranch(1).Value + chkBranch(2).Value = 0 Then
      MsgBox "請勾選所別！", vbExclamation
      Exit Sub
   End If
   
   bolError = False
   Me.Enabled = False
   Label2.Tag = Label2.Caption
   Label2.Visible = True
   AddListItem "更新分所程式開始...................."
   
   List1.Clear
   
   If index = 0 Then
      If Dir(stTempFolder, vbDirectory) = "" Then
         MkDir stTempFolder
      ElseIf Dir(stTempFolder & "\*.*") <> "" Then
         Kill stTempFolder & "\*.*"
      End If
      
      AddListItem "讀取北所清單"
         
      '讀取北所清單
      If Dir(stTempFolder & "\filelist1.lst") <> "" Then
         Kill stTempFolder & "\filelist1.lst"
      End If
      FileCopy stNewProcNetPath & "\filelist.lst", stTempFolder & "\filelist1.lst"
   End If
   
   For Each oCheck In chkBranch
      If oCheck.Value = 1 Then
         stFtpIp = ""
         Select Case oCheck.index
         Case 0 '中所
            'stFtpIp = "192.168.0.250": stPwd = "0418"
            stFtpIp = "192.168.2.253"
            'stPwd = "540327" 'Removed by Morgan 2024/7/12 改抓特殊設定
            
         Case 1 '南所
            'stFtpIp = "192.168.0.252": stPwd = "0418"
            stFtpIp = "192.168.3.253"
            'stPwd = "540327" 'Removed by Morgan 2024/7/12 改抓特殊設定
            
         Case 2 '高所
            'stFtpIp = "192.168.0.253": stPwd = "0418"
            stFtpIp = "192.168.4.253"
            'stPwd = "540327" 'Removed by Morgan 2024/7/12 改抓特殊設定
         End Select
         
         If stFtpIp <> "" Then
            oTime = Time
            Label2.Caption = Label2.Tag & "(" & oCheck.Caption & ")"
            AddListItem oCheck.Caption & "(" & stFtpIp & ") 更新開始"
            
            '先用Winsock檢查FTP Server是否存在,否則若未開機直接建FTP連線有時會很久才有回應(Ex.高所 3分多鐘)
            AddListItem "檢查" & oCheck.Caption & " FTP Server 是否存在"
            If Not CheckServerExist(stFtpIp) Then
               bolError = True
               AddListItem oCheck.Caption & " FTP Server 不存在"
            Else
               If index = 0 Then
                  Process
               Else
                  RestorFileList
               End If
            End If
            
            lngInterval = DateDiff("s", oTime, Time)
            AddListItem oCheck.Caption & "更新結束(" & Format(lngInterval \ 3600, "00") & ":" & Format((lngInterval Mod 3600) \ 60, "00") & ":" & Format(lngInterval Mod 60, "00") & ")"
         End If
      End If
   Next
   AddListItem "更新分所程式結束...................."
   Label2.Visible = False
   Label2.Caption = Label2.Tag
   Me.Enabled = True
   
   If bolError Then
      MsgBox "有錯誤發生，請檢查log檔！(" & App.Path & "\" & App.EXEName & "Log)", vbCritical
   End If
End Sub

Private Sub Command3_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   
   If bolActived = False Then
      bolActived = True
      chkBranch(0).Value = 1
      chkBranch(1).Value = 1
      chkBranch(2).Value = 1
      
      '清單呼叫時不必等
      If Called Then
         AutoRun
      Else
         Me.Width = Frame1.Width + 200
         Me.Height = Frame1.Height + 400
         Frame1.Top = 0
         Frame1.Left = 0
         Frame1.BorderStyle = 0
         Frame2.Visible = False
         Me.Top = (Screen.Height - Me.Height) / 2
         Me.Left = (Screen.Width - Me.Width) / 2
      
         Timer1.Enabled = True
         Timer1.Interval = 1000
         lblCountDown = 5
      End If
   End If
End Sub

'Added by Morgan 2024/7/23
Private Sub SetFtpVer()
   Dim stSQL As String, iQ As Integer
   Dim rsQuery As New ADODB.Recordset
   Dim arrValue() As String
   
On Error GoTo ErrHand

   stNewProcNetPath = "\\Linux\polycom\Setup\NewProc"
   stFtpAccB = "74001"
   stFtpPwdB = "540327"
   
   If PUB_Connect2DB Then
      stSQL = "select oMan from SetSpecMan  where oCode='FTP_AccountB'"
      If rsQuery.State <> adStateClosed Then rsQuery.Close
      rsQuery.CursorLocation = adUseClient
      rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsQuery.RecordCount > 0 Then
         arrValue = Split(rsQuery(0), ":")
         stFtpAccB = arrValue(0)
         stFtpPwdB = arrValue(1)
      End If
      
      stSQL = "select oMan from SetSpecMan  where oCode='FTP_VOL_IP_LINUX'"
      If rsQuery.State <> adStateClosed Then rsQuery.Close
      rsQuery.CursorLocation = adUseClient
      rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsQuery.RecordCount > 0 Then
         stNewProcNetPath = "\\" & rsQuery(0) & "\polycom\Setup\NewProc"
      End If
   End If
   
ErrHand:
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
End Sub

Private Sub Form_Load()
   If mlngID = 0 Then mlngID = AddToSystemTray(Picture1.hWnd, WM_MOUSEMOVE, Me.Icon, Me.Caption)
   
   'Modified by Morgan 2024/7/23
   'stNewProcNetPath = "\\Linux\polycom\Setup\NewProc"
   SetFtpVer
   'end 2024/7/23
   
   stNewProcFtpPath = "//PolyCOM/Setup/NewProc"
   
   '可能會用網路磁碟機啟動執行檔,改固定暫存到本機
   'If Left(App.Path, 2) = "\\" Then
      stTempFolder = "C:\" & App.EXEName & "TEMP"
   'Else
   '   stTempFolder = App.Path & "\TEMP"
   'End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = "1" Then Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If mlngID <> 0 Then
      DeleteFromSystemTray mlngID
      mlngID = 0
   End If
End Sub

Private Sub mnuDisplay_Click()
Me.WindowState = "0"
Me.Visible = True
End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
   lblCountDown = lblCountDown - 1
   If lblCountDown < 1 Then
      AutoRun
   End If
End Sub

Private Sub AutoRun()
   cmdCancel.Value = True
   Command1(0).Value = True
   'Command3.Value = True 'Removed by Morgan 2018/8/13 改不自動結束,確認沒問題後再手動結束
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Msg As Long


If Me.ScaleMode = 1 Then
   Msg = X / Screen.TwipsPerPixelX
Else
  
End If
Select Case Msg
      Case WM_MOUSEMOVE '移動滑鼠
          'Label1.Caption = "正在移動滑鼠"
      Case WM_LBUTTONDBLCLK '連點滑鼠左鍵
          'Label1.Caption = "連點滑鼠左鍵"
          Me.WindowState = "0"
          Me.Visible = True
      Case WM_LBUTTONDOWN '按下滑鼠左鍵
          'Label1.Caption = "按下滑鼠左鍵"
      Case WM_LBUTTONUP '放開滑鼠左鍵
          'Label1.Caption = "放開滑鼠左鍵"
      Case WM_RBUTTONDBLCLK '連點滑鼠右鍵
          'Label1.Caption = "連點滑鼠右鍵"
      Case WM_RBUTTONDOWN '按下滑鼠右鍵
          'Label1.Caption = "按下滑鼠右鍵"
          Me.PopupMenu mnuShow, vbPopupMenuLeftAlign + vbPopupMenuRightButton
      Case WM_RBUTTONUP '放開滑鼠右鍵
          ''Label1.Caption = "放開滑鼠右鍵"
End Select
End Sub

Private Sub ShowProgressBar()
   ProgressBar1.Value = lngUploadFileSize '/ 1024
   lbldownState.Caption = Format(lngUploadFileSize / 1024, "#,###") & "  Ｋ / " & Format(lngAllFileSize / 1024, "#,###") & "  Ｋ"
   lblSpeed.Caption = Format(Trim(lngUploadFileSize / lngAllFileSize * 100), "#.00") & " ％"
   DoEvents
End Sub

Private Sub AddListItem(pMsg As String)
   Dim strText As String
   
   strText = Now & vbTab & pMsg
   List1.AddItem strText
   List1.TopIndex = List1.ListCount - 1
   DoEvents
   
   WriteLog strText
End Sub


Private Sub WriteLog(pMsg As String)
   Dim stLogFolder As String, stLogFile As String, ffa As Integer
   
On Error GoTo ErrHnd
      
   stLogFolder = App.Path & "\" & App.EXEName & "Log"
   If Dir(stLogFolder, vbDirectory) = "" Then
      MkDir stLogFolder
   End If
   
   'log保留一年(清除前一年的log)
   stLogFile = stLogFolder & "\" & (Format(Now, "yyyyww") - 100) & ".log"
   If Dir(stLogFile) <> "" Then
      Kill stLogFile
   End If
   stLogFile = stLogFolder & "\" & (Format(Now, "yyyyww")) & ".log"
   
   ffa = FreeFile
   Open stLogFile For Append As ffa
   Print #ffa, pMsg
   
ErrHnd:
   If ffa <> 0 Then Close ffa
End Sub

'右下角圖示用
Private Function AddToSystemTray(ByVal hWnd As Long, _
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

    mcolNID.Add hWnd, CStr(mlngID)

    Shell_NotifyIconA NIM_ADD, nidTemp
   
    AddToSystemTray = mlngID

End Function

'右下角圖示用
Private Sub DeleteFromSystemTray(ByVal vlngID As Long)

Dim nidTemp As NOTIFYICONDATA

With nidTemp
.cbSize = Len(nidTemp)
.hWnd = mcolNID(CStr(vlngID))
.uID = vlngID
.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
End With

Shell_NotifyIconA NIM_DELETE, nidTemp

End Sub

Private Function SaveBranchList(pConnection As Long, Optional pErrMsg As String) As Boolean
   Dim F1 As Integer, ii As Integer
   Dim stFilesListLine As String
   
On Error GoTo OutPort
   
   pErrMsg = ""
   
   AddListItem "建立分所清單"
   
   '建立分所清單
   F1 = FreeFile
   Open stTempFolder & "\filelist3.lst" For Output As F1
   For ii = 1 To UBound(arrList2, 2)
      stFilesListLine = arrList2(0, ii) & "||" & arrList2(1, ii) & "||" & arrList2(2, ii) & IIf(arrList2(3, ii) <> "", "||" & arrList2(3, ii), "")
      Print #F1, stFilesListLine
   Next
   Close #F1
   
   AddListItem "上傳分所清單"
   
   '上傳分所清單
   If UploadFile(pConnection, stTempFolder & "\filelist3.lst", stNewProcFtpPath & "/filelist.lst") Then
      FtpDeleteFile pConnection, stNewProcFtpPath & "/filelist.lst.f2"
   Else
      pErrMsg = "分所清單上傳失敗！"
      GoTo OutPort
   End If
   SaveBranchList = True
   
OutPort:
   If Err.Number <> 0 Then pErrMsg = Err.Description
   If F1 <> 0 Then Close #F1
End Function

'檢查 Server 是否存在(不用Winsock,因為不知為何即使ip不存在也能連線)
Private Function CheckServerExist(IpAddress As String) As Boolean
     
On Error GoTo ErrHnd

   If Ping(IpAddress, "test", ECHO) = 0 Then
      CheckServerExist = True
   End If
   
ErrHnd:

End Function

Private Function Ping(sAddress As String, sDataToSend As String, ECHO As ICMP_ECHO_REPLY) As Long
   
   Dim hPort As Long
   Dim dwAddress As Long
   
   dwAddress = inet_addr(sAddress)
   
   If dwAddress <> INADDR_NONE Then
   
      hPort = IcmpCreateFile()
      
      If hPort Then
      
         Call IcmpSendEcho(hPort, _
                           dwAddress, _
                           sDataToSend, _
                           Len(sDataToSend), _
                           0, _
                           ECHO, _
                           Len(ECHO), _
                           PING_TIMEOUT)

         Ping = ECHO.status
         Call IcmpCloseHandle(hPort)
      
      End If
      
   Else
         Ping = INADDR_NONE
   End If
  
End Function
