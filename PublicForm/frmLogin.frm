VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  '單線固定
   Caption         =   "登入系統"
   ClientHeight    =   4296
   ClientLeft      =   2832
   ClientTop       =   3480
   ClientWidth     =   5376
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2534.674
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   5042.139
   Begin VB.Timer tmrTimeOut 
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4320
      Top             =   180
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.CommandButton CmdNewBK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "新書到"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   389
      Left            =   4440
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   960
      Width           =   799
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "明細(&B)"
      Height          =   390
      Left            =   4356
      TabIndex        =   7
      Top             =   3840
      Width           =   900
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2010
      TabIndex        =   0
      Top             =   150
      Width           =   1598
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   390
      Left            =   2820
      TabIndex        =   2
      Top             =   960
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   390
      Left            =   3653
      TabIndex        =   3
      Top             =   960
      Width           =   780
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  '暫止
      Left            =   2010
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   1598
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2235
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   5145
      _ExtentX        =   9081
      _ExtentY        =   3937
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|上線日期|序號|程式修改公告"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Line Line1 
      X1              =   98.479
      X2              =   4950.226
      Y1              =   850.791
      Y2              =   850.791
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS :僅顯示一個月內上線之公告"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   3300
   End
   Begin VB.Label lblConnect 
      Caption         =   "連線中, 請稍候 . . . . ."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "使用者名稱："
      Height          =   180
      Index           =   0
      Left            =   945
      TabIndex        =   4
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "使用者密碼："
      Height          =   180
      Index           =   1
      Left            =   945
      TabIndex        =   5
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/15 改成Form2.0 ; GrdDataList改字型=新細明體-ExtB
'Memo By Amy 2013/04/11 顯示程式修改公告
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Dim m_blnFirstShow As Boolean
Dim bolSelData As Boolean, BuIsNoData As Boolean '2013/04/11 add griddatalist
Dim i As Integer  '2013/04/11 add

'Added by Morgan 2016/12/28
'Removed by Morgan 2025/7/18 改全域
'Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
'Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Const OF_READWRITE = &H2

Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_SERVICE_FTP = 1
Const INTERNET_FLAG_PASSIVE = &H8000000
Const FTP_IP As String = "192.168.1.253"
Const FTP_Port As String = "21"
'end 2016/12/28
Dim m_LoadTime As Date 'Added by Morgan 2025/7/4

Private Sub cmdCancel_Click()
    '取消登入
    pub_str_LoginSucceeded = "2"
    'Modify by Amy 2013/04/16
    Set cnnConnection = Nothing
    Unload Me
End Sub

Private Sub cmdDetail_Click()
    PubShowNextData
   Exit Sub
End Sub

Private Sub CmdNewBK_Click()
    Dim bolQuery As Boolean
    bolQuery = frm010035_4.QueryRecord
    frm010035_4.Show vbModal
End Sub

Private Sub cmdok_Click()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strPWD As String  'ADD BY SONIA 2014/8/15
        
    Screen.MousePointer = vbHourglass
    If Me.txtUserName.Text = "" Then
        MsgBox "請輸入使用者名稱!!!", vbExclamation + vbOKOnly
        Me.txtUserName.SetFocus
        txtUserName_GotFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
'    If Me.txtPassword.Text = "" Then
'        MsgBox "請輸入使用者密碼!!!", vbExclamation + vbOKOnly
'        Me.txtPassword.SetFocus
'        txtPassword_GotFocus
'        Exit Sub
'    End If
    'Modify by Amy 2013/04/11
    'Me.Height = 2430: DoEvents
    'MoveFormToCenter Me: DoEvents
    grdDataList.Visible = False: DoEvents
    Line1.Visible = False: DoEvents
    Me.Height = 2415: DoEvents
    lblConnect.Top = 921.7: DoEvents
    lblConnect.Visible = True: DoEvents
   
   If cnnConnection.State = adStateClosed Then 'Added by Morgan 2017/4/25 關閉狀態才要建立連線否則會虛增連線數
      If LoginConnectToServer = False Then End
   End If
   
'    strSQLA = "Select * From Staff, Staff_PWD Where ST01=SP01 And ST04='1' And SP01='" & ChgSQL(Me.txtUserName.Text) & "' And SP02='" & ChgSQL(Me.txtPassword.Text) & "' "
'Morgan 2003/11/19
'    strSQLA = "Select * From Staff, Staff_PWD Where ST01=SP01 And ST04='1' And SP01='" & ChgSQL(UCase(Me.txtUserName.Text)) & "' " & IIf(Me.txtPassword.Text = "", " And SP02 Is Null ", " And SP02='" & ChgSQL(Me.txtPassword.Text) & "' ")
    'MODIFY by sonia 2014/8/15 薪資系統在個入密碼後加輸'sal'
    'StrSQLa = "Select * From Staff, Staff_PWD Where ST01=SP01 And ST04='1' And SP01='" & ChgSQL(UCase(Me.txtUserName.Text)) & "' " & IIf(Me.txtPassword.Text = "", " And SP03 Is Null ", " And SP03='" & ChgSQL(Encrypt(Me.txtPassword.Text, True)) & "' ")
    If UCase(App.EXEName) = "SALARY" Or UCase(App.EXEName) = "TESALARY" Then
       'modify by sonia 2015/12/18 婧瑄說薪資系統密碼後三碼不考慮大小寫
       'If Right(Me.txtPassword.Text, 3) <> "sal" Then
       If UCase(Right(Me.txtPassword.Text, 3)) <> UCase("sal") Then
         lblConnect.Visible = False
         pub_str_LoginSucceeded = "2"
         Set adoTaie = Nothing
         'Set cnnRptConn = Nothing 'Removed by Morgan 2025/9/9 沒用了
         'Set adoTemp = Nothing 'Removed by Morgan 2025/9/9 沒用了
         Set adoEng = Nothing
         strUserNum = ""
         MsgBox "使用者名稱與密碼不符，請重新輸入!", , Me.Name
         Me.txtUserName.SetFocus
         SendKeys "{Home}+{End}"
         If BuIsNoData = False Then
             lblConnect.Visible = False: DoEvents
             grdDataList.Visible = True: DoEvents
             Line1.Visible = True: DoEvents
             Me.Height = 4665: DoEvents
         End If
         GoTo ErrorStop
       Else
         strPWD = Left(Me.txtPassword.Text, Len(Me.txtPassword.Text) - 3)
         StrSQLa = "Select * From Staff, Staff_PWD Where ST01=SP01 And ST04='1' And SP01='" & ChgSQL(UCase(Me.txtUserName.Text)) & "' " & IIf(strPWD = "", " And SP03 Is Null ", " And SP03='" & ChgSQL(Encrypt(strPWD, True)) & "' ")
       End If
    Else
       StrSQLa = "Select * From Staff, Staff_PWD Where ST01=SP01 And ST04='1' And SP01='" & ChgSQL(UCase(Me.txtUserName.Text)) & "' " & IIf(Me.txtPassword.Text = "", " And SP03 Is Null ", " And SP03='" & ChgSQL(Encrypt(Me.txtPassword.Text, True)) & "' ")
    End If
    'end 2014/8/15
    
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
      'Add by Morgan 2003/12/30
      strPassWord = Trim(Me.txtPassword.Text)
      
'        strUserNum = Me.txtUserName.Text
        strUserNum = UCase(Me.txtUserName.Text)
        '成功的登入
        pub_str_LoginSucceeded = "1"
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        
        strUserNum = strUserNum
        'edit by nickc 2007/02/05 不用 dll 了
        'Set obj003.Connection = cnnConnection
        'Set objLawDll.Connection = cnnConnection
        If SetUserData_1 = False Then
            End: DoEvents
        End If
        GetGroupDept
        'Marked By Cheng 2004/04/28
'        strSrvDate(1) = Format(ServerDate)
'        strSrvDate(2) = Format(Val(strSrvDate(1)) - 19110000)
        'End
        Set adoTaie = cnnConnection

        'Modified by Morgan 2013/5/8 判斷非財務系統才連以便與財務系統共用
        'Removed by Morgan 2025/9/9 沒用了
        'If Not (UCase(App.EXEName) = "ACCOUNT" Or UCase(App.EXEName) = "TEACCOUNT" Or UCase(App.EXEName) = "CASHER" Or UCase(App.EXEName) = "TECASHER" Or UCase(App.EXEName) = "FINANCE" Or UCase(App.EXEName) = "TEFINANCE") Then
        '    cnnRptConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\db3.mdb"
        '    cnnRptConn.Open
        'End If
        'adoTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\finance.mdb"
        'adoTemp.Open
        'end 2025/9/9

        'Add By Cheng 2003/04/29
'        If UCase(App.EXEName) = "PROMOTER" Or UCase(App.EXEName) = "TEPROMOTER" Then
        'Modified by Morgan 2011/10/31 +PATPRO1
        'Modify by Sindy 2015/11/3 因經理測試時會改執行檔名稱+數字,所以判斷系統名稱開頭
        'If UCase(App.EXEName) = "PROMOTER" Or UCase(App.EXEName) = "TEPROMOTER" Or UCase(App.EXEName) = "PATPRO" Or UCase(App.EXEName) = "TEPATPRO" Or UCase(App.EXEName) = "PATPRO1" Or UCase(App.EXEName) = "TEPATPRO1" Then
        'Modify By Sindy 2021/7/16 + InStr(UCase(App.EXEName), "LAW") > 0
        'Modify By Sindy 2023/6/20 + InStr(UCase(App.EXEName), "TRADEMARK1") > 0
        If InStr(UCase(App.EXEName), "PROMOTER") > 0 Or _
           InStr(UCase(App.EXEName), "TEPROMOTER") > 0 Or _
           InStr(UCase(App.EXEName), "PATPRO") > 0 Or _
           InStr(UCase(App.EXEName), "TEPATPRO") > 0 Or _
           InStr(UCase(App.EXEName), "PATPRO1") > 0 Or _
           InStr(UCase(App.EXEName), "TEPATPRO1") > 0 Or _
           InStr(UCase(App.EXEName), "LAW") > 0 Or _
           InStr(UCase(App.EXEName), "TRADEMARK1") > 0 Then
        '2015/11/3 END
            'Modified by Morgan 2025/10/31
            'adoEng.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\eng1.mdb"
            'adoEng.Open
            If ConnectEng1Mdb() = False Then
               GoTo ErrorStop
            End If
            'end 2025/10/31
        End If
        
        PUB_SetStaffVar
        pub_bolInformCheck = True 'Add by Morgan 2008/7/4
        
        'Added by Morgan 2022/10/11
        'P12專利處程序部門人員在登入系統時增加檢查是否已設定職代表
        If Pub_StrUserSt03 = "P12" Then
            strExc(0) = "select * from ABS001 where b0101='" & strUserNum & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               MsgBox "程序人員尚未設定職代表，不可操作系統，請通知主管！", vbExclamation
               pub_str_LoginSucceeded = "2"
            End If
        End If
        'end 2022/10/11
        
        'Added by Morgan 2023/6/1
        If Pub_strUserST05 = "" Then
            MsgBox "尚未設定您的等級，不可操作系統，請等待電腦中心作業！", vbExclamation
            pub_str_LoginSucceeded = "2"
        End If
        'end 2023/6/1
        
        Unload Me
        'Removed by Morgan 2013/5/8 移到 basStart 以便與財務系統共用
        'mdiMain.Timer2.Interval = 100 'Modified by Morgan 2012/8/10 要放 Form 消失後否則 mdiMain 會不是 Forms(0)
    Else
        '登入失敗
        lblConnect.Visible = False
        pub_str_LoginSucceeded = "2"
        'Modify by Morgan 2006/11/29
        'Me.Height = 1860
        'Modify by Amy 2013/04/11
        'Me.Height = 2040
        'MoveFormToCenter Me
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        'edit by nickc 2007/02/02 不用 dll 了
        'objPublicData.Connection = Nothing
        'Modify by Amy 2013/04/16 cnnConnection關閉使系統維護公告無法使用
        'Set cnnConnection = Nothing
        'edit by nickc 2007/02/05 不用 dll 了
        'Set obj003.Connection = Nothing
        'Set objLawDll.Connection = Nothing
        Set adoTaie = Nothing
        'Set cnnRptConn = Nothing 'Removed by Morgan 2025/9/9 沒用了
        'Set adoTemp = Nothing 'Removed by Morgan 2025/9/9 沒用了
        Set adoEng = Nothing
        
        strUserNum = ""
        MsgBox "使用者名稱與密碼不符，請重新輸入!", , Me.Name
        Me.txtUserName.SetFocus
        SendKeys "{Home}+{End}"
       
        'Modify by Amy 2013/04/19 顯示程式修改公告
        If BuIsNoData = False Then
            lblConnect.Visible = False: DoEvents
            grdDataList.Visible = True: DoEvents
            Line1.Visible = True: DoEvents
            Me.Height = 4665: DoEvents
        End If
    End If
ErrorStop:
    Screen.MousePointer = vbDefault
End Sub

'Added by Morgan 2025/7/4
Private Sub ChkTimeOut()
   '6小時自動結束系統
   If Now > m_LoadTime + (6 / 24) Then
      cmdCancel.Value = True
   End If
End Sub


Private Sub Form_Activate()
    ChkTimeOut 'Added by Morgan 2025/7/4
    
    If m_blnFirstShow = True Then
        m_blnFirstShow = False
        If Me.txtUserName.Text <> "" Then Me.txtPassword.SetFocus
    End If
    
    'Modify by Amy 2013/04/11 顯示程式修改公告
     SetGridWidth
     lblConnect.Visible = False
     If txtPassword.Text <> "" Then txtPassword.Text = "" 'Added by Morgan 2022/7/15
End Sub

'Added by Morgan 2016/12/30
Private Function GetFtpFile(pFtpFile As String, pLocalFile As String) As Boolean
   Dim IsTimeOut As Boolean
   Dim SeekTimer As Long
   Dim stFTP_IP As String
   Dim stAccount As String, iPos As Integer
   Dim stID As String, stPwd As String
   Dim hOpen As Long
   Dim hConnection As Long
   
   stID = "74001"
   stPwd = "74001"
   stAccount = Pub_GetSpecMan("FTP_Account")
   iPos = InStr(stAccount, ":")
   If iPos > 0 Then
      stID = Left(stAccount, iPos - 1)
      stPwd = Mid(stAccount, iPos + 1)
   Else
      stPwd = stAccount
   End If
   
'   If Winsock1.State <> 0 Then Winsock1.Close
'   Winsock1.Connect Ftp_Ip, Ftp_Port
'   IsTimeOut = False
'   SeekTimer = Timer
'   Do While Winsock1.State = 6 And IsTimeOut = False
'      DoEvents
'      If Timer - SeekTimer > 2 Then
'          IsTimeOut = True
'      End If
'   Loop
'
'   If IsTimeOut = False Then
'      If Winsock1.State = 7 Then
'         Winsock1.Close
'      Else
'         Winsock1.Close
'         FTP_Conn = False
'         Exit Function
'      End If
'   Else
'      Winsock1.Close
'      FTP_Conn = False
'      Exit Function
'   End If
'   Winsock1.Close
   
   hOpen = InternetOpen("Taie Login", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
   If hOpen <> 0 Then
      hConnection = InternetConnect(hOpen, FTP_IP, FTP_Port, _
         stID, stPwd, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
      If hConnection <> 0 Then
         GetFtpFile = FtpGetFile(hConnection, pFtpFile, pLocalFile, False, FILE_ATTRIBUTE_ARCHIVE, 2, 0)
         If hConnection <> 0 Then InternetCloseHandle hConnection
      End If
      If hOpen <> 0 Then InternetCloseHandle hOpen
   End If
End Function

'Added by Morgan 2016/12/30
Private Function StrTimeToSys(oStr As String) As SYSTEMTIME
StrTimeToSys.wYear = Mid(oStr, 1, 4)
StrTimeToSys.wMonth = Mid(oStr, 5, 2)
StrTimeToSys.wDay = Mid(oStr, 7, 2)
StrTimeToSys.wHour = Mid(oStr, 9, 2)
StrTimeToSys.wMinute = Mid(oStr, 11, 2)
StrTimeToSys.wSecond = Mid(oStr, 13, 2)
End Function

Private Sub Form_Load()
Dim lngRt As Long
Dim strUser As String * 100
Dim frm As Form
Dim ii As Integer
   
   'Added by Morgan 2025/7/4
   m_LoadTime = Now
   tmrTimeOut.Interval = 10000
   'end 2025/7/4
    
    m_blnFirstShow = True
    For Each frm In Forms
        Select Case frm.Name
        Case "frmLogin"
            '無動作
        Case "mdiMain"
            frm.Timer1.Interval = 0
            frm.Timer2.Interval = 0
            frm.Hide
        Case Else
            Unload frm
        End Select
    Next
    '尚未登入
    pub_str_LoginSucceeded = ""
    'edit by nickc 2005/09/27
    'Me.Height = 1860
    'Modify by Amy 2013/04/11
    'Me.Height = 2040
    MoveFormToCenter Me
    'edit by nickc 2007/02/02 不用 dll 了
    'objPublicData.Connection = Nothing
    
    Set cnnConnection = Nothing
    'edit by nickc 2007/02/05 不用 dll 了
    'Set obj003.Connection = Nothing
    'Set objLawDll.Connection = Nothing
    Set adoTaie = Nothing
    'Set cnnRptConn = Nothing 'Removed by Morgan 2025/9/9 沒用了
    'Set adoTemp = Nothing 'Removed by Morgan 2025/9/9 沒用了
    Set adoEng = Nothing
    Select Case UCase(App.EXEName)
    Case "PATPRO", "TEPATPRO"
        'Modify By Sindy 2018/11/14
        Me.Caption = "登入專利系統"
    Case "PROMOTER", "TEPROMOTER"
        Me.Caption = "登入承辦人，繪圖人員系統"
    'Modify By Sindy 2018/11/14
    Case "PATPRO", "PATPRO1"
        Me.Caption = "登入國外部專利系統"
    Case "TRADEMARK", "TETRADEMARK"
        Me.Caption = "登入商標系統"
    'Add By Sindy 2018/11/14
    Case "TRADEMARK1", "TETRADEMARK1"
        Me.Caption = "登入國外部商標系統"
    '2018/11/14 END
    Case "WRITER", "TEWRITER"
        Me.Caption = "登入收文系統"
    Case "FILE", "TEFILE"
        Me.Caption = "登入檔案室系統"
    Case "LAW", "TELAW"
        Me.Caption = "登入法務系統"
    Case "COMPUTER", "TECOMPUTER"
        Me.Caption = "登入電腦中心系統"
    Case "ACCOUNT", "TEACCOUNT"
        Me.Caption = "登入財務系統"
    Case "CASHER", "TECASHER"
        Me.Caption = "登入分所財務系統"
    Case "FINANCE", "TEFINANCE"
        Me.Caption = "登入帳務系統"
    Case "SALARY", "TESALARY"
        Me.Caption = "登入薪資系統"
    Case "ABSENCE"
        Me.Caption = "登入出缺勤簽核系統"
    Case Else
    End Select
    lngRt = WNetGetUser("", strUser, 10)
    If lngRt = 0 Then
        Me.txtUserName.Text = Trim(strUser)
'        SendKeys "{Tab}"
    End If
    
    'Modify by Amy 2013/04/11 顯示程式修改公告
    Show_PGMBulletin
    Me.Caption = Me.Caption & PUB_GetDbTerminal
    'Add by Amy 2016/10/03 一週內(工作日5日內)有上架新書架顯示新書到鈕
    CmdNewBK.Visible = False
    'Modify by Amy 2016/12/09 測式期間只開放部分人員
    'If UCase(App.EXEName) = "FILE" Or UCase(App.EXEName) = "TEFILE" Then
    If frm010035_4.QueryRecord = True Then
        CmdNewBK.Visible = True
    End If
    Unload frm010035_4 'Add by Amy 2017/02/02
    'End If
    'end 2016/12/09
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '登入失敗或取消
    If pub_str_LoginSucceeded <> "1" Then
        End: DoEvents
    End If
    IsLoginFormIsUnload = True
    Set frmLogin = Nothing
End Sub

'Add by Amy 2013/04/11
Private Sub GrdDataList_Click()
grdDataList.Visible = False
grdDataList.col = 0
grdDataList.row = grdDataList.MouseRow
If grdDataList.MouseRow <> 0 Then
    If grdDataList.Text = "V" Then
     grdDataList.Text = ""
        For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            grdDataList.CellBackColor = QBColor(15)
        Next i
    Else
     grdDataList.Text = "V"
        For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
        Next i
    End If
End If
grdDataList.Visible = True
End Sub

'Added by Morgan 2025/7/4
Private Sub tmrTimeOut_Timer()
   ChkTimeOut
End Sub

Private Sub txtPassword_GotFocus()
    TextInverse Me.txtPassword
End Sub

Private Sub txtUserName_GotFocus()
    TextInverse Me.txtUserName
End Sub

'連接資料庫
Private Function LoginConnectToServer() As Boolean

On Error GoTo ErrHand
    LoginConnectToServer = False
    'edit by nickc 2007/02/02 不用 dll 了
    'objPublicData.Connection = Nothing
    'edit by nickc 2007/02/05 不用 dll 了
    'Set obj003.Connection = Nothing
    'Set objLawDll.Connection = Nothing
    Set adoTaie = Nothing
    'Set cnnRptConn = Nothing 'Removed by Morgan 2025/9/9 沒用了
    'Set adoTemp = Nothing 'Removed by Morgan 2025/9/9 沒用了
    Set adoEng = Nothing
    
    If ConnectToServer_1 = True Then
        'edit by nickc 2007/02/02 不用 dll 了
        'Set cnnConnection = objPublicData.Connection
        LoginConnectToServer = True
        strSrvDate(1) = Format(ServerDate)
        strSrvDate(2) = Format(Val(strSrvDate(1)) - 19110000)
    End If
    Exit Function
ErrHand:
    MsgBox Err.Description, , "資料庫連結失敗..."
End Function

'Add By Cheng 2003/07/04
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

'Add by Amy 2013/04/11
Private Sub SetGridWidth()
    With grdDataList
           .FormatString = .FormatString
           .ColWidth(0) = 200
           .ColAlignment(0) = flexAlignCenterCenter
           .ColWidth(1) = 885
           .ColAlignment(1) = flexAlignCenterCenter
           .ColWidth(2) = 0
      
           .ColWidth(3) = 3700
           .ColAlignment(3) = flexAlignLeftCenter
   End With
End Sub

'Add by Amy 2013/04/11
Private Sub Show_PGMBulletin()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql, StartDate, SysDate, SysKind As String, strWhere As String
    Dim strTmp As String, strTmpA As String
    Dim intStr As Integer, j As Integer
   
    grdDataList.Clear
    grdDataList.Rows = 2
    If LoginConnectToServer = False Then End
    
    '顯示上線日期為系統日(含)向前推一個月(日曆天)的要公佈資料且公佈系統別有包含登入的系統別
    SysDate = ServerDate()
    StartDate = Val(ChangeWDateStringToWString(DateAdd("d", -30, ChangeWStringToWDateString(DBDATE(SysDate)))))
    If Left(UCase(App.EXEName), 2) = "TE" Then
        SysKind = Right(App.EXEName, Len(App.EXEName) - 2)
    Else
        SysKind = App.EXEName
    End If
    
    'Add By Amy 2013/05/10  電腦中心登入顯示全部系統一個月內的修改資料
    If UCase(App.EXEName) = "TECOMPUTER" Then
        strWhere = ""
    Else
        '第一個字大寫
        SysKind = UCase(Left(SysKind, 1)) & Right(SysKind, Len(SysKind) - 1) & ","
        strWhere = " And instr(BU07,'" & SysKind & "') >0 And BU06='1' "
    End If
    
    
    'Modify By Amy 2013/04/24 改 BU01 >=StartDate
    'strSql = "Select '' as V,sqldatet(BU01) as 上線日期, BU02 as 序號,BU05 as 程式修改公告 " & _
                "From PGMbulletin Where BU01 Between '" & StartDate & "' And '" & SysDate & "' And instr(BU07,'" & SysKind & "') >0 " & _
                "And BU06='1' " & _
                "Order By BU01 Desc,BU02"
    'Modify by Amy 2014/02/14 +<=系統日(公告日未到的資料才能先key而不會顯示)
    strSql = "Select '' as V,sqldatet(BU01) as 上線日期, BU02 as 序號,BU05 as 程式修改公告 " & _
                "From PGMbulletin Where BU01 >= '" & StartDate & "' And BU01<='" & SysDate & "' " & strWhere & _
                "Order By BU01 Desc,BU02"
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        BuIsNoData = False
        Set grdDataList.Recordset = rsTmp
            cmdDetail.Enabled = True
            'Modify by Amy 013/04/18 說明折行 Start
            For i = 1 To grdDataList.Rows - 1
                grdDataList.row = i
                grdDataList.col = 3
                'Modify by Amy 2013/10/11 換行取錯資料
                intStr = 1
                If Trim(grdDataList.Text) <> StrToStr(Trim(grdDataList.Text), 20) Then '超過20個字
                    strTmpA = Trim(grdDataList.Text)
                    strExc(0) = ""
                    Do While strTmpA <> ""
                        strTmp = StrToStr(strTmpA, 20) '目前要取的字串
                        strTmpA = Mid(strTmpA, Len(strTmp) + 1) '取完剩的字串
                        strExc(0) = strExc(0) & strTmp
                        If strTmpA <> "" Then strExc(0) = strExc(0) & vbCrLf
                        intStr = intStr + 1
                    Loop
                 'end2013/10/11
'                intStr = -Int(-(Len(Trim(GrdDataList.Text)) / 20)) '無條件進位
'                strTmpA = ""
'                If intStr > 1 Then
'                    For j = 1 To intStr
'                        strTmp = Mid(GrdDataList.Text, (j - 1) * 20 + 1) '目前要取的字串(以字數算)
'                        If j = intStr Then
'                            strTmpA = strTmpA & PUB_StrToStr_byVal(strTmp, 40)
'                        Else
'                            strTmpA = strTmpA & PUB_StrToStr_byVal(strTmp, 40) & vbCrLf '取後的字串(以byte算)
'                        End If
'                    Next j
                    grdDataList.Text = strExc(0) 'strTmpA
                    grdDataList.RowHeight(i) = 210 * intStr
                End If
            Next i
            'Modify by Amy 013/04/18 說明折行 End
            
                '若查詢結果只有一筆資料
                If Me.grdDataList.Rows = 2 Then
                    grdDataList.row = 1
                    grdDataList.col = 1
                    If grdDataList.Text <> "" Then
                        '直接選定
                        bolSelData = True
                        grdDataList.Visible = False
                        grdDataList.row = 1
                        grdDataList.col = 0
                        grdDataList.Text = "V"
                        For i = 0 To grdDataList.Cols - 1
                            grdDataList.col = i
                            grdDataList.CellBackColor = &HFFC0C0
                        Next i
                        grdDataList.Visible = True
                    End If
     
              End If
    Else
        'Modify by Amy 2013/04/18 無資料不顯示Gird
        'SetGridWidth
        'cmdDetail.Enabled = False
        'GrdDataList.Rows = 2
        'MsgBox ("查無資料")
        BuIsNoData = True
        grdDataList.Visible = False: DoEvents
        Line1.Visible = False: DoEvents
        Me.Height = 2415: DoEvents
        lblConnect.Top = 921.7: DoEvents
        Exit Sub
    End If
          Screen.MousePointer = vbDefault
End Sub

'Add by Amy 2013/04/11
Public Sub PubShowNextData()
Dim i As Integer, j As Integer
    For i = 1 To grdDataList.Rows - 1
        grdDataList.col = 0
        grdDataList.row = i
        If Trim(grdDataList.Text) = "V" Then
            Dim Str01, Str02 As String
            grdDataList.col = 0
            grdDataList.Text = ""
            For j = 0 To grdDataList.Cols - 1
                grdDataList.col = j
                grdDataList.CellBackColor = QBColor(15)
            Next j
            '取上線日
            grdDataList.col = 1
            Str01 = grdDataList.Text
            '取序號
             grdDataList.col = 2
             Str02 = grdDataList.Text
             
             If Not IsNull(grdDataList.Text) Then
                Screen.MousePointer = vbHourglass
                frm100131_2.Tag = Str01 & "," & Str02 & ",frmLogin"
                frm100131_2.StrMenu
                Screen.MousePointer = vbDefault
                frm100131_2.Show vbModal
              End If
              If bolToEndByNick = True Then
               Exit For
              End If
           End If
        Next i

End Sub

'Added by Morgan 2025/10/31
'因Win11系統下連線Mdb有時會有錯，改寫函數並增加重試功能。 (錯誤:-2147217843(80040e4d) Not a valid password)
Private Function ConnectEng1Mdb() As Boolean
   Dim iTimes As Integer
On Error GoTo ErrHnd

   adoEng.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\eng1.mdb"
   adoEng.Open
   ConnectEng1Mdb = True
   Exit Function
   
ErrHnd:
   If Err.Number <> 0 Then
      If iTimes > 10 Then
         MsgBox Err.Description, vbCritical
      Else
         iTimes = iTimes + 1
         Sleep 500
         Resume
      End If
   End If
   
End Function

