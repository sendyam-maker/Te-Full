VERSION 5.00
Begin VB.Form frmReopen 
   Caption         =   "重新連線"
   ClientHeight    =   2030
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   3830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2030
   ScaleWidth      =   3830
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer tmrTimeOut 
      Left            =   990
      Top             =   1140
   End
   Begin VB.Timer TmrClose 
      Left            =   360
      Top             =   990
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  '暫止
      Left            =   1380
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   525
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   390
      Left            =   2790
      TabIndex        =   2
      Top             =   1020
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   390
      Left            =   1815
      TabIndex        =   1
      Top             =   1020
      Width           =   900
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1380
      TabIndex        =   3
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "使用者密碼："
      Height          =   270
      Index           =   1
      Left            =   75
      TabIndex        =   6
      Top             =   540
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      Caption         =   "使用者名稱："
      Height          =   270
      Index           =   0
      Left            =   75
      TabIndex        =   5
      Top             =   150
      Width           =   1200
   End
   Begin VB.Label lblConnect 
      Caption         =   "連線中, 請稍候 . . . . ."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmReopen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

Dim m_iLstPointer As Integer '最後指標狀態 Added by Morgan 2012/10/2
Dim m_DS As String 'Added by Morgan 2016/3/22 記錄前次的連線主機
Dim m_LoadTime As Date 'Added by Morgan 2025/7/4

Private Sub cmdCancel_Click()
   cmdCancel.Tag = "C" 'Add By Sindy 2020/6/16
   Unload Me
End Sub

'Create by Morgan 2003/12/30
Private Sub cmdok_Click()
   If Trim(txtPassword.Text) = strPassWord Then
      lblConnect.Visible = True
      Me.Height = 2430
      'Me.Refresh
      DoEvents
      If ReConnect() = True Then
         'Add by Morgan 2009/7/8 沒關機要重新登入系統
         'Modify By Sindy 2010/7/21
         'If strSrvDate(1) <> Format(ServerDate) Then
         If Val(strSrvDate(1)) <> Val(Format(ServerDate)) Then
            MsgBox "系統日期已經變更,請重新登入!!"
            Exit Sub
         End If
         'end 2009/7/8
            
         If SetUserData_1() = False Then
             MsgBox "資料庫變數設定失敗！"
         Else
            mdiMain.bolReOpen = True
            'Add By Cheng 2004/04/20
            '記錄使用者所別
            pub_strUserOffice = PUB_GetST06(strUserNum)
            'End
            pub_bolInformCheck = True 'Add by Morgan 2008/7/4
            'mdiMain.m_blnABSActivated = False 'Add by Morgan 2011/10/25
            Unload Me
         End If
      End If
   Else
      MsgBox "密碼錯誤！", vbCritical
      txtPassword.SetFocus
      Call txtPassword_GotFocus
   End If
End Sub

'Added by Morgan 2025/7/4
Private Sub ChkTimeOut()
   '6小時自動結束系統
   If Now > m_LoadTime + (6 / 24) Then
      cmdCancel.Value = True
   End If
End Sub

'Added by Morgan 2025/7/4
Private Sub Form_Activate()
   ChkTimeOut
End Sub

Private Sub Form_Load()
Dim frm As Form
Dim strShowMsg As String
   
   'Added by Morgan 2025/7/4
   m_LoadTime = Now
   tmrTimeOut.Interval = 5000
   'end 2025/7/4
   
   m_DS = cnnConnection.Properties("Data Source").Value 'Added by Morgan 2016/3/22
   
   'mdiMain.SetFocus
   'Modified by Morgan 2011/12/6 主畫面為最小化時還原成原來大小
   'mdiMain.WindowState = 2
   If mdiMain.WindowState = 1 Then
      If mdiMain.m_wasMaximized = True Then
         mdiMain.WindowState = 2
      Else
         mdiMain.WindowState = 0
      End If
   End If
   
   Me.Height = 1860
   MoveFormToCenter Me, True
   txtUserName.Text = strUserNum
   lblConnect.Visible = False
   TmrClose.Interval = 10000
   'Added by Morgan 2012/10/2
   m_iLstPointer = Screen.MousePointer
   Screen.MousePointer = vbDefault
   'end 2012/10/2
   
'   'Add By Sindy 2020/4/10
'   '檢查若有開著 frm880019.新郵件 則自動關閉
'   '因為DB已斷線了, 再寄出去也無法回寫DB程式, 會資料不一致
'   For Each frm In Forms
'      If frm.Name = frm880019.Name Then
'         Unload frm
'         strShowMsg = "資料庫已斷線！無法寄出【新郵件】已關閉！"
'      End If
'   Next
'   For Each frm In Forms
'      If frm.Name = frm090202_2.Name Then
'         Unload frm
'      End If
'   Next
'   If strShowMsg <> "" Then MsgBox strShowMsg
'   '2020/4/10 END
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Add By Sindy 2020/6/12
   If cmdCancel.Tag = "C" Then
   If InStr(UCase(App.EXEName), "PROMOTER") > 0 Then
      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
         'Modify By Sindy 2025/7/17 柏翰經理提:這個提醒主要是疫情期間隨時可能居家辦公所以提醒要上傳檔案
         '                                     現在已無這需求,故請移除這個提醒
'         'Add By Sindy 2020/3/20 專利處主管,專利處工程師,專利處繪圖要提醒詢問,但排除王副總因為沒承辦案件
'         If (Pub_StrUserSt03 = "P10" Or Pub_StrUserSt03 = "P11" Or Pub_StrUserSt03 = "P13") And _
'            strUserNum <> "71011" Then
'            cmdCancel.Tag = ""
'            If MsgBox("未完成稿件是否已上傳暫存區？", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption & " 重要訊息！") = vbNo Then
''               If Pub_StrUserSt03 <> "P13" Then
''                  ProState = "1"
''                  ProSysState = "1"
''                  frm090201_2.Show
''               End If
'               Cancel = True
'            End If
'         End If
'         '2020/3/20 END
      End If
   End If
   End If
   '2020/6/12 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Morgan 2005/6/30
   '結束也要連線因為有些form離開會列印
   If mdiMain.bolReOpen = False Then
      lblConnect.Visible = True
      Me.Height = 2430
      DoEvents
      If ReConnect() = True Then
         If SetUserData_1() = False Then
            MsgBox "資料庫變數設定失敗！"
         Else
            pub_strUserOffice = PUB_GetST06(strUserNum)
         End If
      End If
   End If
   '2005/6/30 end
   Screen.MousePointer = m_iLstPointer 'Added by Morgan 2012/10/2
   Set frmReopen = Nothing
End Sub

'Add by Morgan 2011/8/26
'若已2小時沒作業且時間為早上7點以前則自動結束
Private Sub tmrTimeOut_Timer()
   'Modified by Morgan 2025/7/4
   'Static ii As Integer
   'ii = ii + 1
   'If ii > 120 And Format(Now, "HH") < 7 Then
   '   cmdCancel.Value = True
   'End If
   ChkTimeOut
   'end 2025/7/4
End Sub

Private Sub TmrClose_Timer()
   'NothingAllGrid 'Removed by Morgan 2024/12/12 先取消,因重新連線後再操作Grid會錯
   If cnnConnection.State = adStateOpen Then
      'Add by Morgan 2005/4/21 若有開財務的Form時不斷線
      If CheckAccForm = False Then
         cnnConnection.Close
      End If
   End If
   TmrClose.Interval = 0
End Sub

Private Function CheckAccForm() As Boolean
   Dim frm As Form
   For Each frm In Forms
      If UCase(Left(frm.Name, 6)) = "FRMACC" Then
         CheckAccForm = True
         Exit For
      End If
   Next
End Function

Private Sub txtPassword_GotFocus()
   TextInverse txtPassword
End Sub

Private Function ReConnect() As Boolean
On Error GoTo ErrHand
   If cnnConnection.State = adStateClosed Then
      'Added by Morgan 2012/7/16 Win7 斷線後連線資訊會被清除,所以重新設定
      cnnConnection.ConnectionTimeout = 60
      cnnConnection.Provider = IIf(strProvider <> "", strProvider, cProvider)
      'Modified by Morgan 2016/3/22 改恢復前次的連線(測試會切換成與M51CON設的不同)
      If m_DS <> "" Then
         cnnConnection.Properties("Data Source").Value = m_DS
      Else
         cnnConnection.Properties("Data Source").Value = ServerName
      End If
      cnnConnection.Properties("User ID").Value = UserName
      cnnConnection.Properties("Password").Value = Password
      'end 2012/7/16
      cnnConnection.Open
      Forms(0).Caption = Left(Forms(0).Caption, InStr(Forms(0).Caption, "(") - 1) & PUB_GetDbTerminal 'Added by Morgan 2016/3/22
   End If
   ReConnect = True
   
   Exit Function
ErrHand:
   MsgBox Err.Description
   
End Function

'Copy from frmLogin 2004/4/12
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

'Add by Morgan 2005/2/18
Private Sub NothingAllGrid()
   Dim frmTemp As Form
   Dim ctlTemp As Control
   For Each frmTemp In Forms
      For Each ctlTemp In frmTemp.Controls
         If TypeName(ctlTemp) = "MSHFlexGrid" Then
            Set ctlTemp.Recordset = Nothing
         End If
      Next
   Next
End Sub

'Add By Sindy 2010/11/25
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
