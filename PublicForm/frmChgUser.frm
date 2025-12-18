VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmChgUser 
   BorderStyle     =   1  '單線固定
   Caption         =   "更改使用者"
   ClientHeight    =   2112
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4392
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2112
   ScaleWidth      =   4392
   Begin TabDlg.SSTab SSTab1 
      Height          =   1965
      Left            =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   4290
      _ExtentX        =   7557
      _ExtentY        =   3471
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmChgUser.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "   "
      TabPicture(1)   =   "frmChgUser.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "Text2"
      Tab(1).Control(2)=   "Label2"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1470
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "確定"
         Height          =   405
         Left            =   -72255
         TabIndex        =   6
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox Text2 
         Height          =   345
         IMEMode         =   3  '暫止
         Left            =   -73875
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   990
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   195
         TabIndex        =   3
         Top             =   630
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "確定"
         Default         =   -1  'True
         Height          =   405
         Left            =   2760
         TabIndex        =   2
         Top             =   600
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "確定後不可再更改使用者"
         Height          =   225
         Left            =   180
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1125
         Width           =   3840
      End
      Begin VB.Label Label3 
         Caption         =   "系統日期："
         Height          =   285
         Left            =   210
         TabIndex        =   9
         Top             =   1500
         Width           =   915
      End
      Begin MSForms.Label Label1 
         Height          =   375
         Left            =   1470
         TabIndex        =   7
         Top             =   660
         Width           =   1185
         Caption         =   "Label1"
         Size            =   "2090;661"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "授權密碼："
         Height          =   180
         Left            =   -74820
         TabIndex        =   5
         Top             =   1065
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmChgUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/2 改成Form2.0 (Label1)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Private Sub Command1_Click()
'Dim oldUserNum As String 'Add By Sindy 2021/6/11
'
''Add By Sindy 2021/6/11 給使用者的測試系統,鎖只能輸入自己部門人員
'oldUserNum = strUserNum
'If InStr(UCase(App.EXEName), "TEST") > 0 Then
'   If Left(PUB_GetST03(oldUserNum), 2) <> Left(PUB_GetST03(Me.Text1.Text), 2) Then
'      Me.Text1.Text = oldUserNum
'      Me.Label1.Caption = GetPrjSalesNM(Text1.Text)
'      MsgBox "只能輸入自己部門人員！"
'      'Exit Sub
'   End If
'End If
''2021/6/11 END

'Added by Morgan 2022/12/29
If txtDate <> "" Then
   If Not CheckIsTaiwanDate(txtDate) Then
      txtDate.SetFocus
      txtDate_GotFocus
      Exit Sub
   End If
End If
'end 2022/12/29
   
strUserNum = Me.Text1.Text
strUserName = GetPrjSalesNM(Text1.Text)

   strUserNum = Me.Text1.Text
   strUser1Num = strUserNum 'Added by Morgan 2018/9/21
'   pub_strUserOffice = PUB_GetST06(strUserNum)
'   Pub_StrUserSt03 = PUB_GetST03(strUserNum)
'   Pub_StrUserSt15 = PUB_GetStaffST15(strUserNum, 1)
   PUB_SetStaffVar 'Modify By Sindy 2014/9/15
   'Modified by Morgan 2017/5/22
   'PUB_SetUserData_1 'Added by Morgan 2015/6/10
   PUB_SetUserData
   'end 2017/5/22
   GetGroupDept 'Modified by Morgan 2015/9/11 要放後面否則strGroup會用到舊的
   'Added by Morgan 2019/5/31
   Systemkind_g = GetSystemKindByNick
   Systemkind_g_P = GetSystemKindByNickP
   Systemkind_g_T = GetSystemKindByNickT
   Systemkind_g_TnoS = GetSystemKindByNickTnoS
   'end 2019/5/31

'Added by Morgan 2022/12/29
If txtDate <> "" Then
   strSrvDate(1) = DBDATE(txtDate)
   strSrvDate(2) = strSrvDate(1) - 19110000
   g_LetterDate = "" 'Added by Morgan 2025/8/25
End If
If Check1.Value = 1 Then
   Forms(0).mnu00(0).Visible = False
   Forms(0).mnuChUser.Visible = False
End If
'end 2022/12/29
 
Unload Me
End Sub

Private Sub Command2_Click()
   If Text2 = Pub_GetSpecMan("解除畫面擷取限制密碼") Then
      Forms(0).Timer3.Enabled = False
      MsgBox "限制已解除！", vbOKOnly + vbInformation
      
   ElseIf Text2 = Pub_GetSpecMan("解除畫面擷取限制密碼2") Then
      If UpdateCode = True Then
         Forms(0).Timer3.Enabled = False
         MsgBox "限制已解除！", vbOKOnly + vbInformation
      Else
         Exit Sub
      End If
   Else
      MsgBox "密碼輸入錯誤！", vbCritical + vbOKOnly
      Exit Sub
   End If
   Unload Me
End Sub

Private Function UpdateCode() As Boolean
   Dim stSQL As String, stPwd As Tristate
   
On Error GoTo ErrHnd

   Randomize 'Added by Morgan 2016/12/1
   stPwd = Int(99999 * Rnd)
   stSQL = "update setSpecMan set oMan='" & stPwd & "' where  ocode='解除畫面擷取限制密碼2'"
   cnnConnection.Execute stSQL
   UpdateCode = True
   Exit Function
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub Form_Activate()
   If SSTab1.Tab = 0 Then
      'Modify By Sindy 2023/1/9
      If Text1.Enabled = True Then
      '2023/1/9 END
         Text1.SetFocus
      End If
      Command1.Default = True
   ElseIf SSTab1.Tab = 1 Then
      Text2.SetFocus
      Command2.Default = True
   End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Text1.Text = strUserNum
Label1 = strUserName
txtDate = ""
SSTab1.TabVisible(1) = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmChgUser = Nothing
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Len(Trim(Text1.Text)) = 0 Then Exit Sub
Label1.Caption = GetPrjSalesNM(Text1.Text)
If Label1.Caption = "" Then
   MsgBox "錯誤編號！", vbCritical
   Text1.SetFocus
   Text1_GotFocus
   Cancel = True
End If
End Sub

Private Sub txtDate_GotFocus()
   TextInverse txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") Then
      KeyAscii = 0
      Beep
   End If
End Sub
