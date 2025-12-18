VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010002 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   5550
   ClientLeft      =   735
   ClientTop       =   1050
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8535
   Begin VB.TextBox textCUID 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      Height          =   288
      Left            =   100
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5200
      Width           =   8300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3420
      TabIndex        =   23
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7170
      TabIndex        =   26
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆(&N)"
      Height          =   400
      Index           =   2
      Left            =   4680
      TabIndex        =   24
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   3
      Left            =   5925
      TabIndex        =   25
      Top             =   70
      Width           =   1200
   End
   Begin VB.Frame fraWindow 
      BorderStyle     =   0  '沒有框線
      Height          =   4512
      Left            =   120
      TabIndex        =   30
      Top             =   600
      Width           =   8175
      Begin VB.TextBox txtDispDate 
         Height          =   288
         Left            =   6600
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1140
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "尋找"
         Height          =   288
         Left            =   4320
         TabIndex        =   2
         Top             =   780
         Width           =   972
      End
      Begin VB.TextBox txtSystem 
         Height          =   288
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   15
         Top             =   3060
         Width           =   492
      End
      Begin VB.Frame fraDate 
         BorderStyle     =   0  '沒有框線
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Top             =   1530
         Width           =   6132
         Begin VB.OptionButton optDateKind 
            Height          =   252
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   252
         End
         Begin VB.OptionButton optDateKind 
            Height          =   252
            Index           =   1
            Left            =   1230
            TabIndex        =   7
            Top             =   0
            Width           =   252
         End
         Begin VB.OptionButton optDateKind 
            Height          =   252
            Index           =   2
            Left            =   2610
            TabIndex        =   9
            Top             =   0
            Width           =   252
         End
         Begin MSForms.TextBox txtCKind 
            Height          =   300
            Index           =   4
            Left            =   240
            TabIndex        =   6
            Top             =   0
            Width           =   555
            VariousPropertyBits=   679493659
            MaxLength       =   3
            Size            =   "979;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCKind 
            Height          =   300
            Index           =   5
            Left            =   1470
            TabIndex        =   8
            Top             =   0
            Width           =   555
            VariousPropertyBits=   679493659
            MaxLength       =   2
            Size            =   "979;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCKind 
            Height          =   300
            Index           =   6
            Left            =   2850
            TabIndex        =   10
            Top             =   0
            Width           =   1092
            VariousPropertyBits=   679493659
            MaxLength       =   7
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "天                        個月                                  日"
            Height          =   180
            Left            =   840
            TabIndex        =   28
            Top             =   30
            Width           =   3330
         End
      End
      Begin VB.Frame fraElse 
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1680
         TabIndex        =   41
         Top             =   3060
         Width           =   1812
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   2
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   18
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   1
            Left            =   960
            MaxLength       =   1
            TabIndex        =   17
            Top             =   0
            Width           =   252
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   0
            Left            =   0
            MaxLength       =   6
            TabIndex        =   16
            Top             =   0
            Width           =   852
         End
      End
      Begin VB.Frame fraTF 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1680
         TabIndex        =   27
         Top             =   3060
         Width           =   1812
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   3
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   22
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   2
            Left            =   1080
            MaxLength       =   1
            TabIndex        =   21
            Top             =   0
            Width           =   252
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   1
            Left            =   720
            MaxLength       =   1
            TabIndex        =   20
            Top             =   0
            Width           =   252
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   0
            Left            =   0
            MaxLength       =   5
            TabIndex        =   19
            Top             =   0
            Width           =   612
         End
      End
      Begin MSForms.TextBox txtCKind 
         Height          =   300
         Index           =   0
         Left            =   5070
         TabIndex        =   54
         Top             =   180
         Width           =   1095
         VariousPropertyBits=   679493659
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCKind 
         Height          =   300
         Index           =   7
         Left            =   1200
         TabIndex        =   53
         Top             =   1860
         Width           =   6135
         VariousPropertyBits=   679493659
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboCaseName 
         Height          =   300
         Left            =   1080
         TabIndex        =   52
         Top             =   3480
         Width           =   7030
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "12400;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboTrademark 
         Height          =   300
         Left            =   1200
         TabIndex        =   12
         Top             =   1860
         Width           =   6132
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10816;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboPatent 
         Height          =   300
         Left            =   1200
         TabIndex        =   11
         Top             =   1860
         Width           =   6132
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10816;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCKind 
         Height          =   300
         Index           =   9
         Left            =   1080
         TabIndex        =   14
         Top             =   2580
         Width           =   1092
         VariousPropertyBits=   679493659
         MaxLength       =   8
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCKind 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   0
         Top             =   420
         Width           =   372
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCKind 
         Height          =   300
         Index           =   3
         Left            =   1320
         TabIndex        =   3
         Top             =   1140
         Width           =   372
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCKind 
         Height          =   300
         Index           =   2
         Left            =   1080
         TabIndex        =   1
         Top             =   780
         Width           =   3132
         VariousPropertyBits=   679493659
         MaxLength       =   32
         Size            =   "5524;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCKind 
         Height          =   300
         Index           =   8
         Left            =   1080
         TabIndex        =   13
         Top             =   2220
         Width           =   372
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblDispDate 
         Caption         =   "發文日："
         Height          =   255
         Left            =   5865
         TabIndex        =   50
         Top             =   1200
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "案件名稱："
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   3480
         Width           =   972
      End
      Begin VB.Label lblLawDate 
         Height          =   252
         Left            =   4560
         TabIndex        =   42
         Top             =   4140
         Width           =   2292
      End
      Begin VB.Label lblOurDate 
         Height          =   252
         Left            =   1080
         TabIndex        =   43
         Top             =   4140
         Width           =   2172
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   8000
         Y1              =   2964
         Y2              =   2964
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8000
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Label Label10 
         Caption         =   "本所案號："
         Height          =   252
         Left            =   120
         TabIndex        =   48
         Top             =   3120
         Width           =   972
      End
      Begin VB.Label Label11 
         Caption         =   "申請人："
         Height          =   252
         Left            =   120
         TabIndex        =   47
         Top             =   3840
         Width           =   852
      End
      Begin VB.Label Label12 
         Caption         =   "本所期限："
         Height          =   252
         Left            =   120
         TabIndex        =   46
         Top             =   4140
         Width           =   972
      End
      Begin MSForms.Label lblPetition 
         Height          =   252
         Left            =   972
         TabIndex        =   45
         Top             =   3840
         Width           =   6372
         VariousPropertyBits=   27
         Size            =   "11239;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         Caption         =   "法定期限："
         Height          =   252
         Left            =   3600
         TabIndex        =   44
         Top             =   4140
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "前一筆新增之收件號："
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "收件日："
         Height          =   252
         Left            =   4320
         TabIndex        =   38
         Top             =   192
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "系統種類：         （1.專利  2.商標  3.法務、顧問  4.服務業務(含植物新品種保護)）"
         Height          =   252
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   6852
      End
      Begin VB.Label Label4 
         Caption         =   "來函號數："
         Height          =   252
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "期限起算日：          （１.次日   ２.當日   ３.無期限   ４.補優先權證明）"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   5595
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "政府機關：         （1.智慧局  2.內政部  3.經濟部  4.行政院  5.行政法院  6.地方法院  7.其他 8.智商法院）"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   2220
         Width           =   8010
      End
      Begin VB.Label Label8 
         Caption         =   "備註（40）："
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "機關文號："
         Height          =   252
         Left            =   120
         TabIndex        =   32
         Top             =   2580
         Width           =   972
      End
      Begin VB.Label lblRecieveCode 
         Height          =   255
         Left            =   2040
         TabIndex        =   31
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "期限："
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   972
      End
   End
End
Attribute VB_Name = "frm010002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/16 Form2.0已修改 txtCKind()/cboCaseName/cboPatent/cboTradeMark
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Modified by Morgan 2021/8/12 智財法院-->智商法院
Option Explicit

'bolLeave判斷離開時，是否要彈出詢問視窗
'LastData上一次存檔時，所輸入之收文日
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim bolLeave As Boolean, LastDate As String, intLeaveKind As Integer
Dim bolIsRun As Boolean
'strReceiveCode上一畫面frm010011勾選的收文號
'intTotalReceive上一畫面frm010011勾選的收文號總數
'intNowReceive現在Query的收文號Index
Dim strReceiveCode() As String, intTotalReceive As Integer, intNowReceive As Integer
'bolIsQuery是否從frm010011啟動
Public bolIsQuery As Boolean
'Modify by Morgan 2006/7/20 --妙
Public mr18 As String, mr19 As String, mr20 As String, mr21 As String, mr22 As String, mr23 As String  '2014/5/2 add by sonia
Public bolCmdSearck As Boolean                                                                         '2014/7/21 add by sonia
Dim m_ApplDate As String 'Add By Sindy 2015/7/28 申請日
Dim intCKind As Integer 'Added by Lydia 2021/11/19 本所案號的系統種類

Private Sub ReadMemo()
   '專利
   cboPatent.AddItem "公開"
   cboPatent.AddItem "言詞辯論"
   cboPatent.AddItem "延期"
   cboPatent.AddItem "延緩"
   cboPatent.AddItem "面詢"
   cboPatent.AddItem "核准"
   cboPatent.AddItem "核駁"
   cboPatent.AddItem "消滅"
   cboPatent.AddItem "被異議不成立"
   cboPatent.AddItem "被舉發不成立"
   cboPatent.AddItem "提出申復"
   cboPatent.AddItem "提出修正"
   cboPatent.AddItem "提出答辯"
   cboPatent.AddItem "提出意見"
   cboPatent.AddItem "提訴願"
   cboPatent.AddItem "進行審查"
   cboPatent.AddItem "準備程序"
   cboPatent.AddItem "補正"
   cboPatent.AddItem "補件"
   cboPatent.AddItem "審查"
   cboPatent.AddItem "應予撤銷"
   cboPatent.AddItem "舉發不成立"
   cboPatent.AddItem "繳證書費"
   cboPatent.AddItem "證書"
   cboPatent.AddItem "讓與"
   cboPatent.AddItem "變更"
   
   '商標
   cboTrademark.AddItem "自撤"
   cboTrademark.AddItem "更正"
   cboTrademark.AddItem "言詞辯論"
   cboTrademark.AddItem "延展"
   cboTrademark.AddItem "延期"
   cboTrademark.AddItem "延緩"
   cboTrademark.AddItem "准予審定"
   cboTrademark.AddItem "核准"
   cboTrademark.AddItem "核駁"
   cboTrademark.AddItem "授權"
   cboTrademark.AddItem "移轉"
   cboTrademark.AddItem "被異議不成立"
   cboTrademark.AddItem "被評定不成立"
   cboTrademark.AddItem "被廢止不成立"
   cboTrademark.AddItem "提出申復"
   cboTrademark.AddItem "提出意見"
   cboTrademark.AddItem "提出意見"
   cboTrademark.AddItem "答辯"
   cboTrademark.AddItem "註冊證"
   cboTrademark.AddItem "準備程序"
   cboTrademark.AddItem "補正"
   cboTrademark.AddItem "應予撤銷"
   cboTrademark.AddItem "應予廢止"
   cboTrademark.AddItem "繳第二期註冊費"
   cboTrademark.AddItem "變更"
End Sub

Private Sub ClearFormToRekey()
Dim i As Integer

   'modify by sonia 91.1.18
   'For i = 2 To 9
   '       txtCKind(i) = ""
   'Next
   'txtCKind(3) = "1"
   txtCKind(2) = ""
   'Modify By Cheng 2002/03/25
   '若非新增模式
   If frm010001_1.intModifyKind <> 0 Then
      For i = 4 To 9
         txtCKind(i) = ""
      Next
   '若為新增模式
   Else
      For i = 9 To 9
         txtCKind(i) = ""
      Next
   End If
   'Modify By Cheng 2002/03/25
   '若為連續新增時, 備註, 期限, 主管機關請保留上一筆輸入的資料
   'txtCKind(8) = "1"
   ClearCode
   CheckChoose
   'txtCKind(7) = ""
   'cboPatent.Text = ""
   'cboTrademark.Text = ""
   'optDateKind(1).Value = True
   'optDateKind(0).Value = True
   
   '2014/5/2 ADD BY SONIA 專利處人員操作隱藏尋找按鈕,在來函號數欄跳離時自動做按尋找動作
   'txtCKind(1).SetFocus
   If Left(Pub_StrUserSt03, 2) = "P1" Then
      CmdSearch.Visible = False
      txtCKind(1) = "1"
      txtCKind(1).Enabled = False
      txtCKind(2).SetFocus
   Else
      CmdSearch.Visible = True
      txtCKind(1).Enabled = True
      txtCKind(1).SetFocus
   End If
   '2014/5/2 END
End Sub

'edit by nickc 2007/08/06 切換輸入法改用API
Private Sub cboPatent_GotFocus()
   OpenIme
End Sub

Private Sub cboPatent_Validate(Cancel As Boolean)
   'Modify By Cheng 2002/03/25
   If Len(Me.cboPatent.Text) <= 0 Then
      MsgBox "請輸入備註資料!!!", vbExclamation
      Cancel = True
   ElseIf CheckLengthIsOK(cboPatent.Text, 40) = False Then
      Cancel = True
      cboPatent.SelStart = 0
      cboPatent.SelLength = Len(cboPatent.Text)
   End If
   'edit by nickc 2007/08/06 切換輸入法改用API
   'If Cancel = False Then CloseIme 'Removed by Morgan 2016/10/20 會造成 Win7 的切換錯誤
End Sub

Private Sub cboTrademark_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'cboTrademark.IMEMode = 1
   OpenIme
End Sub

Private Sub cboTrademark_Validate(Cancel As Boolean)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'cboTrademark.IMEMode = 2
   'Modify By Cheng 2002/03/25
   If Len(Me.cboTrademark.Text) <= 0 Then
      MsgBox "請輸入備註資料!!!", vbExclamation
      Cancel = True
   ElseIf CheckLengthIsOK(cboTrademark.Text, 40) = False Then
      Cancel = True
      cboTrademark.SelStart = 0
      cboTrademark.SelLength = Len(cboTrademark.Text)
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then CloseIme 'Removed by Morgan 2016/10/20 會造成 Win7 的切換錯誤
End Sub

'Add by Morgan 2004/1/29
'檢查案號是否閉卷
Private Function CheckIsClose() As Boolean
Dim intCaseKind As Integer, strSql As String, strCode(1 To 4) As String
Dim rstQuery As New ADODB.Recordset
On Error GoTo ErrHand
    
   strCode(1) = txtSystem.Text
   If txtSystem.Text <> 馬德里案 Then
      strCode(2) = txtCode(0)
      strCode(3) = IIf(txtCode(1) = "", "0", txtCode(1))
      strCode(4) = IIf(txtCode(2) = "", "00", txtCode(2))
   Else
      strCode(2) = txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1))
      strCode(3) = IIf(txtTFCode(2) = "", "0", txtTFCode(2))
      strCode(4) = IIf(txtTFCode(3) = "", "00", txtTFCode(3))
   End If

   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetSystemKind(txtSystem, intCaseKind) Then
   If ClsPDGetSystemKind(txtSystem, intCaseKind) Then
      Select Case intCaseKind
         Case 專利
            strSql = "select 1 from patent where pa01='" & strCode(1) & "' and pa02='" & strCode(2) & "' and pa03='" & strCode(3) & "' and pa04='" & strCode(4) & "' and pa57='Y'"
         Case 商標
            strSql = "select 1 from trademark where tm01='" & strCode(1) & "' and tm02='" & strCode(2) & "' and tm03='" & strCode(3) & "' and tm04='" & strCode(4) & "' and tm29='Y'"
         Case 法務
            strSql = "select 1 from lawcase where lc01='" & strCode(1) & "' and lc02='" & strCode(2) & "' and lc03='" & strCode(3) & "' and lc04='" & strCode(4) & "' and lc08='Y'"
         Case 顧問
            strSql = "select 1 from hirecase where hc01='" & strCode(1) & "' and hc02='" & strCode(2) & "' and hc03='" & strCode(3) & "' and hc04='" & strCode(4) & "' and hc09='Y'"
         Case Else
            strSql = "select 1 from servicepractice where sp01='" & strCode(1) & "' and sp02='" & strCode(2) & "' and sp03='" & strCode(3) & "' and sp04='" & strCode(4) & "' and sp15='Y'"
      End Select
      rstQuery.CursorLocation = adUseClient
      rstQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rstQuery.RecordCount > 0 Then
         CheckIsClose = True
      End If
      rstQuery.Close
      Set rstQuery = Nothing
   End If
   
ErrHand:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, strAuto As String

   Select Case Index
      Case 0 '確定
         If txtSystem <> "" Then
            CheckCaseCode
         Else
            ShowMsg "找不到此本所案號在基本檔之資料"
            Exit Sub
         End If
         If lblOurDate = "" And txtCKind(3) <> "3" Then
            ShowMsg MsgText(1020)
            For i = 0 To 2
               If optDateKind(i).Value Then
                  txtCKind(4 + i).SetFocus
                  Exit For
               End If
            Next
            Exit Sub
         End If
         If cboCaseName.ListCount = 0 Then
            ShowMsg MsgText(1021)
            Exit Sub
         End If
         For i = 0 To 3
            If CheckKeyIn(i) = False Then
               Exit Sub
            End If
         Next
         For i = 7 To 9
            If CheckKeyIn(i) = False Then
               Exit Sub
            End If
         Next
         If i = 10 Then
            If cboPatent.Visible Then
               txtCKind(7) = cboPatent.Text
            ElseIf cboTrademark.Visible Then
               txtCKind(7) = cboTrademark.Text
            End If
            If txtCKind(7) = "" Then
               MsgBox "備註不可空白 !", vbCritical
               Exit Sub
            End If
            'Modify By Sindy 2010/8/17 比對自動編號年度
            'strAuto = "D" + GetTaiwanThisYear
            strAuto = "D" + CompAutoNumberYear(GetTaiwanThisYear)
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
             
            'Add by Morgan 2004/1/29
            '檢查是否已閉卷
            If CheckIsClose = True Then
               'Add By Sindy 2010/11/26
               If (Trim(txtSystem) = "P" Or Trim(txtSystem) = "FCP") And InStr(1, txtCKind(7), "消滅") > 0 Then
                  '不顯示已閉卷
               '2010/11/26 End
               Else
                  If MsgBox("注意！！本案件已閉卷，請先與專業部聯絡，確定是否為此案件之來函！！！", vbOKCancel) <> vbOK Then
                      Exit Sub
                  End If
               End If
            End If
            'Add End ------
            
            'Add by Sindy 2013/11/7 詢問是否要重覆收文
            strSql = "select mr01 from mailRec" & _
                     " where mr02=" & DBDATE(IIf(txtCKind(0) = "", strSrvDate(1), txtCKind(0))) & _
                     " and mr12='" & Trim(txtSystem) & "'" & _
                     " and mr13='" & Trim(txtCode(0)) & "'" & _
                     " and mr14='" & IIf(Trim(txtCode(1)) = "", "0", Trim(txtCode(1))) & "'" & _
                     " and mr15='" & IIf(Trim(txtCode(2)) = "", "00", Trim(txtCode(2))) & "'"
            If frm010001_1.intModifyKind <> 0 Then
               strSql = strSql & " and mr01<>'" & lblRecieveCode & "'"
            End If
            CheckOC3
            AdoRecordSet3.CursorLocation = adUseClient
            AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If AdoRecordSet3.RecordCount > 0 Then
               If MsgBox("此案號同日已有來函資料, 收件號為" & AdoRecordSet3.Fields(0) & "！是否確認要存檔？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Exit Sub
               End If
            End If
            '2013/11/7 END
            
            '2015/1/13 ADD BY SONIA FCT申請中分割來函,若未輸期限才提醒FCT-034627
            If txtSystem = "FCT" And InStr(txtCKind(7), "分割") > 0 And lblOurDate = "" Then
               strSql = "select * from trademark where tm15 is not null and tm01='" & txtSystem & "' and tm02='" & txtCKind(0) & "' and tm03='" & IIf(Trim(txtCode(1)) = "", "0", Trim(txtCode(1))) & "' and tm04='" & IIf(Trim(txtCode(2)) = "", "00", Trim(txtCode(2))) & "'"
               CheckOC3
               AdoRecordSet3.CursorLocation = adUseClient
               AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If AdoRecordSet3.RecordCount > 0 Then   '有tm15為發證後分割則不詢問
               Else
                  If MsgBox("FCT 分割案！請確認是否有期限？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
                     Exit Sub
                  End If
               End If
            End If
            '2015/1/13 END
            
            If SaveDatabase(strAuto) Then
               PUB_SendMailCache 'Added by Morgan 2021/6/18
               
               bolLeave = True
               If frm010001_1.intModifyKind = 0 Then
                  'edit by nickc 2007/05/16 每次都預設為系統日，因為櫃檯都會忘記改回來，秀玲說的
                  'LastDate = txtCKind(0).Text
                  LastDate = ""
                  txtCKind(0).Text = GetTaiwanTodayDate
                  
                  lblRecieveCode = strAuto
                 'ShowMsg MsgText(1023) + strAuto
               End If
               If frm010001_1.intModifyKind = 1 Then
                  Unload Me
               Else
                  ClearFormToRekey
               End If
            End If
         End If
      Case 1 '結束
         intLeaveKind = 0
         If bolIsQuery Then Unload frm010011
         Unload Me
      Case 2 '下一筆
         intNowReceive = intNowReceive + 1
         If intNowReceive = intTotalReceive - 1 Then
            cmdOK(2).Visible = False
            cmdOK(3).Default = True
         End If
         lblRecieveCode = strReceiveCode(intNowReceive)
         ReadCkindDatabaseR
         frm010011.ClearOneRow
      Case 3 '回前畫面
         If bolIsQuery Then
            frm010011.Show
            frm010011.ClearOneRow
         End If
         Unload Me
   End Select
End Sub

Private Function SaveDatabase(ByRef strAuto As String) As Boolean
   If frm010001_1.intModifyKind = 0 Then
      If txtSystem.Text = 馬德里案 Then
   'edit by nickc 2005/09/16 已經搬到 basquery
   '      SaveDatabase = obj001.InsertCKindDatabase(strAuto, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                  IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), ChangeTDateStringToTString(lblOurDate), ChangeTDateStringToTString(lblLawDate))
         SaveDatabase = InsertCKindDatabase(strAuto, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                  IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), ChangeTDateStringToTString(lblOurDate), ChangeTDateStringToTString(lblLawDate))
      Else
           'Modify By Cheng 2004/01/05
           'P案的本所期限不可為非工作天
   '      SaveDatabase = obj001.InsertCKindDatabase(strAuto, txtCKind(0), txtCKind(1), txtCKind(2) _
   '              , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtCode(0), _
   '               IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), ChangeTDateStringToTString(lblOurDate), ChangeTDateStringToTString(lblLawDate))
   'edit by nickc 2005/09/16 已經搬到  basquery
   '      SaveDatabase = obj001.InsertCKindDatabase(strAuto, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtCode(0), _
                  IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), IIf(Me.txtSystem.Text = "P", ChangeWStringToTString(PUB_GetWorkDay1(ChangeTDateStringToTString(lblOurDate), True)), ChangeTDateStringToTString(lblOurDate)), ChangeTDateStringToTString(lblLawDate))
         
         'Modif by Morgan 2007/6/12 加mr24
         If txtDispDate.Visible = True Then
            SaveDatabase = InsertCKindDatabase(strAuto, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtCode(0), _
                  IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), IIf(Me.txtSystem.Text = "P", ChangeWStringToTString(PUB_GetWorkDay1(ChangeTDateStringToTString(lblOurDate), True)), ChangeTDateStringToTString(lblOurDate)), ChangeTDateStringToTString(lblLawDate), txtDispDate)
         Else
            SaveDatabase = InsertCKindDatabase(strAuto, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtCode(0), _
                  IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), IIf(Me.txtSystem.Text = "P", ChangeWStringToTString(PUB_GetWorkDay1(ChangeTDateStringToTString(lblOurDate), True)), ChangeTDateStringToTString(lblOurDate)), ChangeTDateStringToTString(lblLawDate))
         End If
           'End
      End If
   Else
      If txtSystem.Text = 馬德里案 Then
   'edit by nickc 2005/09/16 已經搬到 basquery
   '      SaveDatabase = obj001.UpdateCKindDatabase(lblRecieveCode, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                  IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), ChangeTDateStringToTString(lblOurDate), ChangeTDateStringToTString(lblLawDate))
         SaveDatabase = UpdateCKindDatabase(lblRecieveCode, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                  IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), ChangeTDateStringToTString(lblOurDate), ChangeTDateStringToTString(lblLawDate))
      Else
   '      SaveDatabase = obj001.UpdateCKindDatabase(lblRecieveCode, txtCKind(0), txtCKind(1), txtCKind(2) _
   '              , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtCode(0), _
   '               IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), ChangeTDateStringToTString(lblOurDate), ChangeTDateStringToTString(lblLawDate))
           'Modify By Cheng 2004/01/05
           'P案的本所期限不可為非工作天
   'edit by nickc 2005/09/16 已經搬到 basquery
   '      SaveDatabase = obj001.UpdateCKindDatabase(lblRecieveCode, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtCode(0), _
                  IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), IIf(Me.txtSystem.Text = "P", ChangeWStringToTString(PUB_GetWorkDay1(ChangeTDateStringToTString(lblOurDate), True)), ChangeTDateStringToTString(lblOurDate)), ChangeTDateStringToTString(lblLawDate))
         
         'Modif by Morgan 2007/6/12 加mr24
         If txtDispDate.Visible = True Then
            SaveDatabase = UpdateCKindDatabase(lblRecieveCode, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtCode(0), _
                  IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), IIf(Me.txtSystem.Text = "P", ChangeWStringToTString(PUB_GetWorkDay1(ChangeTDateStringToTString(lblOurDate), True)), ChangeTDateStringToTString(lblOurDate)), ChangeTDateStringToTString(lblLawDate), txtDispDate)
         Else
            SaveDatabase = UpdateCKindDatabase(lblRecieveCode, txtCKind(0), txtCKind(1), txtCKind(2) _
                 , txtCKind(3), txtCKind(4), txtCKind(5), txtCKind(6), txtCKind(7), txtCKind(8), txtCKind(9), txtSystem, txtCode(0), _
                  IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), IIf(Me.txtSystem.Text = "P", ChangeWStringToTString(PUB_GetWorkDay1(ChangeTDateStringToTString(lblOurDate), True)), ChangeTDateStringToTString(lblOurDate)), ChangeTDateStringToTString(lblLawDate))
         End If
           'End
      End If
   End If
End Function

Private Sub cmdSearch_Click()
Dim mr01 As String, mr02 As String, mr03 As String, mr04 As String, mr05 As String, _
         mr06 As String, mr07 As String, mr08 As String, mr09 As String, mr10 As String, _
         mr11 As String, mr12 As String, mr13 As String, mr14 As String, mr15 As String, _
         i As Integer, rt As Boolean, strCodeName As String, strPetition As String
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String, strCustomer As String
Dim adoquery As New ADODB.Recordset

'Add By Cheng 2001/12/13
'Dim cn As New ADODB.Connection 'Removed by Morgan 2017/4/20 沒用
Dim adoquery1 As New ADODB.Recordset
Dim strSql As String

   'Add by Morgan 2007/6/12
   lblDispDate.Visible = False
   txtDispDate.Visible = False
   'end 2007/6/12
   m_ApplDate = "" 'Add By Sindy 2015/7/28 申請日
   'add by nickc 2005/10/05 若是誤按，會卡很久
   If CheckStr(txtCKind(2)) = "" Then MsgBox "來函號數不能空白！", , "錯誤！": Exit Sub
   
   Screen.MousePointer = vbHourglass
   'Modify By Cheng 2003/05/14
   'Set adoquery = obj001.ReadCKindRst(txtCKind(1), txtCKind(2))
   'edit by nickc 2007/02/06 不用 dll 了
   'Set adoquery = obj001.ReadCKindRst_1(txtCKind(1), txtCKind(2))
   Set adoquery = Cls001ReadCKindRst_1(txtCKind(1), txtCKind(2))
   
   If adoquery.RecordCount = 0 Then
      'Add By cheng 2001/12/13
      '再檢查案件進度檔的對造號數
      Select Case txtCKind(1).Text
         Case 專利
               'Modify By Cheng 2003/05/13
   '         strSQL = "select nvl(pa05,nvl(pa06,pa07)) 案件名稱,pa01 本所案號," & _
   '            "pa02 本所案號,pa02 本所案號,decode(pa03, '0', ' ', pa03) 本所案號,decode(pa04, '00', ' ', pa04) 本所案號," & _
   '            "nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,patent,customer" & _
   '            " where cp36=" & CNULL(txtCKind(2).Text) & _
   '            " And cp01 in ('P','FCP')" & _
   '            " And '000'=pa09(+)" & _
   '            " And cp01=pa01(+) And cp02=pa02(+) And cp03=pa03(+) And cp04=pa04(+)" & _
   '            " And substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02 order by pa01,pa02,pa03,pa04"
            '2005/9/29 MODIFY BY SONIA
            'strSQL = "select nvl(pa05,nvl(pa06,pa07)) 案件名稱,pa01 本所案號," & _
            '   "pa02 本所案號,pa02 本所案號,decode(pa03, '0', ' ', pa03) 本所案號,decode(pa04, '00', ' ', pa04) 本所案號," & _
            '   "nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,patent,customer" & _
            '   " where cp36 Like " & CNULL(txtCKind(2).Text & "%") & _
            '   " And cp01 in ('P','FCP')" & _
            '   " And '000'=pa09(+)" & _
            '   " And cp01=pa01(+) And cp02=pa02(+) And cp03=pa03(+) And cp04=pa04(+)" & _
            '   " And substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
            '   " Group By nvl(pa05,nvl(pa06,pa07)),pa01, pa02,decode(pa03, '0', ' ', pa03),decode(pa04, '00', ' ', pa04),nvl(cu04,nvl(cu05,cu06)) " & _
            '   " order by 2, 3, 4, 5, 6 "
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            '2013/8/16 modify by sonia cp36檢查已寫在Cls001ReadCKindRst_1故刪除,另加檢查CP35欄 P-099556之智慧局答辯函加輸訴訟案號存於CP35
            'strSql = "select nvl(pa05,nvl(pa06,pa07)) 案件名稱,pa01 本所案號,pa02 本所案號,pa02 本所案號,decode(pa03, '0', ' ', pa03) 本所案號,decode(pa04, '00', ' ', pa04) 本所案號," & _
               " nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,patent,customer" & _
               " where cp36 Like " & CNULL(txtCKind(2).Text & "%") & " And cp01 in ('P','FCP')" & _
               " And cp01=pa01(+) And cp02=pa02(+) And cp03=pa03(+) And cp04=pa04(+) And '000'=pa09(+)" & _
               " And substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
               " Group By nvl(pa05,nvl(pa06,pa07)),pa01, pa02,decode(pa03, '0', ' ', pa03),decode(pa04, '00', ' ', pa04),nvl(cu04,nvl(cu05,cu06)) " & _
               " union select nvl(sp05,nvl(sp06,sp07)) 案件名稱,sp01 本所案號,sp02 本所案號,sp02 本所案號,decode(sp03, '0', ' ', sp03) 本所案號,decode(sp04, '00', ' ', sp04) 本所案號," & _
               " nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,servicepractice,customer" & _
               " where cp36 Like " & CNULL(txtCKind(2).Text & "%") & _
               " And cp01 not in ('P','FCP','CFP','T','TF','FCT','CFT','L','LA','FCL','CFL','LIN')" & _
               " And cp01=sp01(+) And cp02=sp02(+) And cp03=sp03(+) And cp04=sp04(+) And '000'=sp09(+)" & _
               " And substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) " & _
               " Group By nvl(sp05,nvl(sp06,sp07)),sp01, sp02,decode(sp03, '0', ' ', sp03),decode(sp04, '00', ' ', sp04),nvl(cu04,nvl(cu05,cu06)) " & _
               " order by 2, 3, 4, 5, 6 "
            '2013/8/29 modify by sonia 智商法院案再抓機關文號,故加入第2句 FCP-032929
            'modify by sonia 2019/7/29 +ACS系統類別
            strSql = "select nvl(pa05,nvl(pa06,pa07)) 案件名稱,pa01 本所案號,pa02 本所案號,pa02 本所案號,decode(pa03, '0', ' ', pa03) 本所案號,decode(pa04, '00', ' ', pa04) 本所案號," & _
               " nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,patent,customer" & _
               " where cp35 = " & CNULL(txtCKind(2).Text) & "And cp01 in ('P','FCP')" & _
               " And cp01=pa01(+) And cp02=pa02(+) And cp03=pa03(+) And cp04=pa04(+) And '000'=pa09(+)" & _
               " And substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
               " Group By nvl(pa05,nvl(pa06,pa07)),pa01, pa02,decode(pa03, '0', ' ', pa03),decode(pa04, '00', ' ', pa04),nvl(cu04,nvl(cu05,cu06)) " & _
               " union select nvl(pa05,nvl(pa06,pa07)) 案件名稱,pa01 本所案號,pa02 本所案號,pa02 本所案號,decode(pa03, '0', ' ', pa03) 本所案號,decode(pa04, '00', ' ', pa04) 本所案號," & _
               " nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,patent,customer" & _
               " where cp08 = " & CNULL(txtCKind(2).Text) & "And cp01 in ('P','FCP')" & _
               " And cp01=pa01(+) And cp02=pa02(+) And cp03=pa03(+) And cp04=pa04(+) And '000'=pa09(+)" & _
               " And substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
               " Group By nvl(pa05,nvl(pa06,pa07)),pa01, pa02,decode(pa03, '0', ' ', pa03),decode(pa04, '00', ' ', pa04),nvl(cu04,nvl(cu05,cu06)) " & _
               " union select nvl(sp05,nvl(sp06,sp07)) 案件名稱,sp01 本所案號,sp02 本所案號,sp02 本所案號,decode(sp03, '0', ' ', sp03) 本所案號,decode(sp04, '00', ' ', sp04) 本所案號," & _
               " nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,servicepractice,customer" & _
               " where cp36 Like " & CNULL(txtCKind(2).Text & "%") & _
               " And cp01 not in ('P','FCP','CFP','T','TF','FCT','CFT','L','LA','FCL','CFL','LIN','ACS')" & _
               " And cp01=sp01(+) And cp02=sp02(+) And cp03=sp03(+) And cp04=sp04(+) And '000'=sp09(+)" & _
               " And substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) " & _
               " Group By nvl(sp05,nvl(sp06,sp07)),sp01, sp02,decode(sp03, '0', ' ', sp03),decode(sp04, '00', ' ', sp04),nvl(cu04,nvl(cu05,cu06)) " & _
               " order by 2, 3, 4, 5, 6 "
            '2005/9/29 END
         Case 商標
               'Modify By Cheng 2003/05/13
   '         strSQL = "select nvl(tm05,nvl(tm06,tm07)) 案件名稱,tm01 本所案號," & _
   '            "decode(tm01," + CNULL(馬德里案) + ",substr(tm02,1,5),tm02) 本所案號," & _
   '            "decode(tm01," + CNULL(馬德里案) + ",substr(tm02,6,1),tm02) 本所案號," & _
   '            "decode(tm03, '0', ' ', tm03) 本所案號,decode(tm04, '0', ' ', tm04) 本所案號,nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,trademark,customer " & _
   '            " where cp36 = " & CNULL(txtCKind(2).Text) & " " & _
   '            " And cp01 in ('T', 'FCT') " & _
   '            " And '000' = tm10(+) " & _
   '            " And cp01 = tm01(+) And cp02 = tm02(+) And cp03 = tm03(+) And cp04 = tm04(+) " & _
   '            " AND substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02" & _
   '            " order by tm01,tm02,tm03,tm04"
            '2013/8/16 modify by sonia cp36檢查已寫在Cls001ReadCKindRst_1故刪除
            'strSql = "select nvl(tm05,nvl(tm06,tm07)) 案件名稱,tm01 本所案號," & _
               "decode(tm01," + CNULL(馬德里案) + ",substr(tm02,1,5),tm02) 本所案號," & _
               "decode(tm01," + CNULL(馬德里案) + ",substr(tm02,6,1),tm02) 本所案號," & _
               "decode(tm03, '0', ' ', tm03) 本所案號,decode(tm04, '0', ' ', tm04) 本所案號,nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,trademark,customer " & _
               " where cp36 Like " & CNULL(txtCKind(2).Text & "%") & " " & _
               " And cp01 in ('T', 'FCT') " & _
               " And '000' = tm10(+) " & _
               " And cp01 = tm01(+) And cp02 = tm02(+) And cp03 = tm03(+) And cp04 = tm04(+) " & _
               " AND substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+)" & _
               " Group By nvl(tm05,nvl(tm06,tm07)), tm01, decode(tm01," + CNULL(馬德里案) + ",substr(tm02,1,5),tm02), decode(tm01," + CNULL(馬德里案) + ",substr(tm02,6,1),tm02) " & _
               " ,decode(tm03, '0', ' ', tm03), decode(tm04, '0', ' ', tm04), nvl(cu04,nvl(cu05,cu06)) " & _
               " order by 2, 3, 4, 5, 6 "
            'modify by sonia 2018/4/19 加檢查CP35欄 T-199865之智慧局答辯函加輸訴訟案號存於CP35
            'ShowMsg MsgText(9211)
            'Screen.MousePointer = vbDefault
            'Exit Sub
            strSql = "select nvl(tm05,nvl(tm06,tm07)) 案件名稱,tm01 本所案號,tm02 本所案號,tm02 本所案號,decode(tm03, '0', ' ', tm03) 本所案號,decode(tm04, '00', ' ', tm04) 本所案號," & _
               " nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,trademark,customer" & _
               " where cp35 = " & CNULL(txtCKind(2).Text) & " And cp01 in ('T','FCT')" & _
               " And cp01=tm01(+) And cp02=tm02(+) And cp03=tm03(+) And cp04=tm04(+) And '000'=tm10(+)" & _
               " And substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
               " Group By nvl(tm05,nvl(tm06,tm07)),tm01,tm02,decode(tm03, '0', ' ', tm03),decode(tm04, '00', ' ', tm04),nvl(cu04,nvl(cu05,cu06)) "
            'add by sonia 2021/9/15 +CFT之cp30(申請英文證明存台灣案審定號)CFT-022553(01692168)
            strSql = strSql & "union select nvl(tm05,nvl(tm06,tm07)) 案件名稱,tm01 本所案號,tm02 本所案號,tm02 本所案號,decode(tm03, '0', ' ', tm03) 本所案號,decode(tm04, '00', ' ', tm04) 本所案號," & _
               " nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,trademark,customer" & _
               " where cp30 = " & CNULL(txtCKind(2).Text) & " And cp01='CFT'" & _
               " And cp01=tm01(+) And cp02=tm02(+) And cp03=tm03(+) And cp04=tm04(+) " & _
               " And substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
               " Group By nvl(tm05,nvl(tm06,tm07)),tm01,tm02,decode(tm03, '0', ' ', tm03),decode(tm04, '00', ' ', tm04),nvl(cu04,nvl(cu05,cu06)) "
            'end 2021/9/15
            strSql = strSql & " order by 2, 3, 4, 5, 6 "
            '2018/4/19 end
            '2013/8/16 end
         '2005/9/29 ADD BY SONIA
         'Modify By Sindy 2009/07/24 增加LIN系統類別
         'modify by sonia 2019/7/29 +ACS系統類別
         Case "4"
            strSql = "select nvl(sp05,nvl(sp06,sp07)) 案件名稱,sp01 本所案號," & _
               "sp02 本所案號,sp02 本所案號,decode(sp03, '0', ' ', sp03) 本所案號,decode(sp04, '00', ' ', sp04) 本所案號," & _
               "nvl(cu04,nvl(cu05,cu06)) 申請人 from CaseProgress,servicepractice,customer" & _
               " where cp36 Like " & CNULL(txtCKind(2).Text & "%") & _
               " And cp01 not in ('P','FCP','CFP','T','TF','FCT','CFT','L','LA','FCL','CFL','LIN','ACS')" & _
               " And '000'=sp09(+)" & _
               " And cp01=sp01(+) And cp02=sp02(+) And cp03=sp03(+) And cp04=sp04(+)" & _
               " And substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) " & _
               " Group By nvl(sp05,nvl(sp06,sp07)),sp01, sp02,decode(sp03, '0', ' ', sp03),decode(sp04, '00', ' ', sp04),nvl(cu04,nvl(cu05,cu06)) " & _
               " order by 2, 3, 4, 5, 6 "
         '2005/9/29 END
         Case Else
            ShowMsg MsgText(9211)
            Screen.MousePointer = vbDefault
            Exit Sub
      End Select
      Set adoquery = Nothing
      If adoquery1.State <> adStateClosed Then
         adoquery1.Close
      End If
      
      'Removed by Morgan 2017/4/20 沒用
      'Set cn = Nothing
      'cn.Open "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=M51CON"
      'end 2017/4/20
      
      adoquery1.CursorLocation = adUseClient
      adoquery1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If adoquery1.RecordCount = 1 Then
         txtSystem = "" & adoquery1.Fields(1).Value
         If txtSystem <> 馬德里案 Then
            txtCode(0).Text = "" & adoquery1.Fields(2).Value
            txtCode(1).Text = IIf("" & adoquery1.Fields(4).Value = "0" Or "" & adoquery1.Fields(4).Value = " ", "", "" & adoquery1.Fields(4).Value)
            txtCode(2).Text = IIf("" & adoquery1.Fields(5).Value = "00" Or "" & adoquery1.Fields(5).Value = " ", "", "" & adoquery1.Fields(5).Value)
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
                IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer) = False Then GoTo Err
            'Add By Sindy 2010/5/6
            If Trim(txtSystem) = "" And Trim(txtCode(0)) = "" Then
               ShowMsg MsgText(1030) '找尋不到資料
               Screen.MousePointer = vbDefault
               Exit Sub
            '2010/5/6 End
            Else
               'Modify By Sindy 2015/8/18 + m_ApplDate
               If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
                   IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , , , , m_ApplDate) = False Then GoTo ErrHand
            End If
         Else
            txtTFCode(0).Text = Left("" & adoquery1.Fields(2).Value, 5)
            txtTFCode(1).Text = IIf(Right("" & adoquery1.Fields(2).Value, 1) = "0" Or Right("" & adoquery1.Fields(2).Value, 1) = " ", "", Right("" & adoquery1.Fields(2).Value, 1))
            txtTFCode(2).Text = IIf("" & adoquery1.Fields(4).Value = "0" Or "" & adoquery1.Fields(4).Value = " ", "", "" & adoquery1.Fields(4).Value)
            txtTFCode(3).Text = IIf("" & adoquery1.Fields(5).Value = "00" Or "" & adoquery1.Fields(5).Value = " ", "", "" & adoquery1.Fields(5).Value)
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer) = False Then GoTo Err
            'Add By Sindy 2010/5/6
            If Trim(txtSystem) = "" And Trim(txtTFCode(0)) = "" Then
               ShowMsg MsgText(1030) '找尋不到資料
               Screen.MousePointer = vbDefault
               Exit Sub
            '2010/5/6 End
            Else
               'Modify By Sindy 2015/8/18 + m_ApplDate
               If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                   IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , , , , m_ApplDate) = False Then GoTo ErrHand
            End If
         End If
         SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
         lblPetition = strCustomer
         If adoquery1.State <> adStateClosed Then
            adoquery1.Close
         End If
         'Set cn = Nothing 'Removed by Morgan 2017/4/20 沒用
           Screen.MousePointer = vbDefault
         Exit Sub
      ElseIf adoquery1.RecordCount = 0 Then
         ShowMsg MsgText(9211)
         If adoquery1.State <> adStateClosed Then
            adoquery1.Close
         End If
         'Set cn = Nothing 'Removed by Morgan 2017/4/20 沒用
           Screen.MousePointer = vbDefault
         Exit Sub
      End If
   Else
      If adoquery.RecordCount = 1 Then
            txtSystem = "" & adoquery.Fields(1).Value
            If txtSystem <> 馬德里案 Then
               txtCode(0).Text = "" & adoquery.Fields(2).Value
               txtCode(1).Text = IIf("" & adoquery.Fields(4).Value = "0" Or "" & adoquery.Fields(4).Value = " ", "", "" & adoquery.Fields(4).Value)
               txtCode(2).Text = IIf("" & adoquery.Fields(5).Value = "00" Or "" & adoquery.Fields(5).Value = " ", "", "" & adoquery.Fields(5).Value)
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
                   IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer) = False Then GoTo Err
               'Add By Sindy 2010/5/6
               If Trim(txtSystem) = "" And Trim(txtCode(0)) = "" Then
                  ShowMsg MsgText(1030) '找尋不到資料
                  Screen.MousePointer = vbDefault
                  Exit Sub
               '2010/5/6 End
               Else
                  'Modify By Sindy 2015/8/18 + m_ApplDate
                  If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
                      IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , , , , m_ApplDate) = False Then GoTo ErrHand
               End If
            Else
               txtTFCode(0).Text = Left("" & adoquery.Fields(2).Value, 5)
               txtTFCode(1).Text = IIf(Right("" & adoquery.Fields(2).Value, 1) = "0" Or Right("" & adoquery.Fields(2).Value, 1) = " ", "", Right("" & adoquery.Fields(2).Value, 1))
               txtTFCode(2).Text = IIf("" & adoquery.Fields(4).Value = "0" Or "" & adoquery.Fields(4).Value = " ", "", "" & adoquery.Fields(4).Value)
               txtTFCode(3).Text = IIf("" & adoquery.Fields(5).Value = "00" Or "" & adoquery.Fields(5).Value = " ", "", "" & adoquery.Fields(5).Value)
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                   IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer) = False Then GoTo Err
               'Add By Sindy 2010/5/6
               If Trim(txtSystem) = "" And Trim(txtTFCode(0)) = "" Then
                  ShowMsg MsgText(1030) '找尋不到資料
                  Screen.MousePointer = vbDefault
                  Exit Sub
               '2010/5/6 End
               Else
                  'Modify By Sindy 2015/8/18 + m_ApplDate
                  If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
                      IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , , , , m_ApplDate) = False Then GoTo ErrHand
               End If
            End If
            SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
            lblPetition = strCustomer
            'Add by Morgan 2007/6/12
            If txtSystem = "P" Then
               strExc(0) = "select cp110 from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and cp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' and cp09<'C' and cp27>0 and cp110 is not null order by cp27 desc,cp09 desc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp(0) = "65002" Then
                     lblDispDate.Visible = True
                     txtDispDate.Visible = True
                     txtDispDate = ""
                     txtDispDate.SetFocus
                  End If
               End If
            End If
            'end 2007/6/12
            Screen.MousePointer = vbDefault
            '2014/5/2 add by sonia 專利處人員操作直接做按尋找動作且只能是P案件
            If Left(Pub_StrUserSt03, 2) = "P1" Then
               If txtSystem <> "P" Then
                  ShowMsg "此案本所案號為 " & txtSystem & "-" & txtCode(0) & "-" & IIf(txtCode(1) = "", "0", txtCode(1)) & "-" & IIf(txtCode(2) = "", "00", txtCode(2)) & ", 非專利處案件 !"
                  Exit Sub
               End If
            End If
            '2014/5/2 END
   
           CheckKeyIn (3)
         Exit Sub
      End If
   End If
   
   '2014/5/2 add by sonia 專利處人員操作直接做按尋找動作且只能是P案件
   If Left(Pub_StrUserSt03, 2) = "P1" Then
      'modify by sonia 2014/7/21 P-108163(103302072)
      'If txtSystem <> "P" Then
      If txtSystem <> "" And txtSystem <> "P" Then
         ShowMsg "此案本所案號為 " & txtSystem & "-" & txtCode(0) & "-" & IIf(txtCode(1) = "", "0", txtCode(1)) & "-" & IIf(txtCode(2) = "", "00", txtCode(2)) & ", 非專利處案件 !"
         Exit Sub
      End If
   End If
   '2014/5/2 END
   
   bolCmdSearck = False   'add by sonia 2014/7/21 否則下面CheckKeyIn(2)會跑迴圈
   Screen.MousePointer = vbDefault
   If CheckKeyIn(1) And CheckKeyIn(2) Then
      frm010003.Show vbModal
   Else
      txtCKind(1).SetFocus
   End If
   Exit Sub
ErrHand:
   ShowMsg MsgText(1029)
   bolLeave = True
    Screen.MousePointer = vbDefault
   Unload Me
End Sub

Private Sub Form_Activate()
   'add by nickc 2006/07/11 解決收來函未關掉又突然收接洽單再回來物件不存在的問題
   'edit by nickc 2007/02/06 不用 dll 了
   'If obj001 Is Nothing Then
   '   Set obj001 = CreateObject("prjTaieDll001.cls001")
   '   Set obj001.Connection = cnnConnection
   'End If
   If bolIsRun Then Exit Sub
   bolIsRun = True
   If bolIsQuery Then
      lblRecieveCode = strReceiveCode(0)
      '查詢：只可查詢
      ReadCkindDatabaseR
      fraWindow.Enabled = False
      cmdOK(0).Visible = False
      cmdOK(2).Default = True
      If intNowReceive = intTotalReceive - 1 Then
         cmdOK(2).Visible = False
         cmdOK(3).Default = True
      End If
   Else
      cmdOK(2).Visible = False
      Select Case frm010001_1.intModifyKind
         Case 0
            '新增：可輸入所有資料
            fraWindow.Enabled = True
            If LastDate = "" Then
               txtCKind(0).Text = GetTaiwanTodayDate
            Else
               txtCKind(0).Text = LastDate
            End If
            cmdOK(0).Left = cmdOK(3).Left
            txtCKind(1).SetFocus
         Case 1
            '修改：可輸入所有資料
            ReadCkindDatabaseR 'Modify by Amy 2021/12/15 從下面搬上來
            fraWindow.Enabled = True
            txtCKind(1).SetFocus
            cmdOK(0).Left = cmdOK(2).Left
         Case 2
            '查詢：只可查詢
            'Modify by Amy 2021/12/15 從下面搬上來,因Form2.0元件都在fraWindow,若先執行fraWindow.Enabled = True,會一直彈查無資料訊息
            ReadCkindDatabaseR
            fraWindow.Enabled = False
            cmdOK(0).Visible = False
      End Select
   End If
   ' 90.12.05 modify by louis (再去計算一次本所期限及法定期限)
   If optDateKind(0).Value = True Then
      CheckKeyIn 4
   ElseIf optDateKind(1).Value = True Then
      CheckKeyIn 5
   ElseIf optDateKind(2).Value = True Then
      CheckKeyIn 6
   End If

   '2014/5/2 ADD BY SONIA 專利處人員操作隱藏尋找按鈕,在來函號數欄跳離時自動做按尋找動作
   'Modify by Amy 2021/12/16 改Form2.0 後主管機關來函查詢txtCKind(1).SetFocus 會當掉
   If Me.Caption <> "主管機關來函－查詢" And Me.Caption <> "主管機關來函查詢" Then
      If Left(Pub_StrUserSt03, 2) = "P1" Then
         CmdSearch.Visible = False
         txtCKind(1) = "1"
         txtCKind(1).Enabled = False
         txtCKind(2).SetFocus
      Else
         CmdSearch.Visible = True
         txtCKind(1).Enabled = True
         txtCKind(1).SetFocus
      End If
   End If
   '2014/5/2 END
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   'edit by nickc 2007/02/06 不用 dll 了
   'If obj001 Is Nothing Then
   '   Set obj001 = CreateObject("prjTaieDll001.cls001")
   '   Set obj001.Connection = cnnConnection
   'End If
   
   textCUID.BackColor = &H8000000F  'add by sonia 2014/5/2
   MoveFormToCenter Me
   ReadMemo
   If frm010001_1.intModifyKind = 0 Then
      txtCKind(1) = "1"
      txtCKind(3) = "1"
      txtCKind(8) = "1"
   End If
   optDateKind_Click (0)
   If bolIsQuery Then
      For i = 1 To frm010011.grdDataList.Rows - 1
         If frm010011.grdDataList.TextMatrix(i, 0) <> "" Then
            ReDim Preserve strReceiveCode(j)
            strReceiveCode(j) = frm010011.grdDataList.TextMatrix(i, 1)
            j = j + 1
         End If
      Next
      intTotalReceive = j
      intNowReceive = 0
   End If
   bolIsRun = False
   bolLeave = False
   intLeaveKind = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bolLeave = False And bolIsQuery = False Then
      If frm010001_1.intModifyKind = 0 Or frm010001_1.intModifyKind = 1 Then
         If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
            Cancel = 1
         End If
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bolIsQuery = False Then Where01ToGo intLeaveKind, "frm010001_1"
   bolIsQuery = False
End Sub

Private Sub optDateKind_Click(Index As Integer)
   txtCKind((Index) Mod 3 + 4).Enabled = True
   txtCKind((Index + 1) Mod 3 + 4).Enabled = False
   txtCKind((Index + 2) Mod 3 + 4).Enabled = False
   txtCKind((Index + 1) Mod 3 + 4).Text = ""
   txtCKind((Index + 2) Mod 3 + 4).Text = ""
End Sub

Private Sub txtCKind_Change(Index As Integer)
   Select Case Index
      Case 1, 2
         ClearCode
      Case 4, 5, 6
         lblOurDate.Caption = ""
         lblLawDate = ""
   End Select
End Sub

Private Sub txtCKind_Validate(Index As Integer, Cancel As Boolean)
   '2014/7/21 add by sonia
   bolCmdSearck = False
   If Index = 2 Then bolCmdSearck = True
   'end 2014/7/21
   
   If CheckKeyIn(Index) = False Then
      Cancel = True
      txtCKind_GotFocus (Index)
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCKind(7).IMEMode = 2
   'If Index = 7 And Cancel = False Then CloseIme 'Removed by Morgan 2016/10/20 會造成 Win7 的切換錯誤
End Sub

Private Sub ClearCode()
   txtSystem = ""
   txtCode(0) = ""
   txtCode(1) = ""
   txtCode(2) = ""
   txtTFCode(0) = ""
   txtTFCode(1) = ""
   txtTFCode(2) = ""
   txtTFCode(3) = ""
   cboCaseName.Clear
   lblPetition = ""
End Sub

Private Sub CheckChoose()
   Select Case Val(txtCKind(1))
      Case 1
         cboPatent.Visible = True
         cboTrademark.Visible = False
         txtCKind(7).Visible = False
      Case 2
         cboTrademark.Visible = True
         cboPatent.Visible = False
         txtCKind(7).Visible = False
      Case Else
         txtCKind(7).Visible = True
         cboPatent.Visible = False
         cboTrademark.Visible = False
   End Select
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Boolean
Dim strTemp As String
Dim iDays As Integer 'Added by Morgan 2019/7/12

   Select Case intIndex
      Case 0
         If txtCKind(intIndex) = "" Then
            'Modified by Morgan 2022/9/6
            'CheckKeyIn = True
            'Exit Function
            MsgBox "收件日不可空白!", vbCritical
            'end 2022/9/6
         ElseIf CheckIsTaiwanDate(txtCKind(intIndex).Text) Then
            CheckKeyIn = True
         End If
      Case 1
         If Val(txtCKind(intIndex).Text) < 1 Or Val(txtCKind(intIndex).Text) > 4 Then
            ShowMsg MsgText(1024)
         Else
            CheckKeyIn = True
            CheckChoose
         End If
      Case 3
         If Val(txtCKind(intIndex).Text) < 1 Or Val(txtCKind(intIndex).Text) > 4 Then
            ShowMsg MsgText(1025)
         Else
            CheckKeyIn = True
            If txtCKind(intIndex) = 3 Then
               optDateKind(0).Value = True
               txtCKind(4) = ""
               fraDate.Enabled = False
            'Add By Sindy 2015/7/28 補優先權證明,亦算出期限
            ElseIf txtCKind(intIndex) = 4 Then
               If txtCKind(1) <> "2" Then
                  MsgBox "非商標案，期限不可為「補優先權證明」!", vbCritical + vbOKOnly, MsgText(9001)
                  CheckKeyIn = False
                  Exit Function
               End If
               '補優先權證明,亦算出期限
               If InStr(cboTrademark.Text, "補優先權證明") = 0 Then
                  cboTrademark.Text = "補優先權證明" & IIf(Trim(cboTrademark.Text) <> "", ";" & cboTrademark.Text, "")
               End If
               optDateKind(1).Value = True
               txtCKind(5) = "3"
               If Val(m_ApplDate) > 0 Then
                  strTemp = ChangeWStringToWDateString(m_ApplDate)
                  strTemp = DateAdd("m", 3, strTemp) '申請日+3個月
                  lblLawDate = ChangeTStringToTDateString(ChangeWDateStringToTString(strTemp)) '法定期限
                  lblOurDate = ChangeWStringToTDateString(PUB_GetOurDeadline(lblLawDate)) '本所期限=法定期限-2工作天
               End If
               fraDate.Enabled = False
            '2015/7/28 END
            Else
               fraDate.Enabled = True
               If optDateKind(0).Value = True Then
                  CheckKeyIn 4
               ElseIf optDateKind(1).Value = True Then
                  CheckKeyIn 5
               Else
                  CheckKeyIn 6
               End If
            End If
         End If
      Case 4, 5
         If txtCKind(intIndex).Text <> "" Then
            If IsNumeric(txtCKind(intIndex).Text) = False Then
               ShowMsg MsgText(1026)
            Else
               CheckKeyIn = True
               'Modify by Morgan 2007/6/12
               'If CheckIsTaiwanDate(txtCKind(0)) = False Then
               '   txtCKind(0).SetFocus
               '   Exit Function
               'End If
               'strTemp = ChangeTStringToWDateString(txtCKind(0).Text)
               If txtDispDate.Visible = False Then
                  If CheckIsTaiwanDate(txtCKind(0)) = False Then
                     txtCKind(0).SetFocus
                     Exit Function
                  End If
                  strTemp = ChangeTStringToWDateString(txtCKind(0).Text)
               Else
                  If CheckIsTaiwanDate(txtDispDate) = False Then
                     txtDispDate.SetFocus
                     Exit Function
                  End If
                  strTemp = ChangeTStringToWDateString(txtDispDate.Text)
               End If
               'end 2007/6/12
               
               Select Case intIndex
                            Case 4
                                       strTemp = DateAdd("d", Val(txtCKind(intIndex).Text), strTemp)
                            Case 5
                                       ' 90.12.05 modify by louis (月數的計算有變)
                                       strTemp = ChangeWStringToWDateString(AddMonth(DBDATE(strTemp), Val(txtCKind(intIndex).Text)))
                                       'strTemp = DateAdd("m", Val(txtCKind(intIndex).Text), strTemp)
               End Select
               If txtCKind(3) = 2 Then
                  strTemp = DateAdd("d", -1, strTemp)
               End If
               lblLawDate = ChangeTStringToTDateString(ChangeWDateStringToTString(strTemp))
               'Added by Morgan 2014/10/2
               'Modified by Morgan 2014/11/20 外專改回舊規則
               If strSrvDate(1) >= 台灣案所限新規則啟用日 And txtSystem <> "FCP" And txtSystem <> "FG" Then
                  lblOurDate = ChangeWStringToTDateString(PUB_GetOurDeadline(lblLawDate))
               Else
               'end 2014/10/2
                                    
                  If Val(txtCKind(4)) = 60 Or Val(txtCKind(4)) = 90 Or Val(txtCKind(5)) = 2 Or Val(txtCKind(5)) = 3 Then
                     'Modified by Morgan 2019/7/12
                     'lblOurDate = ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -4, strTemp)))
                     iDays = 4
                     'end 2019/7/12
                  Else
                     'Modified by Morgan 2019/7/12
                     'lblOurDate = ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -2, strTemp)))
                     iDays = 2
                     'end 2019/7/12
                  End If
                  
                  'Added by Morgan 2019/7/12 外專台灣案所限以改工作天計算
                  If strSrvDate(1) >= 外專台灣案所限新規則啟用日 And (txtSystem = "FCP" Or txtSystem = "FG") Then
                     lblOurDate = ChangeWStringToTDateString(PUB_GetFCPOurDeadline(lblLawDate, iDays))
                  Else
                     lblOurDate = ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -1 * iDays, strTemp)))
                  End If
                  'end 2019/7/12
                  
                  'Add by Morgan 2004/3/15
                  'P案本所期限不可顯示為非工作日
                  If txtSystem = "P" And lblOurDate <> "" Then
                    lblOurDate = ChangeTStringToTDateString(ChangeWStringToTString(PUB_GetWorkDay1(ChangeTDateStringToTString(lblOurDate), True)))
                  End If
               End If 'Added by Morgan 2014/10/2
            End If
         Else
            CheckKeyIn = True
            lblOurDate = ""
            lblLawDate = ""
         End If
      Case 6
         If txtCKind(3) = "" Then
            ShowMsg MsgText(1027)
            Exit Function
         End If
         If txtCKind(intIndex).Text = "" Then
            CheckKeyIn = True
         Else
            If CheckIsTaiwanDate(txtCKind(intIndex).Text) Then
               CheckKeyIn = True
               lblLawDate.Caption = txtCKind(intIndex).Text
               lblLawDate.Caption = ChangeTStringToTDateString(lblLawDate.Caption)
               'Added by Morgan 2014/10/2
               'Modified by Morgan 2014/11/20 外專改回舊規則
               If strSrvDate(1) >= 台灣案所限新規則啟用日 And txtSystem <> "FCP" And txtSystem <> "FG" Then
                  lblOurDate = ChangeWStringToTDateString(PUB_GetOurDeadline(lblLawDate))
               Else
               'end 2014/10/2
                  strTemp = ChangeTStringToWDateString(txtCKind(intIndex).Text)
                  If txtSystem = "FCP" Then
                     'Added by Morgan 2019/7/12 外專台灣案所限以改工作天計算
                     If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
                        strTemp = ChangeWStringToTString(PUB_GetFCPOurDeadline(lblLawDate, 4))
                     Else
                     'end 2019/7/12
                     
                        strTemp = ChangeWDateStringToTString(DateAdd("d", -4, strTemp))
                        
                     End If 'Added by Morgan 2019/7/12
                  Else
                     strTemp = ChangeWDateStringToTString(DateAdd("d", -2, strTemp))
                  End If
                  lblOurDate = ChangeTStringToTDateString(strTemp)
                  'Add by Morgan 2004/3/15
                  'P案本所期限不可顯示為非工作日
                  If txtSystem = "P" And lblOurDate <> "" Then
                    lblOurDate = ChangeTStringToTDateString(ChangeWStringToTString(PUB_GetWorkDay1(ChangeTDateStringToTString(lblOurDate), True)))
                  End If
               End If 'Added by Morgan 2014/10/2
            End If
         End If
      Case 7
         'Add By Cheng 2002/03/25
         If Me.txtCKind(intIndex).Visible = True Then
            If Len(Me.txtCKind(intIndex).Text) <= 0 Then
               ShowMsg MsgText(10)
            Else
               CheckKeyIn = CheckLengthIsOK(txtCKind(intIndex), 40)
            End If
         Else
            CheckKeyIn = True
         End If
      Case 8
         '2010/1/29 MODIFY BY SONIA 加智商法院
         If Val(txtCKind(intIndex).Text) < 1 Or Val(txtCKind(intIndex).Text) > 8 Then
            ShowMsg MsgText(1028)
         Else
            CheckKeyIn = True
         End If
      '2014/5/2 add by sonia 專利處人員操作直接做按尋找動作且只能是P案件
      Case 2
         If Left(Pub_StrUserSt03, 2) = "P1" Then
            'modify by sonia 2014/7/21 P-108163(103302072)
            'cmdSearch_Click
            If bolCmdSearck = True Then cmdSearch_Click
            CheckKeyIn = True
         Else
            CheckKeyIn = True
         End If
      '2014/5/2 END
      Case Else
         CheckKeyIn = True
   End Select
   If CheckKeyIn = False And txtCKind(intIndex).Enabled Then txtCKind(intIndex).SetFocus
End Function

Private Sub txtCKind_GotFocus(Index As Integer)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCKind(7).IMEMode = 1
   If Index = 7 Then OpenIme Else CloseIme
   txtCKind(Index).SelStart = 0
   txtCKind(Index).SelLength = Len(txtCKind(Index).Text)
End Sub

Private Sub ReadCkindDatabaseR()
Dim mr01 As String, mr02 As String, mr03 As String, mr04 As String, mr05 As String, _
         mr06 As String, mr07 As String, mr08 As String, mr09 As String, mr10 As String, _
         mr11 As String, mr12 As String, mr13 As String, mr14 As String, mr15 As String, _
         i As Integer, rt As Boolean, strCodeName As String, strPetition As String, mr24 As String
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String, strCustomer As String

   'edit by nickc 2005/09/16 已經搬到 basquery
   'rt = obj001.ReadCKindDataBase(lblRecieveCode, mr02, mr03, mr04, mr05, mr06, mr07, _
          mr08, mr09, mr10, mr11, mr12, mr13, mr14, mr15)
   rt = Cls001ReadCKindDataBase(lblRecieveCode, mr02, mr03, mr04, mr05, mr06, mr07, _
          mr08, mr09, mr10, mr11, mr12, mr13, mr14, mr15, mr24, mr18, mr19, mr20, mr21, mr22, mr23)
   
   If rt Then
      txtCKind(0) = mr02
      txtCKind(1) = mr03
      '判斷是cboPatent or cboTrademark or txtCKind(8) 何者出現
      CheckChoose
      txtCKind(2) = mr04
      txtCKind(3) = mr05
      If mr05 <> 3 Then
         If mr06 <> "" Then
            optDateKind(0).Value = True
         ElseIf mr07 <> "" Then
            optDateKind(1).Value = True
         Else
            optDateKind(2).Value = True
         End If
      End If
      txtCKind(4) = mr06
      txtCKind(5) = mr07
      txtCKind(6) = mr08
      '計算期限
      CheckKeyIn 3
      If cboPatent.Visible Then
         cboPatent.Text = mr09
      ElseIf cboTrademark.Visible Then
         cboTrademark.Text = mr09
      Else
         txtCKind(7) = mr09
      End If
      txtCKind(8) = mr10
      txtCKind(9) = mr11
      txtSystem = mr12
      If mr12 <> 馬德里案 Then
         txtCode(0).Text = mr13
         txtCode(1).Text = IIf(mr14 = "0", "", mr14)
         txtCode(2).Text = IIf(mr15 = "00", "", mr15)
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer) = False Then GoTo Err
         'Modify By Sindy 2015/8/18 + m_ApplDate
         If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
             IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , , , , m_ApplDate) = False Then GoTo ErrHand
      Else
         txtTFCode(0).Text = Left(mr13, 5)
         txtTFCode(1).Text = IIf(Right(mr13, 1) = "0", "", Right(mr13, 1))
         txtTFCode(2).Text = IIf(mr14 = "0", "", mr14)
         txtTFCode(3).Text = IIf(mr15 = "00", "", mr15)
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
             IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer) = False Then GoTo Err
         'Modify By Sindy 2015/8/18 + m_ApplDate
         If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
             IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , , , , m_ApplDate) = False Then GoTo ErrHand
      End If
      'Added by Lydia 2021/11/19 讀取本所案號的系統種類
      intCKind = 0
      If txtSystem <> "" Or txtTFCode(0) <> "" Then
         If ClsPDGetSystemKind(IIf(txtSystem <> "", txtSystem, txtTFCode(0)), intCKind) = True Then
         End If
      End If
      'end 2021/11/19
      
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
      lblPetition = strCustomer
      If mr24 <> "" Then
         txtDispDate.Visible = True
         lblDispDate.Visible = True
         txtDispDate = TransDate(mr24, 1)
      Else
         txtDispDate.Visible = False
         lblDispDate.Visible = False
      End If
      '2014/5/2 add by sonia
      '更新CreateID及UpdateID
      UpdateCUID mr18, mr19, mr20, mr21, mr22, mr23
      '2014/5/2 end
   Else
ErrHand:
      ShowMsg MsgText(1029)
      bolLeave = True
      Unload Me
   End If
End Sub

Private Sub txtCode_Change(Index As Integer)
   lblPetition = ""
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index))
   CloseIme
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   If Index = 2 Then CheckCaseCode
End Sub

Private Sub txtDispDate_GotFocus()
   TextInverse txtDispDate
   CloseIme
End Sub

'Add by Morgan 2007/6/12
Private Sub txtDispDate_Validate(Cancel As Boolean)
   If txtDispDate <> "" Then
      If CheckIsTaiwanDate(txtDispDate) = False Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtSystem_GotFocus()
   txtSystem.SelStart = 0
   txtSystem.SelLength = Len(txtSystem)
   CloseIme
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
   If txtSystem <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetSystemKind(txtSystem.Text) = False Then
      If ClsPDGetSystemKind(txtSystem.Text) = False Then
         Cancel = True
         txtSystem_GotFocus
      ' 90.12.05 modify by louis (加重新計算本所期限及法定期限)
      Else
         If optDateKind(2).Value = True And Not IsEmptyText(txtCKind(6)) Then
            CheckKeyIn 6
         End If
      End If
   End If
End Sub

Private Sub txtTFCode_Change(Index As Integer)
   cboCaseName.Clear
   lblPetition = ""
End Sub

Private Sub txtSystem_Change()
   If txtSystem.Text = 馬德里案 Then
      fraTF.Visible = True
      fraElse.Visible = False
   Else
      fraTF.Visible = False
      fraElse.Visible = True
   End If
   lblPetition = ""
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Public Sub CheckCaseCode()
Dim rt As Boolean, strCodeName1 As String, strCodeName2 As String, strCodeName3 As String, strPetition As String
   
   m_ApplDate = "" 'Add By Sindy 2015/7/28 申請日
   If txtSystem.Text <> 馬德里案 Then
      'edit by nickc 2007/02/02 不用 dll 了
      'rt = objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
          IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCodeName1, strCodeName2, strCodeName3, strPetition)
      'Modify By Sindy 2015/8/18 + m_ApplDate
      rt = ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
          IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCodeName1, strCodeName2, strCodeName3, strPetition, , , , , m_ApplDate)
   Else
      'edit by nickc 2007/02/02 不用 dll 了
      'rt = objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
          IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCodeName1, strCodeName2, strCodeName3, strPetition)
      'Modify By Sindy 2015/8/18 + m_ApplDate
      rt = ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
          IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCodeName1, strCodeName2, strCodeName3, strPetition, , , , , m_ApplDate)
   End If
   If txtCode(1) = "0" Then
      txtCode(1) = ""
   End If
   If txtCode(2) = "00" Then
      txtCode(2) = ""
   End If
   If rt Then
      SetNameToCombo cboCaseName, strCodeName1, strCodeName2, strCodeName3
      lblPetition = strPetition
      If txtCKind(3) = 4 Then CheckKeyIn 3 'Add By Sindy 2015/7/28
      'Added by Lydia 2021/11/19 讀取本所案號的系統種類
      intCKind = 0
      If txtSystem <> "" Or txtTFCode(0) <> "" Then
         If ClsPDGetSystemKind(IIf(txtSystem <> "", txtSystem, txtTFCode(0)), intCKind) = True Then
         End If
      End If
      'end 2021/11/19
   Else
      cboCaseName.Clear
   End If
End Sub

Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
   If Index = 3 Then CheckCaseCode
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   'Add by Amy 2021/12/16 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
        strControlButton = MsgText(602)
         Exit Function
    End If

   For Each objTxt In Me.txtCKind
      If objTxt.Enabled = True Then
         Cancel = False
         txtCKind_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   For Each objTxt In Me.txtCode
      If objTxt.Enabled = True Then
         Cancel = False
         txtCode_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   If Me.txtSystem.Enabled = True Then
      Cancel = False
      txtSystem_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   For Each objTxt In Me.txtTFCode
      If objTxt.Enabled = True Then
         Cancel = False
         txtTFCode_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   'Added by Lydia 2021/11/19 增加對本所案號和系統種類的檢查;DB0010192因為是後面輸入本所案號再帶入查詢，所以增加檢查系統種類不符彈訊息不可存檔。
'   If txtSystem <> "" Or txtTFCode(0) <> "" Then
'       If intCKind <> txtCKind(1) Then
'           MsgBox "系統種類與本所案號不符！", vbExclamation
'           txtCKind(1).SetFocus
'           Call txtCKind_GotFocus(1)
'           Exit Function
'       End If
'   End If
   'end 2021/11/19
   TxtValidate = True
End Function

'add by sonia 2014/5/2
'更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef mr18 As String, ByRef mr19 As String, ByRef mr20 As String, ByRef mr21 As String, ByRef mr22 As String, ByRef mr23 As String)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(mr18) = False Then
      If IsEmptyText(mr18) = False Then
         strCName = GetStaffName(mr18, True)
      End If
   End If
   If IsNull(mr19) = False Then
      If IsEmptyText(mr19) = False Then
         strTemp = TAIWANDATE(mr19)
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(mr20) = False Then
      If IsEmptyText(mr20) = False Then
         strTemp = mr20
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(mr21) = False Then
      If IsEmptyText(mr21) = False Then
         strUName = GetStaffName(mr21, True)
      End If
   End If
   If IsNull(mr22) = False Then
      If IsEmptyText(mr22) = False Then
         strTemp = TAIWANDATE(mr22)
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(mr23) = False Then
      If IsEmptyText(mr23) = False Then
         strTemp = mr23
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " : " & strCDate & " " & _
              " : " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " : " & strUDate & " " & _
              " : " & strUTime
              
End Sub
'2014/5/2 end

'Move by Lydia 2016/12/15 從basQuery移回來
'新增CKind至資料庫
Private Function InsertCKindDatabase(ByRef mr01 As String, _
             ByRef mr02 As String, ByRef mr03 As String, ByRef mr04 As String, ByRef mr05 As String, _
             ByRef mr06 As String, ByRef mr07 As String, ByRef mr08 As String, _
             ByRef mr09 As String, ByRef mr10 As String, ByRef mr11 As String, ByRef mr12 As String, _
             ByRef mr13 As String, ByRef mr14 As String, ByRef mr15 As String, ByRef mr16 As String, _
             ByRef mr17 As String, Optional ByRef mr24 As String) As Boolean
Dim strSql As String
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo ErrHand

   cnnConnection.BeginTrans
   
   mr01 = mr01 + GetMDno(mr01) 'Modified by Lydia 2016/12/16 流水號存在acc1r0
   mr02 = ChangeTStringToWString(mr02)
   mr08 = ChangeTStringToWString(mr08)
   mr16 = ChangeTStringToWString(mr16)
   mr17 = ChangeTStringToWString(mr17)
   mr24 = ChangeTStringToWString(mr24)
   strSql = "insert into mailrec (mr01,mr02,mr03,mr04,mr05,mr06,mr07,mr08,mr09,mr10,mr11,mr12,mr13," + _
          "mr14,mr15,mr16,mr17,mr24) values (" + CNULL(mr01) + "," + CNULL(mr02) + "," + CNULL(mr03) + "," + CNULL(mr04) _
          + "," + CNULL(mr05) + "," + CNULL(mr06) + "," + CNULL(mr07) + "," + CNULL(mr08) + "," + CNULL(mr09) + "," + _
          CNULL(mr10) + "," + CNULL(mr11) + "," + CNULL(mr12) + "," + CNULL(mr13) + "," + CNULL(mr14) + "," + _
          CNULL(mr15) + "," + CNULL(mr16) + "," + CNULL(mr17) + "," + CNULL(mr24) + ")"
   cnnConnection.Execute strSql
   
   'Added by Morgan 2021/6/17
   'T與FCT共同控管案件通知
   If mr03 = "2" Then
      'Modified by Morgan 2022/1/14 增加案件,改抓系統特殊設定(檢查註冊號就好,阿妙確認她們只會輸註冊號)
      'If (mr04 = "01922108" Or mr04 = "01922109" Or mr04 = "106064194" Or mr04 = "106064195") Then
      strExc(0) = Pub_GetSpecMan("T與FCT共同管控案件")
      If InStr(";" & strExc(0) & ";", ";" & mr04 & ";") > 0 Then
      'end 2022/1/14
         PUB_2SysCaseInform mr12, mr13, mr14, mr15, mr01, 1
      End If
   End If
   'end 2021/6/17
   
   cnnConnection.CommitTrans
   InsertCKindDatabase = True

Exit Function
ErrHand:
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
   InsertCKindDatabase = False
   MsgBox "(" & Err.Number & ")" & Err.Description, vbExclamation + vbOKOnly, "新增來函紀錄檔動作失敗"

End Function

'Added by Lydia 2016/12/16 櫃台收文的主管機關來函抓最大流水號,記錄在Acc1r0
Private Function GetMDno(ByRef pNo As String) As String
Dim inA As Integer
Dim rsA1 As New ADODB.Recordset
Dim strQ As String

    strQ = "select nvl(a1r04,0) mno from acc1r0 where a1r01='MD' and a1r02=" & Left(strSrvDate(1), 4) & " and a1r03=1 "
    inA = 1
    Set rsA1 = ClsLawReadRstMsg(inA, strQ)
    If inA = 1 Then
       GetMDno = Format(Val(rsA1.Fields("mno")) + 1, "000000")
       strQ = "update acc1r0 set a1r04=" & Val(GetMDno) & " where a1r01='MD' and a1r02=" & Left(strSrvDate(1), 4) & " and a1r03=1 "
    Else
       GetMDno = "000001"
       If strSrvDate(1) < "20170101" Then
        strQ = "select max(mr01) mno from mailrec where substr(mr01,1," & Len(pNo) & ")=" & CNULL(pNo) & " and mr02>=" & Left(strSrvDate(1), 4) & "0000"
        inA = 1
        Set rsA1 = ClsLawReadRstMsg(inA, strQ)
        If inA = 1 Then
           If Trim("" & rsA1.Fields("mno")) <> "" Then
              GetMDno = Format(Val(Right(rsA1.Fields("mno"), 6)) + 1, "000000")
           End If
        End If
       End If
       
       strQ = "insert into acc1r0 (a1r01,a1r02,a1r03,a1r04) values ('MD'," & Left(strSrvDate(1), 4) & ",1," & Val(GetMDno) & ") "
    End If
    cnnConnection.Execute strQ
    Set rsA1 = Nothing
End Function
