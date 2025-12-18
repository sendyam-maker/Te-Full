VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210148_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "新客戶建檔"
   ClientHeight    =   6550
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6550
   ScaleWidth      =   8950
   Tag             =   "加班資料"
   Begin VB.TextBox txtCRL49 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   237
      Text            =   "收據公司"
      Top             =   780
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.TextBox txtAgent 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   236
      Text            =   "代理人"
      Top             =   780
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton CmdQueryCus 
      Caption         =   "申請人查詢"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   3360
      TabIndex        =   234
      Top             =   2340
      Width           =   1065
   End
   Begin VB.CheckBox ChkCRL66 
      BackColor       =   &H0000C000&
      Caption         =   "已簽准"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   233
      Top             =   2460
      Width           =   945
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H008080FF&
      Caption         =   "急件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   232
      Top             =   2460
      Width           =   765
   End
   Begin VB.ComboBox cboReason 
      Height          =   300
      ItemData        =   "frm210148_2.frx":0000
      Left            =   960
      List            =   "frm210148_2.frx":0007
      Locked          =   -1  'True
      TabIndex        =   231
      Text            =   "cboReason"
      Top             =   1440
      Width           =   2595
   End
   Begin VB.TextBox txtF0316 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   3870
      Locked          =   -1  'True
      TabIndex        =   213
      Top             =   450
      Width           =   600
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   4
      Left            =   7470
      Locked          =   -1  'True
      MousePointer    =   1  '箭號形狀
      TabIndex        =   191
      Text            =   "申請人5"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   3
      Left            =   5700
      Locked          =   -1  'True
      MousePointer    =   1  '箭號形狀
      TabIndex        =   190
      Text            =   "申請人4"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   2
      Left            =   3960
      Locked          =   -1  'True
      MousePointer    =   1  '箭號形狀
      TabIndex        =   189
      Text            =   "申請人3"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   1
      Left            =   2190
      Locked          =   -1  'True
      MousePointer    =   1  '箭號形狀
      TabIndex        =   188
      Text            =   "申請人2"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   420
      Locked          =   -1  'True
      MousePointer    =   1  '箭號形狀
      TabIndex        =   187
      Text            =   "申請人1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton CmdAddCus 
      BackColor       =   &H00C0C0FF&
      Caption         =   "建客戶檔"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4860
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   2340
      Width           =   885
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   3870
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   780
      Width           =   1665
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   780
      Width           =   1665
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "檢視接洽單"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   2880
      TabIndex        =   1
      Top             =   30
      Width           =   1065
   End
   Begin VB.TextBox txtF0309 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   7215
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   1665
   End
   Begin VB.TextBox txtCRL78 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   450
      Width           =   600
   End
   Begin VB.TextBox txtF0301 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton cmdQueryNext 
      Caption         =   "查詢下一筆(&N)"
      Height          =   360
      Left            =   6630
      TabIndex        =   4
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "退智權人員(&B)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5280
      TabIndex        =   3
      Top             =   30
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "完成(&O)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4350
      TabIndex        =   2
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8040
      TabIndex        =   5
      Top             =   30
      Width           =   885
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3825
      Left            =   60
      TabIndex        =   20
      Top             =   2730
      Width           =   8805
      _ExtentX        =   15522
      _ExtentY        =   6756
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "申請人1"
      TabPicture(0)   =   "frm210148_2.frx":0017
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(21)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(20)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(19)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(17)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(22)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(94)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(77)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(12)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(13)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(14)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(15)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(18)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(85)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(92)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtCRA11(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtCRA11(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtCRA11(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtCRA11(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtCRA11(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtCRA11(6)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(78)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(84)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TxtCRA11(7)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxtCRA04(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TxtCRA1(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TxtCRA1(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtCRA1(4)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtCRA1(5)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtCRA1(6)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtCRA1(7)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtCRA1(8)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TxtCRA1(9)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtCRA1(11)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "TxtCRA1(10)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "ChkCRA26(0)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TxtCRA(1)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Frame1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtCRL51"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "申請人2"
      TabPicture(1)   =   "frm210148_2.frx":0033
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(6)"
      Tab(1).Control(1)=   "Label1(7)"
      Tab(1).Control(2)=   "Label1(8)"
      Tab(1).Control(3)=   "Label1(9)"
      Tab(1).Control(4)=   "Label1(10)"
      Tab(1).Control(5)=   "Label1(11)"
      Tab(1).Control(6)=   "Label1(16)"
      Tab(1).Control(7)=   "Label1(23)"
      Tab(1).Control(8)=   "Label1(24)"
      Tab(1).Control(9)=   "Label1(25)"
      Tab(1).Control(10)=   "Label1(26)"
      Tab(1).Control(11)=   "Label1(27)"
      Tab(1).Control(12)=   "Label1(28)"
      Tab(1).Control(13)=   "Label1(29)"
      Tab(1).Control(14)=   "Label1(30)"
      Tab(1).Control(15)=   "TxtCRA21(6)"
      Tab(1).Control(16)=   "TxtCRA21(5)"
      Tab(1).Control(17)=   "TxtCRA21(3)"
      Tab(1).Control(18)=   "TxtCRA21(4)"
      Tab(1).Control(19)=   "TxtCRA21(2)"
      Tab(1).Control(20)=   "TxtCRA21(1)"
      Tab(1).Control(21)=   "Label1(82)"
      Tab(1).Control(22)=   "Label1(86)"
      Tab(1).Control(23)=   "TxtCRA21(7)"
      Tab(1).Control(24)=   "TxtCRA04(2)"
      Tab(1).Control(25)=   "TxtCRA2(10)"
      Tab(1).Control(26)=   "TxtCRA2(11)"
      Tab(1).Control(27)=   "TxtCRA2(9)"
      Tab(1).Control(28)=   "TxtCRA2(8)"
      Tab(1).Control(29)=   "TxtCRA2(7)"
      Tab(1).Control(30)=   "TxtCRA2(6)"
      Tab(1).Control(31)=   "TxtCRA2(5)"
      Tab(1).Control(32)=   "TxtCRA2(4)"
      Tab(1).Control(33)=   "TxtCRA2(3)"
      Tab(1).Control(34)=   "TxtCRA2(2)"
      Tab(1).Control(35)=   "TxtCRA(2)"
      Tab(1).Control(36)=   "ChkCRA26(1)"
      Tab(1).Control(37)=   "Frame2"
      Tab(1).ControlCount=   38
      TabCaption(2)   =   "申請人3"
      TabPicture(2)   =   "frm210148_2.frx":004F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(31)"
      Tab(2).Control(1)=   "Label1(32)"
      Tab(2).Control(2)=   "Label1(33)"
      Tab(2).Control(3)=   "Label1(34)"
      Tab(2).Control(4)=   "Label1(35)"
      Tab(2).Control(5)=   "Label1(36)"
      Tab(2).Control(6)=   "Label1(37)"
      Tab(2).Control(7)=   "Label1(38)"
      Tab(2).Control(8)=   "Label1(39)"
      Tab(2).Control(9)=   "Label1(40)"
      Tab(2).Control(10)=   "Label1(41)"
      Tab(2).Control(11)=   "Label1(42)"
      Tab(2).Control(12)=   "Label1(43)"
      Tab(2).Control(13)=   "Label1(44)"
      Tab(2).Control(14)=   "Label1(45)"
      Tab(2).Control(15)=   "TxtCRA31(6)"
      Tab(2).Control(16)=   "TxtCRA31(5)"
      Tab(2).Control(17)=   "TxtCRA31(3)"
      Tab(2).Control(18)=   "TxtCRA31(4)"
      Tab(2).Control(19)=   "TxtCRA31(2)"
      Tab(2).Control(20)=   "TxtCRA31(1)"
      Tab(2).Control(21)=   "Label1(79)"
      Tab(2).Control(22)=   "Label1(87)"
      Tab(2).Control(23)=   "TxtCRA31(7)"
      Tab(2).Control(24)=   "TxtCRA04(3)"
      Tab(2).Control(25)=   "TxtCRA3(2)"
      Tab(2).Control(26)=   "TxtCRA3(3)"
      Tab(2).Control(27)=   "TxtCRA3(4)"
      Tab(2).Control(28)=   "TxtCRA3(5)"
      Tab(2).Control(29)=   "TxtCRA3(6)"
      Tab(2).Control(30)=   "TxtCRA3(7)"
      Tab(2).Control(31)=   "TxtCRA3(8)"
      Tab(2).Control(32)=   "TxtCRA3(9)"
      Tab(2).Control(33)=   "TxtCRA3(11)"
      Tab(2).Control(34)=   "TxtCRA3(10)"
      Tab(2).Control(35)=   "ChkCRA26(2)"
      Tab(2).Control(36)=   "TxtCRA(3)"
      Tab(2).Control(37)=   "Frame3"
      Tab(2).ControlCount=   38
      TabCaption(3)   =   "申請人4"
      TabPicture(3)   =   "frm210148_2.frx":006B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(46)"
      Tab(3).Control(1)=   "Label1(47)"
      Tab(3).Control(2)=   "Label1(48)"
      Tab(3).Control(3)=   "Label1(49)"
      Tab(3).Control(4)=   "Label1(50)"
      Tab(3).Control(5)=   "Label1(51)"
      Tab(3).Control(6)=   "Label1(52)"
      Tab(3).Control(7)=   "Label1(53)"
      Tab(3).Control(8)=   "Label1(54)"
      Tab(3).Control(9)=   "Label1(55)"
      Tab(3).Control(10)=   "Label1(56)"
      Tab(3).Control(11)=   "Label1(57)"
      Tab(3).Control(12)=   "Label1(58)"
      Tab(3).Control(13)=   "Label1(59)"
      Tab(3).Control(14)=   "Label1(60)"
      Tab(3).Control(15)=   "TxtCRA41(6)"
      Tab(3).Control(16)=   "TxtCRA41(5)"
      Tab(3).Control(17)=   "TxtCRA41(3)"
      Tab(3).Control(18)=   "TxtCRA41(4)"
      Tab(3).Control(19)=   "TxtCRA41(2)"
      Tab(3).Control(20)=   "TxtCRA41(1)"
      Tab(3).Control(21)=   "Label1(80)"
      Tab(3).Control(22)=   "Label1(88)"
      Tab(3).Control(23)=   "TxtCRA41(7)"
      Tab(3).Control(24)=   "TxtCRA04(4)"
      Tab(3).Control(25)=   "TxtCRA4(2)"
      Tab(3).Control(26)=   "TxtCRA4(3)"
      Tab(3).Control(27)=   "TxtCRA4(4)"
      Tab(3).Control(28)=   "TxtCRA4(5)"
      Tab(3).Control(29)=   "TxtCRA4(6)"
      Tab(3).Control(30)=   "TxtCRA4(7)"
      Tab(3).Control(31)=   "TxtCRA4(8)"
      Tab(3).Control(32)=   "TxtCRA4(9)"
      Tab(3).Control(33)=   "TxtCRA4(11)"
      Tab(3).Control(34)=   "TxtCRA4(10)"
      Tab(3).Control(35)=   "ChkCRA26(3)"
      Tab(3).Control(36)=   "TxtCRA(4)"
      Tab(3).Control(37)=   "Frame4"
      Tab(3).ControlCount=   38
      TabCaption(4)   =   "申請人5"
      TabPicture(4)   =   "frm210148_2.frx":0087
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(61)"
      Tab(4).Control(1)=   "Label1(62)"
      Tab(4).Control(2)=   "Label1(63)"
      Tab(4).Control(3)=   "Label1(64)"
      Tab(4).Control(4)=   "Label1(65)"
      Tab(4).Control(5)=   "Label1(66)"
      Tab(4).Control(6)=   "Label1(67)"
      Tab(4).Control(7)=   "Label1(68)"
      Tab(4).Control(8)=   "Label1(69)"
      Tab(4).Control(9)=   "Label1(70)"
      Tab(4).Control(10)=   "Label1(71)"
      Tab(4).Control(11)=   "Label1(72)"
      Tab(4).Control(12)=   "Label1(73)"
      Tab(4).Control(13)=   "Label1(74)"
      Tab(4).Control(14)=   "Label1(75)"
      Tab(4).Control(15)=   "TxtCRA51(6)"
      Tab(4).Control(16)=   "TxtCRA51(5)"
      Tab(4).Control(17)=   "TxtCRA51(3)"
      Tab(4).Control(18)=   "TxtCRA51(4)"
      Tab(4).Control(19)=   "TxtCRA51(2)"
      Tab(4).Control(20)=   "TxtCRA51(1)"
      Tab(4).Control(21)=   "Label1(81)"
      Tab(4).Control(22)=   "Label1(89)"
      Tab(4).Control(23)=   "TxtCRA51(7)"
      Tab(4).Control(24)=   "TxtCRA04(5)"
      Tab(4).Control(25)=   "TxtCRA5(2)"
      Tab(4).Control(26)=   "TxtCRA5(3)"
      Tab(4).Control(27)=   "TxtCRA5(4)"
      Tab(4).Control(28)=   "TxtCRA5(5)"
      Tab(4).Control(29)=   "TxtCRA5(6)"
      Tab(4).Control(30)=   "TxtCRA5(7)"
      Tab(4).Control(31)=   "TxtCRA5(8)"
      Tab(4).Control(32)=   "TxtCRA5(9)"
      Tab(4).Control(33)=   "TxtCRA5(11)"
      Tab(4).Control(34)=   "TxtCRA5(10)"
      Tab(4).Control(35)=   "ChkCRA26(4)"
      Tab(4).Control(36)=   "TxtCRA(5)"
      Tab(4).Control(37)=   "Frame5"
      Tab(4).ControlCount=   38
      Begin VB.TextBox txtCRL51 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   235
         Text            =   "txtCRL51"
         Top             =   390
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.Frame Frame5 
         Height          =   255
         Left            =   -71820
         TabIndex        =   205
         Top             =   390
         Width           =   2200
         Begin VB.OptionButton Option5 
            Caption         =   "新客戶"
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   207
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option5 
            Caption         =   "舊客戶"
            Height          =   285
            Index           =   1
            Left            =   1140
            TabIndex        =   206
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Height          =   255
         Left            =   -71820
         TabIndex        =   202
         Top             =   390
         Width           =   2200
         Begin VB.OptionButton Option4 
            Caption         =   "新客戶"
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   204
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "舊客戶"
            Height          =   285
            Index           =   1
            Left            =   1140
            TabIndex        =   203
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   255
         Left            =   -71820
         TabIndex        =   199
         Top             =   390
         Width           =   2200
         Begin VB.OptionButton Option3 
            Caption         =   "新客戶"
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   201
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "舊客戶"
            Height          =   285
            Index           =   1
            Left            =   1140
            TabIndex        =   200
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Height          =   255
         Left            =   3180
         TabIndex        =   196
         Top             =   390
         Width           =   2200
         Begin VB.OptionButton Option1 
            Caption         =   "舊客戶"
            Height          =   285
            Index           =   1
            Left            =   1140
            TabIndex        =   198
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "新客戶"
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   197
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   255
         Left            =   -71820
         TabIndex        =   193
         Top             =   390
         Width           =   2200
         Begin VB.OptionButton Option2 
            Caption         =   "新客戶"
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   195
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "舊客戶"
            Height          =   285
            Index           =   1
            Left            =   1140
            TabIndex        =   194
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox TxtCRA 
         Height          =   300
         Index           =   5
         Left            =   -73890
         MaxLength       =   9
         TabIndex        =   137
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtCRA 
         Height          =   300
         Index           =   4
         Left            =   -73890
         MaxLength       =   9
         TabIndex        =   111
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtCRA 
         Height          =   300
         Index           =   3
         Left            =   -73890
         MaxLength       =   9
         TabIndex        =   85
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox ChkCRA26 
         Caption         =   "有對造"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   4
         Left            =   -74820
         TabIndex        =   84
         Top             =   390
         Width           =   1000
      End
      Begin VB.CheckBox ChkCRA26 
         Caption         =   "有對造"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   3
         Left            =   -74820
         TabIndex        =   83
         Top             =   390
         Width           =   1000
      End
      Begin VB.CheckBox ChkCRA26 
         Caption         =   "有對造"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   2
         Left            =   -74820
         TabIndex        =   82
         Top             =   390
         Width           =   1000
      End
      Begin VB.CheckBox ChkCRA26 
         Caption         =   "有對造"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   1
         Left            =   -74820
         TabIndex        =   66
         Top             =   390
         Width           =   1000
      End
      Begin VB.TextBox TxtCRA 
         Height          =   300
         Index           =   2
         Left            =   -73890
         MaxLength       =   9
         TabIndex        =   65
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtCRA 
         Height          =   300
         Index           =   1
         Left            =   1110
         MaxLength       =   9
         TabIndex        =   40
         Text            =   "TxtCRA(1)"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox ChkCRA26 
         Caption         =   "有對造"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   390
         Width           =   1000
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   10
         Left            =   -73890
         TabIndex        =   147
         Top             =   2460
         Width           =   800
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   11
         Left            =   -73890
         TabIndex        =   146
         Top             =   2790
         Width           =   800
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   9
         Left            =   -69570
         TabIndex        =   145
         Top             =   1380
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   8
         Left            =   -73890
         TabIndex        =   144
         Top             =   1380
         Width           =   2500
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "4410;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   7
         Left            =   -67590
         TabIndex        =   143
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   6
         Left            =   -69570
         TabIndex        =   142
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   5
         Left            =   -73890
         TabIndex        =   141
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   4
         Left            =   -67590
         TabIndex        =   140
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   3
         Left            =   -69570
         TabIndex        =   139
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA5 
         Height          =   300
         Index           =   2
         Left            =   -71820
         TabIndex        =   138
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   10
         Left            =   -73890
         TabIndex        =   121
         Top             =   2460
         Width           =   800
         VariousPropertyBits=   679495711
         MaxLength       =   9
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   11
         Left            =   -73890
         TabIndex        =   120
         Top             =   2790
         Width           =   800
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   9
         Left            =   -69570
         TabIndex        =   119
         Top             =   1380
         Width           =   1215
         VariousPropertyBits=   679495711
         MaxLength       =   9
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   8
         Left            =   -73890
         TabIndex        =   118
         Top             =   1380
         Width           =   2500
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "4410;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   7
         Left            =   -67590
         TabIndex        =   117
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         MaxLength       =   9
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   6
         Left            =   -69570
         TabIndex        =   116
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         MaxLength       =   9
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   5
         Left            =   -73890
         TabIndex        =   115
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         MaxLength       =   9
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   4
         Left            =   -67590
         TabIndex        =   114
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         MaxLength       =   9
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   3
         Left            =   -69570
         TabIndex        =   113
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         MaxLength       =   9
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA4 
         Height          =   300
         Index           =   2
         Left            =   -71820
         TabIndex        =   112
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   10
         Left            =   -73890
         TabIndex        =   95
         Top             =   2460
         Width           =   800
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   11
         Left            =   -73890
         TabIndex        =   94
         Top             =   2790
         Width           =   800
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   9
         Left            =   -69570
         TabIndex        =   93
         Top             =   1380
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   8
         Left            =   -73890
         TabIndex        =   92
         Top             =   1380
         Width           =   2500
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "4410;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   7
         Left            =   -67590
         TabIndex        =   91
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   6
         Left            =   -69570
         TabIndex        =   90
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   5
         Left            =   -73890
         TabIndex        =   89
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   4
         Left            =   -67590
         TabIndex        =   88
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   3
         Left            =   -69570
         TabIndex        =   87
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA3 
         Height          =   300
         Index           =   2
         Left            =   -71820
         TabIndex        =   86
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   2
         Left            =   -71820
         TabIndex        =   64
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   3
         Left            =   -69570
         TabIndex        =   63
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   4
         Left            =   -67590
         TabIndex        =   62
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   5
         Left            =   -73890
         TabIndex        =   61
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   6
         Left            =   -69570
         TabIndex        =   60
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   7
         Left            =   -67590
         TabIndex        =   59
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   8
         Left            =   -73890
         TabIndex        =   58
         Top             =   1380
         Width           =   2500
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "4410;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   9
         Left            =   -69570
         TabIndex        =   57
         Top             =   1380
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   11
         Left            =   -73890
         TabIndex        =   56
         Top             =   2790
         Width           =   800
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA2 
         Height          =   300
         Index           =   10
         Left            =   -73890
         TabIndex        =   55
         Top             =   2460
         Width           =   800
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   10
         Left            =   1110
         TabIndex        =   54
         Top             =   2460
         Width           =   800
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   11
         Left            =   1110
         TabIndex        =   53
         Top             =   2790
         Width           =   800
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "1411;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   9
         Left            =   5430
         TabIndex        =   48
         Top             =   1380
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   8
         Left            =   1110
         TabIndex        =   47
         Top             =   1380
         Width           =   2500
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "4410;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   7
         Left            =   7410
         TabIndex        =   46
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   6
         Left            =   5430
         TabIndex        =   45
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   5
         Left            =   1110
         TabIndex        =   44
         Top             =   1050
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   4
         Left            =   7410
         TabIndex        =   43
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   3
         Left            =   5430
         TabIndex        =   42
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA1 
         Height          =   300
         Index           =   2
         Left            =   3180
         TabIndex        =   41
         Top             =   720
         Width           =   1215
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA04 
         Height          =   300
         Index           =   2
         Left            =   -69210
         TabIndex        =   230
         Top             =   390
         Width           =   1900
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "3351;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA04 
         Height          =   300
         Index           =   3
         Left            =   -69210
         TabIndex        =   229
         Top             =   390
         Width           =   1900
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "3351;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA04 
         Height          =   300
         Index           =   4
         Left            =   -69210
         TabIndex        =   228
         Top             =   390
         Width           =   1900
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "3351;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA04 
         Height          =   300
         Index           =   5
         Left            =   -69210
         TabIndex        =   227
         Top             =   390
         Width           =   1900
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "3351;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA04 
         Height          =   300
         Index           =   1
         Left            =   5790
         TabIndex        =   226
         Top             =   390
         Width           =   1900
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "3351;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA21 
         Height          =   600
         Index           =   7
         Left            =   -73890
         TabIndex        =   225
         Top             =   3120
         Width           =   7520
         VariousPropertyBits=   679495711
         Size            =   "13264;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA31 
         Height          =   600
         Index           =   7
         Left            =   -73890
         TabIndex        =   224
         Top             =   3120
         Width           =   7520
         VariousPropertyBits=   679495711
         Size            =   "13264;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA41 
         Height          =   600
         Index           =   7
         Left            =   -73890
         TabIndex        =   223
         Top             =   3120
         Width           =   7520
         VariousPropertyBits=   679495711
         Size            =   "13264;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA51 
         Height          =   600
         Index           =   7
         Left            =   -73890
         TabIndex        =   222
         Top             =   3120
         Width           =   7520
         VariousPropertyBits=   679495711
         Size            =   "13264;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "英文地址："
         Height          =   255
         Index           =   89
         Left            =   -74820
         TabIndex        =   221
         Top             =   3120
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "英文地址："
         Height          =   255
         Index           =   88
         Left            =   -74820
         TabIndex        =   220
         Top             =   3120
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "英文地址："
         Height          =   255
         Index           =   87
         Left            =   -74820
         TabIndex        =   219
         Top             =   3120
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "英文地址："
         Height          =   255
         Index           =   86
         Left            =   -74820
         TabIndex        =   218
         Top             =   3120
         Width           =   1005
      End
      Begin MSForms.TextBox TxtCRA11 
         Height          =   600
         Index           =   7
         Left            =   1110
         TabIndex        =   217
         Top             =   3120
         Width           =   7520
         VariousPropertyBits=   679495711
         Size            =   "13264;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "英文地址："
         Height          =   255
         Index           =   84
         Left            =   180
         TabIndex        =   216
         Top             =   3120
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "與                                            為關係企業"
         Height          =   260
         Index           =   81
         Left            =   -69570
         TabIndex        =   212
         Top             =   450
         Width           =   3320
      End
      Begin VB.Label Label1 
         Caption         =   "與                                            為關係企業"
         Height          =   260
         Index           =   80
         Left            =   -69570
         TabIndex        =   211
         Top             =   450
         Width           =   3320
      End
      Begin VB.Label Label1 
         Caption         =   "與                                            為關係企業"
         Height          =   260
         Index           =   79
         Left            =   -69570
         TabIndex        =   210
         Top             =   450
         Width           =   3320
      End
      Begin VB.Label Label1 
         Caption         =   "與                                            為關係企業"
         Height          =   260
         Index           =   78
         Left            =   5460
         TabIndex        =   209
         Top             =   450
         Width           =   3320
      End
      Begin VB.Label Label1 
         Caption         =   "與                                            為關係企業"
         Height          =   260
         Index           =   82
         Left            =   -69540
         TabIndex        =   208
         Top             =   450
         Width           =   3310
      End
      Begin MSForms.TextBox TxtCRA51 
         Height          =   300
         Index           =   1
         Left            =   -73530
         TabIndex        =   186
         Top             =   1800
         Width           =   3615
         VariousPropertyBits=   679495711
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA51 
         Height          =   300
         Index           =   2
         Left            =   -73530
         TabIndex        =   185
         Top             =   2130
         Width           =   3615
         VariousPropertyBits=   679495711
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA51 
         Height          =   300
         Index           =   4
         Left            =   -68370
         TabIndex        =   184
         Top             =   2130
         Width           =   1995
         VariousPropertyBits=   679495711
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA51 
         Height          =   300
         Index           =   3
         Left            =   -68370
         TabIndex        =   183
         Top             =   1800
         Width           =   1995
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA51 
         Height          =   300
         Index           =   5
         Left            =   -73080
         TabIndex        =   182
         Top             =   2460
         Width           =   6705
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA51 
         Height          =   300
         Index           =   6
         Left            =   -73080
         TabIndex        =   181
         Top             =   2790
         Width           =   6705
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA41 
         Height          =   300
         Index           =   1
         Left            =   -73530
         TabIndex        =   180
         Top             =   1800
         Width           =   3615
         VariousPropertyBits=   679495711
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA41 
         Height          =   300
         Index           =   2
         Left            =   -73530
         TabIndex        =   179
         Top             =   2130
         Width           =   3615
         VariousPropertyBits=   679495711
         MaxLength       =   9
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA41 
         Height          =   300
         Index           =   4
         Left            =   -68370
         TabIndex        =   178
         Top             =   2130
         Width           =   1995
         VariousPropertyBits=   679495711
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA41 
         Height          =   300
         Index           =   3
         Left            =   -68370
         TabIndex        =   177
         Top             =   1800
         Width           =   1995
         VariousPropertyBits=   679495711
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA41 
         Height          =   300
         Index           =   5
         Left            =   -73080
         TabIndex        =   176
         Top             =   2460
         Width           =   6705
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA41 
         Height          =   300
         Index           =   6
         Left            =   -73080
         TabIndex        =   175
         Top             =   2790
         Width           =   6705
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA31 
         Height          =   300
         Index           =   1
         Left            =   -73530
         TabIndex        =   174
         Top             =   1800
         Width           =   3615
         VariousPropertyBits=   679495711
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA31 
         Height          =   300
         Index           =   2
         Left            =   -73530
         TabIndex        =   173
         Top             =   2130
         Width           =   3615
         VariousPropertyBits=   679495711
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA31 
         Height          =   300
         Index           =   4
         Left            =   -68370
         TabIndex        =   172
         Top             =   2130
         Width           =   1995
         VariousPropertyBits=   679495711
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA31 
         Height          =   300
         Index           =   3
         Left            =   -68370
         TabIndex        =   171
         Top             =   1800
         Width           =   1995
         VariousPropertyBits=   679495711
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA31 
         Height          =   300
         Index           =   5
         Left            =   -73080
         TabIndex        =   170
         Top             =   2460
         Width           =   6705
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA31 
         Height          =   300
         Index           =   6
         Left            =   -73080
         TabIndex        =   169
         Top             =   2790
         Width           =   6705
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA21 
         Height          =   300
         Index           =   1
         Left            =   -73530
         TabIndex        =   168
         Top             =   1800
         Width           =   3615
         VariousPropertyBits=   679495711
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA21 
         Height          =   300
         Index           =   2
         Left            =   -73530
         TabIndex        =   167
         Top             =   2130
         Width           =   3615
         VariousPropertyBits=   679495711
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA21 
         Height          =   300
         Index           =   4
         Left            =   -68370
         TabIndex        =   166
         Top             =   2130
         Width           =   1995
         VariousPropertyBits=   679495711
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA21 
         Height          =   300
         Index           =   3
         Left            =   -68370
         TabIndex        =   165
         Top             =   1800
         Width           =   1995
         VariousPropertyBits=   679495711
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA21 
         Height          =   300
         Index           =   5
         Left            =   -73080
         TabIndex        =   164
         Top             =   2460
         Width           =   6705
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA21 
         Height          =   300
         Index           =   6
         Left            =   -73080
         TabIndex        =   163
         Top             =   2790
         Width           =   6705
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "申請地址："
         Height          =   255
         Index           =   75
         Left            =   -74820
         TabIndex        =   162
         Top             =   2790
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡地址："
         Height          =   255
         Index           =   74
         Left            =   -74820
         TabIndex        =   161
         Top             =   2490
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   73
         Left            =   -70260
         TabIndex        =   160
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "傳真："
         Height          =   180
         Index           =   72
         Left            =   -68220
         TabIndex        =   159
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "電話："
         Height          =   180
         Index           =   71
         Left            =   -70260
         TabIndex        =   158
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   180
         Index           =   70
         Left            =   -72540
         TabIndex        =   157
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   69
         Left            =   -74820
         TabIndex        =   156
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID No.："
         Height          =   180
         Index           =   68
         Left            =   -74820
         TabIndex        =   155
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "國   籍："
         Height          =   255
         Index           =   67
         Left            =   -70260
         TabIndex        =   154
         Top             =   1410
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   255
         Index           =   66
         Left            =   -69780
         TabIndex        =   153
         Top             =   2115
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail："
         Height          =   255
         Index           =   65
         Left            =   -74820
         TabIndex        =   152
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "申請人(中文)："
         Height          =   255
         Index           =   64
         Left            =   -74820
         TabIndex        =   151
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   255
         Index           =   63
         Left            =   -74820
         TabIndex        =   150
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "代表人(中文)："
         Height          =   255
         Index           =   62
         Left            =   -69780
         TabIndex        =   149
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手機："
         Height          =   180
         Index           =   61
         Left            =   -68220
         TabIndex        =   148
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "申請地址："
         Height          =   195
         Index           =   60
         Left            =   -74820
         TabIndex        =   136
         Top             =   2790
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡地址："
         Height          =   195
         Index           =   59
         Left            =   -74820
         TabIndex        =   135
         Top             =   2490
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   58
         Left            =   -70260
         TabIndex        =   134
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "傳真："
         Height          =   180
         Index           =   57
         Left            =   -68220
         TabIndex        =   133
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "電話："
         Height          =   180
         Index           =   56
         Left            =   -70260
         TabIndex        =   132
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   180
         Index           =   55
         Left            =   -72540
         TabIndex        =   131
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   54
         Left            =   -74820
         TabIndex        =   130
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID No.："
         Height          =   240
         Index           =   53
         Left            =   -74820
         TabIndex        =   129
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "國   籍："
         Height          =   195
         Index           =   52
         Left            =   -70260
         TabIndex        =   128
         Top             =   1410
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   195
         Index           =   51
         Left            =   -69780
         TabIndex        =   127
         Top             =   2115
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail："
         Height          =   195
         Index           =   50
         Left            =   -74820
         TabIndex        =   126
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "申請人(中文)："
         Height          =   195
         Index           =   49
         Left            =   -74820
         TabIndex        =   125
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   225
         Index           =   48
         Left            =   -74820
         TabIndex        =   124
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "代表人(中文)："
         Height          =   195
         Index           =   47
         Left            =   -69780
         TabIndex        =   123
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手機："
         Height          =   240
         Index           =   46
         Left            =   -68220
         TabIndex        =   122
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "申請地址："
         Height          =   255
         Index           =   45
         Left            =   -74820
         TabIndex        =   110
         Top             =   2790
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡地址："
         Height          =   255
         Index           =   44
         Left            =   -74820
         TabIndex        =   109
         Top             =   2490
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   43
         Left            =   -70260
         TabIndex        =   108
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "傳真："
         Height          =   180
         Index           =   42
         Left            =   -68220
         TabIndex        =   107
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "電話："
         Height          =   180
         Index           =   41
         Left            =   -70260
         TabIndex        =   106
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   180
         Index           =   40
         Left            =   -72540
         TabIndex        =   105
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   39
         Left            =   -74820
         TabIndex        =   104
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID No.："
         Height          =   180
         Index           =   38
         Left            =   -74820
         TabIndex        =   103
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "國   籍："
         Height          =   255
         Index           =   37
         Left            =   -70260
         TabIndex        =   102
         Top             =   1410
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   255
         Index           =   36
         Left            =   -69780
         TabIndex        =   101
         Top             =   2115
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail："
         Height          =   255
         Index           =   35
         Left            =   -74820
         TabIndex        =   100
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "申請人(中文)："
         Height          =   255
         Index           =   34
         Left            =   -74820
         TabIndex        =   99
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   255
         Index           =   33
         Left            =   -74820
         TabIndex        =   98
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "代表人(中文)："
         Height          =   255
         Index           =   32
         Left            =   -69780
         TabIndex        =   97
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手機："
         Height          =   180
         Index           =   31
         Left            =   -68220
         TabIndex        =   96
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手機："
         Height          =   180
         Index           =   30
         Left            =   -68220
         TabIndex        =   81
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "代表人(中文)："
         Height          =   255
         Index           =   29
         Left            =   -69780
         TabIndex        =   80
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   255
         Index           =   28
         Left            =   -74820
         TabIndex        =   79
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "申請人(中文)："
         Height          =   255
         Index           =   27
         Left            =   -74820
         TabIndex        =   78
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail："
         Height          =   255
         Index           =   26
         Left            =   -74820
         TabIndex        =   77
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   255
         Index           =   25
         Left            =   -69780
         TabIndex        =   76
         Top             =   2115
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "國   籍："
         Height          =   255
         Index           =   24
         Left            =   -70260
         TabIndex        =   75
         Top             =   1410
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID No.："
         Height          =   180
         Index           =   23
         Left            =   -74820
         TabIndex        =   74
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   16
         Left            =   -74820
         TabIndex        =   73
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   180
         Index           =   11
         Left            =   -72540
         TabIndex        =   72
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "電話："
         Height          =   180
         Index           =   10
         Left            =   -70260
         TabIndex        =   71
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "傳真："
         Height          =   180
         Index           =   9
         Left            =   -68220
         TabIndex        =   70
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   8
         Left            =   -70260
         TabIndex        =   69
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡地址："
         Height          =   255
         Index           =   7
         Left            =   -74820
         TabIndex        =   68
         Top             =   2490
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "申請地址："
         Height          =   255
         Index           =   6
         Left            =   -74820
         TabIndex        =   67
         Top             =   2790
         Width           =   1305
      End
      Begin MSForms.TextBox TxtCRA11 
         Height          =   300
         Index           =   6
         Left            =   1920
         TabIndex        =   21
         Top             =   2790
         Width           =   6700
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA11 
         Height          =   300
         Index           =   5
         Left            =   1920
         TabIndex        =   24
         Top             =   2460
         Width           =   6700
         VariousPropertyBits=   679495711
         Size            =   "11818;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA11 
         Height          =   300
         Index           =   3
         Left            =   6630
         TabIndex        =   52
         Top             =   1800
         Width           =   1995
         VariousPropertyBits=   679495711
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA11 
         Height          =   300
         Index           =   4
         Left            =   6630
         TabIndex        =   51
         Top             =   2130
         Width           =   1995
         VariousPropertyBits=   679495711
         Size            =   "3519;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA11 
         Height          =   300
         Index           =   2
         Left            =   1470
         TabIndex        =   50
         Top             =   2130
         Width           =   3615
         VariousPropertyBits=   679495711
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TxtCRA11 
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   49
         Top             =   1800
         Width           =   3615
         VariousPropertyBits=   679495711
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "申請地址："
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   25
         Top             =   2790
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "聯絡地址："
         Height          =   255
         Index           =   92
         Left            =   180
         TabIndex        =   39
         Top             =   2490
         Width           =   1000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   85
         Left            =   4740
         TabIndex        =   38
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手機："
         Height          =   180
         Index           =   18
         Left            =   6780
         TabIndex        =   37
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "傳真："
         Height          =   180
         Index           =   15
         Left            =   6780
         TabIndex        =   36
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "電話："
         Height          =   180
         Index           =   14
         Left            =   4740
         TabIndex        =   35
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   180
         Index           =   13
         Left            =   2460
         TabIndex        =   34
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   12
         Left            =   180
         TabIndex        =   33
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID No.："
         Height          =   180
         Index           =   77
         Left            =   180
         TabIndex        =   32
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "國   籍："
         Height          =   255
         Index           =   94
         Left            =   4740
         TabIndex        =   31
         Top             =   1410
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   255
         Index           =   22
         Left            =   5220
         TabIndex        =   30
         Top             =   2115
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail："
         Height          =   255
         Index           =   17
         Left            =   180
         TabIndex        =   29
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "申請人(中文)："
         Height          =   255
         Index           =   19
         Left            =   180
         TabIndex        =   28
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "　　　(英文)："
         Height          =   255
         Index           =   20
         Left            =   180
         TabIndex        =   27
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "代表人(中文)："
         Height          =   255
         Index           =   21
         Left            =   5220
         TabIndex        =   26
         Top             =   1830
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   83
      Left            =   2940
      TabIndex        =   215
      Top             =   480
      Width           =   900
   End
   Begin MSForms.TextBox txtF0316_N 
      Height          =   285
      Left            =   4500
      TabIndex        =   214
      Top             =   450
      Width           =   1300
      VariousPropertyBits=   679495711
      ScrollBars      =   3
      Size            =   "2293;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "您的意見："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   76
      Left            =   30
      TabIndex        =   192
      Top             =   1920
      Width           =   900
   End
   Begin MSForms.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   18
      Top             =   1110
      Width           =   7905
      VariousPropertyBits=   679495711
      ScrollBars      =   3
      Size            =   "13944;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   4
      Left            =   30
      TabIndex        =   19
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   3
      Left            =   2940
      TabIndex        =   17
      Top             =   810
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   30
      TabIndex        =   15
      Top             =   810
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "退回原因："
      Height          =   180
      Index           =   130
      Left            =   30
      TabIndex        =   13
      Top             =   1500
      Width           =   915
   End
   Begin MSForms.TextBox txtNote 
      Height          =   510
      Left            =   960
      TabIndex        =   0
      Top             =   1770
      Width           =   7905
      VariousPropertyBits=   -1466939365
      ScrollBars      =   3
      Size            =   "13944;900"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL78_N 
      Height          =   285
      Left            =   1590
      TabIndex        =   12
      Top             =   450
      Width           =   1300
      VariousPropertyBits=   679495711
      ScrollBars      =   3
      Size            =   "2293;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "目前表單狀態："
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   5940
      TabIndex        =   11
      Top             =   480
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "接洽單編號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "填表人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frm210148_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2022/08/31 Form2.0已修改 txtCrl78/txtCrl78_N/text3/txtNote/TxtCRA11~X1/TxtCRA1~5
Option Explicit

Public m_PrevForm As Form  '前一畫面
Dim RsQ As New ADODB.Recordset, oTxt, strArr
Dim strQ As String, intQ As Integer, i As Integer, m_F0309 As String, m_F0308 As String, strUpdDate As String, strUpdTime As String
Dim oCheck, IsSpec As Boolean 'Add by Amy 2022/10/26

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cboReason_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 'Add by Amy 2022/12/21 先設只能下拉
End Sub

Private Sub cmdBack_Click() '退智權人員
   Dim intItem As Integer, m_F0309 As String
   Dim stMsg As String, stContent As String, stSubject As String, stTO As String, stArr
   Dim stTemp As String, stTemp1 As String, stTemp2 As String, stTemp3 As String
  
   If FormCheck(2) = False Then
       Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   m_F0309 = Flow_退回
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   cnnConnection.BeginTrans
   
   '簽核檔-退回記錄
   strSql = "update FLOW002 set " & _
            "F0205='" & strUpdDate & "'" & _
            ",F0206='" & strUpdTime & "'" & _
            ",F0207='2',F0204='" & strUserNum & "'" & _
            " where F0201='" & txtF0301 & "' and F0202='A3' and F0207 is null "
     cnnConnection.Execute strSql
   
   '表單主檔
   'Modify by Amy 2022/10/19 原:F0307='" & strUserNum & "'
   'Modify by Amy 2022/10/27 +F0305
   strSql = ""
   If cboReason.Enabled = True Then
        strSql = "F0305='" & Left(cboReason, 2) & "',"
   End If
   strSql = "update FLOW003 set " & strSql & _
           "F0307='A7'" & _
           ",F0308='" & txtF0316 & "'" & _
           ",F0309='" & m_F0309 & "'" & _
           " where F0301='" & txtF0301 & "'"
   cnnConnection.Execute strSql
   
   '流程備註檔
   If Trim(txtNote.Text) <> MsgText(601) Then
       strSql = GetInsertFLOW004Sql(Trim(txtF0301), strUserNum, strUpdDate, strUpdTime, m_F0309, ChgSQL(Trim(txtNote.Text)))
       cnnConnection.Execute strSql
   End If
   
   'Add by Amy 2022/10/26 勾選特例回寫接洽單申請人檔
   If IsSpec = True Then
        For Each oCheck In ChkCRA26
        If oCheck.Enabled = True And oCheck.Value = 1 Then
                strSql = "Update ConsultRecApp Set cra26='Y' Where cra01='" & txtF0301 & "' And cra02='" & oCheck.Index + 1 & "' "
                cnnConnection.Execute strSql
            End If
        Next
   End If
   'end 2022/10/26
   cnnConnection.CommitTrans
   
   '發Mail通知當事人
   stContent = GetEMailContent_Flow(txtF0301, stSubject)
   If Trim(txtNote.Text) <> "" Then
       stSubject = stSubject & "；退回原因：" & Trim(txtNote.Text)
   End If
   'Modify By Sindy 2025/5/8 填單人員非智權人員時,退回的收受者掛填單人員,副本為下一處理人員(智權人員)
   If txtCRL78 <> txtF0316 Then
      PUB_SendMail strUserNum, txtCRL78, "", stSubject, stContent, , , , , , txtF0316
   Else
   '2025/5/8 END
      PUB_SendMail strUserNum, txtF0316, "", stSubject, stContent
   End If
   
   '申請人編號有值(已建)且為新客戶
   If FormCheck(3, stMsg) = False Then
       stContent = "": stTO = Pub_GetSpecMan("程式管理人員")
       'Modify by Amy 2023/02/14 修改主旨-秀玲
       stSubject = "櫃台已建新客戶，但接洽單被退回給智權人員，請追蹤接洽單後續，確認新客戶是否要保留！"
       stArr = Split(stMsg, ",")
       For i = LBound(stArr) To UBound(stArr)
           intItem = Right(stArr(i), 1)
           Call SetText(3, TxtCRA(intItem), stTemp1) '編號
           Select Case intItem
                Case 1
                    Call SetText(3, TxtCRA11(1), stTemp2) '中文名稱
                    Call SetText(3, TxtCRA11(2), stTemp3) '英文名稱
                Case 2
                    Call SetText(3, TxtCRA21(1), stTemp2) '中文名稱
                    Call SetText(3, TxtCRA21(2), stTemp3) '英文名稱
                Case 3
                    Call SetText(3, TxtCRA31(1), stTemp2) '中文名稱
                    Call SetText(3, TxtCRA31(2), stTemp3) '英文名稱
                Case 4
                    Call SetText(3, TxtCRA41(1), stTemp2) '中文名稱
                    Call SetText(3, TxtCRA41(2), stTemp3) '英文名稱
                Case 5
                    Call SetText(3, TxtCRA51(1), stTemp2) '中文名稱
                    Call SetText(3, TxtCRA51(2), stTemp3) '英文名稱
           End Select
           If stTemp1 <> MsgText(601) Then
                stContent = stContent & "接洽單編號：" & txtF0301 & vbCrLf & vbCrLf
                stContent = stContent & Text4(intItem - 1).Text & vbCrLf & _
                                "編　　號：" & stTemp1 & vbCrLf & _
                                "中文名稱：" & stTemp2 & vbCrLf & _
                                "英文名稱：" & stTemp3 & vbCrLf & vbCrLf
           End If
       Next i
       If stContent <> MsgText(601) Then
           PUB_SendMail strUserNum, stTO, "", stSubject, stContent
       End If
   End If
   Screen.MousePointer = vbDefault
   
   Me.txtF0301 = ""
   m_PrevForm.Show
   m_PrevForm.PubShowNextData
   If Me.txtF0301 = "" Then
       Unload Me
   End If
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox "退回失敗！" & vbCrLf & Err.Description
End Sub

Private Function FormCheck(ByVal intChoose As Integer, Optional ByRef stReplyMsg As String = "") As Boolean
    Dim strMsg As String
    Dim strSpecTxt As String 'Add by Amy 2022/10/26
    Dim strChkCusNo As String, intItem As String, stArr 'Add by Amy 2022/12/16
    
    FormCheck = False
'*** 要先檢查的 ***
    IsSpec = False
    '完成/退回智權人員
    If intChoose <> 3 Then
        For Each oCheck In ChkCRA26
            If oCheck.Enabled = True And oCheck.Value = 1 Then
                IsSpec = True
                strSpecTxt = strSpecTxt & "/" & Text4(oCheck.Index)
            End If
        Next
        '按完成鈕時,有勾選有對造不可存檔(櫃台勾)
        If intChoose = 1 And IsSpec = True Then
            MsgBox Replace(Mid(strSpecTxt, 2), "/", vbCrLf) & vbCrLf & "有勾選「有對造」需退智權人員！", vbExclamation
            Exit Function
        End If
    End If
    
'*** 完成/退回智權人員-Mail用 ***
    If intChoose = 1 Or intChoose = 3 Then
        For Each oTxt In TxtCRA
            If ChkNewCus(oTxt.Index) = True Then
                '新客戶且 未 建立新客戶資料(編號為空)
                If intChoose = 1 Then
                    'Modify by Amy 2022/12/26 +記錄 strChkCusNo
                    oTxt.Text = ChangeCustomerL(oTxt.Text) 'Add by Amy 2023/01/07 未輸滿9碼自動補上
                    If oTxt.Text = MsgText(601) Then
                        strMsg = strMsg & "," & Text4(oTxt.Index - 1).Text
                    ElseIf oTxt.Locked = False Then
                        strChkCusNo = strChkCusNo & "," & Text4(oTxt.Index - 1).Text
                    End If
                '新客戶且 已 建立新客戶資料(編號有值)
                'Modify by Amy 2022/12/16 + oTxt.Locked = true 客戶編號已回寫接洽單
                ElseIf intChoose = 3 And oTxt.Text <> MsgText(601) And oTxt.Locked = True Then
                    strMsg = strMsg & "," & Text4(oTxt.Index - 1).Text
                End If
            End If
        Next
        If strChkCusNo <> MsgText(601) Then strChkCusNo = Mid(strChkCusNo, 2) 'Add by Amy 2022/12/16
        If strMsg <> MsgText(601) Then
            '完成
            If intChoose = 1 Then
                MsgBox Replace(strMsg, ",", vbCrLf) & vbCrLf & "尚未建立客戶資料,不可存檔！", vbExclamation
            '退回智權人員-Mail用
            Else
                stReplyMsg = Mid(strMsg, 2)
            End If
            Exit Function
        End If
        'Add by Amy 2022/12/16 完成時,若客戶編號未鎖住(表自行輸入),需檢查資料是否與客戶檔一致
        If intChoose = 1 Then
            strMsg = ""
            stArr = Split(strChkCusNo, ",")
            For i = LBound(stArr) To UBound(stArr)
               intItem = Right(stArr(i), 1)
               strExc(1) = ""
               If ChkCusSame(intItem, strExc(1)) = False Then
                    strMsg = strMsg & ";" & stArr(i) & vbCrLf & strExc(1)
               End If
            Next i
            If strMsg <> MsgText(601) Then
                MsgBox Replace(Mid(strMsg, 2), ";", vbCrLf) & ",不可存檔！", vbExclamation
                Exit Function
            End If
        End If
        'end 2022/12/16
'*** 退回智權人員 ***
    ElseIf intChoose = 2 Then
        'Add By Sindy 2022/10/24
        If cboReason.Text = MsgText(601) Then
           MsgBox "退回原因不可空白！", vbExclamation
           cboReason.SetFocus
           Exit Function
        End If
        '2022/10/24 END
        'Add by Amy 2022/10/26
        'Modify by Amy 2022/11/02 櫃台勾選特例,原因只能是「案件對造」
        If IsSpec = False Then
            'Modify by Amy 2022/11/15 +原因未被鎖住(案件對造會鎖)
            If Left(cboReason, 2) = "01" And cboReason.Locked = False Then
                MsgBox "退回原因是「案件對造」需至少勾選一個特例！", vbExclamation
                Exit Function
            End If
        Else
            If Left(cboReason, 2) <> "01" Then
                MsgBox "勾選特例，退回原因只能是「案件對造」！", vbExclamation
                Exit Function
            End If
        End If
        'end 2022/11/02
        'end 2022/10/26
        
        If Trim(txtNote.Text) = MsgText(601) Then
            MsgBox "您的意見不可空白！", vbExclamation
            txtNote.SetFocus
            Exit Function
        End If
    End If
    
'*** 都要檢查的 或放於最後做 ***
    If PUB_CheckFormExist("frm140401") = True Then
        MsgBox "請先關閉「客戶基本資料維護」！", vbExclamation
        Exit Function
    End If
    
    If intChoose <> 3 Then
        'TextBox, ComboBox 是否含有Unicode文字
        If PUB_ChkUniText(Me) = False Then
            Exit Function
        End If
    End If
    
    '退回智權人員時,有勾選有對造(櫃台勾)意見寫入對造項次
    If intChoose = 2 And IsSpec = True Then
        txtNote = Mid(strSpecTxt, 2) & " 有對造;" & txtNote
    End If
    FormCheck = True
End Function

Private Sub cmdExit_Click()
    m_PrevForm.QueryData
    m_PrevForm.Show
    Unload Me
End Sub

Private Sub CmdAddCus_Click()
    Dim intItem As Integer, bolNewCus As Boolean
    Dim stNation As String 'Add by Amy 2022/12/16
    Dim stReceiptCmp As String 'Add by Amy 2024/11/04 收據公司別(出名公司)
   
    intItem = SSTab1.Tab + 1
    If ChkNewCus(intItem) = True Then
        bolNewCus = True
    End If
    '頁籤有資料且客戶編號為空
    If bolNewCus = True And TxtCRA(intItem) = MsgText(601) Then
        'Add by Amy 2022/12/16 接洽單國籍中文名有資料,帶編號至客戶檔
        Select Case intItem
            Case 1
                Call SetText(3, TxtCRA1(9), stNation)
            Case 2
                Call SetText(3, TxtCRA2(9), stNation)
            Case 3
                Call SetText(3, TxtCRA3(9), stNation)
            Case 4
                Call SetText(3, TxtCRA4(9), stNation)
            Case 5
                Call SetText(3, TxtCRA5(9), stNation)
        End Select
        If stNation <> MsgText(601) Then
            stNation = Pub_GetField("Nation", "na03='" & stNation & "'", "na01")
        End If
        'end 2022/12/16
        
        CmdAddCus.Tag = intItem - 1 '記錄傳至客戶檔的頁籤
        Call frm140401.SetParent(Me, "Add " & txtF0301 & "-" & intItem)
        'Add by Amy 2023/05/16 申請國家[非]台灣且收據公司別(出名公司)為J且操作 申請人1,收據公司需帶入客戶基本檔
        'Modify by Amy 2024/11/04 排除ACS系統別,不回寫,且改用共用函數 (有修改要確認cmdOK_Click是否也要修改)
        If Left(Text2, 3) <> 台灣國家代號 And txtCRL49 = "J" And intItem = 1 And SystemNumber(Text1, 1) <> "ACS" Then
            Call GetReceiptCmp("", "", SystemNumber(Text1, 1), Left(Text2, 3), , "", Me.Name, 1, stReceiptCmp)
            If stReceiptCmp <> MsgText(601) Then
               frm140401.m_Crl49JCmp = ",'" & stReceiptCmp & "' as Crl49JCmp "
            End If
        End If
        'end 2024/11/04
        frm140401.m_Cra04 = TxtCRA04(intItem) '關係企業中文
        frm140401.textCU10 = stNation 'Add by Amy 2022/12/16
        frm140401.Show
    End If
End Sub

Private Sub cmdFile_Click()
    frm090801_Q.SetParent Me
    frm090801_Q.m_blnCallPrint = True
    frm090801_Q.Text5 = txtF0301
    Call frm090801_Q.cmdok_Click(4)
    frm090801_Q.ZOrder 'Add By Sindy 2022/11/16
    frm090801_Q.Show
End Sub

Private Sub cmdok_Click()
   Dim stReceiptField As String 'Add by Amy 2024/11/04
On Error GoTo ErrHand
   
   If FormCheck(1) = False Then
       Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   cnnConnection.BeginTrans
   
   'Add by Amy 2022/12/19 先回寫客戶編號至接洽單檔
   For i = 1 To 5
        If TxtCRA(i).Locked = False And TxtCRA(i) <> MsgText(601) Then
            strSql = "Update ConsultRecApp Set cra05='" & Left(TxtCRA(i), 8) & "',cra06='" & Mid(TxtCRA(i), 9, 1) & "' " & _
                        " Where cra01='" & txtF0301 & "' And cra02='" & i & "' "
            cnnConnection.Execute strSql
            'Add by Amy 2024/11/04 申請人1同一天收文多筆,非ACS案且申請國家不同,且與目前客戶檔對應之收據公司別(出名公司)不同時,更新客戶檔對應收據公司別
            If Left(Text2, 3) <> 台灣國家代號 And i = 1 And SystemNumber(Text1, 1) <> "ACS" _
              And txtCRL49 <> GetReceiptCmp(Left(TxtCRA(i), 8), Mid(TxtCRA(i), 9, 1), SystemNumber(Text1, 1), Left(Text2, 3), True, "", Me.Name, 0, stReceiptField) Then
               strSql = "Update Customer Set " & stReceiptField & "=" & CNULL(ChgSQL(txtCRL49)) & _
                              " Where cu01='" & Left(TxtCRA(i), 8) & "' And cu02='" & Mid(TxtCRA(i), 9, 1) & "' "
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql
            End If
            'end 2024/11/04
        End If
   Next i
   
   '簽核檔
   strSql = "update FLOW002 set " & _
               "F0205='" & strUpdDate & "'" & _
               ",F0206='" & strUpdTime & "'" & _
               ",F0207='1',F0204='" & strUserNum & "'" & _
               " where F0201='" & txtF0301 & "' and F0202='A3' and F0207 is null"
   cnnConnection.Execute strSql
   
   '流程備註檔
   'Modify by Amy 2022/10/07 +if 有資料才寫入
   If Trim(txtNote.Text) <> MsgText(601) Then
        strSql = GetInsertFLOW004Sql(Trim(txtF0301), strUserNum, strUpdDate, strUpdTime, m_F0309, ChgSQL(Trim(txtNote.Text)))
        cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2022/10/4
   '讀取下一處理人員
   If GetNextProPerson_Flow(Trim(txtF0301), Trim(txtF0316), m_F0308, m_F0309) = False Then GoTo ErrHand
   '2022/10/4 END
   
   cnnConnection.CommitTrans
   PUB_SendMailCache '發信(因電子收文)
   Screen.MousePointer = vbDefault
   
   Me.txtF0301 = ""
   m_PrevForm.Show
   m_PrevForm.PubShowNextData
   If Me.txtF0301 = "" Then
      Unload Me
   End If
   Exit Sub
    
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox "存檔失敗！" & vbCrLf & Err.Description
End Sub

'Add by Amy 2022/11/15 申請人查詢
Private Sub CmdQueryCus_Click()
    Dim intItem As Integer, bolNewCus As Boolean
    Dim stValue As String
    
    intItem = SSTab1.Tab + 1
    If ChkNewCus(intItem) = True Then
        bolNewCus = True
    End If
    '頁籤有資料且客戶編號為空
    If bolNewCus = True And TxtCRA(intItem) = MsgText(601) Then
        If PUB_CheckFormExist("frm100102_1") = True Then
            Unload frm100102_1
        End If
        Select Case intItem
            Case 1
                stValue = TxtCRA11(1)
            Case 2
                stValue = TxtCRA21(1)
            Case 3
                stValue = TxtCRA31(1)
            Case 4
                stValue = TxtCRA41(1)
            Case 5
                stValue = TxtCRA51(1)
        End Select
        frm100102_1.IsSearchNew = True
        frm100102_1.Caption = "申請人查詢(查新客戶-含對造)"
        frm100102_1.Option2(1).Value = True
        frm100102_1.Text2 = stValue
        'Mark by Amy 2023/01/03 上線後櫃台說不先查,避免太慢
'        If stValue <> MsgText(601) Then
'            frm100102_1.cmdSearch_Click
'        End If
        frm100102_1.Show
    End If
End Sub

Private Sub cmdQueryNext_Click()
    Me.txtF0301 = ""
    m_PrevForm.Show
    m_PrevForm.PubShowNextData
    If Me.txtF0301 = "" Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
End Sub

Private Sub ReadReason()
    Me.cboReason.Clear
    Me.cboReason.AddItem ""
    strQ = "Select * From AllCode Where AC01='12' Order by ac02 "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Do While Not RsQ.EOF
            Me.cboReason.AddItem "" & RsQ("ac02") & "--" & RsQ("ac03")
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
End Sub

Public Sub QueryData()
    Dim RsQ2 As New ADODB.Recordset
    Dim strQ2 As String, intQ2 As Integer
    Dim intCRA02 As Integer, strVal As String
    Dim strNewCusIndex As String 'Add by Amy 2023/03/22
    
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    FormClear '清空欄位值
    ReadReason
    Call SetLock(True)
    
    'Modify By Sindy 2022/11/4 + ,crl59
    'Modify by Amy 2022/11/08 CRL59 改為CRL90
    'Modify by Amy 2023/01/06 +CRL60||CRL61
    'Modify by Amy 2023/05/16 +CRL49
    strQ = "Select F0301,crl78,C.st02 as CreateP,crl07||'-'||Decode(crl08,null,'',crl08||'-'||crl09||'-'||crl10) as CaseNo,crl15||' '||na03 as crl15,crl17,crl66,F0305,F0306,F0316" & _
            ",S.st02 as SalesN,F0308,F0309,Decode(F0309," & ShowFlow表單狀態中文 & ") as F0309NM,crl90,CRL60||CRL61 as AgNo,crl49 " & _
             "From Flow003,ConsultRecordList,AllCode,Staff C,Nation,Staff S " & _
             "Where F0301='" & txtF0301.Text & "' And F0301=CRL01(+) And AC01(+)='12' And F0305=AC02(+) And Crl78=C.st01(+) And F0316=S.st01(+) " & _
             "And crl15=na01(+) "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 0 Then
        ShowNoData
    Else
        txtCRL78 = "" & RsQ.Fields("crl78") '填表人
        txtCRL78_N = "" & RsQ.Fields("CreateP")
        Text1 = "" & RsQ.Fields("CaseNo")
        Text2 = "" & RsQ.Fields("crl15") '申請國家
        Text3 = "" & RsQ.Fields("crl17") '案件名稱(主題)
        txtAgent = "" & RsQ.Fields("AgNo") 'Add by Amy 2023/01/06 代理人
        txtCRL49 = "" & RsQ.Fields("crl49") 'Add by Amy 2023/05/16 收據公司別
        Call SetReason("" & RsQ.Fields("F0305")) '退回原因
        'Add by Amy 2022/10/27 原因不是對造,才可選
        If "" & RsQ.Fields("F0305") = "01" Then
            cboReason.BackColor = &H8000000F
        Else
            cboReason.Locked = False
        End If
        If "" & RsQ.Fields("crl66") = "Y" Then
            ChkCRL66.Value = 1 '已核准
            ChkCRL66.BackColor = &HC000& 'Add by Amy 2022/10/27 綠色
        Else
            ChkCRL66.BackColor = &H8000000F 'Add By Sindy 2022/11/4 灰色
        End If
        'Add By Sindy 2022/11/4
        'Modify by Amy 2022/11/08 原:CRL59
        If "" & RsQ.Fields("crl90") = "Y" Then
            Check11.Value = 1 '急件
            Check11.BackColor = &H8080FF '紅色
        Else
            Check11.BackColor = &H8000000F '灰色
        End If
        '2022/11/4 END
        txtF0316 = "" & RsQ.Fields("F0316") '智權編號
        txtF0316_N = "" & RsQ.Fields("SalesN")
        m_F0308 = "" & RsQ.Fields("F0308") '下一處理人員 Add By Sindy 2022/10/4
        m_F0309 = "" & RsQ.Fields("F0309") '目前狀態 Add By Sindy 2022/10/4
        txtF0309 = m_F0309 & RsQ.Fields("F0309NM") '畫面上狀態
        
'*** 申請人資料 ***
        strQ2 = "Select cra05||cra06 as CuNo,cra04,cra10,cra13,cra15,cra12,cra14,cra18,cra17,cra11,cra22,cra19" & _
                        ",cra07,cra08,cra09,cra24,cra23,cra20,cra21,cra26,cra02,cra03 From ConsultRecApp " & _
                    "Where CRA01='" & txtF0301.Text & "' Order by CRA02 "
        intQ2 = 1
        Set RsQ2 = ClsLawReadRstMsg(intQ2, strQ2)
        If intQ2 = 1 Then
            Do While Not RsQ2.EOF
                intCRA02 = Val("" & RsQ2.Fields("cra02")) '項目
                
                For i = 0 To RsQ2.Fields.Count - 1
                    strVal = "" & RsQ2.Fields(i)
                    '項目
                    If UCase(RsQ2.Fields(i).Name) = UCase("cra02") Then
                        Text4(intCRA02 - 1).Visible = True
                    '申請人編號
                    ElseIf UCase(RsQ2.Fields(i).Name) = UCase("CuNo") Then
                        TxtCRA(intCRA02) = strVal
                    '有對造
                    ElseIf UCase(RsQ2.Fields(i).Name) = UCase("cra26") Then
                        If "" & RsQ2.Fields(i) = "Y" Then
                            ChkCRA26(intCRA02 - 1).Value = vbChecked
                        End If
                        'Add by Amy 2022/10/26 舊客戶 或 新客戶且有對造,鎖住
                        If "" & RsQ2.Fields("cra03") <> "Y" Or "" & RsQ2.Fields(i) = "Y" Then
                            ChkCRA26(intCRA02 - 1).Enabled = False
                        End If
                    '是否新客戶
                    ElseIf UCase(RsQ2.Fields(i).Name) = UCase("cra03") Then
                        Call SetNewCusOpt(intCRA02, "" & RsQ2.Fields("cra03"))
                        'Add by Amy 2023/03/22 新客戶且未設定第一筆新客戶index
                        If "" & RsQ2.Fields("cra03") = "Y" And strNewCusIndex = MsgText(601) Then
                            strNewCusIndex = intCRA02
                        End If
                    '關係企業中文
                    ElseIf UCase(RsQ2.Fields(i).Name) = UCase("cra04") Then
                        TxtCRA04(intCRA02) = strVal
                    ElseIf i >= 2 And i <= 11 Then
                        Select Case intCRA02
                            Case 1
                                Call SetText(1, TxtCRA1(i), strVal)
                            Case 2
                                Call SetText(1, TxtCRA2(i), strVal)
                            Case 3
                                Call SetText(1, TxtCRA3(i), strVal)
                            Case 4
                                Call SetText(1, TxtCRA4(i), strVal)
                            Case 5
                                Call SetText(1, TxtCRA5(i), strVal)
                        End Select
                    '可能有UniCode欄位
                    ElseIf i > 11 Then
                        Select Case intCRA02
                            Case 1
                                Call SetText(1, TxtCRA11(i - 11), strVal)
                            Case 2
                                Call SetText(1, TxtCRA21(i - 11), strVal)
                            Case 3
                                Call SetText(1, TxtCRA31(i - 11), strVal)
                            Case 4
                                Call SetText(1, TxtCRA41(i - 11), strVal)
                            Case 5
                                Call SetText(1, TxtCRA51(i - 11), strVal)
                        End Select
                    End If
                Next i
                RsQ2.MoveNext
            Loop
        End If
        Set RsQ2 = Nothing
'*** End 申請人資料 ***
    End If
    Set RsQ = Nothing
    
    SSTab1.Tab = Val(strNewCusIndex) - 1 'Add by Amy 2023/03/22 預設目前尚未建立之新客戶頁籤
    Screen.MousePointer = vbDefault
    Me.Enabled = True
End Sub

Private Sub FormClear()
    txtCRL78.Text = Empty
    txtCRL78_N.Text = Empty
    Text1.Text = Empty '本所案號
    Text2.Text = Empty '申請國家
    Text3.Text = Empty '案件名稱
    txtAgent.Text = Empty '代理人
    cboReason.Clear '退回原因
    txtNote.Text = Empty '您的意見
        
    For Each oTxt In TxtCRA
        oTxt.Text = Empty
    Next
    
    For Each oTxt In TxtCRA1
        oTxt.Text = Empty
    Next
  
    For Each oTxt In TxtCRA11
        oTxt.Text = Empty
    Next
    
    For Each oTxt In TxtCRA2
        oTxt.Text = Empty
    Next
    
    For Each oTxt In TxtCRA21
        oTxt.Text = Empty
    Next
    
    For Each oTxt In TxtCRA3
        oTxt.Text = Empty
    Next
    
    For Each oTxt In TxtCRA31
        oTxt.Text = Empty
    Next
    For Each oTxt In TxtCRA4
        oTxt.Text = Empty
    Next
    
    For Each oTxt In TxtCRA41
        oTxt.Text = Empty
    Next
    
    For Each oTxt In TxtCRA5
        oTxt.Text = Empty
    Next
    
    For Each oTxt In TxtCRA51
        oTxt.Text = Empty
    Next
    
End Sub

Private Sub SetReason(ByVal stItem As String)
    Dim ii As Integer
    For ii = 0 To Me.cboReason.ListCount - 1
        If Left(Me.cboReason.List(ii), 2) = stItem Then
            Me.cboReason = Me.cboReason.List(ii)
            Exit For
        End If
    Next ii
End Sub

Private Sub SetLock(bolYN As Boolean)
    '退回原因
    cboReason.Locked = bolYN
    
    '已簽准
    ChkCRL66.Enabled = Not bolYN
    'Add By Sindy 2022/11/4
    '急件
    Check11.Enabled = Not bolYN
    '2022/11/4 END
    
    '編號
    For Each oTxt In TxtCRA
        oTxt.Locked = bolYN
    Next
    
    '申請人1
    For Each oTxt In TxtCRA1
        oTxt.Locked = bolYN
    Next
    For Each oTxt In TxtCRA11
        oTxt.Locked = bolYN
    Next
    
    '申請人2
    For Each oTxt In TxtCRA2
        oTxt.Locked = bolYN
    Next
    For Each oTxt In TxtCRA21
        oTxt.Locked = bolYN
    Next
    
    '申請人3
    For Each oTxt In TxtCRA3
        oTxt.Locked = bolYN
    Next
    For Each oTxt In TxtCRA31
        oTxt.Locked = bolYN
    Next
    
    '申請人4
    For Each oTxt In TxtCRA4
        oTxt.Locked = bolYN
    Next
    For Each oTxt In TxtCRA41
        oTxt.Locked = bolYN
    Next
    
    '申請人5
    For Each oTxt In TxtCRA5
        oTxt.Locked = bolYN
    Next
    For Each oTxt In TxtCRA51
        oTxt.Locked = bolYN
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add by Amy 2022/11/15 離開時,有開啟的表單都要關閉
    If PUB_CheckFormExist("frm140401") = True Then
        Unload frm140401
    End If
    If PUB_CheckFormExist("frm100102_1") = True Then
        Unload frm100102_1
    End If
    If PUB_CheckFormExist("frm090801_Q") = True Then
        Unload frm090801_Q
    End If
    'end 2022/11/15
    Set m_PrevForm = Nothing
    Set frm210148_2 = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If PUB_CheckFormExist("frm140401") = True Then
        SSTab1.Tab = CmdAddCus.Tag
        Exit Sub
    End If
End Sub

'頁籤-申請人x字樣
Private Sub Text4_Click(Index As Integer)
    If PUB_CheckFormExist("frm140401") = True Then
        Index = CmdAddCus.Tag
    End If
    SSTab1.Tab = Index
End Sub

Private Sub SetNewCusOpt(ByVal intItem As Integer, ByVal stNewCus As String)
    Select Case intItem
        Case 1
            Frame1.Enabled = False
            If stNewCus = "Y" Then
                Option1(0).Value = True
            Else
                Option1(1).Value = True
            End If
        Case 2
            Frame2.Enabled = False
            If stNewCus = "Y" Then
                Option2(0).Value = True
            Else
                Option2(1).Value = True
            End If
        Case 3
            Frame3.Enabled = False
            If stNewCus = "Y" Then
                Option3(0).Value = True
            Else
                Option3(1).Value = True
            End If
        Case 4
            Frame4.Enabled = False
            If stNewCus = "Y" Then
                Option4(0).Value = True
            Else
                Option4(1).Value = True
            End If
        Case 5
            Frame5.Enabled = False
            If stNewCus = "Y" Then
                Option5(0).Value = True
            Else
                Option5(1).Value = True
            End If
    End Select
    'Modify by Amy 2022/12/16 +編號欄為空(新客戶多案,客戶檔已建,其他多案要可輸編號)
    If stNewCus = "Y" And TxtCRA(intItem) = MsgText(601) Then
        Call SetText(2, TxtCRA(intItem))
    End If
End Sub

Private Function ChkNewCus(ByVal intItem As Integer) As Boolean
    ChkNewCus = False
    Select Case intItem
        Case 1
            If Option1(0).Value = True Then
                ChkNewCus = True
            End If
        Case 2
            If Option2(0).Value = True Then
                ChkNewCus = True
            End If
        Case 3
            If Option3(0).Value = True Then
                ChkNewCus = True
            End If
        Case 4
            If Option4(0).Value = True Then
                ChkNewCus = True
            End If
        Case 5
            If Option5(0).Value = True Then
                ChkNewCus = True
            End If
    End Select
End Function

Private Sub SetText(intChoose As Integer, stIndex As Object, Optional ByRef stVal As String)
    Select Case intChoose
        Case 1 '設定值
            stIndex = stVal
        Case 2 '不鎖住
            stIndex.Locked = False
        Case 3 '取值
            stVal = stIndex
    End Select
End Sub

Private Sub TxtCRA_GotFocus(Index As Integer)
    TextInverse TxtCRA(Index)
End Sub

'Add by Amy 2022/12/15
Private Sub TxtCRA_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'同一新客戶多新案,已建新客戶要能輸客戶編號
Private Sub TxtCRA_Validate(Index As Integer, Cancel As Boolean)
    Dim stMsg As String
    
    If TxtCRA(Index) = MsgText(601) Then Exit Sub
    
    TxtCRA(Index) = TxtCRA(Index) & String(9 - Len(TxtCRA(Index)), "0")
    If TxtCRA(Index).Locked = False Then
        If ChkCusSame(Index, stMsg) = False Then
            Cancel = True
            MsgBox stMsg, vbExclamation
            TxtCRA(Index).SetFocus
            TxtCRA_GotFocus (Index)
        End If
    End If
    
End Sub

'判斷畫面上資料是否與客戶檔資料(目前檢查 畫面客戶編號、中/英文名稱、ID No要與客戶檔一致)
Private Function ChkCusSame(ByVal Index As Integer, ByRef stMsg As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    Dim stNo As String, stCName As String, stEName As String, stID As String
    Dim stTP(1) As String 'Add by Amy 2023/07/04
    
    ChkCusSame = False: stMsg = ""
    stNo = TxtCRA(Index)
    Select Case Index
        Case 1
            stID = TxtCRA1(5)
            stCName = TxtCRA11(1)
            stEName = TxtCRA11(2)
        Case 2
            stID = TxtCRA2(5)
            stCName = TxtCRA21(1)
            stEName = TxtCRA21(2)
        Case 3
            stID = TxtCRA3(5)
            stCName = TxtCRA31(1)
            stEName = TxtCRA31(2)
        Case 4
            stID = TxtCRA4(5)
            stCName = TxtCRA41(1)
            stEName = TxtCRA41(2)
        Case 5
            stID = TxtCRA5(5)
            stCName = TxtCRA51(1)
            stEName = TxtCRA51(2)
    End Select
    strQ = "Select * From Customer Where cu01='" & Left(stNo, 8) & "' And cu02='" & Mid(stNo, 9, 1) & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        'Modify by Amy 2023/01/06 MCT收文時,建新客戶先存接洽單智權人員,當按下確認鈕,判斷FC代理人有管控智權人員(fa120),更新智權人員為MCTF0X編號
        '                          導致第二筆收文檢查時,出現智權人員不一致無法存檔
        If txtF0316 <> "" & RsQ.Fields("cu13") Then
            stMsg = stMsg & ";智權人員"
            '有代理人且客戶智權人員為MCTF開頭者,檢查接洽單智權人員為MCTF0X人員者,可存檔
            If txtAgent <> MsgText(601) And Left("" & RsQ.Fields("cu13"), 4) = "MCTF" Then
                If ChkMCTF0XSales(RsQ.Fields("cu13"), txtF0316) = True Then
                    stMsg = Replace(stMsg, ";智權人員", "")
                End If
            End If
        End If
        'end 2023/01/06
        If "" & RsQ.Fields("cu11") <> stID Then stMsg = stMsg & ";ID No."
        If "" & RsQ.Fields("cu04") <> stCName Then stMsg = stMsg & ";中文名稱"
        'Modify by Amy 2023/01/31 原Trim 若英文切字第二個字(CO., LTD.)與第一個字(CLOUDRICHES DIGITAL TECHNOLOGY)前不會存空白,接洽單CLOUDRICHES DIGITAL TECHNOLOGY CO., LTD.
        'Modify by Amy 2023/06/27  改抓ReplaceSign DB函數
        'Modify by Amy 2023/07/04 原使用strExc(1)及strExc(2),第2筆X87667000 傳至strExc(1)=Pub_GetField("..."),stMsg也會回傳資料
        stTP(0) = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL("" & RsQ.Fields("cu05") & RsQ.Fields("cu88") & RsQ.Fields("cu89") & RsQ.Fields("cu90")) & "')))")
        stTP(1) = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(stEName) & "')))")
        'If Pub_ReplaceSign(False, "" & RsQ.Fields("cu05") & RsQ.Fields("cu88") & RsQ.Fields("cu89") & RsQ.Fields("cu90")) <> Pub_ReplaceSign(False, stEName) Then
        If stTP(0) <> stTP(1) Then
        'end 2023/06/27
        'end 2023/07/04
            stMsg = stMsg & ";英文名稱"
        End If
        
        If stMsg = MsgText(601) Then
            ChkCusSame = True
        Else
            stMsg = Replace(Mid(stMsg, 2), ";", "、") & vbCrLf & "與客戶檔資料不一致"
        End If
    Else
        stMsg = "無此客戶,請確認！"
    End If
    Set RsQ = Nothing
End Function


