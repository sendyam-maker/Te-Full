VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060102 
   BorderStyle     =   1  '單線固定
   Caption         =   "新案建檔"
   ClientHeight    =   8064
   ClientLeft      =   828
   ClientTop       =   972
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8064
   ScaleWidth      =   8880
   Begin VB.CommandButton CmdPA174 
      BackColor       =   &H00C0FFFF&
      Caption         =   "特殊字"
      Height          =   280
      Left            =   180
      Style           =   1  '圖片外觀
      TabIndex        =   294
      Top             =   1410
      Width           =   840
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "電子送件暫存區"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   276
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "外文本"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   275
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "一案兩請資料(&D)"
      Height          =   400
      Index           =   6
      Left            =   4140
      TabIndex        =   46
      Top             =   60
      Width           =   1530
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申5"
      Height          =   375
      Index           =   8
      Left            =   3630
      TabIndex        =   45
      Top             =   60
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申4"
      Height          =   375
      Index           =   7
      Left            =   3180
      TabIndex        =   44
      Top             =   60
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申3"
      Height          =   375
      Index           =   6
      Left            =   2730
      TabIndex        =   43
      Top             =   60
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申2"
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   42
      Top             =   60
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "優先權資料"
      Height          =   405
      Index           =   4
      Left            =   5760
      TabIndex        =   47
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人參考資料(&A)"
      Height          =   375
      Index           =   3
      Left            =   6915
      TabIndex        =   41
      Top             =   510
      Width           =   1710
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人參考資料(&C)"
      Height          =   375
      Index           =   1
      Left            =   5100
      TabIndex        =   40
      Top             =   510
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3504
      TabIndex        =   4
      Top             =   510
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7800
      TabIndex        =   49
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   6915
      TabIndex        =   48
      Top             =   60
      Width           =   840
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   3
      Top             =   570
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   2
      Top             =   570
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   1
      Top             =   570
      Width           =   855
   End
   Begin VB.TextBox text1 
      Height          =   270
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
      Top             =   570
      Width           =   495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6195
      Left            =   120
      TabIndex        =   127
      Top             =   1830
      Width           =   8565
      _ExtentX        =   15113
      _ExtentY        =   10922
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "案件資料"
      TabPicture(0)   =   "frm060102.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label13"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label15"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label19"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label23"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label25"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label17"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label22"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label26(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblPA(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblPA(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(166)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(165)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(164)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label29"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label26(5)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label33"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text19"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text10"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label35"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label37"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(16)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text9"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text11"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text12"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text13"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text14"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text15"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text16"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text17"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text18"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text8"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text24"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Combo3"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtPA(151)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtPA(152)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtPA(155)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtPA(154)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtPA(153)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text25"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cmdPrtContact"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Frame1"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Combo5"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtPA(63)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "ChkAddTct"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtPA(156)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "CmdINST"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "ChkPA175"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtPA(178)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtCP142"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "cmdQueryTCT"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).ControlCount=   57
      TabCaption(1)   =   "申請人/代理人"
      TabPicture(1)   =   "frm060102.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPA(159)"
      Tab(1).Control(1)=   "Text23"
      Tab(1).Control(2)=   "Text21"
      Tab(1).Control(3)=   "Text30"
      Tab(1).Control(4)=   "Text20"
      Tab(1).Control(5)=   "Text32"
      Tab(1).Control(6)=   "Text29"
      Tab(1).Control(7)=   "Text28"
      Tab(1).Control(8)=   "Text27"
      Tab(1).Control(9)=   "Text26"
      Tab(1).Control(10)=   "Text22"
      Tab(1).Control(11)=   "Text31"
      Tab(1).Control(12)=   "Text33(10)"
      Tab(1).Control(13)=   "Text33(9)"
      Tab(1).Control(14)=   "Text33(14)"
      Tab(1).Control(15)=   "Text33(13)"
      Tab(1).Control(16)=   "Text33(12)"
      Tab(1).Control(17)=   "Text33(11)"
      Tab(1).Control(18)=   "Label27(12)"
      Tab(1).Control(19)=   "Label27(11)"
      Tab(1).Control(20)=   "Label27(8)"
      Tab(1).Control(21)=   "Label27(7)"
      Tab(1).Control(22)=   "Label27(6)"
      Tab(1).Control(23)=   "Label27(5)"
      Tab(1).Control(24)=   "Label27(4)"
      Tab(1).Control(25)=   "Label27(3)"
      Tab(1).Control(26)=   "Label27(2)"
      Tab(1).Control(27)=   "Label27(1)"
      Tab(1).Control(28)=   "Label27(0)"
      Tab(1).Control(29)=   "Label24"
      Tab(1).Control(30)=   "Label20"
      Tab(1).Control(31)=   "Label14(0)"
      Tab(1).Control(32)=   "Label12"
      Tab(1).Control(33)=   "Label10"
      Tab(1).Control(34)=   "Label47"
      Tab(1).Control(35)=   "Label46"
      Tab(1).Control(36)=   "Label44"
      Tab(1).Control(37)=   "Label43"
      Tab(1).Control(38)=   "Label42"
      Tab(1).Control(39)=   "Label41"
      Tab(1).Control(40)=   "Label39"
      Tab(1).Control(41)=   "Label38"
      Tab(1).Control(42)=   "Label36"
      Tab(1).Control(43)=   "Label34"
      Tab(1).Control(44)=   "Label32"
      Tab(1).Control(45)=   "Label30"
      Tab(1).Control(46)=   "Label28"
      Tab(1).Control(47)=   "Label26(0)"
      Tab(1).ControlCount=   48
      TabCaption(2)   =   "聯絡人"
      TabPicture(2)   =   "frm060102.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label49"
      Tab(2).Control(1)=   "Label51"
      Tab(2).Control(2)=   "Label53"
      Tab(2).Control(3)=   "Label55"
      Tab(2).Control(4)=   "Label57"
      Tab(2).Control(5)=   "Label59"
      Tab(2).Control(6)=   "Label61"
      Tab(2).Control(7)=   "Label63"
      Tab(2).Control(8)=   "Label65"
      Tab(2).Control(9)=   "Label16"
      Tab(2).Control(10)=   "Text33(2)"
      Tab(2).Control(11)=   "Text33(3)"
      Tab(2).Control(12)=   "Text33(4)"
      Tab(2).Control(13)=   "Text33(5)"
      Tab(2).Control(14)=   "Text33(6)"
      Tab(2).Control(15)=   "Text33(7)"
      Tab(2).Control(16)=   "Text33(8)"
      Tab(2).Control(17)=   "Text33(0)"
      Tab(2).Control(18)=   "Text33(1)"
      Tab(2).Control(19)=   "Text33(15)"
      Tab(2).Control(20)=   "Combo1(1)"
      Tab(2).Control(21)=   "Combo1(0)"
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "副本資料"
      TabPicture(3)   =   "frm060102.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label67"
      Tab(3).Control(1)=   "Label68"
      Tab(3).Control(2)=   "Label69"
      Tab(3).Control(3)=   "Label71"
      Tab(3).Control(4)=   "Label72"
      Tab(3).Control(5)=   "Label73"
      Tab(3).Control(6)=   "Label75"
      Tab(3).Control(7)=   "Label76"
      Tab(3).Control(8)=   "Label27(9)"
      Tab(3).Control(9)=   "Label27(10)"
      Tab(3).Control(10)=   "Text47"
      Tab(3).Control(11)=   "Text46"
      Tab(3).Control(12)=   "Text43"
      Tab(3).Control(13)=   "Text45"
      Tab(3).Control(14)=   "Text42"
      Tab(3).Control(15)=   "Text44"
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "代表人１"
      TabPicture(4)   =   "frm060102.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label5(8)"
      Tab(4).Control(1)=   "Label5(7)"
      Tab(4).Control(2)=   "Label5(6)"
      Tab(4).Control(3)=   "Label5(5)"
      Tab(4).Control(4)=   "Label5(4)"
      Tab(4).Control(5)=   "Label5(3)"
      Tab(4).Control(6)=   "Label14(1)"
      Tab(4).Control(7)=   "Label18(2)"
      Tab(4).Control(8)=   "Label5(24)"
      Tab(4).Control(9)=   "Label5(25)"
      Tab(4).Control(10)=   "Label5(26)"
      Tab(4).Control(11)=   "Label5(27)"
      Tab(4).Control(12)=   "Label5(28)"
      Tab(4).Control(13)=   "Label5(29)"
      Tab(4).Control(14)=   "Label14(2)"
      Tab(4).Control(15)=   "Label18(1)"
      Tab(4).Control(16)=   "Label5(33)"
      Tab(4).Control(17)=   "Label5(34)"
      Tab(4).Control(18)=   "Label5(35)"
      Tab(4).Control(19)=   "Label14(3)"
      Tab(4).Control(20)=   "txtCaseField(53)"
      Tab(4).Control(21)=   "txtCaseField(52)"
      Tab(4).Control(22)=   "txtCaseField(51)"
      Tab(4).Control(23)=   "txtCaseField(50)"
      Tab(4).Control(24)=   "txtCaseField(49)"
      Tab(4).Control(25)=   "txtCaseField(48)"
      Tab(4).Control(26)=   "txtCaseField(47)"
      Tab(4).Control(27)=   "txtCaseField(46)"
      Tab(4).Control(28)=   "txtCaseField(45)"
      Tab(4).Control(29)=   "txtCaseField(44)"
      Tab(4).Control(30)=   "txtCaseField(43)"
      Tab(4).Control(31)=   "txtCaseField(42)"
      Tab(4).Control(32)=   "txtCaseField(41)"
      Tab(4).Control(33)=   "txtCaseField(40)"
      Tab(4).Control(34)=   "txtCaseField(39)"
      Tab(4).Control(35)=   "Combo2(4)"
      Tab(4).Control(36)=   "Combo2(3)"
      Tab(4).Control(37)=   "Combo2(2)"
      Tab(4).Control(38)=   "Combo2(1)"
      Tab(4).Control(39)=   "Combo2(0)"
      Tab(4).ControlCount=   40
      TabCaption(5)   =   "代表人２"
      TabPicture(5)   =   "frm060102.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label5(21)"
      Tab(5).Control(1)=   "Label5(20)"
      Tab(5).Control(2)=   "Label5(19)"
      Tab(5).Control(3)=   "Label5(18)"
      Tab(5).Control(4)=   "Label5(17)"
      Tab(5).Control(5)=   "Label5(16)"
      Tab(5).Control(6)=   "Label14(6)"
      Tab(5).Control(7)=   "Label18(4)"
      Tab(5).Control(8)=   "Label5(15)"
      Tab(5).Control(9)=   "Label5(14)"
      Tab(5).Control(10)=   "Label5(13)"
      Tab(5).Control(11)=   "Label5(12)"
      Tab(5).Control(12)=   "Label5(11)"
      Tab(5).Control(13)=   "Label5(10)"
      Tab(5).Control(14)=   "Label14(5)"
      Tab(5).Control(15)=   "Label18(3)"
      Tab(5).Control(16)=   "Label5(9)"
      Tab(5).Control(17)=   "Label5(2)"
      Tab(5).Control(18)=   "Label5(1)"
      Tab(5).Control(19)=   "Label14(4)"
      Tab(5).Control(20)=   "txtCaseField(68)"
      Tab(5).Control(21)=   "txtCaseField(67)"
      Tab(5).Control(22)=   "txtCaseField(66)"
      Tab(5).Control(23)=   "txtCaseField(65)"
      Tab(5).Control(24)=   "txtCaseField(64)"
      Tab(5).Control(25)=   "txtCaseField(63)"
      Tab(5).Control(26)=   "txtCaseField(62)"
      Tab(5).Control(27)=   "txtCaseField(61)"
      Tab(5).Control(28)=   "txtCaseField(60)"
      Tab(5).Control(29)=   "txtCaseField(59)"
      Tab(5).Control(30)=   "txtCaseField(58)"
      Tab(5).Control(31)=   "txtCaseField(57)"
      Tab(5).Control(32)=   "txtCaseField(56)"
      Tab(5).Control(33)=   "txtCaseField(55)"
      Tab(5).Control(34)=   "txtCaseField(54)"
      Tab(5).Control(35)=   "Combo2(9)"
      Tab(5).Control(36)=   "Combo2(8)"
      Tab(5).Control(37)=   "Combo2(7)"
      Tab(5).Control(38)=   "Combo2(6)"
      Tab(5).Control(39)=   "Combo2(5)"
      Tab(5).ControlCount=   40
      TabCaption(6)   =   "發明人"
      TabPicture(6)   =   "frm060102.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame3"
      Tab(6).Control(1)=   "Frame2"
      Tab(6).Control(2)=   "cmdUpdRow"
      Tab(6).Control(3)=   "cmdDelRow"
      Tab(6).Control(4)=   "cmdAddRow"
      Tab(6).Control(5)=   "txtIN11"
      Tab(6).Control(6)=   "Combo4"
      Tab(6).Control(7)=   "GRD1"
      Tab(6).Control(8)=   "GRDtmp"
      Tab(6).Control(9)=   "Lb_IN11N"
      Tab(6).Control(10)=   "Lb_IN11"
      Tab(6).Control(11)=   "Lb_Inv(3)"
      Tab(6).Control(12)=   "Lb_Inv(2)"
      Tab(6).Control(13)=   "Lb_Inv(1)"
      Tab(6).Control(14)=   "Lb_Inv(0)"
      Tab(6).ControlCount=   15
      TabCaption(7)   =   "翻譯"
      TabPicture(7)   =   "frm060102.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraTrans03"
      Tab(7).Control(1)=   "fraTrans02"
      Tab(7).Control(2)=   "fraTrans01"
      Tab(7).ControlCount=   3
      Begin VB.CommandButton cmdQueryTCT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "命名作業查詢"
         Height          =   280
         Left            =   7056
         Style           =   1  '圖片外觀
         TabIndex        =   319
         Top             =   0
         Width           =   1500
      End
      Begin VB.TextBox txtCP142 
         Height          =   270
         Left            =   5250
         MaxLength       =   7
         TabIndex        =   10
         Top             =   345
         Width           =   975
      End
      Begin VB.TextBox txtPA 
         Height          =   270
         Index           =   178
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   28
         Top             =   2970
         Width           =   495
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   855
         Left            =   -73920
         TabIndex        =   313
         Top             =   780
         Width           =   6225
         Begin MSForms.TextBox txtInvField 
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   316
            Top             =   570
            Width           =   6135
            VariousPropertyBits=   671105051
            BackColor       =   -2147483644
            MaxLength       =   40
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtInvField 
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   315
            Top             =   285
            Width           =   6135
            VariousPropertyBits=   671105051
            BackColor       =   -2147483644
            MaxLength       =   70
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtInvField 
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   314
            Top             =   0
            Width           =   6135
            VariousPropertyBits=   671105051
            BackColor       =   -2147483644
            MaxLength       =   70
            Size            =   "7223;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.CheckBox ChkPA175 
         Caption         =   "有序列表"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   3900
         TabIndex        =   37
         Top             =   4485
         Width           =   1035
      End
      Begin VB.CommandButton CmdINST 
         BackColor       =   &H00FFFFC0&
         Caption         =   "各項指示"
         Height          =   280
         Left            =   7500
         Style           =   1  '圖片外觀
         TabIndex        =   16
         Top             =   930
         Width           =   960
      End
      Begin VB.TextBox txtPA 
         Height          =   270
         Index           =   156
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   31
         Top             =   3540
         Width           =   495
      End
      Begin VB.CheckBox ChkAddTct 
         Caption         =   "重新產生命名記錄"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5940
         TabIndex        =   38
         Top             =   4485
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtPA 
         Height          =   270
         Index           =   63
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   36
         Top             =   4425
         Width           =   495
      End
      Begin VB.Frame fraTrans03 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   615
         Left            =   -74880
         TabIndex        =   281
         Top             =   2320
         Width           =   2655
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   24
            Left            =   1200
            TabIndex        =   201
            Top             =   0
            Width           =   495
            VariousPropertyBits=   671105051
            MaxLength       =   4
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   25
            Left            =   1200
            TabIndex        =   202
            Top             =   328
            Width           =   495
            VariousPropertyBits=   671105051
            MaxLength       =   4
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label26 
            Caption         =   "外文本說明書:              頁"
            Height          =   165
            Index           =   6
            Left            =   0
            TabIndex        =   283
            Top             =   53
            Width           =   2235
         End
         Begin VB.Label Label26 
            Caption         =   "外文本圖示:                  頁"
            Height          =   165
            Index           =   7
            Left            =   0
            TabIndex        =   282
            Top             =   381
            Width           =   2115
         End
      End
      Begin VB.Frame fraTrans02 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   330
         Left            =   -74880
         TabIndex        =   263
         Top             =   440
         Width           =   8175
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   23
            Left            =   1200
            TabIndex        =   287
            Top             =   -8
            Width           =   855
            VariousPropertyBits=   671105051
            MaxLength       =   6
            Size            =   "1508;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   19
            Left            =   3570
            TabIndex        =   190
            Top             =   -8
            Width           =   495
            VariousPropertyBits=   671105051
            MaxLength       =   3
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   20
            Left            =   5640
            TabIndex        =   191
            Top             =   -8
            Width           =   1455
            VariousPropertyBits=   671105051
            MaxLength       =   12
            Size            =   "2566;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label26 
            Caption         =   "原文字數:"
            Height          =   165
            Index           =   2
            Left            =   0
            TabIndex        =   289
            Top             =   45
            Width           =   795
         End
         Begin VB.Label Label26 
            Caption         =   "相似度:            %"
            Height          =   165
            Index           =   3
            Left            =   2970
            TabIndex        =   278
            Top             =   45
            Width           =   1395
         End
         Begin VB.Label Label26 
            Caption         =   "相似案號:"
            Height          =   165
            Index           =   4
            Left            =   4800
            TabIndex        =   277
            Top             =   45
            Width           =   795
         End
      End
      Begin VB.Frame fraTrans01 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame4"
         Height          =   4755
         Left            =   -74880
         TabIndex        =   264
         Top             =   400
         Width           =   8175
         Begin VB.Frame fraTrans04 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame3"
            Height          =   690
            Left            =   60
            TabIndex        =   290
            Top             =   3780
            Width           =   7935
            Begin VB.ComboBox Combo6 
               Height          =   300
               Left            =   1170
               TabIndex        =   209
               Text            =   "Combo6"
               Top             =   330
               Width           =   3000
            End
            Begin MSForms.TextBox txtTF 
               Height          =   285
               Index           =   37
               Left            =   1170
               TabIndex        =   208
               Top             =   30
               Width           =   6000
               VariousPropertyBits=   671105051
               Size            =   "10583;503"
               FontName        =   "新細明體-ExtB"
               FontHeight      =   180
               FontCharSet     =   136
               FontPitchAndFamily=   34
            End
            Begin VB.Label Label26 
               Caption         =   "翻譯瑕疵備註:"
               Height          =   165
               Index           =   14
               Left            =   0
               TabIndex        =   291
               Top             =   75
               Width           =   1245
            End
         End
         Begin VB.CheckBox Chk06 
            Caption         =   "固定報價"
            Height          =   255
            Left            =   0
            TabIndex        =   206
            Top             =   2640
            Width           =   1215
         End
         Begin VB.CheckBox Chk05 
            Caption         =   "暫不翻譯"
            Height          =   255
            Left            =   0
            TabIndex        =   198
            Top             =   1275
            Width           =   1455
         End
         Begin VB.TextBox txtPA 
            Height          =   285
            Index           =   62
            Left            =   1200
            MaxLength       =   1
            TabIndex        =   284
            Top             =   2640
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox Chk04 
            Caption         =   "中說4個月不得延"
            Height          =   255
            Left            =   0
            TabIndex        =   192
            Top             =   398
            Width           =   1935
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "比對結果"
            Height          =   330
            Index           =   2
            Left            =   7200
            TabIndex        =   194
            Top             =   360
            Width           =   915
         End
         Begin VB.CheckBox Chk03 
            Caption         =   "未提申先翻譯"
            Height          =   255
            Left            =   0
            TabIndex        =   195
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox Chk01 
            Caption         =   "待比對(檔案存在，請取消)"
            Height          =   255
            Left            =   4800
            TabIndex        =   193
            Top             =   360
            Width           =   2655
         End
         Begin VB.ComboBox cboSource 
            Height          =   300
            ItemData        =   "frm060102.frx":00E0
            Left            =   5640
            List            =   "frm060102.frx":00E2
            TabIndex        =   199
            Text            =   "cboSource"
            Top             =   1200
            Width           =   1300
         End
         Begin VB.ComboBox cboTarget 
            Height          =   300
            ItemData        =   "frm060102.frx":00E4
            Left            =   5640
            List            =   "frm060102.frx":00E6
            TabIndex        =   200
            Text            =   "cboTarget"
            Top             =   1530
            Width           =   1300
         End
         Begin VB.CheckBox Chk02 
            Caption         =   "待英文本翻譯"
            Height          =   255
            Left            =   5640
            TabIndex        =   203
            Top             =   1965
            Width           =   1455
         End
         Begin VB.CommandButton CmdAddCP 
            Caption         =   "案件進度"
            Height          =   300
            Left            =   6960
            TabIndex        =   205
            Top             =   2265
            Width           =   980
         End
         Begin MSForms.TextBox txtTF 
            Height          =   570
            Index           =   36
            Left            =   1230
            TabIndex        =   207
            Top             =   3150
            Width           =   6765
            VariousPropertyBits=   -1466941413
            ScrollBars      =   2
            Size            =   "11933;1005"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   34
            Left            =   1440
            TabIndex        =   285
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
            VariousPropertyBits=   671105051
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   33
            Left            =   1800
            TabIndex        =   279
            Top             =   360
            Visible         =   0   'False
            Width           =   495
            VariousPropertyBits=   671105051
            MaxLength       =   3
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   1
            Left            =   3390
            TabIndex        =   274
            Top             =   420
            Visible         =   0   'False
            Width           =   495
            VariousPropertyBits=   671105051
            MaxLength       =   3
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   32
            Left            =   6000
            TabIndex        =   197
            Top             =   832
            Width           =   855
            VariousPropertyBits=   671105051
            MaxLength       =   32
            Size            =   "1508;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   31
            Left            =   1440
            TabIndex        =   272
            Top             =   832
            Visible         =   0   'False
            Width           =   495
            VariousPropertyBits=   671105051
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   29
            Left            =   6480
            TabIndex        =   271
            Top             =   480
            Visible         =   0   'False
            Width           =   495
            VariousPropertyBits=   671105051
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   26
            Left            =   3480
            TabIndex        =   196
            Top             =   832
            Width           =   855
            VariousPropertyBits=   671105051
            MaxLength       =   7
            Size            =   "1508;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   30
            Left            =   5640
            TabIndex        =   204
            Top             =   2280
            Width           =   1300
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "2293;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   28
            Left            =   6960
            TabIndex        =   266
            Top             =   1545
            Visible         =   0   'False
            Width           =   495
            VariousPropertyBits=   671105051
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtTF 
            Height          =   285
            Index           =   27
            Left            =   6960
            TabIndex        =   265
            Top             =   1215
            Visible         =   0   'False
            Width           =   495
            VariousPropertyBits=   671105051
            Size            =   "873;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label26 
            Caption         =   "翻譯特殊指示:"
            Height          =   165
            Index           =   13
            Left            =   60
            TabIndex        =   288
            Top             =   3180
            Width           =   1185
         End
         Begin VB.Label Label31 
            Caption         =   "對外翻用-"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3920
            TabIndex        =   280
            Top             =   1275
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "只交Claims期限:"
            Height          =   165
            Index           =   12
            Left            =   4680
            TabIndex        =   273
            Top             =   885
            Width           =   1275
         End
         Begin VB.Label Label26 
            Caption         =   "翻譯語種:"
            Height          =   165
            Index           =   10
            Left            =   4800
            TabIndex        =   270
            Top             =   1605
            Width           =   795
         End
         Begin VB.Label Label26 
            Caption         =   "交稿期限:"
            Height          =   165
            Index           =   8
            Left            =   2640
            TabIndex        =   269
            Top             =   885
            Width           =   795
         End
         Begin VB.Label Label26 
            Caption         =   "原文語種:"
            Height          =   165
            Index           =   9
            Left            =   4800
            TabIndex        =   268
            Top             =   1275
            Width           =   795
         End
         Begin VB.Label Label26 
            Caption         =   "英文本收文號:"
            Height          =   165
            Index           =   11
            Left            =   4440
            TabIndex        =   267
            Top             =   2340
            Width           =   1275
         End
      End
      Begin VB.ComboBox Combo5 
         Height          =   276
         ItemData        =   "frm060102.frx":00E8
         Left            =   5550
         List            =   "frm060102.frx":00EA
         TabIndex        =   35
         Top             =   4125
         Visible         =   0   'False
         Width           =   2500
      End
      Begin VB.Frame Frame2 
         Caption         =   "移動順序:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -70200
         TabIndex        =   260
         Top             =   1650
         Width           =   2025
         Begin VB.CommandButton cmdUp 
            Caption         =   "▲"
            Height          =   255
            Left            =   960
            TabIndex        =   140
            Top             =   90
            Width           =   375
         End
         Begin VB.CommandButton cmdDown 
            Caption         =   "▼"
            Height          =   255
            Left            =   1410
            TabIndex        =   141
            Top             =   90
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   6270
         TabIndex        =   259
         Top             =   360
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "之前"
            Height          =   195
            Index           =   1
            Left            =   690
            TabIndex        =   12
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton Option1 
            Caption         =   "當天"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton Option1 
            Caption         =   "之後"
            Height          =   195
            Index           =   2
            Left            =   1380
            TabIndex        =   13
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdUpdRow 
         Caption         =   "修改"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72270
         TabIndex        =   139
         Top             =   1710
         Width           =   735
      End
      Begin VB.CommandButton cmdDelRow 
         Caption         =   "刪除"
         Height          =   285
         Left            =   -73110
         TabIndex        =   138
         Top             =   1710
         Width           =   735
      End
      Begin VB.CommandButton cmdAddRow 
         Caption         =   "加入"
         Height          =   285
         Left            =   -73935
         TabIndex        =   137
         Top             =   1710
         Width           =   735
      End
      Begin VB.CommandButton cmdPrtContact 
         Caption         =   "發文簡易聯絡單"
         Enabled         =   0   'False
         Height          =   280
         Left            =   6960
         TabIndex        =   15
         Top             =   630
         Width           =   1500
      End
      Begin VB.TextBox txtIN11 
         Height          =   270
         Left            =   -67200
         MaxLength       =   3
         TabIndex        =   136
         Top             =   705
         Width           =   400
      End
      Begin VB.ComboBox Combo4 
         Height          =   276
         ItemData        =   "frm060102.frx":00EC
         Left            =   -73920
         List            =   "frm060102.frx":00EE
         Style           =   2  '單純下拉式
         TabIndex        =   135
         Top             =   450
         Width           =   6135
      End
      Begin VB.TextBox Text25 
         Height          =   270
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   34
         Top             =   4125
         Width           =   495
      End
      Begin VB.TextBox txtPA 
         Height          =   270
         Index           =   159
         Left            =   -73005
         MaxLength       =   20
         TabIndex        =   57
         Top             =   2475
         Width           =   2100
      End
      Begin VB.TextBox txtPA 
         Height          =   270
         Index           =   153
         Left            =   5796
         MaxLength       =   1
         TabIndex        =   20
         Top             =   1830
         Width           =   495
      End
      Begin VB.TextBox txtPA 
         Height          =   270
         Index           =   154
         Left            =   7785
         MaxLength       =   1
         TabIndex        =   21
         Top             =   1830
         Width           =   495
      End
      Begin VB.TextBox txtPA 
         Height          =   270
         Index           =   155
         Left            =   5550
         MaxLength       =   1
         TabIndex        =   23
         Top             =   2100
         Width           =   495
      End
      Begin VB.TextBox txtPA 
         Height          =   270
         Index           =   152
         Left            =   5550
         MaxLength       =   3
         TabIndex        =   27
         Top             =   2700
         Width           =   495
      End
      Begin VB.TextBox txtPA 
         Height          =   270
         Index           =   151
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   26
         Top             =   2700
         Width           =   495
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         ItemData        =   "frm060102.frx":00F0
         Left            =   5550
         List            =   "frm060102.frx":0100
         TabIndex        =   33
         Top             =   3810
         Width           =   2500
      End
      Begin VB.TextBox Text24 
         Height          =   270
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   22
         Top             =   2130
         Width           =   495
      End
      Begin VB.TextBox Text23 
         Height          =   270
         Left            =   -73275
         MaxLength       =   1
         TabIndex        =   64
         Top             =   4560
         Width           =   270
      End
      Begin VB.TextBox Text21 
         Height          =   270
         Left            =   -69060
         MaxLength       =   8
         TabIndex        =   66
         Top             =   3690
         Width           =   975
      End
      Begin VB.TextBox Text30 
         Height          =   270
         Left            =   -73620
         MaxLength       =   8
         TabIndex        =   61
         Top             =   3675
         Width           =   975
      End
      Begin VB.TextBox Text20 
         Height          =   270
         Left            =   -69060
         MaxLength       =   8
         TabIndex        =   65
         Top             =   3390
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   8
         Top             =   345
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Index           =   0
         Left            =   -73440
         Style           =   2  '單純下拉式
         TabIndex        =   68
         Top             =   372
         Width           =   6800
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Index           =   1
         Left            =   -73440
         Style           =   2  '單純下拉式
         TabIndex        =   72
         Top             =   1635
         Width           =   6800
      End
      Begin VB.TextBox Text44 
         Height          =   270
         Left            =   -73380
         MaxLength       =   8
         TabIndex        =   81
         Top             =   972
         Width           =   1095
      End
      Begin VB.TextBox Text42 
         Height          =   270
         Left            =   -73380
         MaxLength       =   8
         TabIndex        =   80
         Top             =   372
         Width           =   1095
      End
      Begin VB.TextBox Text32 
         Height          =   270
         Left            =   -73275
         MaxLength       =   1
         TabIndex        =   63
         Top             =   4260
         Width           =   270
      End
      Begin VB.TextBox Text29 
         Height          =   270
         Left            =   -73620
         MaxLength       =   30
         TabIndex        =   60
         Top             =   3375
         Width           =   2055
      End
      Begin VB.TextBox Text28 
         Height          =   270
         Left            =   -73620
         MaxLength       =   8
         TabIndex        =   59
         Top             =   3075
         Width           =   975
      End
      Begin VB.TextBox Text27 
         Height          =   270
         Left            =   -73620
         MaxLength       =   8
         TabIndex        =   58
         Top             =   2775
         Width           =   975
      End
      Begin VB.TextBox Text18 
         Height          =   270
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   32
         Top             =   3825
         Width           =   495
      End
      Begin VB.TextBox Text17 
         Height          =   270
         Left            =   5550
         MaxLength       =   1
         TabIndex        =   30
         Top             =   3255
         Width           =   495
      End
      Begin VB.TextBox Text16 
         Height          =   270
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   29
         Top             =   3255
         Width           =   495
      End
      Begin VB.TextBox Text15 
         Height          =   270
         Left            =   5550
         MaxLength       =   3
         TabIndex        =   25
         Top             =   2415
         Width           =   495
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   24
         Top             =   2415
         Width           =   495
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   5550
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1545
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1860
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1830
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   2970
         MaxLength       =   8
         TabIndex        =   9
         Top             =   345
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   3495
         Left            =   -74940
         TabIndex        =   258
         Top             =   2040
         Width           =   8355
         _ExtentX        =   14732
         _ExtentY        =   6160
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "V|發明人編號|中文名稱|英文名稱|日文名稱|國籍|申請人1"
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
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRDtmp 
         Height          =   825
         Left            =   -74970
         TabIndex        =   261
         Top             =   1710
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1820
         _ExtentY        =   1461
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "V|發明人編號|中文名稱|英文名稱|日文名稱|國籍|申請人1"
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
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.TextBox Text26 
         Height          =   276
         Left            =   -73860
         TabIndex        =   56
         Top             =   2160
         Width           =   7128
         VariousPropertyBits=   671107099
         MaxLength       =   100
         Size            =   "12573;487"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text22 
         Height          =   276
         Left            =   -69480
         TabIndex        =   67
         Top             =   3960
         Width           =   2952
         VariousPropertyBits=   671107099
         MaxLength       =   35
         Size            =   "5207;487"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text31 
         Height          =   276
         Left            =   -73620
         TabIndex        =   62
         Top             =   3960
         Width           =   2955
         VariousPropertyBits=   671107099
         MaxLength       =   120
         Size            =   "5212;487"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "指定送件日期:"
         Height          =   180
         Index           =   16
         Left            =   4050
         TabIndex        =   318
         Top             =   390
         Width           =   1125
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   0
         Left            =   -73350
         TabIndex        =   85
         Top             =   450
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   5
         Left            =   -73350
         TabIndex        =   105
         Top             =   450
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   6
         Left            =   -73350
         TabIndex        =   109
         Top             =   1560
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   7
         Left            =   -73350
         TabIndex        =   113
         Top             =   2715
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   8
         Left            =   -73350
         TabIndex        =   117
         Top             =   3840
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   9
         Left            =   -73350
         TabIndex        =   121
         Top             =   4995
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   -73350
         TabIndex        =   89
         Top             =   1589
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   2
         Left            =   -73350
         TabIndex        =   93
         Top             =   2728
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   3
         Left            =   -73350
         TabIndex        =   97
         Top             =   3867
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   4
         Left            =   -73350
         TabIndex        =   101
         Top             =   5006
         Width           =   6135
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13652;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "證書形式:                           (1:電子 2:紙本)"
         Height          =   180
         Left            =   240
         TabIndex        =   317
         Top             =   3015
         Width           =   3135
      End
      Begin VB.Label Label35 
         Caption         =   "複製：Ctrl+C鍵　貼上：Ctrl+V鍵　剪下：Ctrl+X鍵　全選：Ctrl+A鍵"
         ForeColor       =   &H00FF00FF&
         Height          =   165
         Left            =   1050
         TabIndex        =   312
         Top             =   5910
         Width           =   5745
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   15
         Left            =   -73440
         TabIndex        =   76
         Top             =   2820
         Width           =   6810
         VariousPropertyBits=   671105051
         Size            =   "12012;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   10
         Left            =   -73860
         TabIndex        =   51
         Top             =   675
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   9
         Left            =   -73860
         TabIndex        =   50
         Top             =   375
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   1
         Left            =   -73440
         TabIndex        =   70
         Top             =   945
         Width           =   6795
         VariousPropertyBits=   671105051
         Size            =   "11986;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   0
         Left            =   -73440
         TabIndex        =   69
         Top             =   672
         Width           =   1995
         VariousPropertyBits=   671105051
         Size            =   "3519;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   14
         Left            =   -73860
         TabIndex        =   55
         Top             =   1875
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   13
         Left            =   -73860
         TabIndex        =   54
         Top             =   1575
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   12
         Left            =   -73860
         TabIndex        =   53
         Top             =   1275
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   11
         Left            =   -73860
         TabIndex        =   52
         Top             =   975
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   8
         Left            =   -73440
         TabIndex        =   79
         Top             =   3855
         Width           =   4155
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "7329;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   7
         Left            =   -73440
         TabIndex        =   78
         Top             =   3555
         Width           =   6795
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "11986;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   6
         Left            =   -73440
         TabIndex        =   77
         Top             =   3255
         Width           =   1935
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "3413;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   5
         Left            =   -73440
         TabIndex        =   75
         Top             =   2535
         Width           =   6810
         VariousPropertyBits=   671105051
         Size            =   "12012;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   4
         Left            =   -73440
         TabIndex        =   74
         Top             =   2235
         Width           =   6795
         VariousPropertyBits=   671105051
         Size            =   "11986;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   3
         Left            =   -73440
         TabIndex        =   73
         Top             =   1935
         Width           =   1995
         VariousPropertyBits=   671105051
         Size            =   "3519;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   285
         Index           =   2
         Left            =   -73440
         TabIndex        =   71
         Top             =   1215
         Width           =   6810
         VariousPropertyBits=   671105051
         Size            =   "12012;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   54
         Left            =   -73350
         TabIndex        =   106
         Top             =   735
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   55
         Left            =   -73350
         TabIndex        =   107
         Top             =   1020
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   56
         Left            =   -73350
         TabIndex        =   108
         Top             =   1305
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   57
         Left            =   -73350
         TabIndex        =   110
         Top             =   1845
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   58
         Left            =   -73350
         TabIndex        =   111
         Top             =   2130
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   59
         Left            =   -73350
         TabIndex        =   112
         Top             =   2430
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   60
         Left            =   -73350
         TabIndex        =   114
         Top             =   3000
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   61
         Left            =   -73350
         TabIndex        =   115
         Top             =   3285
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   62
         Left            =   -73350
         TabIndex        =   116
         Top             =   3555
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   63
         Left            =   -73350
         TabIndex        =   118
         Top             =   4125
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   64
         Left            =   -73350
         TabIndex        =   119
         Top             =   4410
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   65
         Left            =   -73350
         TabIndex        =   120
         Top             =   4710
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   66
         Left            =   -73350
         TabIndex        =   122
         Top             =   5280
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   67
         Left            =   -73350
         TabIndex        =   123
         Top             =   5565
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   68
         Left            =   -73350
         TabIndex        =   124
         Top             =   5850
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   39
         Left            =   -73350
         TabIndex        =   86
         Top             =   746
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   40
         Left            =   -73350
         TabIndex        =   87
         Top             =   1027
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   41
         Left            =   -73350
         TabIndex        =   88
         Top             =   1308
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   42
         Left            =   -73350
         TabIndex        =   90
         Top             =   1885
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   43
         Left            =   -73350
         TabIndex        =   91
         Top             =   2166
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   44
         Left            =   -73350
         TabIndex        =   92
         Top             =   2447
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   45
         Left            =   -73350
         TabIndex        =   94
         Top             =   3024
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   46
         Left            =   -73350
         TabIndex        =   95
         Top             =   3305
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   47
         Left            =   -73350
         TabIndex        =   96
         Top             =   3586
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   48
         Left            =   -73350
         TabIndex        =   98
         Top             =   4163
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   49
         Left            =   -73350
         TabIndex        =   99
         Top             =   4444
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   50
         Left            =   -73350
         TabIndex        =   100
         Top             =   4725
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   51
         Left            =   -73350
         TabIndex        =   102
         Top             =   5302
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   52
         Left            =   -73350
         TabIndex        =   103
         Top             =   5583
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   53
         Left            =   -73350
         TabIndex        =   104
         Top             =   5850
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text45 
         Height          =   285
         Left            =   -73380
         TabIndex        =   295
         TabStop         =   0   'False
         Top             =   1290
         Width           =   6825
         VariousPropertyBits=   671105055
         MaxLength       =   2000
         Size            =   "12039;503"
         BorderColor     =   16777215
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text43 
         Height          =   285
         Left            =   -73380
         TabIndex        =   296
         TabStop         =   0   'False
         Top             =   660
         Width           =   6825
         VariousPropertyBits=   671105055
         MaxLength       =   2000
         Size            =   "12039;503"
         BorderColor     =   16777215
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text46 
         Height          =   495
         Left            =   -74760
         TabIndex        =   297
         Top             =   1860
         Width           =   8205
         VariousPropertyBits=   -1466941413
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "14473;873"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text47 
         Height          =   495
         Left            =   -74760
         TabIndex        =   311
         Top             =   2640
         Width           =   8205
         VariousPropertyBits=   -1466941413
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "14464;873"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   12
         Left            =   -68040
         TabIndex        =   310
         Top             =   3720
         Width           =   1455
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2566;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   11
         Left            =   -68040
         TabIndex        =   309
         Top             =   3420
         Width           =   1455
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2566;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   8
         Left            =   -72600
         TabIndex        =   308
         Top             =   3720
         Width           =   1875
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3307;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   7
         Left            =   -72600
         TabIndex        =   307
         Top             =   3150
         Width           =   5895
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10398;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   6
         Left            =   -72600
         TabIndex        =   306
         Top             =   2820
         Width           =   5895
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10398;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   5
         Left            =   -72720
         TabIndex        =   305
         Top             =   1920
         Width           =   6015
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10610;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   4
         Left            =   -72720
         TabIndex        =   304
         Top             =   1620
         Width           =   6015
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10610;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   3
         Left            =   -72720
         TabIndex        =   303
         Top             =   1320
         Width           =   6015
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10610;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   2
         Left            =   -72720
         TabIndex        =   302
         Top             =   1020
         Width           =   6015
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10610;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   1
         Left            =   -72720
         TabIndex        =   301
         Top             =   720
         Width           =   6015
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10610;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   0
         Left            =   -72720
         TabIndex        =   300
         Top             =   420
         Width           =   6015
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "10610;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   10
         Left            =   -72210
         TabIndex        =   299
         Top             =   1020
         Width           =   5535
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   180
         Index           =   9
         Left            =   -72210
         TabIndex        =   298
         Top             =   420
         Width           =   5535
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9763;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text10 
         Height          =   645
         Left            =   1080
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   900
         Width           =   6405
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "11298;1138"
         BorderColor     =   16777215
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text19 
         Height          =   1125
         Left            =   1050
         TabIndex        =   39
         Top             =   4740
         Width           =   7365
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12991;1984"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "年費特殊管制:                   (Y:年費續辦:有別於Y / X設定  N:寄證書/二核後年費不續辦  空白:視Y / X設定)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   240
         TabIndex        =   292
         Top             =   3540
         Width           =   7995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶有提供彩圖:               (Y:有)"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   286
         Top             =   4470
         Width           =   2445
      End
      Begin VB.Label Label26 
         Caption         =   "設計案屬性:"
         Height          =   165
         Index           =   5
         Left            =   3900
         TabIndex        =   262
         Top             =   4200
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Lb_IN11N 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "Lb_IN11N"
         Height          =   180
         Left            =   -67500
         TabIndex        =   257
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Lb_IN11 
         AutoSize        =   -1  'True
         Caption         =   "國籍:"
         Height          =   180
         Left            =   -67680
         TabIndex        =   256
         Top             =   705
         Width           =   405
      End
      Begin VB.Label Lb_Inv 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   3
         Left            =   -74400
         TabIndex        =   255
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label Lb_Inv 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   2
         Left            =   -74400
         TabIndex        =   254
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Lb_Inv 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   1
         Left            =   -74400
         TabIndex        =   253
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Lb_Inv 
         AutoSize        =   -1  'True
         Caption         =   "發明人1"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   252
         Top             =   510
         Width           =   630
      End
      Begin VB.Label Label29 
         Caption         =   "是否電子送件:                   (Y:是)"
         Height          =   255
         Left            =   240
         TabIndex        =   251
         Top             =   4125
         Width           =   2565
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID:"
         Height          =   180
         Left            =   -74850
         TabIndex        =   250
         Top             =   2490
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "定稿份數:"
         Height          =   180
         Index           =   164
         Left            =   4980
         TabIndex        =   249
         Top             =   1860
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "請款單份數:"
         Height          =   180
         Index           =   165
         Left            =   6750
         TabIndex        =   248
         Top             =   1860
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email 同時寄紙本:                    (Y:是)"
         Height          =   180
         Index           =   166
         Left            =   3915
         TabIndex        =   247
         Top             =   2145
         Width           =   2760
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "年費折扣:                                  %"
         Height          =   180
         Index           =   1
         Left            =   3900
         TabIndex        =   246
         Top             =   2700
         Width           =   2430
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "領證折扣:                          %"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   245
         Top             =   2700
         Width           =   2070
      End
      Begin VB.Label Label26 
         Caption         =   "工程師組別:"
         Height          =   165
         Index           =   1
         Left            =   3900
         TabIndex        =   244
         Top             =   3885
         Width           =   1395
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "以EMail通知:                   （Y:是   D:僅D/N）"
         Height          =   180
         Left            =   240
         TabIndex        =   243
         Top             =   2130
         Width           =   3330
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人:                (Y:印)"
         Height          =   180
         Left            =   3900
         TabIndex        =   242
         Top             =   1575
         Width           =   2775
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請/翻譯折扣:                         %"
         Height          =   180
         Left            =   3900
         TabIndex        =   241
         Top             =   2445
         Width           =   2430
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "年費自動代繳:                        (Y:自動代繳)"
         Height          =   180
         Left            =   3900
         TabIndex        =   240
         Top             =   3270
         Width           =   3270
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "C類收文是否請款:          (N:否)"
         Height          =   180
         Left            =   -74835
         TabIndex        =   239
         Top             =   4560
         Width           =   2340
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日):"
         Height          =   180
         Left            =   -74760
         TabIndex        =   238
         Top             =   2850
         Width           =   1245
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人5"
         Height          =   180
         Index           =   3
         Left            =   -74205
         TabIndex        =   237
         Top             =   5006
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   35
         Left            =   -73845
         TabIndex        =   236
         Top             =   5310
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   34
         Left            =   -73845
         TabIndex        =   235
         Top             =   5595
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   33
         Left            =   -73845
         TabIndex        =   234
         Top             =   5880
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人4"
         Height          =   180
         Index           =   1
         Left            =   -74205
         TabIndex        =   233
         Top             =   3867
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人3"
         Height          =   180
         Index           =   2
         Left            =   -74205
         TabIndex        =   232
         Top             =   2728
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   29
         Left            =   -73845
         TabIndex        =   231
         Top             =   3006
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   28
         Left            =   -73845
         TabIndex        =   230
         Top             =   3290
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   27
         Left            =   -73845
         TabIndex        =   229
         Top             =   3574
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   26
         Left            =   -73845
         TabIndex        =   228
         Top             =   4142
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   25
         Left            =   -73845
         TabIndex        =   227
         Top             =   4426
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   24
         Left            =   -73845
         TabIndex        =   226
         Top             =   4710
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人2"
         Height          =   180
         Index           =   2
         Left            =   -74205
         TabIndex        =   225
         Top             =   1589
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人1"
         Height          =   180
         Index           =   1
         Left            =   -74205
         TabIndex        =   224
         Top             =   450
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   3
         Left            =   -73845
         TabIndex        =   223
         Top             =   734
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   4
         Left            =   -73845
         TabIndex        =   222
         Top             =   1018
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   5
         Left            =   -73845
         TabIndex        =   221
         Top             =   1302
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   6
         Left            =   -73845
         TabIndex        =   220
         Top             =   1870
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   7
         Left            =   -73845
         TabIndex        =   219
         Top             =   2154
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   8
         Left            =   -73845
         TabIndex        =   218
         Top             =   2438
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人10"
         Height          =   180
         Index           =   4
         Left            =   -74205
         TabIndex        =   217
         Top             =   4995
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   1
         Left            =   -73845
         TabIndex        =   216
         Top             =   5287
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   2
         Left            =   -73845
         TabIndex        =   215
         Top             =   5568
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   9
         Left            =   -73845
         TabIndex        =   214
         Top             =   5850
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人9"
         Height          =   180
         Index           =   3
         Left            =   -74205
         TabIndex        =   213
         Top             =   3840
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人8"
         Height          =   180
         Index           =   5
         Left            =   -74205
         TabIndex        =   212
         Top             =   2715
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   10
         Left            =   -73845
         TabIndex        =   211
         Top             =   3015
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   11
         Left            =   -73845
         TabIndex        =   210
         Top             =   3285
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   12
         Left            =   -73845
         TabIndex        =   189
         Top             =   3570
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   13
         Left            =   -73845
         TabIndex        =   188
         Top             =   4140
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   14
         Left            =   -73845
         TabIndex        =   187
         Top             =   4410
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   15
         Left            =   -73845
         TabIndex        =   186
         Top             =   4695
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人7"
         Height          =   180
         Index           =   4
         Left            =   -74205
         TabIndex        =   185
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人6"
         Height          =   180
         Index           =   6
         Left            =   -74205
         TabIndex        =   184
         Top             =   450
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   16
         Left            =   -73845
         TabIndex        =   183
         Top             =   735
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   17
         Left            =   -73845
         TabIndex        =   182
         Top             =   1005
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   18
         Left            =   -73845
         TabIndex        =   181
         Top             =   1290
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   19
         Left            =   -73845
         TabIndex        =   180
         Top             =   1860
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   20
         Left            =   -73845
         TabIndex        =   179
         Top             =   2130
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   21
         Left            =   -73845
         TabIndex        =   178
         Top             =   2415
         Width           =   345
      End
      Begin VB.Label Label14 
         Caption         =   "年費聯絡人:"
         Height          =   255
         Index           =   0
         Left            =   -70560
         TabIndex        =   177
         Top             =   3990
         Width           =   1485
      End
      Begin VB.Label Label12 
         Caption         =   "年費D/N列印對象:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   176
         Top             =   3690
         Width           =   1485
      End
      Begin VB.Label Label10 
         Caption         =   "D/N固定列印對象:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   175
         Top             =   3390
         Width           =   1485
      End
      Begin VB.Label Label76 
         Caption         =   "實體副本收受人彼所案號2:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   174
         Top             =   2412
         Width           =   2292
      End
      Begin VB.Label Label75 
         Caption         =   "實體副本收受人彼所案號1:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   173
         Top             =   1632
         Width           =   2292
      End
      Begin VB.Label Label73 
         Caption         =   "實體副本聯絡人:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   172
         Top             =   1272
         Width           =   1332
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -72360
         TabIndex        =   171
         Top             =   1152
         Width           =   45
      End
      Begin VB.Label Label71 
         Caption         =   "實體副本收受人:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   170
         Top             =   972
         Width           =   1332
      End
      Begin VB.Label Label69 
         Caption         =   "副本聯絡人:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   169
         Top             =   672
         Width           =   1092
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -72360
         TabIndex        =   168
         Top             =   432
         Width           =   45
      End
      Begin VB.Label Label67 
         Caption         =   "副本收受人:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   167
         Top             =   372
         Width           =   1092
      End
      Begin VB.Label Label65 
         Caption         =   "實體聯絡人(日):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   166
         Top             =   3885
         Width           =   1335
      End
      Begin VB.Label Label63 
         Caption         =   "實體聯絡人(英):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   165
         Top             =   3585
         Width           =   1335
      End
      Begin VB.Label Label61 
         Caption         =   "實體聯絡人(中):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   164
         Top             =   3285
         Width           =   1335
      End
      Begin VB.Label Label59 
         Caption         =   "聯絡人2(日):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   163
         Top             =   2555
         Width           =   1095
      End
      Begin VB.Label Label57 
         Caption         =   "聯絡人2(英):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   162
         Top             =   2260
         Width           =   1095
      End
      Begin VB.Label Label55 
         Caption         =   "聯絡人2(中):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   161
         Top             =   1965
         Width           =   1095
      End
      Begin VB.Label Label53 
         Caption         =   "聯絡人1(日):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   160
         Top             =   1272
         Width           =   1092
      End
      Begin VB.Label Label51 
         Caption         =   "聯絡人1(英):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   159
         Top             =   972
         Width           =   1092
      End
      Begin VB.Label Label49 
         Caption         =   "聯絡人1(中):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   158
         Top             =   672
         Width           =   1092
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "信函是否列印Title:          (Y:印)"
         Height          =   180
         Left            =   -74835
         TabIndex        =   157
         Top             =   4290
         Width           =   2355
      End
      Begin VB.Label Label46 
         Caption         =   "客戶案件案號:"
         Height          =   255
         Left            =   -74835
         TabIndex        =   156
         Top             =   3975
         Width           =   1215
      End
      Begin VB.Label Label44 
         Caption         =   "年費請款對象:"
         Height          =   255
         Left            =   -74835
         TabIndex        =   155
         Top             =   3675
         Width           =   1215
      End
      Begin VB.Label Label43 
         Caption         =   "年費彼所案號:"
         Height          =   255
         Left            =   -74835
         TabIndex        =   154
         Top             =   3375
         Width           =   1215
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   -72480
         TabIndex        =   153
         Top             =   3615
         Width           =   45
      End
      Begin VB.Label Label41 
         Caption         =   "年費代理人:"
         Height          =   255
         Left            =   -74835
         TabIndex        =   152
         Top             =   3075
         Width           =   1095
      End
      Begin VB.Label Label39 
         Caption         =   "固定請款對象:"
         Height          =   255
         Left            =   -74835
         TabIndex        =   151
         Top             =   2775
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "彼所案號:"
         Height          =   252
         Left            =   -74835
         TabIndex        =   150
         Top             =   2172
         Width           =   852
      End
      Begin VB.Label Label36 
         Caption         =   "代理人:"
         Height          =   252
         Left            =   -74835
         TabIndex        =   149
         Top             =   1872
         Width           =   852
      End
      Begin VB.Label Label34 
         Caption         =   "申請人5:"
         Height          =   252
         Left            =   -74835
         TabIndex        =   148
         Top             =   1572
         Width           =   852
      End
      Begin VB.Label Label32 
         Caption         =   "申請人4:"
         Height          =   252
         Left            =   -74835
         TabIndex        =   147
         Top             =   1272
         Width           =   852
      End
      Begin VB.Label Label30 
         Caption         =   "申請人3:"
         Height          =   252
         Left            =   -74835
         TabIndex        =   146
         Top             =   972
         Width           =   852
      End
      Begin VB.Label Label28 
         Caption         =   "申請人2:"
         Height          =   252
         Left            =   -74835
         TabIndex        =   145
         Top             =   672
         Width           =   852
      End
      Begin VB.Label Label26 
         Caption         =   "申請人1:"
         Height          =   252
         Index           =   0
         Left            =   -74835
         TabIndex        =   144
         Top             =   372
         Width           =   852
      End
      Begin VB.Label Label25 
         Caption         =   "進度備註:"
         Height          =   255
         Left            =   240
         TabIndex        =   143
         Top             =   4740
         Width           =   855
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "年費單筆不跑:                   (Y:不跑)"
         Height          =   180
         Left            =   240
         TabIndex        =   142
         Top             =   3840
         Width           =   2685
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "領證自動代繳:                   (Y:自動代繳)"
         Height          =   180
         Left            =   240
         TabIndex        =   134
         Top             =   3300
         Width           =   3045
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "全部折扣:                          %"
         Height          =   180
         Left            =   240
         TabIndex        =   133
         Top             =   2415
         Width           =   2160
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文:                          (1.中 2.英 3.日)"
         Height          =   180
         Left            =   240
         TabIndex        =   132
         Top             =   1530
         Width           =   3165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "後續准駁簡單報告:              (Y:核准以及C類來函簡單報告)"
         Height          =   180
         Left            =   240
         TabIndex        =   131
         Top             =   1860
         Width           =   4536
      End
      Begin VB.Label Label8 
         Caption         =   "案件備註:"
         Height          =   255
         Left            =   240
         TabIndex        =   130
         Top             =   930
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "法定期限:"
         Height          =   255
         Left            =   2145
         TabIndex        =   129
         Top             =   372
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "本所期限:"
         Height          =   252
         Left            =   240
         TabIndex        =   128
         Top             =   372
         Width           =   852
      End
   End
   Begin VB.CheckBox ChkPA174 
      Caption         =   "有特殊字"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   0
      TabIndex        =   293
      Top             =   1170
      Width           =   1035
   End
   Begin MSForms.TextBox Text7 
      Height          =   300
      Left            =   1440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1500
      Width           =   7185
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "12674;529"
      BorderColor     =   16777215
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text6 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1185
      Width           =   7185
      VariousPropertyBits=   671105051
      MaxLength       =   250
      Size            =   "12674;529"
      BorderColor     =   16777215
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   870
      Width           =   7185
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "12674;529"
      BorderColor     =   16777215
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   1050
      TabIndex        =   125
      Top             =   1170
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      Height          =   180
      Index           =   0
      Left            =   1050
      TabIndex        =   126
      Top             =   1410
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   1050
      TabIndex        =   84
      Top             =   930
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   120
      TabIndex        =   83
      Top             =   930
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   82
      Top             =   570
      Width           =   765
   End
End
Attribute VB_Name = "frm060102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/3 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'整理 by Morgan 2005/8/10
Option Explicit

'Modify by Morgan 2006/10/19
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim m_CP60 As String, m_CP09 As String, m_CP10 As String 'Memo by Lydia 2022/08/04 新案(CP31=Y)的請款單號,收文號,案件性質
Dim m_CP05 As String 'Added by Lydia 2022/08/04 新案(CP31=Y)的收文日
Dim intWhere As Integer, bolExist As Boolean
'優先權資料
Dim strPriority(1 To 5) As String 'Modify by Amy 2014/04/11 +pd08,pd09

Dim m_901CP09 As String 'Memo by Lydia 2018/04/23 是否為提申前告代
Dim m_203CP09 As String 'Added by Lydia 2018/04/23 是否為提申前主動修正
'Add by Morgan 2008/11/14
Dim m_otxt As Object '共用物件
Dim strName As String 'Add by Amy 2013/05/20
Dim pPrevRow As Integer 'Add By Sindy 2014/11/6
Dim m_CP27 As String 'Added by Morgan 2015/10/19
Dim m_CP57 As String 'Add by Sindy 2023/4/18
'Added by Lydia 2016/03/15 發明人輸入比對兼自動代入(模糊比對)
' 宣告發明人
Private Type INVENTOR
   iN01 As String
   iN02 As String
   iN04 As String
   IN05 As String
   IN06 As String
End Type
Dim m_InventorList() As INVENTOR
Dim m_InventorListCount As Integer
'end 2016/03/15
Dim m_bolChkPA91 As Boolean '是否檢查案件備註 Added by Morgan 2017/8/22
Dim m_TCT01 As String 'Added by Lydia 2017/11/14 FCP案件命名檔PK
Dim m_TCT04 As String 'Added by Lydia 2017/11/28 命名作業-已分案主管
Dim m_TCT10 As String 'Added by Lydia 2018/06/12 命名作業-命名人員
Dim m_TCT27 As String, m_TCT28 As String 'Added by Lydia 2018/06/07 命名作業-欲翻譯人員
'Dim m_TF23 As Double, m_TF19 As Double, m_TF20 As String 'Added by Lydia 2017/12/01 記錄原文字數、相似度、相似案號 'Remove by Lydia 2018/06/07 翻譯分案無紙化
'Added by Lydia 2018/03/01 處理FCP新案命名欄位
Private Const m_FS As Integer = 5  '輸入欄位對應Table欄位的起始位置
'Modified by Lydia 2023/02/16 改成共用常數m_FE=>TF_TCT, m_NotFS=>TF_TCTnotFS
'Private Const m_FE As Integer = 119 '輸入欄位對應Table欄位的終止位置
'Private Const m_NotFS As String = "112,113,114,115" 'Added by Lydia 2018/03/01 排除不修改的欄位
'end 2023/02/16
Dim m_TCT16 As String, m_TCT17 As String 'Added by Lydia 2018/04/16 記錄命名記錄的中文、英文名稱
'Added by Lydia 2018/06/07 翻譯分案無紙化
Dim m_TF01 As String, m_TF01pty As String '記錄中說進度和案件性質
Dim m_TF01cp14 As String, m_TF01cp27 As String  '中說-承辦, 發文日
Dim m_TF01cp06 As String 'Added by Lydia 2022/08/04 中說-所限
Dim m_TF01cp60 As String 'Added by Lydia 2019/06/28 中說-請款單號
Dim m_TCT23 As String, m_TCT24 As String '命名-相似案號和相似度
Dim m_GrpManList As String   '所有工程師主管(含F編號)
Dim m_PrevForm As Form '前一畫面
Dim m_Case(1 To 4)  As String '前一畫面-傳入案號
Dim strResPath As String   '上傳相似比對結果存放路徑
Dim FCP檢視中說必輸原文字數 As String 'Added by Lydia 2019/06/28
'Added by Lydia 2020/02/21
Public bolAskPA174 As Boolean '存檔前檢查有修改案件名稱，將原始檔之維護word檔自動打開，是否有上傳
Dim bolUpdTCN23 As Boolean, m_TCN19 As String  'Added by Lydia 2023/03/07 外專新案認領：取消認領階段，並且通知認領工程

'Added by Lydia 2018/06/07 前一畫面
Public Sub SetParent(ByRef pForm As Form, ByVal pCaseNo As String)
    Set m_PrevForm = pForm
    Call ChgCaseNo(pCaseNo, m_Case)
End Sub

'Add By Sindy 2014/11/10
Private Sub cmdAddRow_Click()
Dim bolChk As Boolean
Dim ii As Integer
Dim Cancel As Boolean
Dim rsTmp  As New ADODB.Recordset
Dim strNo As String
   
   'Added by Lydia 2019/01/31 有新案的客戶編號後補,又要輸入發明人資料,在按確定時程式會出錯
   If Trim(Text33(9)) = "" Then
        MsgBox "請輸入申請人1編號！", vbCritical, "資料檢核"
        Exit Sub
   End If
   'end 2019/01/31
   
   '檢查發明人
   strExc(1) = Replace(Right(Combo4.Text, 11), ")", "")
   If strExc(1) = "" Then
      If txtInvField(0) = "" And txtInvField(1) = "" And txtInvField(2) = "" Then
         Exit Sub
      Else
         '判斷國籍是否有輸入
         If txtIN11.Visible = True Then
            If txtIN11 = "" Then
                MsgBox "請輸入國藉！", vbExclamation
                SSTab1.Tab = 6
                txtIN11.SetFocus
                Exit Sub
            Else
                Cancel = False
                txtIN11_Validate Cancel
                If Cancel = True Then
                  SSTab1.Tab = 6
                  txtIN11.SetFocus
                  TextInverse txtIN11
                  Exit Sub
                End If
            End If
         End If
         '判斷客戶發明人檔是否有重覆資料:發明人會有造字無法存檔時會加空白,所以改在語法內trim
         'Modified by Morgan 2017/12/19 申請人為更名前編號(9碼)時會有錯
         'strNo = Text33(9) & String(8 - Len(Text33(9)), "0")
         If Len(Text33(9)) < 8 Then
            strNo = Text33(9) & String(8 - Len(Text33(9)), "0")
         Else
            strNo = Left(Text33(9), 8)
         End If
         'end 2017/12/19
         strSql = "Select * From Inventor Where IN01=" & CNULL(strNo) & " and (rtrim(IN04)=rtrim('" + ChgSQL(txtInvField(0)) & "')" & _
                  " OR upper(rtrim(IN05))=rtrim('" & ChgSQL(UCase(txtInvField(1))) & "') OR rtrim(IN06)=rtrim('" & ChgSQL(txtInvField(2)) & "'))"
         Set rsTmp = ClsPDReadRst(strSql)
         If Not rsTmp.EOF Then
            If Trim(txtInvField(0)) = Trim("" & rsTmp.Fields("IN04")) Then
               If MsgBox("發明人名稱中文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  rsTmp.Close
                  txtInvField(0).SetFocus
                  TextInverse txtInvField(0)
                  Exit Sub
               End If
            End If
            If Trim(UCase(txtInvField(1))) = UCase(Trim("" & rsTmp.Fields("IN05"))) Then
               If MsgBox("發明人名稱英文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  rsTmp.Close
                  txtInvField(1).SetFocus
                  TextInverse txtInvField(1)
                  Exit Sub
               End If
            End If
            If Trim(txtInvField(2)) = Trim("" & rsTmp.Fields("IN06")) Then
               If MsgBox("發明人名稱日文相同, 是否確定存檔 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  rsTmp.Close
                  txtInvField(2).SetFocus
                  TextInverse txtInvField(2)
                  Exit Sub
               End If
            End If
         End If
         rsTmp.Close
         
         If Trim(txtInvField(0)) <> "" Then
            bolChk = True
            For ii = 1 To GRD1.Rows - 1
               If GRD1.TextMatrix(ii, 2) = Trim(txtInvField(0)) Then
                  bolChk = False
                  Exit For
               End If
            Next ii
            If Not bolChk Then
               MsgBox "發明人中文名稱不可重覆 !", vbCritical
               Combo4.SetFocus
               Exit Sub
            End If
         End If
         
         If Trim(txtInvField(1)) <> "" Then
            bolChk = True
            For ii = 1 To GRD1.Rows - 1
               If UCase(GRD1.TextMatrix(ii, 3)) = Trim(UCase(txtInvField(1))) Then
                  bolChk = False
                  Exit For
               End If
            Next ii
            If Not bolChk Then
               MsgBox "發明人英文名稱不可重覆 !", vbCritical
               Combo4.SetFocus
               Exit Sub
            End If
         End If
         
         If Trim(txtInvField(2)) <> "" Then
            bolChk = True
            For ii = 1 To GRD1.Rows - 1
               If GRD1.TextMatrix(ii, 4) = Trim(txtInvField(2)) Then
                  bolChk = False
                  Exit For
               End If
            Next ii
            If Not bolChk Then
               MsgBox "發明人日文名稱不可重覆 !", vbCritical
               Combo4.SetFocus
               Exit Sub
            End If
         End If
      End If
   Else
      bolChk = True
      For ii = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(ii, 1) = strExc(1) Then
            bolChk = False
            Exit For
         End If
      Next ii
      If Not bolChk Then
         MsgBox "發明人不可重覆 !", vbCritical
         Combo4.SetFocus
         Exit Sub
      End If
   End If
   If Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) <> "" Or Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 6) <> "" Then
      GRD1.AddItem ""
   End If
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) = strExc(1)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 2) = txtInvField(0)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 3) = txtInvField(1)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 4) = txtInvField(2)
   If strExc(1) = "" Then
      'Modified by Lydia 2024/12/03
      'Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 5) = txtIN11
      Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 5) = Lb_IN11N
      Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 7) = txtIN11
      'end 2024/12/03
      Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 6) = strNo '申請人1_ID
   End If
   cmdAddRow.Tag = "I" '記錄有異動資料
   '清空欄位
   Combo4.ListIndex = 0
   For Each m_otxt In txtInvField
      m_otxt.Text = ""
      m_otxt.Tag = "" 'Add By Sindy 2015/3/5
   Next
   txtIN11.Text = "" 'Add By Sindy 2015/12/4
   Lb_IN11N = "" 'Added by Lydia 2024/12/03
End Sub

'Add By Sindy 2014/11/10
Private Sub cmdDelRow_Click()
   If pPrevRow <= 0 Then Exit Sub
   GRD1.col = 0
   GRD1.row = pPrevRow
   If GRD1.CellBackColor <> &HFFC0C0 Then Exit Sub
   If pPrevRow = 1 And GRD1.Rows = 2 Then
      GRD1.TextMatrix(pPrevRow, 0) = ""
      GRD1.TextMatrix(pPrevRow, 1) = ""
      GRD1.TextMatrix(pPrevRow, 2) = ""
      GRD1.TextMatrix(pPrevRow, 3) = ""
      GRD1.TextMatrix(pPrevRow, 4) = ""
      GRD1.TextMatrix(pPrevRow, 5) = ""
      GRD1.TextMatrix(pPrevRow, 6) = ""
   Else
      If pPrevRow > 0 Then
         Call GRD1.RemoveItem(pPrevRow)
      Else
         Exit Sub
      End If
   End If
   pPrevRow = pPrevRow - 1
   cmdDelRow.Tag = "D" '記錄有異動資料
   '清空欄位
   Combo4.ListIndex = 0
   For Each m_otxt In txtInvField
      m_otxt.Text = ""
      m_otxt.Tag = "" 'Add By Sindy 2015/3/5
   Next
   txtIN11.Text = "" 'Add By Sindy 2015/12/4
   Lb_IN11N = "" 'Added by Lydia 2024/12/03
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim tmpBol As Boolean 'Added by Lydia 2018/04/23

   Select Case Index
      Case 0 '確定
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         'Add by Morgan 2008/8/18
         m_901CP09 = ""
         m_203CP09 = "" 'Added by Lydia 2018/04/23
         
         'Modified by Lydia 2021/05/12 案件提申後，不管制彈提醒
         'If Text8 <> "" Then
         If (pa(1) = "FCP" Or pa(1) = "P") And Text8 <> "" And pa(10) = "" Then
            'Modified by Lydia 2023/06/06 +CP06
            strExc(0) = "select CP06,cp48,cp09 from CaseProgress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='901' and cp27 is null and cp57 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Modified by Lydia 2018/04/23 可能有多筆
'               If IsNull(RsTemp.Fields(0)) Or Val(DBDATE(Text8)) < Val("" & RsTemp.Fields(0)) Then
'                  If MsgBox("是否為提申前告代？", vbYesNo) = vbYes Then
'                      m_901CP09 = RsTemp.Fields(1)
'                  End If
'               End If
               tmpBol = False 'Added by Lydia 2018/04/23
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                    If tmpBol = False Then
                        'Modified by Lydia 2023/06/06 改成本所期限並且大於新案所限會彈訊息
                        'If IsNull(RsTemp.Fields(0)) Or Val(DBDATE(Text8)) < Val("" & RsTemp.Fields(0)) Then
                         '  If MsgBox("是否為提申前告代？", vbYesNo) = vbYes Then
                        If IsNull(RsTemp.Fields("CP06")) Or Val(DBDATE(Text8)) < Val("" & RsTemp.Fields("CP06")) Then
                           'Modified by Lydia 2023/07/24 請增加「取消」選項，若點選取消則回新案建檔;預設選項在「取消」
                           'If MsgBox("是否為提申前告代？" & vbCrLf & "收文號：" & RsTemp.Fields("CP09"), vbYesNo) = vbYes Then
                           intI = MsgBox("是否為提申前告代？" & vbCrLf & "收文號：" & RsTemp.Fields("CP09"), vbInformation + vbYesNoCancel + vbDefaultButton3)
                           If intI = vbCancel Then
                              tmpBol = False
                              Exit Sub
                           ElseIf intI = vbYes Then
                           'end 2023/07/24
                        'end 2023/06/06
                               tmpBol = True
                               m_901CP09 = m_901CP09 & RsTemp.Fields("CP09") & ","
                           End If
                        End If
                    Else
                        'Modified by Lydia 2023/06/06
                        'If IsNull(RsTemp.Fields(0)) Or Val(DBDATE(Text8)) < Val("" & RsTemp.Fields(0)) Then
                        '    m_901CP09 = m_901CP09 & RsTemp.Fields(1) & ","
                        If IsNull(RsTemp.Fields("CP06")) Or Val(DBDATE(Text8)) < Val("" & RsTemp.Fields("CP06")) Then
                            m_901CP09 = m_901CP09 & RsTemp.Fields("CP09") & ","
                        'end 2023/06/06
                        End If
                    End If
                    RsTemp.MoveNext
               Loop
               'end 2018/04/23
            End If
            'Added by Lydia 2018/04/23 是否為提申前主動修正
            'Modified by Lydia 2023/06/06 +CP06
            strExc(0) = "select CP06,cp48,cp09 from CaseProgress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='203' and cp27 is null and cp57 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               tmpBol = False
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                    If tmpBol = False Then
                        'Modified by Lydia 2023/06/06 改成本所期限並且大於新案所限會彈訊息
                        'If IsNull(RsTemp.Fields(0)) Or Val(DBDATE(Text8)) < Val("" & RsTemp.Fields(0)) Then
                        '   If MsgBox("是否為提申前主動修正？", vbYesNo) = vbYes Then
                        If IsNull(RsTemp.Fields("CP06")) Or Val(DBDATE(Text8)) < Val("" & RsTemp.Fields("CP06")) Then
                           'Modified by Lydia 2023/07/24 請增加「取消」選項，若點選取消則回新案建檔;預設選項在「取消」
                           'If MsgBox("是否為提申前主動修正？" & vbCrLf & "收文號：" & RsTemp.Fields("CP09"), vbYesNo) = vbYes Then
                           intI = MsgBox("是否為提申前主動修正？" & vbCrLf & "收文號：" & RsTemp.Fields("CP09"), vbInformation + vbYesNoCancel + vbDefaultButton3)
                           If intI = vbCancel Then
                              tmpBol = False
                              Exit Sub
                           ElseIf intI = vbYes Then
                           'end 2023/07/24
                        'end 2023/06/06
                               tmpBol = True
                               m_203CP09 = m_203CP09 & RsTemp.Fields("CP09") & ","
                           End If
                        End If
                    Else
                        'Modified by Lydia 2023/06/06
                        'If IsNull(RsTemp.Fields(0)) Or Val(DBDATE(Text8)) < Val("" & RsTemp.Fields(0)) Then
                        '    m_203CP09 = m_203CP09 & RsTemp.Fields(1) & ","
                        If IsNull(RsTemp.Fields("CP06")) Or Val(DBDATE(Text8)) < Val("" & RsTemp.Fields("CP06")) Then
                            m_203CP09 = m_203CP09 & RsTemp.Fields("CP09") & ","
                        'end 2023/06/06
                        End If
                    End If
               RsTemp.MoveNext
               Loop
            End If
            'end 2018/04/23
         End If
         
         'Added by Morgan 2013/7/9 一案兩請檢查
         '同一申請人同一天收文但不同專利種類時，若未建立關聯則提醒使用者。
         If pa(1) = "FCP" And (m_CP10 = "101" Or m_CP10 = "102") Then
            If PUB_DualCaseExist(pa, strExc(1)) = True Then
               If PUB_DualCaseRelationExist(pa) = False Then
                  If MsgBox("本案與 " & strExc(1) & " 案可能為一案兩申請且尚未建立關聯，確定要繼續？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
               End If
            End If
         End If
         'end 2013/7/9
         
          'Added by Morgan 2014/11/10
         'Modified by Morgan 2014/11/19 所限改為法限的前2個日曆天
         If Text8 <> "" And Text9 <> "" And (Text8 <> Text8.Tag Or Text9 <> Text9.Tag) Then
            'Modified by Morgan 2023/5/24
            'If Text8 <> TransDate(CompDate(2, -2, Text9), 1) Then
            '   If MsgBox("本所期限非法定期限的前2天，是否確定？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            If Text8 <> TransDate(PUB_GetFCPOurDeadline(Text9, 2), 1) Then
               If MsgBox("本所期限非法定期限的前2個工作天，是否確定？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            'end 2023/5/24
                  Text8.SetFocus
                  Exit Sub
               End If
            End If
         End If
         'end 2014/11/19
         'end 2014/11/10
         
        'Added by Lydia 2018/06/07 若有交稿期限新案建檔輸入時，行事曆自動新增期限
        If txtTF(26).Text <> txtTF(26).Tag Or txtTF(32).Text <> txtTF(32).Tag Then
            strExc(0) = "select sc01,02 from staff_calendar where sc05='" & pa(1) & "' and sc06='" & pa(2) & "' and sc07='" & pa(3) & "' and sc08='" & pa(4) & "' " & _
                              "and sc18 is null and instr(sc04,'譯者') > 0  and instr(sc04,'交稿期限') > 0 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                MsgBox "國外部行事曆已有譯者翻譯/Claims交稿期限，請自行解除 !", vbInformation
            End If
        End If
        'end 2018/06/07
        
        'Added by Lydia 2017/11/14 FCP案件命名電子化
'Modified by Lydia 2018/06/07 翻譯分案無紙化：改成按下尋找就抓資料
'        If strSrvDate(1) >= FCP案件命名啟用日 And Text1.Text = "FCP" And Combo3.Tag <> Combo3.Text Then
'        'memo by Lydia 2018/04/11 若發文後,改組別仍要清空資料 =>拿掉 AND CP158=0
'           strExc(0) = "SELECT TCT01,TCT02,TCT03,TCT04,TCT07,TCT10,CP06,CP07 FROM CaseProgress,TransCaseTitle " & _
'                        "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' " & _
'                        "AND CP10 IN (" & NewCasePtyList & ") AND CP159=0 AND CP09=TCT01 ORDER BY CP09 DESC"
'           intI = 1
'           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'           If intI = 1 Then
'              RsTemp.MoveFirst
'              '已有案件命名資料檔,避免系統上線時新舊混合
'              If "" & RsTemp.Fields("TCT01") <> "" Then
'                  'Modified by Lydia 2017/12/27 可變更
'                  'If "" & RsTemp.Fields("TCT07") & RsTemp.Fields("TCT10") <> "" Then
'                  '       MsgBox "工程師主管已分案給命名人員，不可變更工程師組別 !", vbCritical
'                  '       Combo3.SetFocus
'                  '       Exit Sub
'                  'ElseIf "" & RsTemp.Fields("TCT04") <> "" Then
'                  '   If MsgBox("確定要更改工程師組別嗎？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
'                     If "" & RsTemp.Fields("TCT04") <> "" Then 'Added by Lydia 2018/03/27 命名作業人工流程變更: Phoebe告知與Jack討論後,所有新案在櫃台收文時一律輸入"退程序",等到Gill確定組別後做分案作業,再通知程序人員到新案建檔設定工程師組別。
'                           If MsgBox("更改工程師組別會清空命名記錄檔除案件名稱以外的內容，確定要繼續嗎？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
'                        'end 2017/12/28
'                              Combo3.SetFocus
'                              Exit Sub
'                           End If
'                     End If
'                  'End If 'Remark by Lydia 2017/12/28
'                  m_TCT01 = "" & RsTemp.Fields("TCT01")
'                  m_TCT04 = "" & RsTemp.Fields("TCT04") 'Added by Lydia 2017/11/28
'              End If
'           End If
'        End If
'        'end 2017/11/14
        'Added by Lydia 2023/03/07 外專新案認領：若處於認領階段則取消該階段(TCN23=9)，將認領期限更新為系統日期+時間，並且Email認領工程師。
        bolUpdTCN23 = False: m_TCN19 = ""
        If strSrvDate(1) >= 外專新案認領啟用日 And m_TCT01 <> "" And Combo3.Tag <> Combo3.Text Then
            strExc(0) = "select tct04,tcn20,tcn21,tcn23,tcn16,tcn17,tcn18,st01,decode(st16,'1',nvl(tcn19,'Y'),tcn19) as grpno " & _
                             "from transcasetitle , trackingcasename ,staff  where tct01='" & m_TCT01 & "' and tct01=tcn05(+) and tcn03=st01(+) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_TCT04 = "" & RsTemp.Fields("tct04")
               If "" & RsTemp.Fields("tcn16") = "Y" Then
                   MsgBox "本案處於暫不認領，請勿變更工程師組別!", vbExclamation + vbOKOnly, "外專新案認領階段"
                   Combo3.SetFocus
                   Exit Sub
               End If
               If "" & RsTemp.Fields("tcn23") <> "9" Then
                  bolUpdTCN23 = True
                  If m_TCT04 = "" And "" & RsTemp.Fields("tcn23") <> "" And "" & RsTemp.Fields("tcn21") <> "99999999" Then
                     If MsgBox("變更工程師組別會取消外專新案認領，是否繼續存檔？", vbExclamation + vbYesNo + vbDefaultButton2, "外專新案認領階段") = vbNo Then
                         Combo3.SetFocus
                         Exit Sub
                     End If
                     m_TCN19 = "" & RsTemp.Fields("GRPNO") '英文組預設=Y
                  End If
               End If
            End If
        End If
        'end 2023/03/07
        'Modified by Lydia 2018/05/09 +FMP案
        'If m_TCT01 <> "" And m_TCT04 <> "" And Text1.Text = "FCP" And Combo3.Tag <> Combo3.Text Then
        If m_TCT01 <> "" And m_TCT04 <> "" And (text1.Text = "FCP" Or text1.Text = "P") And Combo3.Tag <> Combo3.Text Then
               '2018/03/27 命名作業人工流程變更: Phoebe告知與Jack討論後,所有新案在櫃台收文時一律輸入"退程序",等到Gill確定組別後做分案作業,再通知程序人員到新案建檔設定工程師組別。
               If MsgBox("更改工程師組別會清空命名記錄檔除案件名稱以外的內容，確定要繼續嗎？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                   Combo3.SetFocus
                   Exit Sub
               End If
        End If
'end 2018/06/07
        
         'Added by Lydia 2018/07/06 待比對發mail通知
         If txtTF(29).Tag = "" And Chk01.Value = 1 And Chk01.Visible = True And m_TCT01 <> "" Then
               If Combo3.Tag <> Combo3.Text Then
                    MsgBox "修改工程師組別不可與待比對同時進行 !", vbCritical, "取消待比對"
                    Chk01.Value = 0
                    Exit Sub
               End If
               If m_TCT10 = "" Then
                    MsgBox "請先等工程師主管指定命名人員 !", vbCritical, "取消待比對"
                    Chk01.Value = 0
                    Exit Sub
               End If
               If txtTF(20).Text = "" Then
                    MsgBox "請記得在內文輸入相似案號 !", vbInformation
                    strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(4) = ""   'Added by Lydia 2018/09/28 沒有相似案號要清空(ex.FCP-59643)
               End If
                '開啟Email畫面
                If txtTF(20).Text <> "" Then
                    Call ChgCaseNo(txtTF(20).Text, strExc)
                End If
                strExc(5) = ChangeTStringToTDateString(TransDate(CompWorkDay(6, strSrvDate(1)), 1)) '比對期限:系統日+5個工作天
                frm880019.txtSubject = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & " 待比對，期限：" & strExc(5)
                'Modified by Lydia 2018/09/19 +PDF
                strExc(0) = "請回到命名作業畫面上傳比對結果檔案" & pa(1) & pa(2) & ".RES.doc或PDF檔，並且註明相似度。" & vbCrLf
                strExc(0) = strExc(0) & "P.S若命名作業已完成，請到本案的卷宗區點選命名記錄(RCD.menu)進入已確認-命名作業畫面。" & vbCrLf & vbCrLf
                strExc(0) = strExc(0) & "比對期限：" & strExc(5) & vbCrLf
                If strExc(2) <> "" Then
                     strExc(0) = strExc(0) & "相似案號：" & IIf(strExc(2) <> "", strExc(1) & strExc(2) & IIf(strExc(3) & strExc(4) <> "000", strExc(3) & strExc(4), ""), "") & vbCrLf
                Else
                     strExc(0) = strExc(0) & "相似案號：(請輸入相似案號)" & vbCrLf
                End If
                frm880019.txtContent = strExc(0)
                frm880019.txtReceiver = m_TCT10
                frm880019.SetParent Me
                frm880019.Show vbModal
                tmpBol = frm880019.m_bolDone '是否傳送成功
                Unload frm880019
                If tmpBol = False Then
                    MsgBox "送信失敗，請重新Email !", vbCritical, "取消待比對"
                    Chk01.Value = 0
                    Exit Sub
                End If
         End If
         'end 2018/07/06
         
         'Added by Lydia 2018/08/27 人工勾選固定報價，彈訊息
         If txtPA(62) = "Y" And txtPA(62).Tag = "" Then
              MsgBox "請回報固定報價之編號給Sharon", vbInformation
         End If
         'end 2018/08/27
         
         'Add By Sindy 2023/4/18 檢查指定送件日相關欄位
         If Val(txtCP142.Text) > 0 Then
            If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
               MsgBox "有輸入指定送件日，當天或之前或之後請擇一。", vbExclamation
               Exit Sub
            End If
         Else
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
         End If
         '2023/4/18 END
         
         'Add by Sindy 2021/11/3 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         'Modified by Lydia 2021/11/05 改成可選擇; 因為無法輸入的案件名稱需要保留問號
         'If PUB_ChkUniText(Me) = False Then
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         '2021/11/3 END

         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Added by Lydia 2018/04/10 存檔後發mail
         PUB_SendMailCache
         
         FormClear
        '保留原輸入的系統類別
         Text2.Text = ""
         Text3.Text = ""
         Text4.Text = ""
         text1.SetFocus
         Me.Text2.SetFocus
         SSTab1.Tab = 0 'Add By Sindy 2014/12/10
         
         'Added by Lydia 2018/05/21 翻譯分案無紙化：存檔後,關閉
         If TypeName(m_PrevForm) = "frm060122" Then
               Unload Me
         End If
         'end 2018/05/21
      Case 2 '結束
         Unload Me
      Case 1 '申請人參考資料
         If Text33(9) <> "" Then
            If mdiMain.mnuTitle(10).Enabled = True Then
               Me.Enabled = False
               frm100101_11.Show
               frm100101_11.Tag = ChangeCustomerL(Text33(9)) '傳申請人代號
               frm100101_11.StrMenu
               Me.Enabled = True
            Else
               MsgBox "請先關閉共同查詢畫面！"
            End If
         Else
            MsgBox "請輸入申請人1 !", vbInformation
            SSTab1.Tab = 1
            Text33(9).SetFocus
         End If
      Case 3 '代理人參考資料
         If Text33(14) <> "" Then
            If mdiMain.mnuTitle(10).Enabled = True Then
               Me.Enabled = False
               frm100101_10.Show
               frm100101_10.Tag = ChangeCustomerL(Text33(14)) '傳代理人代號
               frm100101_10.StrMenu
               Me.Enabled = True
            Else
               MsgBox "請先關閉共同查詢畫面！"
            End If
         Else
            MsgBox "請輸入代理人 !", vbInformation
            SSTab1.Tab = 1
            Text33(14).SetFocus
         End If
         
        'Add by Morgan 2004/2/11
        '優先權資料
        Case 4
            'Modify by Sindy 2019/1/23 + pa(1) & pa(2) & pa(3) & pa(4)
            ModifyPriority strPriority(1), strPriority(2), strPriority(3), , , pa(1) & pa(2) & pa(3) & pa(4), , , strPriority(4), strPriority(5) 'Modify by Amy 2014/04/11 +pd08,pd09
            
      'Add by Morgan 2011/6/13
      Case 5, 6, 7, 8 '申請人參考資料
         If Text33(5 + Index) <> "" Then
            If mdiMain.mnuTitle(10).Enabled = True Then
               Me.Enabled = False
               frm100101_11.Show
               frm100101_11.Tag = ChangeCustomerL(Text33(5 + Index))  '傳申請人代號
               frm100101_11.StrMenu
               Me.Enabled = True
            Else
               MsgBox "請先關閉共同查詢畫面！"
            End If
         End If
   End Select
End Sub

'Add by Amy 2014/07/04 新申請(101,102,103,125)已發文者可重印簡易聯絡單
Private Sub cmdPrtContact_Click()
    'Modify by Amy 2016/04/29 +傳案件性質
    'Modified by Lydia 2020/02/10 +外部呼叫=True
    'Modiofy By Sindy 2022/6/14 + m_strContactSheetA4:記錄簡易聯絡單資料
    'Modified by Lydia 2023/05/19 改模組
    'm_strContactSheetA4 = frm060104_3.PrintContactSheetA4(m_CP09, Text1, Text2, Text3, Text4, m_CP10, True)
    m_strContactSheetA4 = PUB_FCPPrintContactSheetA4(True, m_CP09, text1, Text2, Text3, Text4, m_CP10, True)
End Sub
'end 2014/07/04

'Add By Sindy 2017/3/15 向上移
Private Sub cmdUp_Click()
Dim ii As Integer, jj As Integer
   
   If pPrevRow > 1 And GRD1.Rows - 1 > 0 Then
      If GRD1.TextMatrix(pPrevRow, 0) <> "" Then '點選的資料列有資料
         '記錄暫存Grid
         GRDtmp.Clear
         Call SetGrd(GRDtmp): GRDtmp.Visible = False
         For ii = 1 To GRD1.Rows - 1
            If ii > 1 Then GRDtmp.AddItem ""
            For jj = 0 To GRD1.Cols - 1
               GRDtmp.TextMatrix(ii, jj) = GRD1.TextMatrix(ii, jj)
            Next jj
         Next ii
         GRD1.Enabled = False
         '處理上移後的上方資料
         For ii = 1 To pPrevRow - 2
            For jj = 0 To GRD1.Cols - 1
               GRD1.TextMatrix(ii, jj) = GRDtmp.TextMatrix(ii, jj)
            Next jj
         Next ii
         '對換資料列
         For jj = 0 To GRD1.Cols - 1
            GRD1.TextMatrix(pPrevRow - 1, jj) = GRDtmp.TextMatrix(pPrevRow, jj)
         Next jj
         For jj = 0 To GRD1.Cols - 1
            GRD1.TextMatrix(pPrevRow, jj) = GRDtmp.TextMatrix(pPrevRow - 1, jj)
         Next jj
         '處理上移後的下方資料
         For ii = pPrevRow + 1 To GRD1.Rows - 1
            For jj = 0 To GRD1.Cols - 1
               GRD1.TextMatrix(ii, jj) = GRDtmp.TextMatrix(ii, jj)
            Next jj
         Next ii
         cmdAddRow.Tag = "I" '記錄有異動資料
         Call SetGrd1SelRow(pPrevRow - 1)
         GRD1.Enabled = True
      End If
   ElseIf pPrevRow = 1 Then
      MsgBox "已到第一筆！", vbCritical + vbOKOnly, MsgText(9001)
   Else
      MsgBox "欲移動資料項目，請選擇一筆資料！", vbCritical + vbOKOnly, MsgText(9001)
   End If
End Sub

'Add By Sindy 2017/3/15 向下移
Private Sub cmdDown_Click()
Dim ii As Integer, jj As Integer
   
   If (pPrevRow > 0 And pPrevRow < GRD1.Rows - 1) And GRD1.Rows - 1 > 0 Then
      If GRD1.TextMatrix(pPrevRow, 0) <> "" Then '點選的資料列有資料
         '記錄暫存Grid
         GRDtmp.Clear
         Call SetGrd(GRDtmp): GRDtmp.Visible = False
         For ii = 1 To GRD1.Rows - 1
            If ii > 1 Then GRDtmp.AddItem ""
            For jj = 0 To GRD1.Cols - 1
               GRDtmp.TextMatrix(ii, jj) = GRD1.TextMatrix(ii, jj)
            Next jj
         Next ii
         GRD1.Enabled = False
         '處理下移後的上方資料
         For ii = 1 To pPrevRow - 1
            For jj = 0 To GRD1.Cols - 1
               GRD1.TextMatrix(ii, jj) = GRDtmp.TextMatrix(ii, jj)
            Next jj
         Next ii
         '對換資料列
         For jj = 0 To GRD1.Cols - 1
            GRD1.TextMatrix(pPrevRow + 1, jj) = GRDtmp.TextMatrix(pPrevRow, jj)
         Next jj
         For jj = 0 To GRD1.Cols - 1
            GRD1.TextMatrix(pPrevRow, jj) = GRDtmp.TextMatrix(pPrevRow + 1, jj)
         Next jj
         '處理下移後的下方資料
         For ii = pPrevRow + 2 To GRD1.Rows - 1
            For jj = 0 To GRD1.Cols - 1
               GRD1.TextMatrix(ii, jj) = GRDtmp.TextMatrix(ii, jj)
            Next jj
         Next ii
         cmdAddRow.Tag = "I" '記錄有異動資料
         Call SetGrd1SelRow(pPrevRow + 1)
         GRD1.Enabled = True
      End If
   ElseIf pPrevRow = GRD1.Rows - 1 Then
      MsgBox "已到最末筆！", vbCritical + vbOKOnly, MsgText(9001)
   Else
      MsgBox "欲移動資料項目，請選擇一筆資料！", vbCritical + vbOKOnly, MsgText(9001)
   End If
End Sub

'Add By Sindy 2017/3/15
Private Sub SetGrd1SelRow(intSelRow As Integer)
Dim nRow As Integer, nCol As Integer
Dim iCol As Integer
   
   With GRD1
      .Visible = False
      nRow = intSelRow
      If nRow > 0 Then
         nCol = .col
         If pPrevRow > 0 Then
            If pPrevRow <> nRow Then
               .row = pPrevRow
               .TextMatrix(pPrevRow, 0) = ""
               If .FixedCols > 0 Then
                  .col = .FixedCols - 1
                  .CellBackColor = .BackColorFixed
                  .CellForeColor = .ForeColor
               End If
               For iCol = .FixedCols To .Cols - 1
                  .col = iCol
                  .CellBackColor = .BackColor
               Next
            End If
         End If
         If nRow > 0 Then
            .row = nRow
            .TextMatrix(nRow, 0) = "V"
            If .FixedCols > 0 Then
               .col = .FixedCols - 1
               .CellBackColor = .BackColorSel
               .CellForeColor = .ForeColorSel
            End If
            For iCol = .FixedCols To .Cols - 1
              .col = iCol
              .CellBackColor = &HFFC0C0
            Next
         End If
         pPrevRow = intSelRow
         Call SetCombo4Data(.TextMatrix(nRow, 1))
         If .TextMatrix(nRow, 1) = "" Then
            txtInvField(0) = .TextMatrix(nRow, 2)
            txtInvField(1) = .TextMatrix(nRow, 3)
            txtInvField(2) = .TextMatrix(nRow, 4)
            'Modifed by Lydia 2024/12/03
            'txtIN11 = .TextMatrix(nRow, 5)
            txtIN11 = .TextMatrix(nRow, 7)
            Lb_IN11N = .TextMatrix(nRow, 5)
            'end 2024/12/03
            cmdUpdRow.Enabled = True
            cmdAddRow.Enabled = False
         End If
      End If
      .Visible = True
   End With
End Sub

'Add By Sindy 2015/3/5
Private Sub cmdUpdRow_Click()
      Me.GRD1.TextMatrix(pPrevRow, 2) = txtInvField(0)
      Me.GRD1.TextMatrix(pPrevRow, 3) = txtInvField(1)
      Me.GRD1.TextMatrix(pPrevRow, 4) = txtInvField(2)
      'Modified by Lydia 2024/12/03
      'Me.GRD1.TextMatrix(pPrevRow, 5) = txtIN11
      Me.GRD1.TextMatrix(pPrevRow, 5) = Lb_IN11N
      Me.GRD1.TextMatrix(pPrevRow, 7) = txtIN11
      'end 2024/12/03
      cmdUpdRow.Enabled = False
End Sub

Private Sub Combo1_Click(Index As Integer)
 Dim i As Integer, strTmp As String
   If Combo1(Index) = "" Then
      For i = 0 To 2
         Text33(i + Index * 3) = ""
      Next
      'Added by Lydia 2017/05/02 無聯絡人資料才清空
      If Trim(Text33(0) & Text33(1) & Text33(2) & Text33(3) & Text33(4) & Text33(5)) = "" Then
         Text33(15) = ""
      End If
      'end 2017/05/02
      Exit Sub
   End If
   
   strTmp = Mid(Combo1(Index).Text, InStr(Combo1(Index).Text, "-") + 1, 1)
   Select Case text1
      Case "FCP", "P", "CFP"
         If pa(75) <> "" Then
            Select Case strTmp
               Case "1"
                  strExc(1) = "FA07,FA08,FA09,FA78"
               Case "2"
                  strExc(1) = "FA52,FA53,FA54,FA78"
            End Select
         Else
            Select Case strTmp
               Case "1"
                  strExc(1) = "CU58,CU59,CU60,CU114"
               Case "2"
                  strExc(1) = "CU61,CU62,CU63,CU114"
            End Select
         End If
      
      Case "FG", "PS", "CPS"
         If pa(26) <> "" Then
            Select Case strTmp
               Case "1"
                  strExc(1) = "FA07,FA08,FA09,FA78"
               Case "2"
                  strExc(1) = "FA52,FA53,FA54,FA78"
            End Select
         Else
            Select Case strTmp
               Case "1"
                  strExc(1) = "CU58,CU59,CU60,CU114"
               Case "2"
                  strExc(1) = "CU61,CU62,CU63,CU114"
            End Select
         End If
   End Select
   
   strExc(2) = ChgFagent(Left(Combo1(Index).Text, InStr(Combo1(Index).Text, "-") - 1))
   strExc(3) = ChgCustomer(Left(Combo1(Index).Text, InStr(Combo1(Index).Text, "-") - 1))
   Select Case text1
      Case "FCP", "P", "CFP"
         If pa(75) <> "" Then
            strExc(0) = "SELECT " & strExc(1) & " FROM FAGENT WHERE " & strExc(2)
         Else
            strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & strExc(3)
         End If

      Case "FG", "PS", "CPS"
         If pa(26) <> "" Then
            strExc(0) = "SELECT " & strExc(1) & " FROM FAGENT WHERE " & strExc(2)
         Else
            strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & strExc(3)
         End If

   End Select
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Remove by Morgan 2006/10/19 改與FCP一致
      'Select Case Text1
      '   Case "FCP"
            For i = 0 To 2
               If Not IsNull(RsTemp.Fields(i)) Then
                  Text33(i + Index * 3) = RsTemp.Fields(i)
               Else
                  Text33(i + Index * 3) = ""
               End If
            Next
            'Modified by Lydia 2017/05/02 有值才代入,避免清掉聯絡人部門(日)
            'Text33(15) = "" & RsTemp.Fields(3) 'Add by Morgan 2006/10/19
            If "" & RsTemp.Fields(3) <> "" Then Text33(15) = RsTemp.Fields(3)
            'end 2017/05/02
'         Case "FG"
'            If Not IsNull(rsTemp.Fields(0)) Then Text33(1) = rsTemp.Fields(0)
      'End Select
   End If
End Sub

'Morgan 2003/11/24
Private Sub Combo2_Click(Index As Integer)

   Dim i As Integer, strTmp As String
   
   If Combo2(Index) = "" Then
      For i = 0 To 2
         txtCaseField(i + (Index + 1) * 3 + 36) = ""
      Next
      Exit Sub
   End If
   
   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For i = 0 To 2
         
         If Not IsNull(RsTemp.Fields(i)) Then
            txtCaseField(i + (Index + 1) * 3 + 36) = RsTemp.Fields(i)
         Else
            txtCaseField(i + (Index + 1) * 3 + 36) = ""
         End If
         
      Next
   End If
End Sub

'2010/1/8 ADD BY SONIA
Private Sub Combo3_Validate(Cancel As Boolean)
'Memo by Lydia 2019/05/09 FCP-061061衍生設計案的命名記錄處理方式
'ex.FCP-61061~63因為母案是舊案，所以程序沒有在新案建檔設工程師組別，直到中說(201,209,210)進度設承辦人(F2X)Trigger會自動更新ST16為基本檔的工程師組別。問過敏莉，目前暫無提申後又改組的情況。
'end 2019/05/09
   If Combo3 <> "" Then
      Combo3 = Left(Combo3, 1) + "." + PUB_GetFCPGrpName(Left(Combo3, 1))
      If Combo3 = Left(Combo3, 1) + "." Then
         Combo3 = Left(Combo3, 1)
         Cancel = True
         Combo3.SetFocus
      End If
      'Added by Lydia 2018/08/15 (中說未發文) 改工程師組別,改預設語系
      If fraTrans01.Visible = True And m_TF01pty & m_TF01cp27 = "201" And Left(Combo3.Tag, 1) <> Left(Combo3.Text, 1) Then
           If Left(Combo3.Text, 1) = "3" Then
                cboSource.ListIndex = 1 '日文
           Else
                cboSource.ListIndex = 0 '英文
           End If
           Call cboSource_Validate(False)
           If Trim(cboTarget.Text) = "" Then 'Added by Lydia 2019/11/28 判斷無設定值,才預設翻譯語種
                If pa(1) = "P" Then 'P案預設為簡體中文
                    cboTarget.ListIndex = 1
                Else
                    cboTarget.ListIndex = 0
                End If
                Call cboTarget_Validate(False)
           End If
      End If
      'end 2018/08/15
   End If
End Sub
'2010/1/8 end

'Add by Amy 2013/05/20
Private Sub Combo4_Click()
   Dim strMain As String, i As Integer
   strMain = Replace(Right(Combo4.Text, 11), ")", "")
   For i = 0 To 2
      txtInvField(i).Text = ""
      txtInvField(i).Tag = "" 'Add By Sindy 2015/3/5
   Next
   Call InvFieldEnabled(True)  'Added by Lydia 2022/03/25 控制發明人欄位是否可點選
   
   cmdUpdRow.Enabled = False 'Add By Sindy 2015/3/5
   cmdAddRow.Enabled = True 'Add By Sindy 2015/3/5
   If Len(strMain) > 0 Then
'      cmdUpdRow.Enabled = True 'Add By Sindy 2015/3/5
      If ClsLawGetInventor(strMain, strExc) Then
         For i = 0 To 2
            txtInvField(i).Text = strExc(i + 1)
            txtInvField(i).Tag = txtInvField(i).Text 'Add By Sindy 2015/3/5
         Next
         'Add By Sindy 2015/12/4
         txtIN11 = strExc(6)
         Call txtIN11_Validate(False)
         '2015/12/4 END
         Call InvFieldEnabled(False)  'Added by Lydia 2022/03/25 控制發明人欄位是否可點選
      End If
   End If
End Sub

'Added by Lydia 2019/10/25 翻譯瑕疵備註之選單
Private Sub Combo6_Validate(Cancel As Boolean)
    If Combo6.Tag <> Combo6.Text Then
       txtTF(37).Text = txtTF(37).Text & IIf(Trim(txtTF(37).Text) <> "", "、", "") & Combo6.Text
    End If
    Combo6.Tag = Combo6.Text
End Sub

Private Sub Command1_Click()
 Dim i As Integer
 
 SSTab1.Tab = 0 'Add By Amy 2013/05/20
 cmdPrtContact.Enabled = False 'Add by Amy 2014/07/04
      
 'Modify By Amy 2013/05/14 修改輸入少於6碼 會顯示案件進度的錯誤
 'If Text1.Text = "" Or Text2.Text = "" Then
   If text1.Text = "" Or Len(Text2.Text) <> 6 Then
      MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
      text1.SetFocus
      Exit Sub
   End If
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
   pa(1) = text1:   pa(2) = Text2
   pa(3) = Text3:   pa(4) = Text4
   FormClear
   'Add By Sindy 2014/11/10
   cmdAddRow.Tag = ""
   cmdDelRow.Tag = ""
   GRD1.Clear
   Call SetGrd(GRD1)
   '2014/11/10 END
   pPrevRow = 0 'Added by Lydia 2024/12/03
   
   'Added by Morgan 2015/7/21 考慮案件換業務區,FMP案抓最新收文業務區判斷 P-106433
   If pa(1) <> "FCP" And pa(1) <> "FG" Then
      strExc(0) = "select cp12 from CaseProgress where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " order by cp66 desc,cp67 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Left(RsTemp("cp12"), 1) <> "F" Then
            MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
            text1.SetFocus
            Exit Sub
         End If
      End If
   End If
   'end 2015/7/21
      
   Select Case pa(1)
      Case "FCP", "P", "CFP"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then
            'Add by Amy 2013/05/20
            SSTab1.TabEnabled(6) = True
'            SSTab1.TabEnabled(7) = True
            PatentShow
         End If
      Case "FG", "PS", "CPS"
         If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
            'Add by Amy 2013/05/20
            SSTab1.TabEnabled(6) = False
'            SSTab1.TabEnabled(7) = False
            ServiceShow
         End If
   End Select
   If Left(pa(26), 6) = "X27766" And Text44 <> "" And Text46 = "" And Text47 = "" Then
      Text46 = "*Murata's reference number for the U.S. Patent application is"
      Text47 = "*Corresponding Japanese Patent Application number"
   End If
   m_CP60 = "": m_CP09 = "": m_CP27 = "": m_CP10 = ""
   m_CP05 = "" 'Added by Lydia 2022/08/04
   m_CP57 = "" 'Add by Sindy 2023/4/18
   'Modify By Amy 2013/05/14 增加電子送件欄位
   'strExc(0) = "SELECT CP06,CP07,CP64,CP10,CP37,CP38,CP39,CP60,CP09,CP12 FROM CaseProgress WHERE " & _
      ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP31='Y'"
   'Modify by Amy 2014/07/04 +CP27
   'Modify By Sindy 2015/12/29 + CP141,CP142
   'Modified by Lydia 2022/08/04 +CP05
   strExc(0) = "SELECT CP06,CP07,CP64,CP10,CP37,CP38,CP39,CP60,CP09,CP12,CP118,CP27,CP141,CP142,CP164,CP05,CP57 FROM CaseProgress WHERE " & _
      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP31='Y'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         If IsNull(.Fields(0)) = False Then Text8 = TransDate(.Fields(0), 1)
         If IsNull(.Fields(1)) = False Then Text9 = TransDate(.Fields(1), 1)
         Text8.Tag = Text8 'Added by Morgan 2014/11/19
         Text9.Tag = Text9 'Added by Morgan 2014/11/19
         If IsNull(.Fields(2)) = False Then Text19 = .Fields(2)
         Text19.Tag = Text19 'Added by Lydia 2018/05/09 新案進度的備註
         If IsNull(.Fields(3)) = False Then m_CP10 = .Fields(3)
         m_CP05 = "" & .Fields("CP05") 'Added by Lydia 2022/08/04
         If pa(1) = "FCP" And pa(23) <> "1" Then
            If IsNull(.Fields(4)) = False Then Text5 = .Fields(4)
            If IsNull(.Fields(5)) = False Then Text6 = .Fields(5)
            If IsNull(.Fields(6)) = False Then Text7 = .Fields(6)
         End If
         Text5.Tag = Text5.Text: Text6.Tag = Text6.Text: Text7.Tag = Text7.Text 'Added by Lydia 2018/05/10
         
         m_CP60 = "" & .Fields("CP60")
         m_CP09 = "" & .Fields("CP09")
         If IsNull(.Fields("CP118")) = False Then Text25 = .Fields("CP118")
         Text25.Tag = Text25 'Added by Lydia 2018/05/09 新案進度的電子送件
         
         'Add By Sindy 2023/4/18
         'Modified by Lydia 2021/11/03 C類來函客戶指定送件日：不會有CP141 (對智慧局)
         'If "" & .Fields("CP141") = "3" And Val("" & .Fields("CP142")) > 0 Then
         If Val("" & .Fields("CP142")) > 0 Then
            txtCP142.Text = Val("" & .Fields("CP142")) - 19110000
         End If
         If "" & .Fields("CP164") = "1" Then
            Option1(0).Value = True
         ElseIf "" & .Fields("CP164") = "2" Then
            Option1(1).Value = True
         ElseIf "" & .Fields("CP164") = "3" Then
            Option1(2).Value = True
         End If
         '2023/4/18 END
      End With
      cmdOK(0).Enabled = True
      cmdOK(1).Enabled = True
      cmdOK(3).Enabled = True
       'Added by Lydia 2018/06/27
      cmdOK(4).Enabled = True
      Command2(6).Enabled = True
      
      'Add by Amy 2014/07/04 新申請(101,102,103,125)已發文者可重印簡易聯絡單
      If InStr("101,102,103,125", m_CP10) > 0 And Not IsNull(RsTemp.Fields("CP27")) Then
            cmdPrtContact.Enabled = True
      End If
      'end 2014/07/04
      m_CP27 = "" & RsTemp.Fields("CP27") 'Added by Morgan 2015/10/19
      m_CP57 = "" & RsTemp.Fields("CP57") 'Add by Sindy 2023/4/18
      'Added by Lydia 2018/06/07 翻譯分案無紙化：開啟資料夾
      cmdOpen(0).Enabled = True
      cmdOpen(1).Enabled = True
      'end 2018/06/07
   End If
    
   'Added by Lydia 2017/05/17 有新案翻譯,顯示原文字數、相似度、相似案號
   'Modified by Lydia 2018/06/01 檢視中說和核對中說也要帶入相似度和相似案號
   'strExc(0) = "SELECT CP09,B.* FROM CaseProgress A,TransFee B WHERE " & ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='201' AND CP159=0 AND CP09=TF01(+)"
   'Modified by Lydia 2019/06/28 +CP60
   strExc(0) = "SELECT CP09,CP10,CP14,CP27,CP60,B.* FROM CaseProgress A,TransFee B " & _
                     "WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 in ('201','209','235') AND CP159=0 AND CP09=TF01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If "" & RsTemp.Fields("CP09") <> "" Then
         'Modified by Lydia 2018/06/07 翻譯分案無紙化：欄位改成txtTF
'         Text33(16).Visible = True: Text33(17).Visible = True: Text33(18).Visible = True
'         Label26(2).Visible = True: Label26(3).Visible = True: Label26(4).Visible = True
'         Text33(16).Text = "" & RsTemp.Fields("TF23")
'         Text33(16).Tag = "" & RsTemp.Fields("CP09") '記錄收文號
'         Text33(17).Text = "" & RsTemp.Fields("TF19")
'         Text33(18).Text = "" & RsTemp.Fields("TF20")
'         'Added by Lydia 2017/12/01 記錄原本相似度、相似案號
'         m_TF23 = Val("" & RsTemp.Fields("TF23"))
'         m_TF19 = Val("" & RsTemp.Fields("TF19"))
'         m_TF20 = "" & RsTemp.Fields("TF20")
'         'end 2017/12/01
         m_TF01 = "" & RsTemp.Fields("CP09")
         m_TF01pty = "" & RsTemp.Fields("CP10")
         m_TF01cp14 = "" & RsTemp.Fields("CP14")
         m_TF01cp27 = "" & RsTemp.Fields("CP27")
         m_TF01cp60 = "" & RsTemp.Fields("CP60") 'Added by Lydia 2019/06/28
         fraTrans01.Visible = True
         fraTrans02.Visible = True
         fraTrans03.Visible = True
         For Each m_otxt In txtTF
             Select Case m_otxt.Index
                  'Modified by Lydia 2018/09/28 +只交Claim期限 32
                  Case 26, 32 '交稿期限
                       m_otxt.Text = TransDate("" & RsTemp.Fields("TF" & Format(m_otxt.Index, "00")), 1)
                  Case Else
                       m_otxt.Text = "" & RsTemp.Fields("TF" & Format(m_otxt.Index, "00"))
             End Select
             m_otxt.Tag = m_otxt.Text
         Next
         
         '原文語種
         If "" & RsTemp.Fields("TF27") <> "" Then
               cboSource.ListIndex = Val(RsTemp.Fields("TF27")) - 1
         End If
         cboSource.Tag = cboSource.Text
         '翻譯語種
         If "" & RsTemp.Fields("TF28") <> "" Then
               cboTarget.ListIndex = Val(RsTemp.Fields("TF28")) - 1
         End If
         cboTarget.Tag = cboTarget.Text
         
         '有新案翻譯:預設非日文組預設"英文翻中文"，日文組預設"日文翻中文"
         'Modified by Lydia 2018/08/10 有工程師組別才預設
         'Modified by Lydia 2018/08/15 debug (ex.FCP-059389 Gill分工程師組別,預設為英文)
         'If m_TF01pty = "201" And Trim(Combo3.Text) = "" Then
         If m_TF01pty = "201" And Trim(Combo3.Text) <> "" Then
            If "" & RsTemp.Fields("TF27") = "" Then
                  If Combo3.Text <> "" And Left(Combo3.Text, 1) = "3" Then
                         cboSource.ListIndex = 1 '日文
                  Else
                         cboSource.ListIndex = 0 '英文
                  End If
                  Call cboSource_Validate(False)
            End If
            If "" & RsTemp.Fields("TF28") = "" Then
                  If pa(1) = "P" Then 'P案預設為簡體中文
                      cboTarget.ListIndex = 1
                  Else
                      cboTarget.ListIndex = 0
                  End If
                  Call cboTarget_Validate(False)
            End If
         End If
         '待比對
         Chk01.Value = 0
         If "" & RsTemp.Fields("TF29") = "Y" Then
              Chk01.Value = 1
         End If
         Chk01.Tag = "" & RsTemp.Fields("TF29")
         '待英文本翻譯／英文本收文號
         Chk02.Value = 0
         If "" & RsTemp.Fields("TF30") = "Y" Then
              Chk02.Value = 1
              txtTF(30).Text = ""
         End If
         '未提申先翻譯
         Chk03.Value = 0
         If "" & RsTemp.Fields("TF31") = "Y" Then
              Chk03.Value = 1
         End If
         Chk03.Tag = "" & RsTemp.Fields("TF31")
         '中說4個月不得延
         Chk04.Value = 0
         If "" & RsTemp.Fields("TF33") = "Y" Then
              Chk04.Value = 1
         End If
         'Added by Lydia 2018/08/24 暫不翻譯
         Chk05.Value = 0
         If "" & RsTemp.Fields("TF34") = "Y" Then
              Chk05.Value = 1
         End If
         'end 2018/08/24
         
         If m_TF01pty = "201" Then '限新案翻譯可輸入全部欄位
             fraTrans01.Enabled = True
             fraTrans02.Enabled = True
             fraTrans03.Enabled = True
             'fraTrans04.Enabled = True 'Added by Lydia 2019/10/25 翻譯瑕疵備註
         Else '其他翻譯-只可輸入相似度和外文本頁數
             fraTrans01.Enabled = False
             fraTrans02.Enabled = True
             fraTrans03.Enabled = True
             'fraTrans04.Enabled = False 'Added by Lydia 2019/10/25 翻譯瑕疵備註
         End If
         'end 2018/06/07
      End If
   End If
   'end 2016/05/17
   
   'Added by Lydia 2018/06/07 翻譯分案無紙化：先抓工程師-命名作業的資料
    strExc(0) = "SELECT TCT01,TCT02,TCT03,TCT04,TCT07,TCT10,CP06,CP07,CP27,TCT23,TCT24,TCT16,TCT17,TCT27,TCT28 " & _
                      "FROM CaseProgress,TransCaseTitle " & _
                      "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' " & _
                      "AND CP10 IN (" & NewCasePtyList & ") AND CP159=0 AND CP09=TCT01 ORDER BY CP09 DESC"
    intI = 1
    strExc(2) = ""
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
         m_TCT01 = "" & RsTemp.Fields("TCT01")
         m_TCT04 = "" & RsTemp.Fields("TCT04")
         m_TCT10 = "" & RsTemp.Fields("TCT10")
         m_TCT16 = "" & RsTemp.Fields("TCT16")
         m_TCT17 = "" & RsTemp.Fields("TCT17")
         m_TCT23 = "" & RsTemp.Fields("TCT23")
         m_TCT24 = "" & RsTemp.Fields("TCT24")
         m_TCT27 = "" & RsTemp.Fields("TCT27")
         m_TCT28 = "" & RsTemp.Fields("TCT28")
         '代入相似度
         If Val(txtTF(19).Text) = 0 And Val(m_TCT24) <> 0 Then
              txtTF(19).Text = m_TCT24
         End If
         '代入相似案號
         If Trim(txtTF(20).Text) = "" And m_TCT23 <> "" Then
              txtTF(20).Text = m_TCT23
         End If
    End If
    
   'Added by Lydia 2019/07/04 FCP衍生設計新案是否要走新案命名
            '108/5/8 發現FCP-61061~63因為母案是舊案，所以程序沒有在新案建檔設工程師組別，直到中說(201,209,210)進度設承辦人(F2X)Trigger會自動更新ST16為基本檔的工程師組別。問過敏莉，目前暫無提申後又改組的情況。
   '修改1.在分案輸入母案案號時，原本複製母案名稱和優先權等資料，現在一併將命名記錄刪除；
   '修改2.若有需要變更名稱者，請程序在新案建檔勾選重新產生命名記錄存檔後，再重設工程師組別並且發命名通知Email；
   'Remove by Lydia 2019/07/30 修改
   '1.在分案作業輸入母案案號時，一併將母案的組別帶到子案；取消刪除命名記錄之作業。
   '2.若母案是新案則子案有需要變更名稱者，請程序在新案建檔再重設工程師組別並且發命名分組通知Email；
   '3.若母案是舊案則待子案發文時，檢查命名記錄尚未分組才刪除。
   'If Text1 = "FCP" And m_CP10 = "125" And m_TCT01 = "" Then
   '    ChkAddTct.Visible = True
   'End If
   'end 2019/07/30
   
    '預設待比對(未發文,未分案)
    'If Chk01.Visible = True Then 'Memo by Lydia 2018/08/07 待比對控制,工程師協調8/13上線
        'Modified by Lydia 2018/10/05 FCP案工程師主管也有認翻譯;FMP案才有預設主管的情況
        'If SSTab1.TabVisible(7) = True And Chk01.Value = vbUnchecked And m_TCT10 <> "" And txtTF(19) <> "" And txtTF(20) <> "" And m_TF01pty = "201" _
                And Val(m_TF01cp27) = 0 And (m_TF01cp14 = "" Or (m_TF01cp14 <> "" And InStr(m_GrpManList, m_TF01cp14) > 0)) Then
        'Modified by Lydia 2022/11/08 有相似度或相似案號就檢查; ex.FCP-67931只有相似案號沒有相似度
        'If SSTab1.TabVisible(7) = True And Chk01.Value = vbUnchecked And m_TCT10 <> "" And txtTF(19) <> "" And txtTF(20) <> "" And m_TF01pty = "201" _
                And Val(m_TF01cp27) = 0 And (m_TF01cp14 = "" Or (pa(1) = "P" And m_TF01cp14 <> "" And InStr(m_GrpManList, m_TF01cp14) > 0)) Then
        If SSTab1.TabVisible(7) = True And Chk01.Value = vbUnchecked And m_TCT10 <> "" And (txtTF(19) <> "" Or txtTF(20) <> "") And m_TF01pty = "201" _
                And Val(m_TF01cp27) = 0 And (m_TF01cp14 = "" Or (pa(1) = "P" And m_TF01cp14 <> "" And InStr(m_GrpManList, m_TF01cp14) > 0)) Then
             'Modified by Lydia 2018/09/19 +PDF
             'If Dir(strResPath & "\" & pa(1) & "*" & pa(2) & ".res.doc*") = "" Then
             'Modified by Ldia 2019/04/15 案號可以6碼或5碼
             'strExc(1) = Dir(strResPath & "\" & pa(1) & "*" & pa(2) & ".res.doc*")
             'If strExc(1) = "" Then strExc(1) = Dir(strResPath & "\" & pa(1) & "*" & pa(2) & ".res.pdf")
             strExc(1) = Dir(strResPath & "\" & pa(1) & "*" & Val(pa(2)) & ".res.doc*")
             If strExc(1) = "" Then strExc(1) = Dir(strResPath & "\" & pa(1) & "*" & pa(2) & ".res.pdf")
             'end 2019/04/15
             If strExc(1) = "" Then
             'end 2018/09/19
                 MsgBox "工程師尚未上傳比對結果，預設為待比對 !", vbCritical
                 Chk01.Value = vbChecked
                 SSTab1.Tab = 7
             End If
        End If
    'End If
    'end 2018/06/07
End Sub

'Added by Morgan 2013/7/5
Private Sub Command2_Click(Index As Integer)
   Load frm040109_1
   Set frm040109_1.frmParent = Me
   frm040109_1.txtCode(0) = pa(1)
   frm040109_1.txtCode(1) = pa(2)
   frm040109_1.txtCode(2) = pa(3)
   frm040109_1.txtCode(3) = pa(4)
   frm040109_1.txtCode(8) = "1"
   If frm040109_1.ChkExist = True Then
      frm040109_1.Move frm040109_1.Left, frm040109_1.Top - 550
      If MsgBox("該案已建立一案兩請關聯是否修改？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         Unload frm040109_1
      Else
         Me.Hide
         frm040109_1.txtCode(8) = "2"
         frm040109_1.Show
         m_bolChkPA91 = True 'Added by Morgan 2017/8/22
      End If
   'Added by Lydia 2018/06/25 提示
   Else
      If MsgBox("在建立一案兩請關聯前，發明案的發明人、代表人、優先權資料是否輸入完畢？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
          Unload frm040109_1
      End If
   'end 2018/06/25
   End If
End Sub

Private Sub Form_Activate()
   'Added by Morgan 2017/8/22 刪除一案兩請資料可能會回寫案件備註,要檢查是否有變更
   If m_bolChkPA91 Then
      strExc(0) = "select pa91 from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If pa(91) <> "" & RsTemp(0) Then
            If pa(91) <> Text10 Then
               MsgBox "案件備註已更新，請確認！", vbExclamation
            End If
            Text10 = "" & RsTemp(0)
         End If
      End If
      m_bolChkPA91 = False
   End If
   'end 2017/8/22
End Sub

Private Sub Form_Load()
   'Memo by Amy 2025/08/06  不續辦但准通知 改為 後續准駁簡單報告
   'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    txtCaseField(39).MaxLength = Pub_MaxCEL10
    txtCaseField(40).MaxLength = Pub_MaxCEL11
    txtCaseField(42).MaxLength = Pub_MaxCEL10
    txtCaseField(43).MaxLength = Pub_MaxCEL11
    txtCaseField(45).MaxLength = Pub_MaxCEL10
    txtCaseField(46).MaxLength = Pub_MaxCEL11
    txtCaseField(48).MaxLength = Pub_MaxCEL10
    txtCaseField(49).MaxLength = Pub_MaxCEL11
    txtCaseField(51).MaxLength = Pub_MaxCEL10
    txtCaseField(52).MaxLength = Pub_MaxCEL11
    txtCaseField(54).MaxLength = Pub_MaxCEL10
    txtCaseField(55).MaxLength = Pub_MaxCEL11
    txtCaseField(57).MaxLength = Pub_MaxCEL10
    txtCaseField(58).MaxLength = Pub_MaxCEL11
    txtCaseField(60).MaxLength = Pub_MaxCEL10
    txtCaseField(61).MaxLength = Pub_MaxCEL11
    txtCaseField(63).MaxLength = Pub_MaxCEL10
    txtCaseField(64).MaxLength = Pub_MaxCEL11
    txtCaseField(66).MaxLength = Pub_MaxCEL10
    txtCaseField(67).MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   MoveFormToCenter Me
   intWhere = 國外_FC
   FormClear
   SSTab1.Tab = 0
   SendKeys "{Tab}"
     
   ReDim pa(TF_PA) '陣列大小改用全域變數
   
   'Added by Lydia 2017/11/17 設計案屬性
   Combo5.Clear
   For intI = 1 To 4
      Combo5.AddItem intI & "." & PUB_GetCaseAttributeName(Trim(intI), "3")
   Next
   'end 2017/11/17
   
   'Added by Lydia 2018/06/07 翻譯分案無紙化-語種設定
   cboSource.Clear
   cboSource.AddItem "1." & Pub_GetTransFeeL("1", "1")
   cboSource.AddItem "2." & Pub_GetTransFeeL("1", "2")
   cboSource.AddItem "3." & Pub_GetTransFeeL("1", "3")
   cboSource.AddItem "4." & Pub_GetTransFeeL("1", "4")  'Added by Lydia 2024/02/21
   cboTarget.Clear
   cboTarget.AddItem "1." & Pub_GetTransFeeL("2", "1")
   cboTarget.AddItem "2." & Pub_GetTransFeeL("2", "2")
   fraTrans01.BackColor = &H8000000F
   fraTrans02.BackColor = &H8000000F
   fraTrans03.BackColor = &H8000000F
   fraTrans04.BackColor = &H8000000F 'Added by Lydia 2019/10/25 翻譯瑕疵備註
   m_GrpManList = Pub_GetSt16Man(True)  '所有工程師主管(含F編號)
   strResPath = Pub_GetSpecMan("FCP相似比對結果暫存")
   'end 2018/06/07
   
   Frame3.BackColor = &H8000000F 'Added by Lydia 2022/03/25
   
   'Added by Lydia 2018/05/21 翻譯分案無紙化 (前一畫面-傳入案號)
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" And m_Case(1) & m_Case(2) <> "" Then
        text1.Text = m_Case(1)
        Text2.Text = m_Case(2)
        Text3.Text = m_Case(3)
        Text4.Text = m_Case(4)
       Call Command1_Click
   End If
   
   FCP檢視中說必輸原文字數 = Pub_GetSpecMan("FCP檢視中說必輸原文字數") 'Added by Lydia 2019/06/28
   
   'Added by Lydia 2019/10/25 翻譯瑕疵備註之選單
   Combo6.Clear
   Combo6.AddItem "漏譯"
   Combo6.AddItem "誤譯"
   Combo6.AddItem "贅字太多"
   Combo6.AddItem "語句不通順"
   Combo6.AddItem "其他(自行輸入內容)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Added by Lydia 2016/03/15 發明人輸入比對兼自動代入(模糊比對)
    ' 刪除串列結構
    If m_InventorListCount > 0 Then
       Erase m_InventorList
    End If
    m_InventorListCount = 0
   'end 2016/03/15
   
   'Added by Lydia 2018/05/21 翻譯分案無紙化 (前一畫面-傳入案號)
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
       If TypeName(m_PrevForm) = "frm060122" Then
             m_PrevForm.cmdState = 0
             Call m_PrevForm.PubShowNextData
       End If
       m_PrevForm.Show
   End If
   
   Set frm060102 = Nothing
'   Unload frm060102_1
'   Unload frm060102_2
End Sub

Private Sub PatentShow()
Dim i As Integer, j As Integer
Dim strTmp As String 'Add by Amy 2013/05/20
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   Text5 = pa(5):   Text6 = pa(6):   Text7 = pa(7):   Text10 = pa(91)
   Text11 = pa(89): Text12 = pa(78): Text13 = pa(85): Text14 = pa(49)
   Text15 = pa(50): Text16 = pa(71): Text17 = pa(70): Text18 = pa(107)
   Text23 = pa(146): Text24 = pa(142)
   
   '2008/10/15 add by Toni 增加FCP工程師組別
   If pa(150) = "" Then
      Combo3 = ""
   Else
      Combo3 = pa(150) + "." + PUB_GetFCPGrpName(pa(150))
   End If
   'end 200/1015
   Combo3.Tag = Combo3.Text 'Added by Lydia 2017/11/14 FCP案件命名電子化
   
   'Added by Lydia 2017/11/17 設計案屬性
   If pa(1) = "FCP" And pa(8) = "3" Then
      Label26(5).Visible = True
      Combo5.Visible = True
      If pa(158) = "" Then
         Combo5.Text = ""
      Else
         Combo5.Text = pa(158) + "." + PUB_GetCaseAttributeName(pa(158), "3")
      End If
      Combo5.Tag = Combo5.Text
   Else
      Label26(5).Visible = False
      Combo5.Visible = False
   End If
   'end 2017/11/17
   
   '代表人１
   For i = 79 To 84
      txtCaseField(i - 40) = pa(i)
   Next
   '代表人２
   For i = 109 To 132
      txtCaseField(i - 64) = pa(i)
   Next
      
   For i = 0 To 9
      Combo2(i).Clear
      Combo2(i).AddItem ""
   Next
   
   For i = 26 To 30
      If pa(i) <> "" Then
         'Modified by Morgan 2017/1/12 若無英文要帶日文--何淑華
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(pa(i))
         strExc(0) = "SELECT NVL(CU40,NVL(CU41,CU39)) CU40,NVL(CU43,NVL(CU44,CU42)) CU43,NVL(CU46,NVL(CU47,CU45)) CU46,NVL(CU49,NVL(CU50,CU48)) CU49,NVL(CU52,NVL(CU53,CU51)) CU52,NVL(CU55,NVL(CU56,CU54)) CU55 FROM CUSTOMER WHERE " & ChgCustomer(pa(i))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2((i - 26) * 2).AddItem pa(i) & "-" & j & strExc(0)
               Combo2((i - 26) * 2 + 1).AddItem pa(i) & "-" & j & strExc(0)
            Next
         End If
      End If
   Next
        
   For i = 0 To 5
      Text33(i) = pa(i + 51)
   Next
   Text33(15) = pa(139) 'Add by Morgan 2006/10/19 聯絡人部門(日)
   For i = 9 To 13
      If pa(i + 17) <> "" Then
         Text33(i) = pa(i + 17)
         ChgType i
         'Add by Amy 2013/05/20 抓取申請人編號 (for  增加發明人)
        If Len(pa(i + 17)) < 9 Then
            strTmp = strTmp & "'" & Left(pa(i + 17), 8) & String(8 - Len(pa(i + 17)), "0") & "',"
        'Added by Morgan 2015/11/20
        Else
            strTmp = strTmp & "'" & Left(pa(i + 17), 8) & "',"
        'end 2015/11/20
        End If
      End If
   Next
   
   'Add by Amy 2013/05/20 增加發明人
   If strTmp <> "" Then strTmp = Left(strTmp, Len(strTmp) - 1)
   GetCombo4Data strTmp
'   For i = 60 To 69
'      If pa(i) <> "" Then
'         SetCombo4Data i - 60, GetInventorName(pa(i))
'      Else
'         Combo4(i - 60).ListIndex = 0
'      End If
'   Next
'   'end 2013/05/20
   'Add By Sindy 2014/11/10
   'Modified by Lydia 2024/12/03 +IN11
   StrSQLa = "SELECT '' as V,pi06 as 發明人編號,in04 as 中文名稱,in05 as 英文名稱,in06 as 日文名稱,na03 as 國籍,'' as 申請人1, IN11" & _
             " from PatentInventor,Inventor,nation" & _
             " where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
             " and substr(pi06,1,8)=in01(+) and substr(pi06,9,2)=in02(+)" & _
             " and in11=na01(+)" & _
             " order by pi05 asc"
   If rsA.State <> adStateClosed Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      Set GRD1.Recordset = rsA
   End If
   '2014/11/10 END
   
   If pa(75) <> "" Then
      Text33(14) = pa(75)
      ChgType 14
      
      Select Case pa(85)
         Case 1
            strExc(0) = "FA07,FA52"
         Case 2
            strExc(0) = "FA08,FA53"
         Case 3
            strExc(0) = "FA09,FA54"
         Case Else
            strExc(0) = "FA08,FA53"
      End Select
      
      strExc(0) = "SELECT " & strExc(0) & " FROM FAGENT WHERE " & ChgFagent(pa(75))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If IsNull(RsTemp.Fields(0)) Then
            strExc(0) = ""
         Else
            strExc(0) = "-" & RsTemp.Fields(0)
         End If
         Combo1(0).AddItem pa(75) & "-1" & strExc(0)
         Combo1(1).AddItem pa(75) & "-1" & strExc(0)
         If IsNull(RsTemp.Fields(1)) Then
            strExc(0) = ""
         Else
            strExc(0) = "-" & RsTemp.Fields(1)
         End If
         Combo1(0).AddItem pa(75) & "-2" & strExc(0)
         Combo1(1).AddItem pa(75) & "-2" & strExc(0)
      End If
   Else
      For i = 26 To 30
         If pa(i) <> "" Then
            Select Case pa(85)
               Case 1
                  strExc(0) = "CU58,CU61"
               Case 2
                  strExc(0) = "CU59,CU62"
               Case 3
                  strExc(0) = "CU60,CU63"
               Case Else
                  strExc(0) = "CU59,CU62"
            End Select
            strExc(0) = "SELECT " & strExc(0) & " FROM CUSTOMER WHERE " & ChgCustomer(pa(i))
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               For j = 1 To 2
                  If IsNull(RsTemp.Fields(j - 1)) Then
                     strExc(0) = ""
                  Else
                     strExc(0) = "-" & RsTemp.Fields(j - 1)
                  End If
                  Combo1(0).AddItem pa(i) & "-" & j & strExc(0)
                  Combo1(1).AddItem pa(i) & "-" & j & strExc(0)
               Next
            End If
         End If
      Next
   End If
   If Combo1(0).ListCount > 0 And Text33(0) = "" And Text33(1) = "" And Text33(2) = "" Then Combo1(0).ListIndex = 0
   If Combo1(1).ListCount > 0 And Text33(3) = "" And Text33(4) = "" And Text33(5) = "" Then Combo1(1).ListIndex = 0
   Text27 = pa(88): ChgType (27):    Text28 = pa(76): ChgType (28)
   Text29 = pa(106): Text26 = pa(77)
   Text30 = pa(105): ChgType (30)
   Text31 = pa(48): Text32 = pa(90)
    Me.Text20.Text = pa(133): ChgType (20)
    Me.Text21.Text = pa(134): ChgType (21)
    Me.Text22.Text = pa(135)
   For i = 98 To 100
      Text33(i - 92) = pa(i)
   Next
   Text42 = pa(86): ChgType (42):    Text44 = pa(101): ChgType (44)
   Text43 = pa(87): Text45 = pa(102): Text46 = pa(103):  Text47 = pa(104)
   
   'Add by Morgan 2008/11/14
   '新增欄位改用陣列idex對應欄位序號，以後新增就不必再改。
   For Each m_otxt In txtPA
      m_otxt = pa(m_otxt.Index)
      m_otxt.Tag = m_otxt.Text 'Added by Lydia 2018/08/27
   Next
   'end 2008/11/14
   'Added by Lydia 2018/10/17
   Label1(1).Visible = True
   txtPA(63).Visible = True
   'end 2018/10/17
   
   'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地
   If pa(1) = "FCP" Or pa(1) = "P" Then
       ' 已放在原始檔區
       If PUB_ChkCPExist(pa, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
            cmdOpen(0).Caption = "原始檔"
            cmdOpen(0).Tag = strExc(1)
       Else
            cmdOpen(0).Caption = "外文本"
            cmdOpen(0).Tag = ""
       End If
      'Added by Lydia 2020/02/21 預設「名稱有特殊字」
      ChkPA174.Visible = True
      CmdPA174.Visible = True
      If pa(174) = "Y" Then
          ChkPA174.Value = vbChecked
          ChkPA174.Tag = pa(174)
      End If
      'end 2020/02/21
      'Added by Lydia 2021/04/09 預設「有序列表」
      ChkPA175.Visible = True
      If pa(175) = "Y" Then
          ChkPA175.Value = vbChecked
          ChkPA175.Tag = pa(175)
      End If
      'end 2021/04/09
   End If
   'end 2020/01/20
   
   'Modify by Amy 2014/04/11 +pd08,pd09
   If Not ClsPDReadPriority(pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)) Then
        MsgBox "優先權資料讀取失敗！！"
   End If
   
   'Added by Lydia 2018/08/24 固定報價
   If txtPA(62).Text = "Y" Then
       Chk06.Value = 1
   End If
   
End Sub

'Add By Sindy 2014/11/10
Private Sub SetGrd(tmpGrd As MSHFlexGrid)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   'Modified by Lydia 2024/12/03 + 7 => IN11
   '                        0    1             2           3           4           5       6          7
   arrGridHeadText = Array("V", "發明人編號", "中文名稱", "英文名稱", "日文名稱", "國籍", "申請人1", "IN11")
   arrGridHeadWidth = Array(200, 1100, 2000, 2000, 2000, 800, 0, 0)
   tmpGrd.Visible = False
   tmpGrd.Cols = UBound(arrGridHeadText) + 1
   tmpGrd.Rows = 2
   For iRow = 0 To tmpGrd.Cols - 1
      tmpGrd.row = 0
      tmpGrd.col = iRow
      tmpGrd.Text = arrGridHeadText(iRow)
      tmpGrd.ColWidth(iRow) = arrGridHeadWidth(iRow)
      tmpGrd.CellAlignment = flexAlignCenterCenter
   Next
   tmpGrd.Visible = True
End Sub

Private Sub ServiceShow()
 Dim i As Integer
   Text5 = pa(5):   Text6 = pa(6):   Text7 = pa(7):   Text10 = pa(18)
   Text12 = pa(33): Text13 = pa(34): Text14 = pa(31)
   Text33(1) = pa(30)
   Text33(4) = pa(75) 'Add by Morgan 2007/8/13
   If pa(8) <> "" Then Text33(9) = pa(8): ChgType 9
   If pa(58) <> "" Then Text33(10) = pa(58): ChgType 10
   If pa(59) <> "" Then Text33(11) = pa(59): ChgType 11
   
   Text33(12).Enabled = False:   Text33(13).Enabled = False
   
   For i = 0 To 8
      Text33(i).Enabled = False
   Next
   Text33(1).Enabled = True
   Text33(4).Enabled = True 'Add by Morgan 2007/8/13
   
   'Text10.Enabled = False 'Removed by Morgan 2011/11/4
   Text11.Enabled = False
   Text15.Enabled = False:   Text16.Enabled = False
   Text17.Enabled = False:   Text18.Enabled = False
   Text28.Enabled = False:   Text29.Enabled = False
   Text30.Enabled = False:   Text32.Enabled = False
   Text44.Enabled = False:   Text45.Enabled = False
   Text46.Enabled = False:   Text47.Enabled = False
    Me.Text21.Enabled = False: Me.Text22.Enabled = False
   Combo1(1).Enabled = False
   
   If pa(26) <> "" Then
      Text33(14) = pa(26)
      ChgType 14
      strExc(0) = "SELECT FA08,FA53 FROM FAGENT WHERE " & ChgFagent(pa(26))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If IsNull(RsTemp.Fields(0)) Then
            strExc(0) = ""
         Else
            strExc(0) = "-" & RsTemp.Fields(0)
         End If
         Combo1(0).AddItem pa(26) & "-1" & strExc(0)
         If IsNull(RsTemp.Fields(1)) Then
            strExc(0) = ""
         Else
            strExc(0) = "-" & RsTemp.Fields(1)
         End If
         Combo1(0).AddItem pa(26) & "-2" & strExc(0)
      End If
   Else
      For i = 8 To 66
         Select Case i
            Case 8, 58, 59, 65, 66
               If pa(i) <> "" Then
                  strExc(0) = "SELECT CU59,CU62 FROM CUSTOMER WHERE " & ChgCustomer(pa(i))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If IsNull(RsTemp.Fields(0)) Then
                        strExc(0) = ""
                     Else
                        strExc(0) = "-" & RsTemp.Fields(0)
                     End If
                     Combo1(0).AddItem pa(i) & "-1" & strExc(0)
                     If IsNull(RsTemp.Fields(1)) Then
                        strExc(0) = ""
                     Else
                        strExc(0) = "-" & RsTemp.Fields(1)
                     End If
                     Combo1(0).AddItem pa(i) & "-2" & strExc(0)
                  End If
               End If
         End Select
      Next
   End If
   
   Text26 = pa(27)
   Text31 = pa(29)
   Text42 = pa(35): ChgType (42)
   Text43 = pa(36)
   Text27 = pa(37): ChgType (27)
   Text20.Text = pa(67): ChgType (20)
    
   'Modify by Morgan 2009/9/7 +工程師組別
   If pa(79) = "" Then
      Combo3 = ""
   Else
      Combo3 = pa(79) + "." + PUB_GetFCPGrpName(pa(79))
   End If
   'end 200/1015
   
   Text24 = pa(80) 'Added by Morgan 2017/6/28
   txtPA(153) = pa(81) 'Added by Morgan 2017/6/28
   txtPA(154) = pa(82) 'Added by Morgan 2017/6/28
   txtPA(155) = pa(83) 'Added by Morgan 2017/6/28
   txtPA(159) = pa(84) 'Add by Morgan 2010/11/9
   'Added by Lydia 2018/10/17
   Label1(1).Visible = False
   txtPA(63).Visible = False
End Sub

Private Function FormSave() As Boolean
   Dim i As Integer, ii As Integer
   Dim strTmp(1 To 3) As String
   Dim stUpdates As String
   Dim m_PI06 As String 'Add by Amy 2013/05/20
   'Add by Morgan 2011/6/13
   Dim bolChkMemo605 As Boolean, bolChkMemo416 As Boolean
   Dim strOldMemo605 As String, strOldMemo416 As String
   Dim strNewMemo605 As String, strNewMemo416 As String
   Dim tmpArr As Variant 'Added by Lydia 2018/04/23
   Dim rsAD As New ADODB.Recordset 'Added by Lydia 2018/09/27
   Dim strErrMsg As String, strPassSql As String 'Added by Lydia 2020/03/17
   Dim m_UpdPA63TCT118 As String 'Added by Lydia 2024/10/22
   Dim intPI05 As Integer 'Add By Sindy 2025/2/11
   
On Error GoTo CheckingErr

   cnnConnection.BeginTrans
   
   'Add by Morgan 2011/6/13
   '若有期限則於更新資料前紀錄原來的備註
   'Modified by Morgan 2012/2/2 +pa75
   strExc(0) = "select pa26,pa27,pa28,pa29,pa30,pa75,np07 from nextprogress,patent" & _
      " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07 in (416,605)" & _
      " and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp("np07") = "416" Then
         bolChkMemo416 = True
         'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
         'For intI = 0 To 4
         '   If Not IsNull(RsTemp(intI)) Then
         '      'Modified by Morgan 2012/2/2
         '      'Modified by Morgan 2013/9/11 改抓設定檔
         '      'strOldMemo416 = PUB_Get416Memo(ChangeCustomerL(RsTemp(intI)), ChangeCustomerL("" & RsTemp("pa75")))
         '      strOldMemo416 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp("pa75")), ChangeCustomerL(RsTemp(intI)))
         '      If strOldMemo416 <> "" Then Exit For
         '   End If
         'Next
         strOldMemo416 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp.Fields("pa75")), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
         'end 2022/08/02
         
      ElseIf RsTemp("np07") = "605" Then
         bolChkMemo605 = True
         strExc(9) = PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1")
         'Modified by Morgan 2012/6/4 +pa26
         'Modified by Morgan 2013/9/11 改抓設定檔
         'strOldMemo605 = PUB_Get605Memo(strExc(9), RsTemp("pa26"), pa(1) & pa(2) & pa(3) & pa(4))
         'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
         'strOldMemo605 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp("pa26"))
         strOldMemo605 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
      End If
   End If
   'end 2011/6/13

   stUpdates = ""
   Select Case text1
      Case "FCP", "P", "CFP"
         'Modified by Lydia 2018/05/09 有修改才Update+判斷
'         pa(5) = Text5: stUpdates = stUpdates & ",pa05=" & CNULL(ChgSQL(pa(5)))
'         pa(6) = Text6: stUpdates = stUpdates & ",pa06=" & CNULL(ChgSQL(pa(6)))
'         pa(7) = Text7: stUpdates = stUpdates & ",pa07=" & CNULL(ChgSQL(pa(7)))
'         pa(48) = Text31: stUpdates = stUpdates & ",pa48=" & CNULL(ChgSQL(pa(48)))
'         'add by toni 2008/10/15 增加FCP工程師組別
'         pa(150) = Left(Combo3, 1): stUpdates = stUpdates & ",pa150=" & CNULL(ChgSQL(pa(150)))
'         'end 2008/10/15
         If pa(5) <> Text5 Then pa(5) = Text5: stUpdates = stUpdates & ",pa05=" & CNULL(ChgSQL(pa(5)))
         If pa(6) <> Text6 Then pa(6) = Text6: stUpdates = stUpdates & ",pa06=" & CNULL(ChgSQL(pa(6)))
         If pa(7) <> Text7 Then pa(7) = Text7: stUpdates = stUpdates & ",pa07=" & CNULL(ChgSQL(pa(7)))
         If pa(48) <> Text31 Then pa(48) = Text31: stUpdates = stUpdates & ",pa48=" & CNULL(ChgSQL(pa(48)))
         If pa(150) <> Trim(Left(Combo3, 1)) Then pa(150) = Left(Combo3, 1): stUpdates = stUpdates & ",pa150=" & CNULL(ChgSQL(pa(150)))
         'end 2018/05/09
         
         'Added by Lydia 2017/11/17 設計案屬性
         If Combo5.Visible = True And Combo5.Text <> Combo5.Tag Then
            pa(158) = Left(Combo5, 1): stUpdates = stUpdates & ",pa158=" & CNULL(ChgSQL(pa(158)))
         End If
         'end 2017/11/17
         
         If pa(49) <> Text14 Then 'Added by Lydia 2018/05/09 有修改才Update+判斷
            If Text14 = "" Then
               stUpdates = stUpdates & ",pa49=NULL"
            Else
               pa(49) = Text14: stUpdates = stUpdates & ",pa49=" & pa(49)
            End If
         End If 'end 2018/05/09
         
         If pa(50) <> Text15 Then 'Added by Lydia 2018/05/09 有修改才Update+判斷
            If Text15 = "" Then
               stUpdates = stUpdates & ",pa50=NULL"
            Else
               pa(50) = Text15: stUpdates = stUpdates & ",pa50=" & pa(50)
            End If
         End If 'end 2018/05/09
         'Modified by Lydia 2018/05/09 有修改才Update+判斷
         If pa(51) <> Text33(0) Then pa(51) = Text33(0): stUpdates = stUpdates & ",pa51=" & CNULL(ChgSQL(pa(51)))
         If pa(52) <> Text33(1) Then pa(52) = Text33(1): stUpdates = stUpdates & ",pa52=" & CNULL(ChgSQL(pa(52)))
         If pa(53) <> Text33(2) Then pa(53) = Text33(2): stUpdates = stUpdates & ",pa53=" & CNULL(ChgSQL(pa(53)))
         If pa(54) <> Text33(3) Then pa(54) = Text33(3): stUpdates = stUpdates & ",pa54=" & CNULL(ChgSQL(pa(54)))
         If pa(55) <> Text33(4) Then pa(55) = Text33(4): stUpdates = stUpdates & ",pa55=" & CNULL(ChgSQL(pa(55)))
         If pa(56) <> Text33(5) Then pa(56) = Text33(5): stUpdates = stUpdates & ",pa56=" & CNULL(ChgSQL(pa(56)))
         If pa(70) <> Text17 Then pa(70) = Text17: stUpdates = stUpdates & ",pa70=" & CNULL(ChgSQL(pa(70)))
         If pa(71) <> Text16 Then pa(71) = Text16: stUpdates = stUpdates & ",pa71=" & CNULL(ChgSQL(pa(71)))
         If pa(75) <> Text33(14) Then pa(75) = Text33(14): stUpdates = stUpdates & ",pa75=" & CNULL(ChangeCustomerL(pa(75)))
         If pa(76) <> Text28 Then pa(76) = Text28: stUpdates = stUpdates & ",pa76=" & CNULL(ChangeCustomerL(pa(76)))
         If pa(77) <> Text26 Then pa(77) = Text26: stUpdates = stUpdates & ",pa77=" & CNULL(ChgSQL(pa(77)))
         If pa(78) <> Text12 Then pa(78) = Text12: stUpdates = stUpdates & ",pa78=" & CNULL(ChgSQL(pa(78)))
         If pa(85) <> Text13 Then pa(85) = Text13: stUpdates = stUpdates & ",pa85=" & CNULL(ChgSQL(pa(85)))
         If pa(86) <> Text42 Then pa(86) = Text42: stUpdates = stUpdates & ",pa86=" & CNULL(ChangeCustomerL(pa(86)))
         If pa(87) <> Text43 Then pa(87) = Text43: stUpdates = stUpdates & ",pa87=" & CNULL(ChgSQL(pa(87)))
         If pa(88) <> Text27 Then pa(88) = Text27: stUpdates = stUpdates & ",pa88=" & CNULL(ChangeCustomerL(pa(88)))
         If pa(89) <> Text11 Then pa(89) = Text11: stUpdates = stUpdates & ",pa89=" & CNULL(ChgSQL(pa(89)))
         If pa(90) <> Text32 Then pa(90) = Text32: stUpdates = stUpdates & ",pa90=" & CNULL(ChgSQL(pa(90)))
         If pa(91) <> Text10 Then pa(91) = Text10: stUpdates = stUpdates & ",pa91=" & CNULL(ChgSQL(pa(91)))
         If pa(98) <> Text33(6) Then pa(98) = Text33(6): stUpdates = stUpdates & ",pa98=" & CNULL(ChgSQL(pa(98)))
         If pa(99) <> Text33(7) Then pa(99) = Text33(7): stUpdates = stUpdates & ",pa99=" & CNULL(ChgSQL(pa(99)))
         If pa(100) <> Text33(8) Then pa(100) = Text33(8): stUpdates = stUpdates & ",pa100=" & CNULL(ChgSQL(pa(100)))
         If pa(101) <> Text44 Then pa(101) = Text44: stUpdates = stUpdates & ",pa101=" & CNULL(ChangeCustomerL(pa(101)))
         If pa(102) <> Text45 Then pa(102) = Text45: stUpdates = stUpdates & ",pa102=" & CNULL(ChgSQL(pa(102)))
         If pa(103) <> Text46 Then pa(103) = Text46: stUpdates = stUpdates & ",pa103=" & CNULL(ChgSQL(pa(103)))
         If pa(104) <> Text47 Then pa(104) = Text47: stUpdates = stUpdates & ",pa104=" & CNULL(ChgSQL(pa(104)))
         If pa(105) <> Text30 Then pa(105) = Text30: stUpdates = stUpdates & ",pa105=" & CNULL(ChangeCustomerL(pa(105)))
         If pa(106) <> Text29 Then pa(106) = Text29: stUpdates = stUpdates & ",pa106=" & CNULL(ChgSQL(pa(106)))
         If pa(107) <> Text18 Then pa(107) = Text18: stUpdates = stUpdates & ",pa107=" & CNULL(ChgSQL(pa(107)))
         If pa(133) <> Text20 Then pa(133) = Text20: stUpdates = stUpdates & ",pa133=" & CNULL(ChangeCustomerL(pa(133)))
         If pa(134) <> Text21 Then pa(134) = Text21: stUpdates = stUpdates & ",pa134=" & CNULL(ChangeCustomerL(pa(134)))
         If pa(135) <> Text22 Then pa(135) = Text22: stUpdates = stUpdates & ",pa135=" & CNULL(ChgSQL(pa(135)))
         If pa(139) <> Text33(15) Then pa(139) = Text33(15): stUpdates = stUpdates & ",pa139=" & CNULL(ChgSQL(pa(139)))
         If pa(146) <> Text23 Then pa(146) = Text23: stUpdates = stUpdates & ",pa146=" & CNULL(ChgSQL(pa(146)))
         If pa(142) <> Text24 Then pa(142) = Text24: stUpdates = stUpdates & ",pa142=" & CNULL(ChgSQL(pa(142)))
         'end 2018/05/09
         'Add by Morgan 2008/11/14
         '新增欄位改用陣列idex對應欄位序號，以後新增就不必再改。
         For Each m_otxt In txtPA
               If pa(m_otxt.Index) <> m_otxt Then 'Added by Lydia 2018/05/09 有修改才Update+判斷
                    pa(m_otxt.Index) = m_otxt
                    stUpdates = stUpdates & ",pa" & m_otxt.Index & " = " & CNULL(ChgSQL(pa(m_otxt.Index)))
                    'Added by Lydia 2024/10/22 英文組的工程師命名記錄依照新案建檔的修改而變更---Sharon
                    If Combo3.Text <> "" And Left(Combo3.Text, 1) <> "3" And m_otxt.Index = 63 And m_TCT01 <> "" Then
                       m_UpdPA63TCT118 = "Update transcasetitle set tct118=" & CNULL(pa(m_otxt.Index)) & " where tct01='" & m_TCT01 & "' "
                    End If
                    'end 2024/10/22
               End If 'end 2018/05/09
         Next
         'end 2008/11/14
         
         'Added by Lydia 2020/02/21 設定「名稱有特殊字」
         pa(174) = IIf(ChkPA174.Value = 0, "", "Y")
         stUpdates = stUpdates & ",pa174=" & CNULL(ChgSQL(pa(174)))
         'Added by Lydia 2021/04/09 設定「有序列表」
         pa(175) = IIf(ChkPA175.Value = 0, "", "Y")
         stUpdates = stUpdates & ",pa175=" & CNULL(ChgSQL(pa(175)))
         
         '申請人1
         If ChangeCustomerL(pa(26)) <> ChangeCustomerL(Text33(9)) Then 'Added by Lydia 2018/05/09 有修改才Update+判斷
            If Text33(9) = "" Then
               pa(26) = "": stUpdates = stUpdates & ",pa26=Null"
               pa(31) = "": stUpdates = stUpdates & ",pa31=Null"
               pa(36) = "": stUpdates = stUpdates & ",pa36=Null"
               pa(41) = "": stUpdates = stUpdates & ",pa41=Null"
               
            ElseIf ChangeCustomerL(pa(26)) <> ChangeCustomerL(Text33(9)) Then
               If ClsPDGetCustomerNameAndAddress(Text33(9).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  If m_CP60 <> "" Then
                     strExc(1) = pa(1)
                     strExc(2) = pa(2)
                     strExc(3) = pa(3)
                     strExc(4) = pa(4)
                     strExc(5) = m_CP60
                     strExc(6) = Text33(9)
                     strExc(7) = strExc(0)
                     If Not ClsLawUpdAcc0k0(strExc(), True) Then
                        Text33(9).SetFocus
                        GoTo CheckingErr
                     End If
                  End If
                  pa(26) = Text33(9): stUpdates = stUpdates & ",pa26=" & CNULL(ChangeCustomerL(pa(26)))
                  pa(31) = strTmp(1): stUpdates = stUpdates & ",pa31=" & CNULL(ChgSQL(pa(31)))
                  pa(36) = strTmp(2): stUpdates = stUpdates & ",pa36=" & CNULL(ChgSQL(pa(36)))
                  pa(41) = strTmp(3): stUpdates = stUpdates & ",pa41=" & CNULL(ChgSQL(pa(41)))
              End If
            End If
         End If 'end 2018/05/09
         
         '申請人2
         If ChangeCustomerL(pa(27)) <> ChangeCustomerL(Text33(10)) Then 'Added by Lydia 2018/05/09 有修改才Update+判斷
            If Text33(10) = "" Then
               pa(27) = "": stUpdates = stUpdates & ",pa27=Null"
               pa(32) = "": stUpdates = stUpdates & ",pa32=Null"
               pa(37) = "": stUpdates = stUpdates & ",pa37=Null"
               pa(42) = "": stUpdates = stUpdates & ",pa42=Null"
               
            ElseIf ChangeCustomerL(pa(27)) <> ChangeCustomerL(Text33(10)) Then
               If ClsPDGetCustomerNameAndAddress(Text33(10).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  pa(27) = Text33(10): stUpdates = stUpdates & ",pa27=" & CNULL(ChangeCustomerL(pa(27)))
                  pa(32) = strTmp(1): stUpdates = stUpdates & ",pa32=" & CNULL(ChgSQL(pa(32)))
                  pa(37) = strTmp(2): stUpdates = stUpdates & ",pa37=" & CNULL(ChgSQL(pa(37)))
                  pa(42) = strTmp(3): stUpdates = stUpdates & ",pa42=" & CNULL(ChgSQL(pa(42)))
               End If
            End If
         End If 'end 2018/05/09
         
         '申請人3
         If ChangeCustomerL(pa(28)) <> ChangeCustomerL(Text33(11)) Then 'Added by Lydia 2018/05/09 有修改才Update+判斷
            If Text33(11) = "" Then
               pa(28) = "": stUpdates = stUpdates & ",pa28=Null"
               pa(33) = "": stUpdates = stUpdates & ",pa33=Null"
               pa(38) = "": stUpdates = stUpdates & ",pa38=Null"
               pa(43) = "": stUpdates = stUpdates & ",pa43=Null"
               
            ElseIf ChangeCustomerL(pa(28)) <> ChangeCustomerL(Text33(11)) Then
               If ClsPDGetCustomerNameAndAddress(Text33(11).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  pa(28) = Text33(11): stUpdates = stUpdates & ",pa28=" & CNULL(ChangeCustomerL(pa(28)))
                  pa(33) = strTmp(1): stUpdates = stUpdates & ",pa33=" & CNULL(ChgSQL(pa(33)))
                  pa(38) = strTmp(2): stUpdates = stUpdates & ",pa38=" & CNULL(ChgSQL(pa(38)))
                  pa(43) = strTmp(3): stUpdates = stUpdates & ",pa43=" & CNULL(ChgSQL(pa(43)))
              End If
            End If
         End If 'end 2018/05/09
         
         '申請人4
         If ChangeCustomerL(pa(29)) <> ChangeCustomerL(Text33(12)) Then 'Added by Lydia 2018/05/09 有修改才Update+判斷
            If Text33(12) = "" Then
               pa(29) = "": stUpdates = stUpdates & ",pa29=Null"
               pa(34) = "": stUpdates = stUpdates & ",pa34=Null"
               pa(39) = "": stUpdates = stUpdates & ",pa39=Null"
               pa(44) = "": stUpdates = stUpdates & ",pa44=Null"
               
            ElseIf ChangeCustomerL(pa(29)) <> ChangeCustomerL(Text33(12)) Then
               If ClsPDGetCustomerNameAndAddress(Text33(12).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  pa(29) = Text33(12): stUpdates = stUpdates & ",pa29=" & CNULL(ChangeCustomerL(pa(29)))
                  pa(34) = strTmp(1): stUpdates = stUpdates & ",pa34=" & CNULL(ChgSQL(pa(34)))
                  pa(39) = strTmp(2): stUpdates = stUpdates & ",pa39=" & CNULL(ChgSQL(pa(39)))
                  pa(44) = strTmp(3): stUpdates = stUpdates & ",pa44=" & CNULL(ChgSQL(pa(44)))
               End If
            End If
         End If 'end 2018/05/09
         
         '申請人5
         If ChangeCustomerL(pa(30)) <> ChangeCustomerL(Text33(13)) Then 'Added by Lydia 2018/05/09 有修改才Update+判斷
            If Text33(13) = "" Then
               pa(30) = "": stUpdates = stUpdates & ",pa30=Null"
               pa(35) = "": stUpdates = stUpdates & ",pa35=Null"
               pa(40) = "": stUpdates = stUpdates & ",pa40=Null"
               pa(45) = "": stUpdates = stUpdates & ",pa45=Null"
               
            ElseIf ChangeCustomerL(pa(30)) <> ChangeCustomerL(Text33(13)) Then
               If ClsPDGetCustomerNameAndAddress(Text33(13).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  pa(30) = Text33(13): stUpdates = stUpdates & ",pa30=" & CNULL(ChangeCustomerL(pa(30)))
                  pa(35) = strTmp(1): stUpdates = stUpdates & ",pa35=" & CNULL(ChgSQL(pa(35)))
                  pa(40) = strTmp(2): stUpdates = stUpdates & ",pa40=" & CNULL(ChgSQL(pa(40)))
                  pa(45) = strTmp(3): stUpdates = stUpdates & ",pa45=" & CNULL(ChgSQL(pa(45)))
               End If
            End If
         End If 'end 2018/05/09

         'Add By Sindy 2014/11/6
         If cmdAddRow.Tag = "I" Or cmdDelRow.Tag = "D" Then '有異動發明人資料
            '全部刪除,重新新增
            strSql = "delete from patentInventor where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4))
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strSql
            Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strSql
            intPI05 = 0 'Add By Sindy 2025/2/11
            For ii = 1 To GRD1.Rows - 1
               '自行輸入則客戶發明人檔IN01=PA26
               If Trim(GRD1.TextMatrix(ii, 1)) = "" And _
                  (Trim(GRD1.TextMatrix(ii, 2)) <> "" Or Trim(GRD1.TextMatrix(ii, 3)) <> "" Or Trim(GRD1.TextMatrix(ii, 4)) <> "") Then
                  'Modified by Morgan 2015/12/14 造字後面可能會加空白不可用Trim
                  'InsInventor m_PI06, pa(26), Trim(GRD1.TextMatrix(ii, 2)), Trim(GRD1.TextMatrix(ii, 3)), Trim(GRD1.TextMatrix(ii, 4)), Trim(GRD1.TextMatrix(ii, 5))
                  'Modified by Lydia 2024/12/03 Trim(GRD1.TextMatrix(ii, 5))=>Trim(GRD1.TextMatrix(ii, 7))
                  InsInventor m_PI06, pa(26), LTrim(GRD1.TextMatrix(ii, 2)), Trim(GRD1.TextMatrix(ii, 3)), LTrim(GRD1.TextMatrix(ii, 4)), Trim(GRD1.TextMatrix(ii, 7))
                  'end 2015/12/14
                  GRD1.TextMatrix(ii, 1) = m_PI06
               End If
               'Add By Sindy 2025/2/11
               If Trim(GRD1.TextMatrix(ii, 1)) <> "" Then
                  intPI05 = intPI05 + 1
               '2025/2/11 END
                  strSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                           CNULL(pa(1)) & "," & CNULL(pa(2)) & "," & CNULL(pa(3)) & "," & CNULL(pa(4)) & "," & intPI05 & ",'" & Trim(GRD1.TextMatrix(ii, 1)) & "')"
                  'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
                  'Pub_SeekTbLog strSql
                  Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
                  cnnConnection.Execute strSql
               End If
            Next ii
         End If
         '2014/11/6 END
         
         For i = 79 To 84 'Memo by Lydia 2018/06/25 個案代表人1-2(PA79~PA84)
            If pa(i) <> txtCaseField(i - 40) Then   'Added by Lydia 2018/05/09 有修改才Update+判斷
                 pa(i) = txtCaseField(i - 40): stUpdates = stUpdates & ",pa" & i & "=" & CNULL(ChgSQL(pa(i)))
            End If 'end 2018/05/09
         Next
         
         For i = 109 To 132 'Memo by Lydia 2018/06/25 個案代表人3-10(PA109~PA132)
             If pa(i) <> txtCaseField(i - 64) Then   'Added by Lydia 2018/05/09 有修改才Update+判斷
                 pa(i) = txtCaseField(i - 64): stUpdates = stUpdates & ",pa" & i & "=" & CNULL(ChgSQL(pa(i)))
             End If 'end 2018/05/09
         Next
         If stUpdates <> "" Then
            stUpdates = Mid(stUpdates, 2)
            strSql = "UPDATE PATENT SET " & stUpdates & " WHERE PA01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
            'Modified by Lydia 2018/10/19 +詳細記錄
            'Pub_SeekTbLog strSql 'Added by Morgan 2011/11/23
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strSql, , True
            'Modified by Lydia 2025/10/30 改用模組判斷
            'Pub_SeekTbLog strSql, , True, , Me.Caption & "(" & Me.Name & ")"
            Pub_SeekTbLog strSql, , PUB_FilterSeekSQL(strSql), , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strSql
         End If
         
         'Added by Lydia 2019/11/27 FCP年費特殊管制PA165=N => 目前案件的年費期限自動上不續辦
         If (pa(1) = "FCP" Or pa(1) = "P") And txtPA(156).Text = "N" And txtPA(156).Tag <> txtPA(156).Text Then
             'Modified by Lydia 2020/03/17 回傳FMP案範圍，發清單通知程序
             'Call Pub_AutoUpdFCP605(pa(1) & pa(2) & pa(3) & pa(4))
             If Pub_AutoUpdFCP605(pa(1) & pa(2) & pa(3) & pa(4), strPassSql, strErrMsg) = False Then
                  GoTo CheckingErr
             End If
             'end 2020/03/17
         End If
         'end 2019/11/27
      
         '新增優先權資料
         'Modify by Amy 2014/04/11 +pd08,pd09
         If Not ClsPDSavePriority(pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)) Then
             GoTo CheckingErr
         End If
         
'Removed by Morgan 2013/1/8 102新法,主動修正無期限,但若有實審未發文時掛相同期限(此狀況都為新案送件後收文情形,分案更新就好--靜芳)
'
'         'Add by Morgan 2007/8/28 若有主動修正未發文時更新期限
'         '1.實審未發文時期限同該程序
'         '2.法限=申請日(最早優先權日)+15個月(新型2個月);所限=法限-4天
'         '法限=申請日(最早優先權日)+15個月(新型2個月);所限=法限-4天
'         'Modified by Morgan 2012/3/5  +206 補充說明,並控制法限遇假日順延(與分案相同)
'         If (pa(8) = "1" Or pa(8) = "2") And pa(10) <> "" Then
'            strExc(0) = "select cp09 from CaseProgress" & _
'               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'               " and cp10 in ('203','206') and cp27 is null and cp57 is null"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strExc(1) = "": strExc(2) = "": strExc(3) = "": strExc(4) = ""
'               '主動修正收文號
'               strExc(1) = RsTemp.Fields(0)
'               '發明
'               If pa(8) = "1" Then
'                  strExc(0) = "select cp06,cp07 from CaseProgress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='416' and cp27 is null and cp57 is null"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  '有實審未發文
'                  If intI = 1 Then
'                     If Not IsNull(RsTemp.Fields(0)) Then
'                        strExc(4) = RsTemp.Fields(0)
'                        strExc(3) = RsTemp.Fields(1)
'                     End If
'                  '無實審未發文
'                  Else
'                     '最早優先權日
'                     If strPriority(2) <> "" Then
'                        strExc(2) = PUB_GetFirstPriDate2(strPriority(2))
'                     '申請日
'                     Else
'                        strExc(2) = DBDATE(pa(10))
'                     End If
'                     strExc(3) = CompDate(1, 15, strExc(2))
'                     strExc(4) = CompDate(2, -4, strExc(3))
'                     strExc(3) = PUB_GetWorkDay1(strExc(3), False) 'Added by Morgan 2012/3/5
'                  End If
'               '新型
'               Else
'                  strExc(2) = DBDATE(pa(10))
'                  strExc(3) = CompDate(1, 2, strExc(2))
'                  strExc(4) = CompDate(2, -4, strExc(3))
'                  strExc(3) = PUB_GetWorkDay1(strExc(3), False) 'Added by Morgan 2012/3/5
'               End If
'               'Modify by Morgan 2009/4/16 大於系統日才要更新 -- FCP-33835
'               'If strExc(3) <> "" Then
'               If Val(strExc(3)) >= Val(strSrvDate(1)) Then
'                  strSql = "update CaseProgress set cp06=" & strExc(4) & ",cp07=" & strExc(3) & " where cp09='" & strExc(1) & "'"
'                  Pub_SeekTbLog strSql 'Added by Morgan 2011/11/23
'                  cnnConnection.Execute strSql, intI
'               End If
'            End If
'         End If
'         'end 2007/8/28
'
'end 2013/1/8
         
      Case "FG", "PS", "CPS"
         'Modified by Lydia 2018/05/09 有修改才Update+判斷
         If pa(5) <> Text5 Then pa(5) = Text5: stUpdates = stUpdates & ",Sp05=" & CNULL(ChgSQL(pa(5)))
         If pa(6) <> Text6 Then pa(6) = Text6: stUpdates = stUpdates & ",Sp06=" & CNULL(ChgSQL(pa(6)))
         If pa(7) <> Text7 Then pa(7) = Text7: stUpdates = stUpdates & ",Sp07=" & CNULL(ChgSQL(pa(7)))
         If pa(8) <> Text33(9) Then pa(8) = Text33(9): stUpdates = stUpdates & ",Sp08=" & CNULL(ChangeCustomerL(pa(8)))
         If pa(18) <> Text10 Then pa(18) = Text10: stUpdates = stUpdates & ",Sp18=" & CNULL(ChgSQL(pa(18)))
         If pa(26) <> Text33(14) Then pa(26) = Text33(14): stUpdates = stUpdates & ",Sp26=" & CNULL(ChangeCustomerL(pa(26)))
         If pa(27) <> Text26 Then pa(27) = Text26: stUpdates = stUpdates & ",Sp27=" & CNULL(ChgSQL(pa(27)))
         If pa(29) <> Text31 Then pa(29) = Text31: stUpdates = stUpdates & ",Sp29=" & CNULL(ChgSQL(pa(29)))
         If pa(30) <> Text33(1) Then pa(30) = Text33(1): stUpdates = stUpdates & ",Sp30=" & CNULL(ChgSQL(pa(30)))
         If pa(71) <> Text33(15) Then pa(71) = Text33(15): stUpdates = stUpdates & ",Sp71=" & CNULL(ChgSQL(pa(71)))
         If pa(75) <> Text33(4) Then pa(75) = Text33(4): stUpdates = stUpdates & ",Sp75=" & CNULL(ChgSQL(pa(75)))
         'end 2018/05/09
         
         If pa(31) <> Text14 Then 'Added by Lydia 2018/05/09 有修改才Update+判斷
            If Text14 = "" Then
               stUpdates = stUpdates & ",Sp31=null"
            Else
               pa(31) = Text14: stUpdates = stUpdates & ",Sp31=" & pa(31)
            End If
         End If 'end 2018/05/09
         
         'Modified by Lydia 2018/05/09 有修改才Update+判斷
         If pa(33) <> Text12 Then pa(33) = Text12: stUpdates = stUpdates & ",Sp33=" & CNULL(ChgSQL(pa(33)))
         If pa(34) <> Text13 Then pa(34) = Text13: stUpdates = stUpdates & ",Sp34=" & CNULL(ChgSQL(pa(34)))
         If pa(35) <> Text42 Then pa(35) = Text42: stUpdates = stUpdates & ",Sp35=" & CNULL(ChangeCustomerL(pa(35)))
         If pa(36) <> Text43 Then pa(36) = Text43: stUpdates = stUpdates & ",Sp36=" & CNULL(ChgSQL(pa(36)))
         If pa(37) <> Text27 Then pa(37) = Text27: stUpdates = stUpdates & ",Sp37=" & CNULL(ChangeCustomerL(pa(37)))
         If pa(58) <> Text33(10) Then pa(58) = Text33(10): stUpdates = stUpdates & ",Sp58=" & CNULL(ChangeCustomerL(pa(58)))
         If pa(59) <> Text33(11) Then pa(59) = Text33(11): stUpdates = stUpdates & ",Sp59=" & CNULL(ChangeCustomerL(pa(59)))
         If ChangeCustomerL(pa(67)) <> ChangeCustomerL(Text20) Then pa(67) = ChangeCustomerL(Text20): stUpdates = stUpdates & ",Sp67=" & CNULL(ChangeCustomerL(pa(67)))
         If pa(79) <> Trim(Left(Combo3, 1)) Then pa(79) = Left(Combo3, 1): stUpdates = stUpdates & ",sp79=" & CNULL(ChgSQL(pa(79)))
         If pa(80) <> Text24 Then pa(80) = Text24: stUpdates = stUpdates & ",sp80=" & CNULL(ChgSQL(pa(80)))
         If pa(81) <> txtPA(153) Then pa(81) = txtPA(153): stUpdates = stUpdates & ",sp81=" & CNULL(ChgSQL(pa(81)))
         If pa(82) <> txtPA(154) Then pa(82) = txtPA(154): stUpdates = stUpdates & ",sp82=" & CNULL(ChgSQL(pa(82)))
         If pa(83) <> txtPA(155) Then pa(83) = txtPA(155): stUpdates = stUpdates & ",sp83=" & CNULL(ChgSQL(pa(83)))
         If pa(84) <> txtPA(159) Then pa(84) = txtPA(159): stUpdates = stUpdates & ",sp84=" & CNULL(ChgSQL(pa(84)))
         'end 2018/05/09
                  
         If stUpdates <> "" Then
            stUpdates = Mid(stUpdates, 2)
            strSql = "UPDATE SERVICEPRACTICE SET " & stUpdates & " WHERE SP01='" & pa(1) & "' and SP02='" & pa(2) & "' and SP03='" & pa(3) & "' and SP04='" & pa(4) & "'"
            'Modified by Lydia 2018/10/19 +詳細記錄
            'Pub_SeekTbLog strSql 'Added by Morgan 2011/11/23
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strSql, , True
            'Modified by Lydia 2025/10/30 改用模組判斷
            'Pub_SeekTbLog strSql, , True, , Me.Caption & "(" & Me.Name & ")"
            Pub_SeekTbLog strSql, , PUB_FilterSeekSQL(strSql), , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strSql
         End If
   End Select
   
   'Modified by Morgan 2012/6/7
   'If pa(1) = "FCP" And pa(23) <> "1" Then
   If (pa(1) = "FCP" Or pa(1) = "P" Or pa(1) = "CFP") And pa(23) <> "1" Then
      strTmp(1) = "CP37=" & CNULL(ChgSQL(Text5)) & ",CP38=" & CNULL(ChgSQL(Text6)) & ",CP39=" & CNULL(ChgSQL(Text7)) & ","
   Else
      strTmp(1) = ""
   End If
   
   'Add By Sindy 2023/4/18 存指定日期
   If m_CP27 = "" And m_CP57 = "" Then
      If txtCP142.Text <> "" Then
         strTmp(1) = strTmp(1) & "CP141='3'," '3.指定日期送件
         strTmp(1) = strTmp(1) & "CP142=" & DBDATE(txtCP142) & ","
         strTmp(1) = strTmp(1) & "CP164='" & IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", "3")) & "',"
         'Modify By Sindy 2023/5/2
         strExc(10) = "客戶指定" & ChangeWStringToTDateString(DBDATE(txtCP142.Text)) & IIf(Option1(0).Value = True, "當天", IIf(Option1(1).Value = True, "之前", IIf(Option1(2).Value = True, "之後", ""))) & "送件"
         If InStr(Trim(Text19.Text), strExc(10)) = 0 Then
         '2023/5/2 END
            Text19 = strExc(10) & ";" & Trim(Text19.Text)
         End If
      Else
         '取消指定日
         strTmp(1) = strTmp(1) & "CP141=null,CP142=null,CP164=null,"
         If InStr(Text19, "客戶指定") > 0 And _
            (InStr(Text19, "當天送件") > 0 Or InStr(Text19, "之前送件") > 0 Or InStr(Text19, "之後送件") > 0) Then
            MsgBox "有取消指定日，請拿掉備註裡的客戶指定資訊！", vbExclamation
            GoTo CheckingErr
         End If
      End If
   End If
   '2023/4/18 END
   
   'Added by Lydia 2018/05/09 新案進度的欄位
   If Text8.Tag <> Text8.Text Then
       strTmp(1) = strTmp(1) & " CP06=" & CNULL(DBDATE(Text8), True) & ","
   End If
   If Text9.Tag <> Text9.Text Then
       strTmp(1) = strTmp(1) & " CP07=" & CNULL(DBDATE(Text9), True) & ","
   End If
   If Text19.Tag <> Text19.Text Then
       strTmp(1) = strTmp(1) & " CP64=" & CNULL(ChgSQL(Text19)) & ","
   End If
   If Text25.Tag <> Text25.Text Then
       strTmp(1) = strTmp(1) & " CP118=" & CNULL(ChgSQL(Text25)) & ","
   End If
   'end 2018/05/09
   
   'Modify By Amy 2013/05/14 新增電子送件欄位(更新)
   'strExc(1) = "UPDATE CaseProgress SET " & strTmp(1) & "CP06=" & DBNullDate(Text8) & ",CP07=" & DBNullDate(Text9) & _
      ",CP64=" & CNULL(ChgSQL(Text19)) & " WHERE " & ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP31='Y'"
   'Modified by Lydia 2016/12/28
   'strExc(1) = "UPDATE CaseProgress SET " & strTmp(1) & "CP06=" & DBNullDate(Text8) & ",CP07=" & DBNullDate(Text9) & _
      ",CP64=" & CNULL(ChgSQL(Text19)) & ",CP118=" & CNULL(ChgSQL(Text25)) & " WHERE " & ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP31='Y'"
   'Pub_SeekTbLog strSql 'Added by Morgan 2011/11/23
   strTmp(1) = Replace(strTmp(1), ",,", ",") 'Added by Lydia 2017/07/28 清除重覆逗點
   'Modified by Lydia 2018/05/09 新案進度有改才更新
   'strExc(1) = "UPDATE CaseProgress SET " & strTmp(1) & "CP06=" & CNULL(DBDATE(Text8)) & ",CP07=" & CNULL(DBDATE(Text9)) & _
      ",CP64=" & CNULL(ChgSQL(Text19)) & ",CP118=" & CNULL(ChgSQL(Text25)) & " WHERE " & ChgCaseProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP31='Y'"
   If strTmp(1) <> "" Then
         strTmp(1) = IIf(Right(strTmp(1), 1) = ",", Mid(strTmp(1), 1, Len(strTmp(1)) - 1), strTmp(1))
         strExc(1) = "UPDATE CaseProgress SET " & strTmp(1) & " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP31='Y'"
         'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
         'Pub_SeekTbLog strExc(1)
         Pub_SeekTbLog strExc(1), , , , Me.Caption & "(" & Me.Name & ")"
         cnnConnection.Execute strExc(1)
        
         'Added by Morgan 2022/4/20
         'FMP案要同步更新主張國際優先權之本所期限 Ex:P-129195 --Sharon
         If pa(1) = "P" Then
            strSql = "update caseprogress a set cp06=(select cp06 from caseprogress b where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP31='Y')" & _
               " where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp10='106'"
            cnnConnection.Execute strSql, intI
         End If
         'end 2022/4/20
   End If
   'end 2018/05/09
   
  'Add by Morgan 2008/8/19 更新告代期限
   If m_901CP09 <> "" Then
      'Added by Lydia 2022/08/04 更新提申前告代和主動修正的計算方式要和工程師命名計算一樣。
      If m_TCT01 <> "" Then
          strExc(0) = PUB_GetTCTbCP48("2", pa(1), pa(2), pa(3), pa(4), pa(9), pa(10), m_CP05, DBDATE(Text8), "901", strExc(1), m_TF01cp06, m_TF01cp27)
      Else
          strExc(1) = DBDATE(Text8) '避免Trigger(Caseprogress_Before)更新=> 本所期限=承辦期限+5個工作天
      'end 2022/08/04
          strExc(0) = DBDATE(Text8)
      End If 'Added by Lydia 2022/08/04
      'Modified by Lydia 2018/04/23 改成多筆
'      strSql = "update CaseProgress set cp48=" & strExc(0) & " where cp09='" & m_901CP09 & "'"
'      Pub_SeekTbLog strSql 'Added by Morgan 2011/11/23
'      cnnConnection.Execute strSql, intI
      tmpArr = Empty
      tmpArr = Split(m_901CP09, ",")
      For ii = 0 To UBound(tmpArr)
          If Trim(tmpArr(ii)) <> "" Then
            'Modified by Lydia 2022/08/04 一併更新本所期限
            strExc(1) = "update CaseProgress set cp06=" & CNULL(strExc(1)) & ", cp48=" & CNULL(strExc(0)) & ",cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 提申前告代;'||cp64 where cp09='" & tmpArr(ii) & "'"
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strExc(1)
            Pub_SeekTbLog strExc(1), , , , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strExc(1), intI
          End If
      Next ii
      'end 2018/04/23
   End If
   'Added by Lydia 2018/04/23 更新主動修正期限
   If m_203CP09 <> "" Then
      'Added by Lydia 2022/08/04 更新提申前告代和主動修正的計算方式要和工程師命名計算一樣。
      If m_TCT01 <> "" Then
          strExc(0) = PUB_GetTCTbCP48("2", pa(1), pa(2), pa(3), pa(4), pa(9), pa(10), m_CP05, DBDATE(Text8), "203", strExc(1), m_TF01cp06, m_TF01cp27)
      Else
          strExc(1) = DBDATE(Text8) '避免Trigger(Caseprogress_Before)更新=> 本所期限=承辦期限+5個工作天
      'end 2022/08/04
          strExc(0) = DBDATE(Text8)
      End If 'Added by Lydia 2022/08/04
      tmpArr = Empty
      tmpArr = Split(m_203CP09, ",")
      For ii = 0 To UBound(tmpArr)
          If Trim(tmpArr(ii)) <> "" Then
            'Modified by Lydia 2022/08/04 一併更新本所期限
            strExc(1) = "update CaseProgress set cp06=" & CNULL(strExc(1)) & ", cp48=" & CNULL(strExc(0)) & ",cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & " 提申前主動修正;'||cp64 where cp09='" & tmpArr(ii) & "'"
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strExc(1)
            Pub_SeekTbLog strExc(1), , , , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strExc(1), intI
          End If
      Next ii
   End If
   'end 2018/04/23
   
   'Add by Morgan 2011/6/13
   '若有期限則於更新資料後清除原來的備註並加入新的備註
   'Modified by Morgan 2012/2/2 +pa75
   strExc(0) = "select pa26,pa27,pa28,pa29,pa30,pa75 from patent where " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If bolChkMemo416 Then
         'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
         'For intI = 0 To 4
         '   If Not IsNull(RsTemp(intI)) Then
         '      'Modified by Morgan 2012/2/2  +pa75
         '      'Modified by Morgan 2013/9/11 改抓設定檔
         '      'strNewMemo416 = PUB_Get416Memo(ChangeCustomerL(RsTemp(intI)), ChangeCustomerL("" & RsTemp("pa75")))
         '      strNewMemo416 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp("pa75")), ChangeCustomerL(RsTemp(intI)))
         '      If strNewMemo416 <> "" Then Exit For
         '   End If
         'Next
         strNewMemo416 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp("pa75")), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
         'end 2022/08/02
         
         If strNewMemo416 <> strOldMemo416 Then
            If strOldMemo416 <> "" Then 'Added by Lydia 2022/08/02
              strSql = "update nextprogress set np15=replace(np15,'" & ChgSQL(strOldMemo416) & "','')" & _
                 " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=416"
              'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
              'Pub_SeekTbLog strSql 'Added by Morgan 2011/11/23
              Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
              cnnConnection.Execute strSql, intI
            End If 'Added by Lydia 2022/08/02
            
            If strNewMemo416 <> "" Then 'Added by Lydia 2022/08/02
              strSql = "update nextprogress set np15='" & ChgSQL(strNewMemo416) & "'||';'||np15" & _
                 " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=416"
              'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
              'Pub_SeekTbLog strSql 'Added by Morgan 2011/11/23
              Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
              cnnConnection.Execute strSql, intI
            End If 'Added by Lydia 2022/08/02
         End If
      
      ElseIf bolChkMemo605 = True Then
         strExc(9) = PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1")
         'Modified by Morgan 2012/6/4 +pa26
         'Modified by Morgan 2013/9/11 改抓設定檔
         'strNewMemo605 = PUB_Get605Memo(strExc(9), RsTemp("pa26"), pa(1) & pa(2) & pa(3) & pa(4))
         'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
         'strNewMemo605 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp("pa26"))
         strNewMemo605 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
         
         If strNewMemo605 <> strOldMemo605 Then
            If strOldMemo605 <> "" Then 'Added by Lydia 2022/08/02
              strSql = "update nextprogress set np15=replace(np15,'" & ChgSQL(strOldMemo605) & "','')" & _
                 " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=605"
              'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
              'Pub_SeekTbLog strSql 'Added by Morgan 2011/11/23
              Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
              cnnConnection.Execute strSql, intI
            End If 'Added by Lydia 2022/08/02
            
            If strNewMemo605 <> "" Then 'Added by Lydia 2022/08/02
              strSql = "update nextprogress set np15='" & ChgSQL(strNewMemo605) & "'||';'||np15" & _
                 " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=605"
              'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
              'Pub_SeekTbLog strSql 'Added by Morgan 2011/11/23
              Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
              cnnConnection.Execute strSql, intI
            End If 'Added by Lydia 2022/08/02
         End If
         
      End If
   End If
   'end 2011/6/13
   
   'Added by Lydia 2017/05/17 有新案翻譯,輸入原文字數、相似度、相似案號,儲存在翻譯費用檔
   'Modified by Lydia 2017/12/01 輸錯案號後,再次修改清空,無法寫入(FCP-57715)
   'If Text33(16).Visible = True And Text33(16).Tag <> "" And Trim(Replace(Text33(16) & Text33(17) & Text33(18), " ", "")) <> "" Then
   'Modified by Lydia 2018/06/07 翻譯分案無紙化
'   If Text33(16).Visible = True And Text33(16).Tag <> "" And (m_TF23 <> Val(Text33(16).Text) Or m_TF19 <> Val(Text33(17).Text) Or m_TF20 <> Trim(Text33(18).Text)) Then
'      strSql = "UPDATE TransFee SET TF23=" & Val(Text33(16)) & " , TF19=" & IIf(Val(Text33(17)) = 100, "NULL", Val(Text33(17))) & " , TF20='" & Trim(Text33(18)) & "' WHERE TF01='" & Text33(16).Tag & "' "
'      cnnConnection.Execute strSql, intI
'      'Modified by Lydia 2017/12/01
'      'If intI = 0 Then
'      If intI = 0 And Trim(Text33(16).Text & Text33(17).Text & Text33(18).Text) <> "" Then
'         strSql = "INSERT INTO TransFee (TF01,TF19,TF20,TF23) VALUES ('" & Text33(16).Tag & "' , " & IIf(Val(Text33(17)) = 100, "NULL", Val(Text33(17))) & " , " & CNULL(Trim(Text33(18))) & " , " & Val(Text33(16)) & ") "
'         cnnConnection.Execute strSql, intI
'      End If
'   End If
   'end 2017/05/17
   'Added by Lydia 2018/06/07 翻譯分案無紙化
   'Modified by Lydia 2022/06/15
   'If fraTrans01.Enabled = True Or fraTrans02.Enabled = True Or fraTrans03.Enabled = True Then
   If (fraTrans01.Visible = True And fraTrans01.Enabled = True) Or (fraTrans02.Visible = True And fraTrans02.Enabled = True) Or (fraTrans03.Visible = True And fraTrans03.Enabled = True) Then
         If txtTF(1).Text <> "" Then '已有翻譯費用檔
                strExc(5) = "": strExc(6) = ""
                For Each m_otxt In txtTF
                      Select Case m_otxt.Index
                             Case 30  '待英文本翻譯／英文本收文號
                                      strExc(6) = ""
                                      If Len(txtTF(30).Text) = 9 Then
                                          strExc(6) = txtTF(30).Text
                                      ElseIf Chk02.Value = 1 Then
                                          strExc(6) = "Y"
                                      End If
                                      If strExc(6) <> txtTF(30).Tag Then
                                           strExc(5) = strExc(5) & ", TF30=" & CNULL(strExc(6))
                                      End If
                             'Added by Lydia 2018/09/12 交稿期限、只交Claims期限
                             Case 26, 32
                                   If m_otxt.Text <> m_otxt.Tag Then
                                        strExc(5) = strExc(5) & ", TF" & Format(m_otxt.Index, "00") & "=" & CNULL(TransDate(m_otxt.Text, 2), True)
                                   End If
                             'Added by Lydia 2019/08/23 翻譯特殊指示
                             'Modified by Lydia 2019/10/25 +翻譯瑕疵備註 txtTF(37)
                             Case 36, 37
                                   If m_otxt.Text <> m_otxt.Tag Then
                                        strExc(5) = strExc(5) & ", TF" & Format(m_otxt.Index, "00") & "=" & CNULL(ChgSQL(PUB_StringFilter(m_otxt.Text)))
                                   End If
                             'end 2019/08/23
                             Case Else
                                   If m_otxt.Text <> m_otxt.Tag Then
                                        'Modifed by Lydia 2018/09/12 debug
                                        'If m_otxt.Index >= 23 And m_otxt.Index <= 28 Then  '數字
                                        '    strExc(5) = strExc(5) & ", TF" & Format(m_otxt.Index, "00") & "=" & CNULL(TransDate(m_otxt.Text, 1))
                                        If m_otxt.Index >= 23 And m_otxt.Index <= 28 Then  '數字
                                            strExc(5) = strExc(5) & ", TF" & Format(m_otxt.Index, "00") & "=" & CNULL(m_otxt.Text, True)
                                        'end 2018/09/12
                                        Else
                                            strExc(5) = strExc(5) & ", TF" & Format(m_otxt.Index, "00") & "=" & CNULL(m_otxt.Text)
                                        End If
                                   End If
                      End Select
                Next
                If strExc(5) <> "" Then
                    strSql = "UPDATE TransFee SET " & Mid(strExc(5), 2) & " WHERE TF01='" & txtTF(1) & "' "
                    cnnConnection.Execute strSql, intI
                End If
         Else
                strExc(5) = "": strExc(6) = ""
                
                For Each m_otxt In txtTF
                      Select Case m_otxt.Index
                             Case 30  '待英文本翻譯／英文本收文號
                                      strExc(7) = ""
                                      If Len(txtTF(30).Text) = 9 Then
                                          strExc(7) = txtTF(30).Text
                                      ElseIf Chk02.Value = 1 Then
                                          strExc(7) = "Y"
                                      End If
                                      If strExc(7) <> "" Then
                                           strExc(5) = strExc(5) & ", TF" & Format(m_otxt.Index, "00")
                                           strExc(6) = strExc(6) & ", " & CNULL(strExc(7))
                                      End If
                             'Added by Lydia 2018/09/12 交稿期限、只交Claims期限
                             Case 26, 32
                                   If m_otxt.Text <> m_otxt.Tag Then
                                        strExc(5) = strExc(5) & ", TF" & Format(m_otxt.Index, "00")
                                        strExc(6) = strExc(6) & ", " & CNULL(TransDate(m_otxt.Text, 2), True)
                                   End If
                             'Added by Lydia 2019/08/23 翻譯特殊指示
                             'Modified by Lydia 2019/10/25 +翻譯瑕疵備註 txtTF(37)
                             Case 36, 37
                                   If m_otxt.Text <> m_otxt.Tag Then
                                        strExc(5) = strExc(5) & ", TF" & Format(m_otxt.Index, "00")
                                        strExc(6) = strExc(6) & ", " & CNULL(ChgSQL(PUB_StringFilter(m_otxt.Text)))
                                   End If
                             'end 2019/08/23
                             Case Else
                                   If m_otxt.Text <> m_otxt.Tag Then
                                        strExc(5) = strExc(5) & ", TF" & Format(m_otxt.Index, "00")
                                        'Modified by Lydia 2018/09/12 debug
                                        'If m_otxt.Index >= 23 And m_otxt.Index <= 28 Then  '數字
                                        '     strExc(6) = strExc(6) & ", " & CNULL(TransDate(m_otxt.Text, 1))
                                        If m_otxt.Index >= 23 And m_otxt.Index <= 28 Then  '數字
                                             strExc(6) = strExc(6) & ", " & CNULL(m_otxt.Text, True)
                                        'end 2018/09/12
                                        Else
                                             strExc(6) = strExc(6) & ", " & CNULL(m_otxt.Text)
                                        End If
                                   End If
                      End Select
                Next
                If strExc(5) <> "" And strExc(6) <> "" Then
                      strSql = "insert into TransFee(TF01 " & strExc(5) & ") values ('" & m_TF01 & "' " & strExc(6) & " ) "
                      cnnConnection.Execute strSql, intI
                End If
         End If
         
         strExc(0) = "": strExc(1) = "": strExc(2) = ""
         '若有交稿期限於新案建檔輸入時，行事曆自動新增期限
         If txtTF(26).Text <> txtTF(26).Tag And txtTF(26).Text <> "" Then
            If strExc(1) = "" Then strExc(1) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
            If strExc(0) = "" Then strExc(0) = Pub_GetSpecMan("M") 'Sharon
            If PUB_AddFCPStaffCalendar(TransDate(txtTF(26).Text, 2), "1", strExc(1) & IIf(strExc(1) <> strExc(0), "," & strExc(0), ""), "譯者翻譯交稿期限", strExc(1), "1", pa(1), pa(2), pa(3), pa(4)) = True Then
            End If
         End If
         '只交Claims期限:比照交稿期限
         If txtTF(32).Text <> txtTF(32).Tag And txtTF(32).Text <> "" Then
            If strExc(1) = "" Then strExc(1) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
            If strExc(0) = "" Then strExc(0) = Pub_GetSpecMan("M") 'Sharon
            If PUB_AddFCPStaffCalendar(TransDate(txtTF(32).Text, 2), "1", strExc(1) & IIf(strExc(1) <> strExc(0), "," & strExc(0), ""), "譯者Claims交稿期限", strExc(1), "1", pa(1), pa(2), pa(3), pa(4)) = True Then
            End If
         End If
         
         '未提申先翻譯：一般案件需要提申後才進入翻譯分案階段，勾選後可進入翻譯分案作業，並且發mail通知Sharon
         If Chk03.Value = 1 And txtTF(31).Text <> txtTF(31).Tag Then
              If strExc(0) = "" Then strExc(0) = Pub_GetSpecMan("M")
              If strExc(0) <> "" Then
                    '主旨
                    strExc(1) = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "") & " 未提申先翻譯"
                    If m_TCT27 <> "" Then
                          strExc(2) = Pub_GetTct27ID(m_TCT10, m_TCT27, m_TCT28, , strExc(3))
                          strExc(1) = strExc(1) & "(命名作業欲翻譯人員：" & strExc(2) & " " & strExc(3) & ")"
                    End If
                    '內文
                    strExc(4) = ""
                    If txtTF(32).Text <> "" Then
                        strExc(4) = strExc(4) & vbCrLf & "只交Claim期限：" & ChangeTStringToTDateString(txtTF(32).Text)
                    End If
                    strExc(4) = strExc(4) & vbCrLf & "交稿期限：" & ChangeTStringToTDateString(txtTF(26).Text)
                    
                    strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                       " values( '" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd')" & _
                       ",to_char(sysdate,'hh24miss'),'" & strExc(1) & "','" & strExc(4) & "',null)"
                    cnnConnection.Execute strSql
                    Sleep 1000
              End If
         End If
   End If
   'end 2018/06/07
   
   'Added by Lydia 2017/11/14 FCP案件命名電子化
   'Modified by Lydia 2018/06/07 有命名作業記錄和修改工程師組別
   'If m_TCT01 <> "" Then 'Memo by Lydia 2018/05/1 有改組別
   If m_TCT01 <> "" And Combo3.Tag <> Combo3.Text Then
      strExc(1) = ""
      'Modified by Lydia 2019/01/09
      'Select Case Left(Combo3.Text, 1)
      '    Case "1": strExc(1) = Pub_GetSpecMan("T")
      '    Case "2": strExc(1) = Pub_GetSpecMan("R")
      '    Case "3": strExc(1) = Pub_GetSpecMan("S")
      '    Case "4": strExc(1) = Pub_GetSpecMan("T1")
      '    'Added by Lydia 2018/03/26 清空為退程序
      '    Case Else: strExc(1) = "B"
      'End Select
      strExc(1) = Pub_GetFCPGrpMan(Left(Combo3.Text, 1))
      If strExc(1) = "" Then strExc(1) = "B"
      'end 2019/01/09
      'Added by Lydia 2023/03/07 外專新案認領：若處於認領階段則取消該階段(TCN23=9)，將認領期限更新為系統日期+時間，並且Email認領工程師。
      If bolUpdTCN23 = True Then
          '一併取消暫不認領TCN16
          strSql = "Update TrackingCaseName Set TCN23='9',TCN16=null, TCN21=to_char(sysdate,'yyyymmdd'), TCN22=substr(lpad(to_char(sysdate,'hh24miss'),6,'0'),1,4) Where TCN05='" & m_TCT01 & "' "
          cnnConnection.Execute strSql
          If m_TCN19 = "Y" Then '通知英文組工程師主管
             strExc(7) = PUB_GetEngGrpMan(strExc(8))
             If strExc(7) <> "" Then
                 'Modified by Lydia 2023/05/26 Email主旨開頭改成模組
                 strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                               " values( '" & strUserNum & "','" & strExc(7) & "',to_char(sysdate,'yyyymmdd')" & _
                               ",to_char(sysdate,'hh24miss'),'" & PUB_GetTCNmTitle(pa(1), pa(2), pa(3), pa(4), pa(10), "", "") & "已分案，無需再認領。','同主旨' )"
                 cnnConnection.Execute strSql
             End If
          End If
      End If
      'end 2023/03/07
      
      If strExc(1) <> "" Then
          'Added by Lydia 2022/10/12 特殊情況之指定職代
          strExc(1) = PUB_GetStateForMan(strExc(1))
         'Modified by Lydia 2017/12/27 更改分案組別-清空命名記錄檔除案件名稱以外的內容
         'strSql = "update transcasetitle set tct04='" & strExc(1) & "', TCT112='" & strUserNum & "', TCT113=" & strSrvDate(1) & ", TCT114=" & Mid(Format(ServerTime, "000000"), 1, 4) & " where tct01='" & m_TCT01 & "' "
         strExc(3) = "": strExc(4) = ""
         If m_TCT04 <> "" Then
              'Modified by Lydia 2018/03/01  改控制
              'For intI = 5 To 111
              For intI = m_FS To TF_TCT
                   If InStr(TF_TCTnotFS, Format(intI, "000")) = 0 Then
              'end 2018/03/01
                        Select Case intI
                              Case 16, 17 '案件名稱
                                      'Added by Lydia 2018/04/16 若有主管確認後,再改變組別
                                      If Trim(Text5.Text & Text6.Text) <> "待命名" And Trim(Text5.Text & Text6.Text) <> Trim(m_TCT16 & m_TCT17) Then
                                          'Modified by Lydia 2018/06/26 排除單引號
                                          'strExc(3) = strExc(3) & ", " & IIf(intI = 16, "TCT16=" & CNULL(Text5.Text), "TCT17=" & CNULL(Text6.Text))
                                          strExc(3) = strExc(3) & ", " & IIf(intI = 16, "TCT16=" & CNULL(ChgSQL(Text5.Text)), "TCT17=" & CNULL(ChgSQL(Text6.Text)))
                                      End If
                                      'end 2018/04/16
                              Case 19 '是否收文主動修正
                                      'Modified by Lydia 2018/03/01 +限B類
                                      'Modified by Lydia 2018/04/20 工程師收文
                                      'If PUB_ChkCPExist(pa, "203", , , , "B") = True Then
                                      'Modified by Lydia 2022/04/28 改成共用模組
                                      'If frm090902_2.ChkCPisExist(pa, "203") Then
                                      If PUB_ChkBCPisExist(pa, "203") Then
                                          strExc(3) = strExc(3) & ", TCT19='Y' "
                                      Else
                                          strExc(3) = strExc(3) & ", TCT19=NULL "
                                      End If
                              Case 20 '告代
                                      'Modified by Lydia 2018/03/01 +限B類
                                      'Modified by Lydia 2018/04/20 工程師收文
                                      'If PUB_ChkCPExist(pa, "901", , strExc(4), , "B") = True Then
                                      'Modified by Lydia 2022/04/28 改成共用模組
                                      'If frm090902_2.ChkCPisExist(pa, "901", strExc(4)) Then
                                      If PUB_ChkBCPisExist(pa, "901", strExc(4)) Then
                                          'Modified by Lydia 2018/04/20 未提申=>提申前
                                          'strExc(3) = strExc(3) & ", TCT20='1' " '修改工程師組別，預設提申後告代
                                          strExc(3) = strExc(3) & IIf(Val(m_CP27) > 0, ", TCT20='1' ", ", TCT20='2' ")
                                      Else
                                          strExc(3) = strExc(3) & ", TCT20=NULL "
                                      End If
                               'Added by Lydia 2018/07/12 相似度和相似案(覆蓋)
                              Case 23 '相似案號
                                      If txtTF(20).Text <> "" Then
                                          strExc(3) = strExc(3) & ", TCT23=" & CNULL(txtTF(20).Text)
                                      Else
                                          strExc(3) = strExc(3) & ", TCT23=NULL "
                                      End If
                              Case 24 '相似度
                                      If txtTF(19).Text <> "" Then
                                          strExc(3) = strExc(3) & ", TCT24=" & CNULL(txtTF(19).Text, True)
                                      Else
                                          strExc(3) = strExc(3) & ", TCT24=NULL "
                                      End If
                              'end 2018/07/12
                              Case Else
                                      strExc(3) = strExc(3) & ", TCT" & IIf(intI < 100, Format(intI, "00"), Format(intI, "000")) & "=NULL "
                        End Select
                   End If
              Next intI
         End If
         If strExc(4) <> "" And Mid(strExc(4), 1, 1) = "B" Then
             strSql = "update CaseProgress set cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & "修改工程師組別：" & PUB_GetFCPGrpName(pa(150)) & "->" & Trim(Mid(Combo3.Tag, 3)) & ";'||cp64 where cp09='" & strExc(4) & "' and cp158=0"
             cnnConnection.Execute strSql, intI
         End If
          'Remove by Lydia 2018/04/23 不清空譯畢期限(因為最初是櫃台收文有組別,所以才設清空)
         'strExc(3) = ", TCT02=null, TCT03=null " & strExc(3)
         
         'Modified by Lydia 2018/03/26 區分-退程序
         'strSql = "update transcasetitle set tct04=" & CNULL(strExc(1)) & strExc(3) & ", TCT112='" & strUserNum & "', TCT113=" & strSrvDate(1) & ", TCT114=" & Mid(Format(ServerTime, "000000"), 1, 4) & " where tct01='" & m_TCT01 & "' "
         strSql = "update transcasetitle set tct04=" & IIf(strExc(1) <> "B", CNULL(strExc(1)), "Null") & strExc(3) & ", TCT112='" & strUserNum & "', TCT113=" & strSrvDate(1) & ", TCT114=" & Mid(Format(ServerTime, "000000"), 1, 4) & " where tct01='" & m_TCT01 & "' "
         'end 2017/12/27
         cnnConnection.Execute strSql, intI
         If intI > 0 Then
            'Added by Lydia 2017/11/28 更改分案組別，通知雙方
            If m_TCT04 <> "" Then
               'Addded by Lydia 2018/04/18 重新分案,清除卷宗區記錄,直到新組別主管確認再次產生
               strSql = "delete from casepaperpdf where cpp01='" & m_TCT01 & "' and instr(cpp02,'" & FCP命名記錄 & "') > 0 "
               cnnConnection.Execute strSql, intI
               'end 2018/04/18
               
               'Added by Lydia 2018/04/25 因為新申請案接洽單可直接收回代902和主動修正203,當命名作業的主管確認自動掛承辦人=命名人員並且上已分案
               '                                        所以改工程師組別時一併清空,直到下次主管確認一併更新
               'Modified by Lydia 2022/04/06 增加收文A類901告代
               'Modified by Lydia 2022/04/28 增加: 加速審查422,高速審查431
               'Modified by Lydia 2023/05/10 保留FCP案的告代901和主動修正203；因FCP新案急件重新認領，修改進度檔若有提申後告代、主動修正再發一次mail通知舊和新承辦人之事。
               'strSql = "Update CaseProgress set cp14=null, cp122=null where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 in ('902','203','901','422','431') "
               strSql = "Update CaseProgress set cp14=null, cp122=null where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 in (" & GetAddStr(IIf(pa(1) = "FCP", Replace(TCTforCP14, "203,901,", ""), TCTforCP14)) & ") "
               cnnConnection.Execute strSql, intI
               '清空-工程師收告代901和主動修正203
               If pa(1) <> "FCP" Then  'Added by Lydia 2023/05/10 保留FCP案
                  strSql = "Update CaseProgress set cp14=null, cp122=null where cp09 in (select cp09 from CaseProgress,staff where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                            " and cp158=0 and cp159=0 and substr(cp09,1,1)='B' and cp10 in ('203','901') and cp65=st01(+) and (st03='F21' or st03='M51') )  "
                  cnnConnection.Execute strSql, intI
               End If 'Added by Lydia 2023/05/10
               'end 2018/04/25
               
               If strExc(1) <> "B" Then 'Added by Lydia 2018/03/26
                   'Remove by Lydia 2018/04/17  原工程師主管改為副本
                   'strExc(1) = m_TCT04 & ";" & strExc(1)
                   strExc(2) = Left(Combo3.Tag, 1) & "-" & Left(Combo3.Text, 1)
               'Added by Lydia 2018/03/26 +區分退程序
               Else
                   strExc(1) = m_TCT04
                   strExc(2) = Left(Combo3.Tag, 1) & "-B"
               End If
               'end 2018/03/26
               'Modified by Lydia 2018/04/17  原工程師主管改為副本
               'If PUB_GetTCTmail(True, 2, pa(1), pa(2), pa(3), pa(4), m_TCT01, "", strExc(1), strExc(2)) Then
               If PUB_GetTCTmail(True, 2, pa(1), pa(2), pa(3), pa(4), m_TCT01, "", strExc(1), strExc(2), , , IIf(strExc(1) <> "B", m_TCT04, "")) Then
               End If
            Else
            'end 2017/11/28
                If strExc(1) <> "B" Then  'Added by Lydia 2018/05/23 排除命名作業沒主管,又改退程序
                    If PUB_GetTCTmail(True, 1, pa(1), pa(2), pa(3), pa(4), m_TCT01, "", strExc(1)) Then
                    End If
                End If
            End If
            'Added by Lydia 2018/09/27 通知已認領翻譯人員
            strSql = "select * from transfeeassign where tfa01 = '" & m_TF01 & "' "
            ii = 1
            Set rsAD = ClsLawReadRstMsg(ii, strSql)
            If ii = 1 Then
                 strSql = "delete from transfeeassign where tfa01='" & m_TF01 & "' "
                 cnnConnection.Execute strSql
                 If "" & rsAD.Fields("tfa04") <> "" And Left("" & rsAD.Fields("tfa04"), 1) <> "F" Then '外翻不通知
                        strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                             " values( '" & strUserNum & "','" & rsAD.Fields("tfa04") & "',to_char(sysdate,'yyyymmdd')" & _
                             ",to_char(sysdate,'hh24miss'),'" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "更改工程師組別，取消認領翻譯" & "','同主旨',null)"
                        cnnConnection.Execute strSql
                 End If
            End If
            'end 2018/09/27
         End If
      End If
   'Added by Lydia 2019/07/04 FCP衍生設計若有需要變更名稱者，請程序在新案建檔勾選重新產生命名記錄存檔後，再重設工程師組別並且發命名通知Email；
   ElseIf ChkAddTct.Value = 1 Then
         strSql = "INSERT INTO TransCaseTitle(TCT01,TCT16,TCT17,TCT24,TCT23,TCT112,TCT113,TCT114) " & _
                    "VALUES ('" & m_CP09 & "','" & ChgSQL(Text5.Text) & "','" & ChgSQL(Text6.Text) & "'," & CNULL(txtTF(19).Text) & ", " & CNULL(txtTF(20).Text) & ",'" & strUserNum & "'," & strSrvDate(1) & "," & Mid(Format(ServerTime, "000000"), 1, 4) & ")"
         cnnConnection.Execute strSql, intI

   'Added by Lydia 2018/05/10 若有修改名稱,覆蓋命名作業名稱
   Else
        strSql = ""
        If Text5.Tag <> Text5.Text Then strSql = strSql & ", tct16=" & CNULL(ChgSQL(Text5.Text))
        If Text6.Tag <> Text6.Text Then strSql = strSql & ", tct17=" & CNULL(ChgSQL(Text6.Text))
        
        'Added by Lydia 2018/07/12 相似度和相似案(覆蓋)
        If Val(m_TCT24) <> Val(txtTF(19).Text) Then strSql = strSql & ", tct24=" & CNULL(txtTF(19).Text, True)
        If m_TCT23 <> txtTF(20).Text Then strSql = strSql & ", tct23=" & CNULL(txtTF(20).Text)
        'end 2018/07/12
        
        If strSql <> "" Then
              strExc(0) = "select tct01,tct16,tct17 from transcasetitle where tct01='" & m_CP09 & "' "
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
              If intI = 1 Then
                   If InStr(UCase(strSql), "TCT23") > 0 Or InStr(UCase(strSql), "TCT24") > 0 Then
                       If txtTF(20).Text <> "" Or Val(txtTF(19).Text) > 0 Then
                           strSql = strSql & ", tct22='Y' "
                       Else
                           strSql = strSql & ", tct22=null "
                       End If
                   End If
                   strSql = "update transcasetitle set " & Mid(strSql, 2) & " where tct01='" & m_CP09 & "' "
                   'Modified by Lydia 2018/10/19 +詳細記錄
                   'Pub_SeekTbLog strSql  '新增log(命名記錄)
                   'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
                   'Modified by Lydia 2025/10/30 改用模組判斷
                   'Pub_SeekTbLog strSql, , True, , Me.Caption & "(" & Me.Name & ")"
                   Pub_SeekTbLog strSql, , PUB_FilterSeekSQL(strSql), , Me.Caption & "(" & Me.Name & ")"
                   cnnConnection.Execute strSql, intI
              End If
        End If
   'end 2018/05/10
   End If
   'end 2017/11/14
   
   'Added by Lydia 2024/10/22 英文組的工程師命名記錄依照新案建檔的修改而變更---Sharon
   If m_UpdPA63TCT118 <> "" Then
      cnnConnection.Execute m_UpdPA63TCT118
   End If
   
   cnnConnection.CommitTrans
            
   FormSave = True
    'Added by Lydia 2020/03/17 FMP案件不自動上年費不續辦，改發清單給程序，由各區程序逐筆產生定稿通知大陸代理人
    If strPassSql <> "" Then
       If PUB_GetP605Email("1", strPassSql, strErrMsg) = False Then
          If strErrMsg <> "" Then
              MsgBox strErrMsg, vbCritical
          End If
       End If
    End If
    'end 2020/03/17
    
    Set rsAD = Nothing 'Added by Lydia 2018/09/27
    Exit Function
    
CheckingErr:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      'Modified by Lydia 2020/03/17
      'MsgBox Err.Description, vbCritical
      MsgBox Err.Description & vbCrLf & strErrMsg, vbCritical
   End If
   Set rsAD = Nothing 'Added by Lydia 2018/09/27
End Function

'Add by Amy 2013/05/20
Private Sub GetCombo4Data(ByVal strTmp As String)
Dim i As Integer
   
   Combo4.Clear
   Combo4.AddItem ""
   If strTmp = "" Then Exit Sub
   'Modified by Lydia 2016/03/15 發明人輸入比對兼自動代入(模糊比對)
'   strExc(0) = "Select " & IIf(Text1 <> "P", " NVL(IN05,NVL(IN04,IN06))||'('||IN01||IN02||')' ", " NVL(IN04,NVL(IN05,IN06))||'('||IN01||IN02||')' ") & " From Inventor Where IN01 IN (" & strTmp & ")"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   With RsTemp
'      If intI = 1 Then
'         Do While Not .EOF
'            Combo4.AddItem .Fields(0)
'            .MoveNext
'         Loop
'      End If
'   End With
   'Modify By Sindy 2018/5/18 將新案建檔發明人欄位的拉選視窗以發明人的英文名稱由小到大排序，若沒建英文抓中文名稱再抓日文名稱
'    strExc(0) = "Select " & IIf(Text1 <> "P", " NVL(IN05,NVL(IN04,IN06))||'('||IN01||IN02||')' ", " NVL(IN04,NVL(IN05,IN06))||'('||IN01||IN02||')' ") & ",IN01, IN02, IN04, IN05, IN06 " & _
'               "FROM INVENTOR WHERE IN01 IN (" & strTmp & ") order by IN01, IN02 "
   'Modify By Sindy 2018/7/17 新案建檔只有外專在使用,依外專提的需求做sort
'    If Text1 = "P" Then
'      strExc(0) = "Select NVL(IN04,NVL(IN05,IN06))||'('||IN01||IN02||')' as sort,IN01, IN02, IN04, IN05, IN06" & _
'                 " FROM INVENTOR WHERE IN01 IN (" & strTmp & ") order by IN01, IN02"
'    Else
      strExc(0) = "Select NVL(IN05,NVL(IN04,IN06))||'('||IN01||IN02||')' as sort,IN01, IN02, IN04, IN05, IN06" & _
                 " FROM INVENTOR WHERE IN01 IN (" & strTmp & ") order by sort"
'    End If
    '2018/5/18 END
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    With RsTemp
       If intI = 1 Then
          Do While Not .EOF
             Combo4.AddItem .Fields(0)
             If RsTemp.AbsolutePosition = 1 Then
                Erase m_InventorList '清空陣列
                ReDim m_InventorList(RsTemp.RecordCount - 1) '定義陣列
                m_InventorListCount = 0
             End If
                strExc(1) = "" & RsTemp.Fields("IN01")
                strExc(2) = "" & RsTemp.Fields("IN02")
                strExc(4) = "" & RsTemp.Fields("IN04")
                strExc(5) = "" & RsTemp.Fields("IN05")
                strExc(6) = "" & RsTemp.Fields("IN06")
                AddInventor strExc(1), strExc(2), strExc(4), strExc(5), strExc(6)
             .MoveNext
          Loop
       End If
    End With
    'end 2016/03/15
End Sub

Private Sub SetCombo4Data(ByVal strData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   
   For nPos = 0 To Combo4.ListCount - 1
      'Modify By Sindy 2015/12/4
      'If Combo4.List(nPos) = strData Then
      'Modify By Sindy 2018/1/23 + And strData <> ""
      If InStr(Combo4.List(nPos), strData) > 0 And strData <> "" Then
      '2015/12/4 END
         bFind = True
         Exit For
      End If
   Next nPos
   If Not bFind Then
      Combo4.AddItem strData
      Combo4.Refresh
      Combo4.ListIndex = Combo4.ListCount - 1
      Call InvFieldEnabled(True)  'Added by Lydia 2022/03/25 控制發明人欄位是否可點選
   Else
      Combo4.ListIndex = nPos
      Call InvFieldEnabled(False)  'Added by Lydia 2022/03/25 控制發明人欄位是否可點選
   End If
End Sub

'Private Function GetInventorName(strIn As String) As String
'Dim rsA  As New ADODB.Recordset
'Dim StrSQLa As String
'
'GetInventorName = ""
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'StrSQLa = "SELECT " & IIf(strSysKind <> "P", " NVL(IN05,NVL(IN04,IN06))||'('||IN01||IN02||')' ", " NVL(IN04,NVL(IN05,IN06))||'('||IN01||IN02||')' ") & " FROM INVENTOR WHERE IN01||IN02='" & strIn & "'"
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   If Not IsNull(rsA.Fields(0).Value) Then
'      GetInventorName = "" & rsA.Fields(0).Value
'   End If
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'
'End Function

'Add by Amy 2013/05/20 更新發明人
Private Sub InsInventor(ByRef m_PI06, ByVal InvNo As String, ByVal InvCh As String, ByVal InvEng As String, ByVal InvJP As String, ByVal IN11 As String)
   Dim strIns As String, m_IN01 As String, m_IN02 As String
   
   'Modified by Morgan 2015/1/6 有更名會錯
   'm_IN01 = InvNo & String(8 - Len(InvNo), "0")
   m_IN01 = Left(ChangeCustomerL(InvNo), 8)
   'end 2015/1/6
   m_IN02 = PUB_GetNewIN02(m_IN01)
   m_PI06 = m_IN01 & m_IN02
   strIns = "Insert Into Inventor (IN01,IN02,IN04,IN05,IN06,IN11) Values(" & CNULL(ChgSQL(m_IN01)) & "," & CNULL(ChgSQL(m_IN02)) & "," & _
               CNULL(ChgSQL(InvCh)) & "," & CNULL(ChgSQL(InvEng)) & "," & CNULL(ChgSQL(InvJP)) & "," & CNULL(ChgSQL(IN11)) & ")"
   cnnConnection.Execute strIns
End Sub
'end 2013/05/20

''Add By Sindy 2015/3/5 修改發明人的中英日名稱時可存檔
'Private Sub UpdateInventor()
'    Dim strUpd As String, m_IN01 As String, m_IN02 As String
'
'    m_IN01 = Left(Combo4.Text, 8)
'    m_IN02 = Mid(Combo4.Text, 9, 2)
'    strUpd = "update Inventor set" & _
'             " IN04=" & CNULL(txtInvField(0)) & ",IN05=" & CNULL(txtInvField(1)) & ",IN06=" & CNULL(txtInvField(2)) & _
'             " where IN01='" & m_IN01 & "' and IN02='" & m_IN02 & "'"
'    cnnConnection.Execute strUpd
'End Sub

'Add By Sindy 2014/11/10
Private Sub Grd1_Click()
Dim nCol As Integer, nRow As Integer
Dim iCol As Integer
   
   With GRD1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow > 0 Then 'And .TextMatrix(nRow, 1) <> "" Then
      nCol = .col
      If pPrevRow > 0 Then
         If pPrevRow <> nRow Then
            .row = pPrevRow
            .TextMatrix(pPrevRow, 0) = ""
            If .FixedCols > 0 Then
               .col = .FixedCols - 1
               .CellBackColor = .BackColorFixed
               .CellForeColor = .ForeColor
            End If
            For iCol = .FixedCols To .Cols - 1
               .col = iCol
               .CellBackColor = .BackColor
            Next
         End If
      End If
   
      If nRow > 0 Then
         .row = nRow
         .TextMatrix(nRow, 0) = "V"
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorSel
            .CellForeColor = .ForeColorSel
         End If
         For iCol = .FixedCols To .Cols - 1
           .col = iCol
           .CellBackColor = &HFFC0C0
         Next
      End If
      .col = nCol
      pPrevRow = nRow
      Call SetCombo4Data(.TextMatrix(nRow, 1))
      'Add By Sindy 2015/3/5
      If .TextMatrix(nRow, 1) = "" Then
         txtInvField(0) = .TextMatrix(nRow, 2)
         txtInvField(1) = .TextMatrix(nRow, 3)
         txtInvField(2) = .TextMatrix(nRow, 4)
         'Modified by Lydia 2024/12/03 (nRow, 5)=>(nRow, 7)
         'txtIN11 = .TextMatrix(nRow, 5)
         Lb_IN11N = .TextMatrix(nRow, 5)
         txtIN11 = .TextMatrix(nRow, 7)
         'end 2024/12/03
         cmdUpdRow.Enabled = True
         cmdAddRow.Enabled = False
      End If
      '2015/3/5 END
   End If
   .Visible = True
   End With
End Sub

Private Sub Text1_GotFocus()
   TextInverse text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   'Modified by Morgan 2012/5/16 +P,PS,CFP,CPS
   'If Text1 <> "FCP" And Text1 <> "FG" Then
   If text1 <> "FCP" And text1 <> "FG" And text1 <> "P" And text1 <> "PS" And text1 <> "CFP" And text1 <> "CPS" Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      Cancel = True
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2022/6/10
   'If KeyAscii <> 89 And KeyAscii <> 8 Then
   If KeyAscii <> 89 And KeyAscii <> 8 Then
   'end 2022/6/10
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text15_GotFocus()
   InverseTextBox Text15
End Sub

Private Sub Text16_GotFocus()
   InverseTextBox Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text17_GotFocus()
   InverseTextBox Text17
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Lydia 2016/08/18 +N
   'Modified by Lydia 2019/11/27 -N
   'If KeyAscii <> 89 And KeyAscii <> 8 And KeyAscii <> 78 Then
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text18_GotFocus()
   InverseTextBox Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text19_GotFocus()
   InverseTextBox Text19
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text20_GotFocus()
    TextInverse Me.Text20
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text20_Validate(Cancel As Boolean)
    If ChgType(20) = False Then Cancel = True: Text20_GotFocus
End Sub

Private Sub Text21_GotFocus()
    TextInverse Me.Text21
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text21_Validate(Cancel As Boolean)
    If ChgType(21) = False Then Cancel = True: Text21_GotFocus
End Sub

Private Sub Text22_GotFocus()
    TextInverse Me.Text22
End Sub

Private Sub Text23_GotFocus()
   TextInverse Text23
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text24_GotFocus()
   TextInverse Text24
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2014/5/30 +68(D)
   If KeyAscii <> 89 And KeyAscii <> 68 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text25_GotFocus()
 TextInverse Text25
End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text26_GotFocus()
   InverseTextBox Text26
End Sub

Private Sub Text27_GotFocus()
   InverseTextBox Text27
End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/12/29
    KeyAscii = UpperCase(KeyAscii)
    'End
End Sub

Private Sub Text28_GotFocus()
   InverseTextBox Text28
End Sub

Private Sub Text28_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/12/29
    KeyAscii = UpperCase(KeyAscii)
    'End
End Sub

Private Sub Text29_GotFocus()
   InverseTextBox Text29
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub

Private Sub Text30_GotFocus()
   InverseTextBox Text30
End Sub

Private Sub Text30_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text31_GotFocus()
   InverseTextBox Text31
End Sub

Private Sub Text32_GotFocus()
   InverseTextBox Text32
End Sub

Private Sub Text32_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> Asc("Y") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text27_Validate(Cancel As Boolean)
   Text27 = UCase(Text27)
   If ChgType(27) = False Then Cancel = True
End Sub

Private Sub Text28_Validate(Cancel As Boolean)
   Text28 = UCase(Text28)
   If ChgType(28) = False Then Cancel = True
End Sub

Private Sub Text30_Validate(Cancel As Boolean)
   Text30 = UCase(Text30)
   If ChgType(30) = False Then Cancel = True
End Sub

Private Sub Text33_Change(Index As Integer)
   'Modified by Morgan 2019/8/13 沒有申請人1按鈕(會錯把優先權資料按鈕隱藏)
   'If Index >= 9 And Index <= 13 Then
   If Index >= 10 And Index <= 13 Then
      If Text33(Index) <> "" Then
         cmdOK(Index - 5).Visible = True
      Else
         cmdOK(Index - 5).Visible = False
      End If
   End If
End Sub

Private Sub Text33_GotFocus(Index As Integer)
   InverseTextBox Text33(Index)
End Sub

Private Sub Text33_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      'Modified by Lydia 2017/05/17 +18(相似案號)
      'Modified by Lydia 2018/06/07 - 18
      Case 9, 10, 11, 12, 13, 14
         KeyAscii = UpperCase(KeyAscii)
      'Added by Lydia 2017/05/17 原文字數、相似度
      'Remove by Lydia 2018/06/07
'      Case 16, 17
'         KeyAscii = Pub_NumAscii(KeyAscii)
       'end 2018/06/07
   End Select
End Sub

Private Sub Text33_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 3
         'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
         'If Not CheckLengthIsOK(Text33(Index), 10) Then
         If Not CheckLengthIsOK(Text33(Index), 30) Then
            Cancel = True
         End If
      Case 1, 4
         'Modified by Lydia 2017/06/14
         'If Not CheckLengthIsOK(Text33(Index), Text33(Index).MaxLength) Then
         If Not CheckLengthIsOK(Text33(Index), 35) Then
            Cancel = True
         End If
      Case 2, 5, 15
         'Modified by Lydia 2017/06/14
         'If Not CheckLengthIsOK(Text33(Index), Text33(Index).MaxLength) Then
         If Not CheckLengthIsOK(Text33(Index), 60) Then
            Cancel = True
         End If
      Case 9, 10, 11, 12, 13, 14
         If Text33(Index) <> "" Then
            If ChgType(Index) = False Then
               Cancel = True
               TextInverse Text33(Index)
            End If
         Else
            Me.Label27(Index - 9).Caption = ""
         End If
         
         '2009/6/24 add by sonia 若輸入9碼且最後一碼不為"0"
         'Modified by Lydia 2015/12/15 Text33(14).MaxLength從8設為9
         If Index = 14 And Cancel = False Then
            If Len(Text33(Index)) = 9 And Right(Text33(Index), 1) <> "0" Then
               MsgBox "此代理人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
               Cancel = True
               TextInverse Text33(Index)
            End If
         End If
         '2009/6/24 end
      'Added by Lydia 2017/05/17 原文字數、相似度
      'Remove by Lydia 2018/06/07 翻譯分案無紙化：欄位改成txtTF
'      Case 16, 17
'         If Text33(Index).Text <> "" Then
'            If Index = 17 And Val(Text33(17).Text) > 100 Then
'               MsgBox "相似度不可大於100！"
'               Cancel = True
'               TextInverse Text33(Index)
'            Else
'               Text33(Index) = Val(Text33(Index))
'            End If
'         End If
'      Case 18 '相似案號
'         If Text33(Index).Text <> "" Then
'            Call ChgCaseNo(Text33(Index), strExc)
'            If ClsPDCheckCaseCodeIsExist(strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
'               Cancel = True
'               TextInverse Text33(Index)
'            Else
'               Text33(Index).Text = strExc(1) & strExc(2) & strExc(3) & strExc(4)
'            End If
'         End If
      'end 2017/05/17
   
   End Select
End Sub
Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text42_GotFocus()
   InverseTextBox Text42
End Sub

Private Sub Text42_Validate(Cancel As Boolean)
   Text42 = UCase(Text42)
   If ChgType(42) = False Then Cancel = True
End Sub

Private Sub Text43_GotFocus()
   InverseTextBox Text43
End Sub

Private Sub Text44_GotFocus()
   InverseTextBox Text44
End Sub

Private Sub Text44_Validate(Cancel As Boolean)
   Text44 = UCase(Text44)
   If ChgType(44) = False Then Cancel = True
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String, strTxt As String, j As Integer
   ChgType = False
   Select Case i
      Case 9, 10, 11, 12, 13
         If ClsPDGetCustomer(Text33(i), strTempName) Then
            Label27(i - 9) = strTempName
            If i = 9 Then
               If m_CP60 <> "" And InStr(ChangeCustomerL(pa(26)), ChangeCustomerL(Text33(i).Text)) = 0 Then
                  strExc(1) = pa(1)
                  strExc(2) = pa(2)
                  strExc(3) = pa(3)
                  strExc(4) = pa(4)
                  strExc(5) = m_CP60
                  strExc(6) = Text33(i)
                  strExc(7) = strTempName
                  If ClsLawUpdAcc0k0(strExc()) Then
                     ChgType = True
                  Else
                     Label27(i - 9) = ""
                  End If
               Else
                  ChgType = True
               End If
            Else
               ChgType = True
            End If
            
         Else
            Label27(i - 9) = ""
         End If
      Case 14
         strExc(1) = Text33(i).Text
         If ClsPDGetAgent(strExc(1), strTempName) Then
            Text33(i).Text = strExc(1)
            Label27(i - 9) = strTempName
            ChgType = True
         Else
            Label27(i - 9) = ""
         End If
      Case 27, 28, 30, 42, 44, 20, 21
         Select Case i
            Case 20
               strTxt = Text20: j = 11
            Case 21
               strTxt = Text21: j = 12
            Case 27
               strTxt = Text27: j = 6
            Case 28
               strTxt = Text28: j = 7
            Case 30
               strTxt = Text30: j = 8
            Case 42
               strTxt = Text42: j = 9
            Case 44
               strTxt = Text44: j = 10
         End Select
         If strTxt = "" Then
            Label27(j) = ""
            ChgType = True
         Else
            If ClsLawLawGetName(strTxt, strTempName) Then
               Label27(j) = strTempName
               Select Case i
                  Case 20
                     Text20 = strTxt
                  Case 21
                     Text21 = strTxt
                  Case 27
                     Text27 = strTxt
                  Case 28
                     Text28 = strTxt
                  Case 30
                     Text30 = strTxt
                  Case 42
                     Text42 = strTxt
                  Case 44
                     Text44 = strTxt
               End Select
               ChgType = True
            End If
         End If
   End Select
End Function

' 清除資料表
Private Sub FormClear()
 Dim i As Integer, txt As Object, Lbl As Object
 
   For Each txt In Text33
      txt.Text = ""
   Next
   For Each Lbl In Label27
      Lbl.Caption = ""
   Next
   
   Text33(12).Enabled = True
   Text33(13).Enabled = True
   For i = 0 To 8
      Text33(i).Enabled = True
   Next
   Text10.Enabled = True
   Text11.Enabled = True
   Text15.Enabled = True
   Text16.Enabled = True
   Text17.Enabled = True
   Text18.Enabled = True
   
   'Add by Morgan 2006/2/7 漏掉了
   Text21.Enabled = True
   Text22.Enabled = True
   
   Text28.Enabled = True
   Text29.Enabled = True
   Text30.Enabled = True
   Text32.Enabled = True
   Text44.Enabled = True
   Text45.Enabled = True
   Text46.Enabled = True
   Text47.Enabled = True

   Combo1(1).Enabled = True
   Combo1(0).Clear
   Combo1(1).Clear
   Combo1(0).AddItem ""
   Combo1(1).AddItem ""
   
   Text5 = "":   Text6 = ""
   Text7 = "":   Text8 = "":   Text9 = "":   Text10 = "":  Text11 = "":  Text12 = ""
   Text13 = "":  Text14 = "":  Text15 = "":  Text16 = "":  Text17 = "":  Text18 = ""
   Text19 = "":  Text26 = "":  Text27 = "":  Text28 = "":  Text29 = "":  Text30 = ""
   Text31 = "":  Text32 = "":  Text42 = "":  Text43 = "":  Text44 = "":  Text45 = ""
   Text46 = "":  Text47 = ""
   Text25 = "" 'Modify By Amy 2013/05/14 增加電子送件
   'Added by Lydia 2018/05/09 新案進度的欄位
   Text19.Tag = "": Text8.Tag = "": Text9.Tag = "": Text25.Tag = ""
   Text5.Tag = "": Text6.Tag = "": Text7.Tag = ""  'Added by Lydia 2018/05/10
   
   Me.Text20.Text = "": Me.Text21.Text = "": Me.Text22.Text = ""
   Text23 = "": Text24 = ""
   'add by toni 2008/10/14 增加FCP工程師組別
   'text34 = ""
   Combo3 = ""
   'end 2008/10/18
   
   'Added by Lydia 2017/11/14 FCP案件命名電子化
   Combo3.Tag = ""
   m_TCT01 = ""
   m_TCT04 = ""
   'Added by Lydia 2018/04/16
   m_TCT16 = ""
   m_TCT17 = ""
   'Added by Lydia 2018/06/12
   m_TCT10 = ""
   m_TCT27 = ""
   m_TCT28 = ""
   m_TCT23 = ""
   m_TCT24 = ""
   
   'Added by Lydia 2017/11/17 設計案屬性
   Combo5.Text = ""
   Combo5.Tag = ""

   'Added by Lydia 2019/10/25 翻譯瑕疵備註
   Combo6.Text = ""
   Combo6.Tag = ""
   
   'Add by Amy 2013/05/20 增加發明人
   SSTab1.TabEnabled(6) = True

   Combo4.Clear
   Combo4.AddItem ""
   For Each m_otxt In txtInvField
      m_otxt.Text = ""
      m_otxt.Tag = "" 'Add By Sindy 2015/3/5
   Next
   'Add By Sindy 2014/12/9
   GRD1.Clear
   Call SetGrd(GRD1)
   '2014/12/9 END
   
   'IN11 國籍
   txtIN11 = ""
   Lb_IN11N = "" 'Add By Sindy 2015/12/4
   
   'Add by Morgan 2008/11/14
   '新增欄位改用陣列idex對應欄位序號，以後新增就不必再改。
   For Each m_otxt In txtPA
      m_otxt = ""
      m_otxt.Tag = ""
   Next
   'end 2008/11/14
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   ChkPA174.Visible = False
   ChkPA174.Value = vbUnchecked: ChkPA174.Tag = ""
   CmdPA174.Visible = False
   bolAskPA174 = False
   'end 2020/02/21
   
   'Added by Lydia 2021/04/09 預設「有序列表」
   ChkPA175.Visible = False
   ChkPA175.Value = vbUnchecked
   ChkPA175.Tag = ""
   'end 2021/04/09
   
   cmdOK(0).Enabled = False
   cmdOK(1).Enabled = False
   cmdOK(3).Enabled = False
   'Added by Lydia 2018/06/27
   cmdOK(4).Enabled = False
   Command2(6).Enabled = False
   
   'Add By Sindy 2023/4/18
   txtCP142.Text = ""
   Option1(0).Value = False
   Option1(1).Value = False
   Option1(2).Value = False
   '2023/4/18 END
   
   'Added by Lydia 2017/05/17 原文字數,相似度和相似案號
   'Modified by Lydia 2018/06/07 翻譯分案無紙化：欄位改成txtTF
'   Text33(16).Visible = False: Text33(17).Visible = False: Text33(18).Visible = False
'   Label26(2).Visible = False: Label26(3).Visible = False: Label26(4).Visible = False
'   'Added by Lydia 2017/12/01
'   m_TF23 = 0
'   m_TF19 = 0
'   m_TF20 = ""
   For Each txt In txtTF
      txt.Text = ""
      txt.Tag = ""
   Next
   fraTrans01.Visible = False
   fraTrans02.Visible = False
   fraTrans03.Visible = False
   m_TF01 = ""
   m_TF01pty = ""
   m_TF01cp14 = ""
   m_TF01cp27 = ""
   m_TF01cp60 = "" 'Added by Lydia 2019/06/28
   cboSource.Text = "": cboSource.Tag = ""
   cboTarget.Text = "": cboTarget.Tag = ""
   Chk01.Value = 0: Chk01.Tag = ""
   Chk02.Value = 0: Chk02.Tag = ""
   Chk03.Value = 0: Chk03.Tag = ""
   Chk04.Value = 0: Chk04.Tag = ""
   cmdOpen(0).Enabled = False
   cmdOpen(1).Enabled = False
   'end 2018/06/07
    
    'Added by Lydia 2018/08/24
    Chk05.Value = 0: Chk05.Tag = ""
    Chk06.Value = 0: Chk06.Tag = ""
    
    'Added by Lydia 2019/07/04
    ChkAddTct.Value = False
    ChkAddTct.Visible = False
    
   'Added by Lydia 2020/01/20 專利案件和English_Vers檔案：判斷檔案上傳目的地
   If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
        cmdOpen(0).Caption = "原始檔"
   Else
        cmdOpen(0).Caption = "外文本"
   End If
   cmdOpen(0).Tag = ""
   'end 2020/01/20
   'Added by Lydia 2023/03/07
   bolUpdTCN23 = False
   m_TCN19 = ""
End Sub

Private Sub Text45_GotFocus()
   InverseTextBox Text45
End Sub

Private Sub Text46_GotFocus()
   InverseTextBox Text46
End Sub

Private Sub Text47_GotFocus()
   InverseTextBox Text47
End Sub

Private Sub Text5_GotFocus()
   InverseTextBox Text5
End Sub
'Added by Morgan 2018/1/19
Private Sub Text5_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Remove by Lydia 2018/04/19 智慧局來函的中文名稱為全形，實際上客戶對全形和半形字不在意
   '               ，所以還是與其他系統一致取消轉全形，已與Phoebe、Sharon確認
   'KeyAscii = Asc(ToWide(Chr(KeyAscii)))
End Sub


Private Sub Text5_Validate(Cancel As Boolean)
   Dim iMaxLen As Integer
'Modify by Morgan 2007/4/30 SP欄位長度改和PA一樣
'   'Add by Morgan 2006/10/11 FG長度另外控制
'   If Text1 = "FG" Then
'      iMaxLen = 140
'   Else
'      iMaxLen = Text5.MaxLength
'   End If
   iMaxLen = Text5.MaxLength
'end 2007/4/30
   
   'Removed by Morgan 2019/10/8 欄位已改char長度等於字數不必再檢查(欄位長度自動會限制內容)
   'If Not CheckLengthIsOK(Text5, iMaxLen) Then
   '   Cancel = True
   'End If
   'end 2019/10/8
   
End Sub

Private Sub Text6_GotFocus()
   InverseTextBox Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   Dim iMaxLen As Integer
'Modify by Morgan 2007/4/30 SP欄位長度改和PA一樣
'   'Add by Morgan 2006/10/11 FG長度另外控制
'   If Text1 = "FG" Then
'      iMaxLen = 60
'   Else
'      iMaxLen = Text6.MaxLength
'   End If
   iMaxLen = Text6.MaxLength
'end 2007/4/30

   'Removed by Morgan 2019/10/8 欄位已改char長度等於字數不必再檢查(欄位長度自動會限制內容)
   'If Not CheckLengthIsOK(Text6, iMaxLen) Then
   '   Cancel = True
   'End If
   'end 2019/10/8
   
End Sub

Private Sub Text7_GotFocus()
   InverseTextBox Text7
End Sub

'2011/9/8 CANCEL BY SONIA 移至
'Private Sub Text7_LostFocus()
'   If Text5 = "" And Text6 = "" And Text7 = "" Then
'      MsgBox "案件名稱不可同時空白 !", vbCritical
'      Text5.SetFocus
'   End If
'End Sub
'2011/9/8 END

Private Sub Text7_Validate(Cancel As Boolean)
Dim iMaxLen As Integer
   '2011/9/8 ADD BY SONIA 自Text7_LostFocus移過來
   If Text5 = "" And Text6 = "" And Text7 = "" Then
      MsgBox "案件名稱不可同時空白 !", vbCritical
      Text5_GotFocus
      Cancel = True
   End If
   '2011/9/8 END
'Modify by Morgan 2007/4/30 SP欄位長度改和PA一樣
'   'Add by Morgan 2006/10/11 FG長度另外控制
'   If Text1 = "FG" Then
'      iMaxLen = 60
'   Else
'      iMaxLen = Text7.MaxLength
'   End If
   iMaxLen = Text7.MaxLength
'end 2007/4/30

   'Removed by Morgan 2019/10/8 欄位已改char長度等於字數不必再檢查(欄位長度自動會限制內容)
   'If Not CheckLengthIsOK(Text7, iMaxLen) Then
   '   Cancel = True
   'End If
   'end 2019/10/8
End Sub

'Add By Sindy 2025/8/18
Private Sub txtCP142_GotFocus()
   InverseTextBox txtCP142
End Sub
Private Sub txtCP142_Validate(Cancel As Boolean)
   If txtCP142 <> "" Then
      If ChkDate(txtCP142) = False Then Cancel = True: TextInverse txtCP142
   End If
End Sub
'2025/8/18 END

Private Sub Text8_GotFocus()
   InverseTextBox Text8
End Sub

'Modify By Sindy 2023/4/18 取消當所限=法限時，為指定當天
''Add By Sindy 2015/12/16
'Private Sub Text8_LostFocus()
'   Check1.Value = 0
'   If Text8 = "" Then Exit Sub
'   If Text8 = Text9 Then
'      Check1.Value = 1
'   End If
'End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 <> "" Then
      If ChkDate(Text8) = False Then Cancel = True: TextInverse Text8
   Else
      If m_CP10 = 年費 Then
         MsgBox "案件性質為年費時，一定要有本所期限 !", vbCritical
         Cancel = True
      End If
   End If
End Sub

Private Sub Text9_GotFocus()
   InverseTextBox Text9
End Sub
'2008/11/26 CANCEL BY SONIA 移至Text9_Validate
'Private Sub Text9_LostFocus()
'   If Text8 <> "" Then
'      If Text8 > Text9 Then
'         MsgBox "本所期限不能大於法定期限 !", vbCritical
'         Text9.SetFocus
'      End If
'   End If
'End Sub
'2008/11/26 END

'Modify By Sindy 2023/4/18 取消當所限=法限時，為指定當天
''Add By Sindy 2015/12/16
'Private Sub Text9_LostFocus()
'   Check1.Value = 0
'   If Text9 = "" Then Exit Sub
'   If Text8 = Text9 Then
'      Check1.Value = 1
'   End If
'End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> "" Then
      '2008/11/26 ADD BY SONIA
      If Val(Text8) > Val(Text9) Then
         MsgBox "本所期限不能大於法定期限 !", vbCritical
         Text9.SetFocus
         Cancel = True
      Else
      '2008/11/26 END
         Cancel = Not ChkDate(Text9)
      End If
      
      'Added by Morgan 2014/11/10
      If Cancel = False Then
         If Text8 = "" Then
            'Modified by Morgan 2014/11/19 所限改為法限的前2個日曆天
            'Text8 = TransDate(PUB_GetOurDeadline(Text9), 1)
            'Modified by Morgan 2023/5/24
            'Text8 = TransDate(CompDate(2, -2, Text9), 1)
            Text8 = TransDate(PUB_GetFCPOurDeadline(Text9, 2), 1)
            'end 2023/5/24
            'end 2014/11/19
         End If
      End If
      'end 2014/11/10
      
   Else
      If m_CP10 = 年費 Then
         MsgBox "案件性質為年費時，一定要有法定期限 !", vbCritical
         Cancel = True
      End If
   End If
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
'Add by Amy 2013/05/20
Dim strNo As String ', strSql As String
'Dim strk1(1 To 100) As String, strk2(1 To 100) As String, strk3(1 To 100) As String
'Dim j As Integer, k1 As Integer, k2 As Integer, k3 As Integer
'Dim rsTmp  As New ADODB.Recordset
'end 2013/05/20
Dim ii As Integer
Dim Cancel As Boolean
      
   'Added by Lydia 2020/02/21 檢查案號是否正確
   If CheckFindPass = False Then
       Exit Function
   End If
   'end 2020/02/21
   
   TxtValidate = False
   If Me.text1.Enabled = True Then
      Cancel = False
      Text1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Added by Morgan 2018/1/19
   'Remove by Lydia 2018/04/19 智慧局來函的中文名稱為全形，實際上客戶對全形和半形字不在意，所以還是與其他系統一致取消轉全形，已與Phoebe、Sharon確認
'   strExc(1) = ToWide(Text5)
'   If strExc(1) <> Text5 Then
'      MsgBox "中文專利名稱有半形字，將自動轉為全形！", vbInformation
'      If Not CheckLengthIsOK(strExc(1), Text5.MaxLength, False) Then
'         MsgBox "轉全形後之中文專利名稱超過長度，請修正！", vbInformation
'         Exit Function
'      Else
'         Text5 = strExc(1)
'         Cancel = False
'         Text5_Validate Cancel
'         If Cancel = True Then
'            Exit Function
'         End If
'      End If
'   End If
   'end 2018/1/19
   
   '2011/9/8 ADD BY SONIA
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   '2011/9/8 END
   If Me.Text20.Enabled = True Then
      Cancel = False
      Text20_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text21.Enabled = True Then
      Cancel = False
      Text21_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text27.Enabled = True Then
      Cancel = False
      Text27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text28.Enabled = True Then
      Cancel = False
      Text28_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text30.Enabled = True Then
      Cancel = False
      Text30_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   strNo = Text33(9)
   For Each objTxt In Text33
      If objTxt.Enabled = True Then
         Cancel = False
         Text33_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
      'Add by Amy 2013/05/20 抓取申請人編號 (for 發明人判斷)
       If objTxt.Index > 9 And objTxt.Index <= 13 Then
            If objTxt <> "" Then
                strNo = strNo & "," & Text33(objTxt.Index)
            End If
       End If
      'end 2013/05/20
   Next
   If Me.Text42.Enabled = True Then
      Cancel = False
      Text42_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text44.Enabled = True Then
      Cancel = False
      Text44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text8.Enabled = True Then
      Cancel = False
      Text8_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text9.Enabled = True Then
      Cancel = False
      Text9_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
      MsgBox "本所案號錯誤，儲存失敗 !", vbCritical
      Exit Function
   End If
   If Text33(9) = "" And Text33(10) = "" And Text33(11) = "" And Text33(12) = "" And _
      Text33(13) = "" And Text33(14) = "" Then
      MsgBox "申請人代理人不可同時空白 !", vbCritical
      Text33(9).SetFocus
      Exit Function
   End If
      
   'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
   If Pub_ChkAppList(strExc(0), Text33(9) & "," & Text33(10) & "," & Text33(11) & "," & Text33(12) & "," & Text33(13)) = False Then
      SSTab1.Tab = 1
      Text33(Val(strExc(0)) + 8).SetFocus
      Text33_GotFocus Val(strExc(0)) + 8
      Exit Function
   End If
   'end 2024/06/14
   
   'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
   For ii = 9 To 14
      strExc(1) = ChangeCustomerL(Text33(ii))
      If ii < 14 Then
         If pa(1) = "FCP" Or pa(1) = "CFP" Or pa(1) = "P" Then
            strExc(2) = ChangeCustomerL(pa(ii + 17))
         Else
            If ii = 9 Then strExc(2) = ChangeCustomerL(pa(8))
            If ii = 10 Then strExc(2) = ChangeCustomerL(pa(58))
            If ii = 11 Then strExc(2) = ChangeCustomerL(pa(59))
            If ii = 12 Then strExc(2) = ChangeCustomerL(pa(65))
            If ii = 13 Then strExc(2) = ChangeCustomerL(pa(66))
         End If
         If strExc(1) <> "" And strExc(1) <> strExc(2) Then
            If GetCustomerAndState(strExc(1), strExc(3), , , , pa(1), strExc(8), False, Me.Name, pa(2), pa(3), pa(4)) = False Then
               SSTab1.Tab = 1
               Text33(ii).SetFocus
               Text33_GotFocus ii
               Exit Function
            End If
         End If
      Else
         If pa(1) = "FCP" Or pa(1) = "CFP" Or pa(1) = "P" Then
            strExc(2) = ChangeCustomerL(pa(75))
         Else
            strExc(2) = ChangeCustomerL(pa(26))
         End If
         If strExc(1) <> "" And strExc(1) <> strExc(2) Then
            If GetAgentAndState(strExc(1), strExc(3), , , , pa(1), strExc(8), False) = False Then
               SSTab1.Tab = 1
               Text33(ii).SetFocus
               Text33_GotFocus ii
               Exit Function
            End If
         End If
      End If
   Next ii
   'end 2024/06/13
   
   'add by toni 2008/10/15 增加FCP工程師組別
   'If Text1 = "FCP" And text34 = "" Then
   'Modify by Morgan 2009/9/9
   'If Text1 = "FCP" And Text2 >= "035187" And Combo3 = "" Then
   'Modified by Lydia 2018/03/05 改到客戶提供文件處理,控制要輸入組別; FCP新案命名從FCP-58447
   'If ((Text1 = "FCP" And Text2 >= "035187") Or (Text1 = "FG" And Text2 >= "000536")) And Combo3 = "" Then
   'Modified by Lydia 2018/05/11 +衍生設計125
   If ((text1 = "FCP" And Text2 >= "035187" And Text2 <= "058443") Or (text1 = "FCP" And Text2 >= "058447" And InStr("101,102,103,125", m_CP10) = 0) Or (text1 = "FG" And Text2 >= "000536")) And Combo3 = "" Then
      MsgBox "請輸入工程師組別", vbCritical
      Cancel = False
      Combo3.SetFocus
      Exit Function
   End If
   'end 2008/10/15
   'Added by Lydia 2019/07/04 FCP衍生設計新案若有需要變更名稱者，請程序在新案建檔勾選重新產生命名記錄存檔後，再重設工程師組別並且發命名通知Email；
   If ChkAddTct.Visible = True And ChkAddTct.Value = 1 And Combo3.Text = "" Then
      MsgBox "請不要同時重新產生命名記錄和設定工程師組別！", vbCritical '若先有組別先改成空白
      Cancel = False
      Combo3.SetFocus
      Exit Function
   End If
   
   '2010/1/8 ADD BY SONIA
   If Combo3.Enabled = True Then
      Cancel = False
      Combo3_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2010/1/8 END
   
   'Added by Lydia 2017/11/17 設計案屬性
   If Combo5.Visible = True Then
      Cancel = False
      Combo5_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 20117/11/17
   
   'Add by Morgan 2011/2/24
   '檢查申請人順序
   If Text33(9).Enabled = True Then
      If (Text33(10) <> "" And Text33(9) = "") Or _
         (Text33(11) <> "" And Text33(10) = "") Or _
         (Text33(12) <> "" And Text33(11) = "") Or _
         (Text33(13) <> "" And Text33(12) = "") Then
         MsgBox "請依順序輸入申請人！"
         
         SSTab1.Tab = 1
         If Text33(9) = "" Then Text33(9).SetFocus: Exit Function
         If Text33(10) = "" Then Text33(10).SetFocus: Exit Function
         If Text33(11) = "" Then Text33(11).SetFocus: Exit Function
         If Text33(12) = "" Then Text33(12).SetFocus: Exit Function
         Exit Function
      End If
   End If
   
   'Added by Lydia 2017/05/17 檢查原文字數和相似度
   'Modified by Lydia 2018/06/07 翻譯分案無紙化：欄位改成txtTF
'   If Text33(16).Visible = True And Text33(16).Tag <> "" And Trim(Replace(Text33(16) & Text33(17) & Text33(18), " ", "")) <> "" Then
'      If Trim(Text33(16)) = "" Then
'         MsgBox "請輸入原文字數！", vbCritical
'         Text33(16).SetFocus
'         Exit Function
'      End If
'      'Memo by Lydia 2017/12/29 淑華要求107/1/1先開放單獨輸入原文字數
'      If Val(Text33(17)) > 100 Then
'         MsgBox "相似度不可大於100！"
'         Text33(17).SetFocus
'         Exit Function
'      End If
'   End If
   'end 2017/05/17
   If (fraTrans01.Enabled = True Or fraTrans02.Enabled = True) And m_TF01 <> "" Then
      If Trim(Replace(txtTF(23) & txtTF(19) & txtTF(20), " ", "")) <> "" Then '有輸入原文字數、相似度和案號才檢查
            'Mark by Lydia 2022/04/08 請取消「針對輸入相似度或相似案號時，需輸入原文字數」，原文字數其他的控制請維持不變(by Sharon)
            'If m_TF01pty = "201" And Trim(txtTF(23)) = "" Then '限新案翻譯
            '   MsgBox "請輸入原文字數！", vbCritical
            '   SSTab1.Tab = 7 'Mark by Lydia 2018/07/18 先隱藏 'ReMark by Lydia 2018/08/07 新案建檔先上線
            '   txtTF(23).SetFocus
            '   Exit Function
            'End If
            'end 2022/04/08
            If Val(txtTF(19)) > 100 Then
               MsgBox "相似度不可大於100！"
               SSTab1.Tab = 7  'Mark by Lydia 2018/07/18 'ReMark by Lydia 2018/08/07 新案建檔先上線
               txtTF(19).SetFocus
               Exit Function
            End If
          'Added by Lydia 2024/05/31 相似案號檢查; 發現FCP-071046的相似案號只有「FCP071020」
          If Trim(txtTF(20)) <> "" Then
              Call txtTF_Validate(20, Cancel)
              If Cancel = True Then
                 SSTab1.Tab = 7
                 txtTF(20).SetFocus
                 Exit Function
              End If
          End If
          'end 2024/05/31
      End If
      'Added by Lydia 2019/06/28 固定請款對象之帳單（LEDES帳單）其請款項目209檢視中說英文敘述後方加上+英文字數
                                      ' 遇到案件代理人為Y48309或其他固定請款對象，並且檢視中說尚未有請款單號則檢查有無輸入原文字數。
      If m_TF01cp60 = "" And fraTrans02.Enabled = True And Text33(14) <> "" And InStr(FCP檢視中說必輸原文字數, ChangeCustomerL(Text33(14))) > 0 And m_TF01pty = "209" Then
           If Trim(txtTF(23)) = "" Then
               MsgBox "請輸入原文字數！", vbCritical, "FCP檢視中說必輸原文字數"
               SSTab1.Tab = 7
               txtTF(23).SetFocus
               Exit Function
           End If
      End If
      
      'Mark by Lydia 2018/07/18 判斷是否隱藏
      If fraTrans01.Enabled = True And SSTab1.TabVisible(7) = True Then
             'Modified by Lydia 2018/08/13 中說發文不受限(舊案)
            If m_TF01pty = "201" And Trim(cboSource.Text) = "" And m_TF01cp27 = "" Then  '限新案翻譯
                 MsgBox "請輸入原文語種！", vbCritical
                 SSTab1.Tab = 7
                 cboSource.SetFocus
                 Exit Function
            Else
                 txtTF(27).Text = Left(cboSource.Text, 1)
            End If
            'Modified by Lydia 2018/08/13 中說發文不受限(舊案)
            If m_TF01pty = "201" And Trim(cboTarget.Text) = "" And m_TF01cp27 = "" Then  '限新案翻譯
                 MsgBox "請輸入翻譯語種！", vbCritical
                 SSTab1.Tab = 7
                 cboTarget.SetFocus
                 Exit Function
            Else
                 txtTF(28).Text = Left(cboTarget.Text, 1)
            End If
            '已有英文本收文號，取消勾選項
            If txtTF(30).Text <> "" Then
                 Chk02.Value = 0
            End If
            
            'Memo by Lydia 2018/08/07 待比對控制,工程師協調8/13上線
            If Val(m_TF01cp27) > 0 And Chk01.Value = 1 And Chk01.Visible = True And txtTF(29).Tag = "" Then
                 MsgBox "新案翻譯已發文，不可勾選待比對 !", vbCritical
                 Chk01.Value = 0
            End If
            If m_TF01cp14 <> "" And InStr(m_GrpManList, m_TF01cp14) = 0 And Chk01.Value = 1 And Chk01.Visible = True And txtTF(29).Tag = "" Then
                 MsgBox "新案翻譯已分案，不可勾選待比對 !", vbCritical
                 Chk01.Value = 0
            End If
            'Modified by Lydia 2024/05/22 未提申先翻譯可以只輸其中一種---Sharon口述
            'If Chk03.Value = 1 And txtTF(26).Text = "" Then
             '   MsgBox "勾選未提申先翻譯，請輸入交稿期限 !", vbCritical
            If Chk03.Value = 1 And Trim(txtTF(26).Text & txtTF(32).Text) = "" Then
                MsgBox "勾選未提申先翻譯，請輸入交稿期限／只交Claims期限 !", vbCritical
                SSTab1.Tab = 7
                txtTF(26).SetFocus
                txtTF_GotFocus 26
                Exit Function
            End If
            If Val(m_CP27) > 0 And Chk03.Value = 1 And txtTF(31).Tag = "" Then
                 MsgBox "新申請案已發文，不可勾選未提申先翻譯 !", vbCritical
                 Chk03.Value = 0
                 txtTF(26).Text = txtTF(26).Tag
                 txtTF(29).Text = txtTF(29).Tag
                 Exit Function
            End If
            'Modified by Lydia 2018/08/28 只交Claim期限和交稿期限不一定要全輸入 +Val(txtTF(26)) > 0
            If Val(txtTF(32)) > 0 And Val(txtTF(32)) >= Val(txtTF(26)) And Val(txtTF(26)) > 0 Then
                MsgBox "只交Claim期限不可晚於交稿期限 !", vbCritical
                SSTab1.Tab = 7
                txtTF(32).SetFocus
                txtTF_GotFocus 32
                Exit Function
            End If
            'Added by Lydia 2019/10/25 翻譯瑕疵備註
            If fraTrans04.Visible = True Then
                Call txtTF_Validate(37, Cancel)
                If Cancel = True Then
                    SSTab1.Tab = 7
                    txtTF(37).SetFocus
                    txtTF_GotFocus 37
                    Exit Function
                End If
            End If
            'Added by Lydia 2025/06/12 因為FCP-073897第一次輸入交稿期限有問題，產生行事曆管制日19110114
            If txtTF(26).Text <> "" Then '交稿期限
                Call txtTF_Validate(26, Cancel)
                If Cancel = True Then
                    SSTab1.Tab = 7
                    txtTF(26).SetFocus
                    txtTF_GotFocus 26
                    Exit Function
                End If
            End If
            If txtTF(32).Text <> "" Then '只交Claims期限
                Call txtTF_Validate(32, Cancel)
                If Cancel = True Then
                    SSTab1.Tab = 7
                    txtTF(32).SetFocus
                    txtTF_GotFocus 32
                    Exit Function
                End If
            End If
            'end 2025/06/12
      End If
   End If
   'end 2018/06/07

''Add by Amy 2013/05/20 檢查發明人
'Memoed by Morgan 2018/1/19 刪除已標記為註解的舊程式碼

'Modify by Sindy 2014/11/10 檢查發明人
If text1 = "FCP" Or text1 = "P" Or text1 = "CFP" Then
   'Added by Morgan 2015/11/18
   '排除有讓與,合併,繼承,專利權讓與發文者
   strExc(0) = "select cp09 from CaseProgress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('701','702','703','708') and cp27>0"
   If intI = 0 Then
   'end 2015/11/18
   
      For ii = 1 To GRD1.Rows - 1
         strExc(0) = Trim(GRD1.TextMatrix(ii, 1))
         If strExc(0) <> "" Then
            If PUB_ChkInventor(strExc(0), strNo) = False Then
                'Modified by Morgan 2014/12/30 strExc(0) 會被回寫/清除
                'If MsgBox("發明人(" & strExc(0) & ")資料與目前申請人不符，是否要繼續！", vbYesNo + vbDefaultButton2) = vbNo Then
                If MsgBox("發明人(" & Trim(GRD1.TextMatrix(ii, 1)) & ")資料與目前申請人不符，是否要繼續！", vbYesNo + vbDefaultButton2) = vbNo Then
                   SSTab1.Tab = 6
                   Exit Function
                End If
            End If
         Else
            'Modified by Morgan 2015/1/8
            'strNo = Text33(9) & String(8 - Len(Text33(9)), "0")
            strNo = Left(ChangeCustomerL(Text33(9)), 8)
            'end 2015/1/8
            If Trim(GRD1.TextMatrix(ii, 6)) <> "" Then
               If Trim(GRD1.TextMatrix(ii, 6)) <> strNo Then
                  If MsgBox("第 " & ii & " 筆發明人資料與目前申請人不符，是否要繼續！", vbYesNo + vbDefaultButton2) = vbNo Then
                     SSTab1.Tab = 6
                     Exit Function
                  End If
               End If
            End If
         End If
      Next ii
      
   End If 'Added by Morgan 2015/11/18
End If
'end 2013/05/20

'Added by Morgan 2015/10/19
'新申請案法定期限不可大於最早優先權日+1年(設計: 6個月)
If pa(1) = "FCP" And InStr("101,102,103", m_CP10) > 0 And m_CP27 = "" And Text9 <> "" Then
   If DBDATE(Text9) < strSrvDate(1) Then
      MsgBox "法定期限不可小於系統日期!!", vbCritical
      SSTab1.Tab = 0
      Text9.SetFocus
      Exit Function
      
   ElseIf strPriority(2) <> "" Then
      strExc(1) = PUB_GetFirstPriDate2(strPriority(2))
      If pa(8) = "3" Then
         intI = 6
         strExc(3) = "設計案新案提申法定期限不可晚於最早優先權日 + 6個月!!"
      Else
         intI = 12
         strExc(3) = "發明/新型案新案提申法定期限不可晚於最早優先權日 + 1年!!"
      End If
      strExc(2) = CompDate(1, intI, strExc(1))
      'Modified by Morgan 2017/3/1 若為假日可順延到下個工作日 Ex.FCP-056104
      'If DBDATE(Text9) > strExc(2) Then
      If DBDATE(Text9) > PUB_GetWorkDay1(strExc(2), False) Then
         MsgBox strExc(3), vbCritical
         SSTab1.Tab = 0
         Text9.SetFocus
         Exit Function
      End If
   End If
End If
'end 2015/10/19
     
   'Added by Lydia 2020/02/21 檢查「名稱有特殊字」
   If text1 = "P" Or text1 = "FCP" Then
       If Pub_GetPA174toFile("2", text1, Text2, Text3, Text4, Me, frm100101_M_1) = True Then
           strExc(1) = "Y"
       Else
           strExc(1) = "N"
       End If
       If ChkPA174.Value = vbUnchecked And strExc(1) = "Y" Then
           If MsgBox("原始檔區已有案件名稱Word檔，請問是否取消「名稱有特殊字」？", vbInformation + vbYesNo + vbDefaultButton2, "檢查資料") = vbNo Then
               Exit Function
           End If
       End If
       If ChkPA174.Value = vbChecked And strExc(1) = "N" Then
           If MsgBox("原始檔區沒有案件名稱Word檔，請問是否繼續作業？", vbInformation + vbYesNo + vbDefaultButton2, "檢查資料") = vbNo Then
               Exit Function
           End If
       End If
       '當「名稱有特殊字」有勾選，並且有修改案件名稱，將原始檔之維護word檔自動打開，並彈訊息提醒。
       If ChkPA174.Value = vbChecked And bolAskPA174 = False Then  '不用再次彈訊息
           If Text5 & Text6 & Text7 <> pa(5) & pa(6) & pa(7) Then
               MsgBox "名稱有特殊字，案件名稱有修改，請一併修改案件名稱Word檔。", vbInformation, "檢查資料"
               Call ProcPA174toFile("Y")
               Exit Function
           End If
       End If
   End If
   'end 2020/02/21
   
   TxtValidate = True
End Function

'Add by Amy 2013/05/20
Private Sub txtIN11_Validate(Cancel As Boolean)
    If txtIN11 = "" Then Exit Sub 'Add By Sindy 2015/12/4
    If Val(txtIN11) >= 1 And Val(txtIN11) <= 8 Then
        MsgBox ("發明人國籍不可輸入 001 - 008")
        Me.Lb_IN11N.Caption = ""
        Cancel = True
    Else
        If ClsPDGetNation(txtIN11, strName) Then
            Me.Lb_IN11N.Caption = strName
        Else
            Me.Lb_IN11N.Caption = ""
            Cancel = True
        End If
   End If
End Sub

Private Sub txtInvField_GotFocus(Index As Integer)
    
    If Combo4 <> "" Then
    'Memo by Lydia 2019/09/26 Sharon: 輸入發明人名稱按Tab的順序為中文->英文->日文->國籍代碼->加入按鈕
'        Lb_IN11.Visible = False
'        txtIN11.Visible = False
'        Lb_IN11N.Visible = False
        Combo4.SetFocus
    Else
    'Memo by Lydia 2019/09/26 Sharon: 輸入發明人名稱按Tab的順序為中文->英文->日文->國籍代碼->加入按鈕
'        Lb_IN11.Visible = True
'        txtIN11.Visible = True
'        Lb_IN11N.Visible = True
        If txtInvField(0) = "" And txtInvField(1) = "" And txtInvField(2) = "" Then
            txtIN11.Text = ""
            Lb_IN11N.Caption = ""
        End If
    End If
End Sub

Private Sub txtInvField_LostFocus(Index As Integer)
'Mark by Lydia 2022/03/25
'    Dim idx As Integer
'    idx = Index \ 3
    If Combo4 = "" And txtInvField(0) = "" And txtInvField(1) = "" And txtInvField(2) = "" Then
'        Lb_IN11.Visible = False
'        txtIN11.Visible = False
'        Lb_IN11N.Visible = False
    End If
'    '滑鼠指標移至下一個Combo4
'    Select Case Index Mod 3
'        Case 2
'            If idx <= 8 Then
'                Combo4(idx + 1).SetFocus
'            End If
'    End Select
'end 2022/03/25

End Sub
''end 2013/05/20

'Add by Morgan 2008/11/14
Private Sub txtPA_GotFocus(Index As Integer)
   TextInverse txtPA(Index)
   CloseIme
End Sub
'Add by Morgan 2008/11/14
Private Sub txtPA_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      'Modify by Morgan 2009/10/28 +153,154
      Case 151, 152, 153, 154
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
      'Add by Morgan 2009/10/28
      'Modified by Lydia 2018/10/17 +PA63
      Case 155, 63
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
      'Added by Lydia 2019/11/27
      Case 156
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then  'Y/N/null
            KeyAscii = 0
            Beep
         End If
      'Added by Morgan 2022/11/30
      Case 178 '證書形式 1:電子 2:紙本
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then  'Y/N/null
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

'Added by Lydia 2016/03/15 發明人輸入比對兼自動代入(模糊比對)
Private Sub txtInvField_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim tRec As Integer, tSearch As Boolean
   Dim tInx As Integer, tSno As Integer 'tInx =combo4(index), tSno=List編號

   Cancel = False
   If IsEmptyText(txtInvField(Index)) = False Then
      If StrLength(txtInvField(Index)) > txtInvField(Index).MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "發明人名稱太長"
         MsgBox strMsg, vbOKOnly + vbCritical, strTit
      Else
         '改為選擇即有發明人或新增發明人->淑華表示用自動帶,若遇到名字相同唯寫法不同,到維護畫面採人工新增
         For tRec = 0 To m_InventorListCount - 1
            If Index = 0 Then '(發明人)中文名稱
               If InStr(m_InventorList(tRec).iN04, txtInvField(Index)) > 0 Then
                 tSearch = True: tInx = Index \ 3: tSno = tRec
                 Cancel = True
                 Exit For
               End If

            ElseIf Index = 1 Then '(發明人)英文名稱
               If InStr(UCase(m_InventorList(tRec).IN05), UCase(txtInvField(Index))) > 0 Then
                 tSearch = True: tInx = Index \ 3: tSno = tRec
                 Cancel = True
                 Exit For
               End If

            ElseIf Index = 2 Then '(發明人)日文名稱
               If InStr(m_InventorList(tRec).IN06, txtInvField(Index)) > 0 Then
                 tSearch = True: tInx = Index \ 3: tSno = tRec
                 Cancel = True
                 Exit For
               End If
            End If
         Next tRec
      End If
   End If

   If Cancel = False Then
      CloseIme
   Else
      If tSearch = True Then
         Combo4.ListIndex = tSno + 1 '讀發明人List=>call combo4_click
         Combo4.SetFocus  '移到比對出的發明人combo List
      End If
   End If
End Sub

'Added by Lydia 2016/03/15 發明人輸入比對兼自動代入(模糊比對)
' 增加發明人
Private Sub AddInventor(ByVal strInventor As String, Optional ByVal mIN02 As String, Optional ByVal mIN04 As String, Optional ByVal mIN05 As String, Optional ByVal mIN06 As String)
Dim strIN01 As String
   
    ' 字串補滿八碼或只取八碼
    If Len(strInventor) > 8 Then
       strIN01 = Mid(strInventor, 1, 8)
    Else
       strIN01 = strInventor & String(8 - Len(strInventor), "0")
    End If
    
     m_InventorList(m_InventorListCount).iN01 = strIN01 '客戶編號(8碼)
     m_InventorList(m_InventorListCount).iN02 = mIN02  '發明人代號
     m_InventorList(m_InventorListCount).iN04 = mIN04  '(發明人)中文名稱
     m_InventorList(m_InventorListCount).IN05 = mIN05  '(發明人)英文名稱
     m_InventorList(m_InventorListCount).IN06 = mIN06  '(發明人)日文名稱
    
     m_InventorListCount = m_InventorListCount + 1
End Sub

'Added by Lydia 2017/11/17 設計案屬性
Private Sub Combo5_Validate(Cancel As Boolean)
   If Combo5 <> "" Then
      Combo5 = Left(Combo5, 1) + "." + PUB_GetCaseAttributeName(Left(Combo5, 1), "3")
      If Combo5 = Left(Combo5, 1) + "." Then
         Combo5 = Left(Combo5, 1)
         Cancel = True
         Combo5.SetFocus
      End If
   End If
End Sub

'Added by Morgan 2018/1/19
'轉全形
Private Function ToWide(pString As String) As String
   pString = Replace(pString, "\", "＼")
   ToWide = StrConv(pString, vbWide)
End Function

'Added by Lydia 2018/06/07
Private Sub cboSource_Validate(Cancel As Boolean)
Dim iR As Integer
   If cboSource.Text <> "" Then
        iR = Val(cboSource.Text)
        'Modified by Lydia 2024/02/21 +4.韓文
        If iR = 0 Or iR > 4 Then
             MsgBox "請輸入1-4的選項！", vbCritical
             Cancel = True
             cboSource.SetFocus
             Exit Sub
        Else
             If cboSource.ListIndex <> iR - 1 Then
                  cboSource.ListIndex = iR - 1
             End If
        End If
        txtTF(27).Text = Left(cboSource.Text, 1)
   Else
        txtTF(27).Text = ""
   End If
End Sub

Private Sub cboTarget_Validate(Cancel As Boolean)
Dim iR As Integer
   If cboTarget.Text <> "" Then
        iR = Val(cboTarget.Text)
        If iR = 0 Or iR > 2 Then
             MsgBox "請輸入1-2的選項！", vbCritical
             Cancel = True
             cboTarget.SetFocus
             Exit Sub
        Else
             If cboTarget.ListIndex <> iR - 1 Then
                  cboTarget.ListIndex = iR - 1
             End If
        End If
        txtTF(28).Text = Left(cboTarget.Text, 1)
   Else
        txtTF(28).Text = ""
   End If
End Sub

Private Sub CmdAddCP_Click()
      Set frm060101_2.fmParent = Me
      frm060101_2.Show
      Me.Hide
End Sub

Private Sub Chk01_Click()
    If Chk01.Value = 1 Then
         txtTF(29).Text = "Y"
    Else
         txtTF(29).Text = ""
    End If
End Sub

Private Sub Chk02_Click()
    If Chk02.Value = 1 Then
         txtTF(30).Text = ""
    End If
End Sub

Private Sub Chk03_Click()
    If Chk03.Value = 1 Then
         txtTF(31).Text = "Y"
    Else
         txtTF(31).Text = ""
    End If
End Sub

Private Sub Chk04_Click()
    If Chk04.Value = 1 Then
         txtTF(33).Text = "Y"
    Else
         txtTF(33).Text = ""
    End If
End Sub

Private Sub txtTF_GotFocus(Index As Integer)
    TextInverse txtTF(Index)
End Sub

Private Sub txtTF_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 20, 30
            KeyAscii = UpperCase(KeyAscii)
      'Added by Lydia 2019/08/23
      'Modified by Lydia 2019/10/25 +翻譯瑕疵備註 txtTF(37)
      Case 36, 37 '翻譯特殊指示
              '排除
      Case Else
            KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub txtTF_Validate(Index As Integer, Cancel As Boolean)
    
    If txtTF(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 19 '相似度
           If txtTF(Index).Text <> "" Then
              If Val(txtTF(Index).Text) > 100 Then
                 MsgBox "相似度不可大於100！"
                 Cancel = True
                 TextInverse txtTF(Index)
              Else
                 txtTF(Index) = Val(txtTF(Index))
              End If
           End If
        Case 20 '相似案號
           If txtTF(Index).Text <> "" Then
              Call ChgCaseNo(txtTF(Index), strExc)
              'Added by Lydia 2019/06/18 在測試程式時,發現沒有檢查重複
              If Trim(strExc(1) & strExc(2)) = Trim(text1 & Text2) Then
                  MsgBox "相似案號不可輸入本案！", vbCritical, "檢核資料"
                 Cancel = True
                 TextInverse txtTF(Index)
                 Exit Sub
              End If
              
              If ClsPDCheckCaseCodeIsExist(strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
                 Cancel = True
                 TextInverse txtTF(Index)
              Else
                 txtTF(Index).Text = strExc(1) & strExc(2) & strExc(3) & strExc(4)
              End If
           End If
        Case 26, 32 '交稿期限,只交Claim期限
            If txtTF(Index) <> "" Then
                If CheckIsTaiwanDate(txtTF(Index).Text) = False Then
                      txtTF_GotFocus Index
                      TextInverse txtTF(Index)
                      Cancel = True
                ElseIf Not ChkWorkDay(DBDATE(txtTF(Index).Text)) Then
                      MsgBox "交稿期限必須是工作天 !"
                      txtTF_GotFocus Index
                      TextInverse txtTF(Index)
                      Cancel = True
                End If
            End If
        Case 30 '英文本收文號
            If txtTF(Index) <> "" Then
                If PUB_ChkCPExist(pa, "202", , txtTF(Index).Text, , "B") = False Then
                      MsgBox "請輸入B類補文件的收文號！", vbCritical
                      txtTF_GotFocus Index
                      TextInverse txtTF(Index)
                      Cancel = True
                Else
                      Chk02.Value = 0
                End If
            End If
        'Added by Lydia 2019/08/23
        Case 36 '翻譯特殊指示
            If txtTF(Index).Text = "" Then Exit Sub
            If Not CheckLengthIsOK(txtTF(Index), 200) Then
               Cancel = True
            End If
        'Added by Lydia 2019/10/25
        Case 37 '翻譯瑕疵備註
            If txtTF(Index).Text = "" Then Exit Sub
            If Not CheckLengthIsOK(txtTF(Index), 100) Then
               Cancel = True
            End If
    End Select
End Sub
'end 2018/06/07

'Added by Lydia 2018/06/07 開啟資料夾(方便看ORI.PDF, 輸入頁數)
Private Sub cmdOpen_Click(Index As Integer)
Dim hLocalFile As Long

On Error GoTo ErrHand01
    
    'Added by Lydia 2020/02/21 檢查案號是否正確
    If CheckFindPass = False Then
        Exit Sub
    End If
    'end 2020/02/21
   
    'Added by Lydia 2020/01/20 開啟[原始檔區]
    If Index = 0 And InStr(cmdOpen(Index).Caption, "原始檔") > 0 Then
        If mdiMain.mnuTitle(10).Enabled = True Then
            If cmdOpen(Index).Tag = "" Then
                MsgBox pa(1) & "-" & pa(2) & "在〔原始檔區〕的English_Vers收文號不存在!", vbInformation
            Else
                frm100101_M.m_strKey = cmdOpen(Index).Tag '多筆總收文號
                frm100101_M.SetParent Me
                If frm100101_M.QueryData = True Then
                   frm100101_M.Show
                   Me.Hide
                End If
            End If
        Else
            MsgBox "請先關閉共同查詢畫面！"
        End If
    Else
    'end 2020/01/20
        strExc(1) = ""
        If Index = 0 Then '外文本=English_vers
            'Remove by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
            'strExc(1) = Pub_GetFCPcaseFilePath(pa(2), , pa(1))
        ElseIf Index = 1 Then
            'Modified by Lydia 2024/07/22 改用變數
            'strExc(1) = "\\Typing2\電子送件暫存區\" & pa(1) & pa(2)
            strExc(1) = "\\" & strTyping2Path & "\電子送件暫存區\" & pa(1) & pa(2)
        ElseIf Index = 2 Then
            strExc(1) = strResPath & "\."
        End If
        If strExc(1) = "" Then Exit Sub
    
        If Dir(strExc(1), vbDirectory) <> "" Then
             ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
        Else
             MsgBox strExc(1) & " 資料夾不存在 ！", vbInformation
        End If
    End If 'Added by Lydia 2020/01/20
    
    Exit Sub
    
ErrHand01:
    If Err.Number <> 0 Then
         MsgBox "無法讀取" & strExc(1) & "，請通知電腦中心！", vbCritical
         Resume Next
    End If
End Sub

'Added by Lydia 2018/08/24
Private Sub Chk05_Click()
    '暫不翻譯
    If Chk05.Value = 1 Then
         txtTF(34).Text = "Y"
    Else
         txtTF(34).Text = ""
    End If
End Sub

Private Sub Chk06_Click()
    '固定報價
    If Chk06.Value = 1 Then
         txtPA(62).Text = "Y"
    Else
         txtPA(62).Text = ""
    End If
End Sub

'Added by Lydia 2020/02/21
Private Function CheckFindPass() As Boolean
   
   CheckFindPass = False
   If text1 & Text2 & Text3 & Text4 <> pa(1) & pa(2) & pa(3) & pa(4) Then
       If pa(1) & pa(2) & pa(3) & pa(4) = "" Then
           MsgBox "請先查詢正確的案號!", vbExclamation, "檢核資料"
       Else
           MsgBox "本所案號與資料的案號" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "不一致，請先查詢正確的案號!", vbExclamation, "檢核資料"
       End If
       Exit Function
   End If
   
   CheckFindPass = True
End Function

'Added by Lydia 2020/02/21
Private Sub CmdPA174_Click()
   Call ProcPA174toFile("N")
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟/維護FCP0xxxxx.新案性質.案件名稱.doc
Private Sub ProcPA174toFile(ByVal pKind As String)
Dim strKind As String

   If CheckFindPass = False Then
        Exit Sub
   Else
        If ChkPA174.Value = vbUnchecked Then
            MsgBox "請先勾選「有特殊字」!", vbInformation + vbOKOnly, Me.Caption
        Else
            If pKind = "Y" Then 'bolAskPA174
                strKind = "3"
            Else
                strKind = "1"
            End If
            If Pub_GetPA174toFile(strKind, Me.text1, Me.Text2, Me.Text3, Me.Text4, Me, frm100101_M_1) = True Then
            End If
        End If
   End If
End Sub

'Added by Lydia 2020/02/21
Public Sub PubShowNextData()
   '原始檔Word檔維護，上傳後直接進入存檔
   If bolAskPA174 = True Then
        Call cmdok_Click(0) '確定->存檔
   End If
End Sub

'Added by Lydia 2020/09/11
Private Sub CmdINST_Click()

     If text1 = "" Or Text2 = "" Then Exit Sub
     
     If PUB_CheckFormExist("frm12040159") Then
         MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
         Exit Sub
     End If
     
     If CheckFindPass = True Then
        frm12040159.SetParent "E", pa(1) & pa(2) & pa(3) & pa(4), Me
        frm12040159.Show
     End If
End Sub

'Added by Lydia 2022/03/25 控制發明人欄位是否可點選
Private Sub InvFieldEnabled(ByVal bEnabled As Boolean)
'點選(帶出)已建檔的發明人中/英/日文名稱欄位會卡住操作=>停留在txtInvField_Validate無法跳離
'根據Form 1.0版本的程式不會有此問題; 因為已建檔的發明人不可再變更,所以用Frame包住名稱欄位控制是否可點選
    Frame3.Enabled = bEnabled
End Sub

'Added by Lydia 2023/07/31 命名作業查詢
Private Sub cmdQueryTCT_Click()
Dim intJ As Integer, strB1 As String, strB2 As String, strB3 As String
Dim rsQD As New ADODB.Recordset

   If text1.Text = "" Or Len(Text2.Text) <> 6 Then
      MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
      text1.SetFocus
      Exit Sub
   End If
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
   If pa(1) & pa(2) & pa(3) & pa(4) <> text1 & Text2 & Text3 & Text4 Then
      MsgBox "請先尋找本所案號的資料！", vbCritical
      Command1.SetFocus
      Exit Sub
   End If
   strB1 = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as caseno " & _
           ",tct01,tct04||s1.st02 as tct04, substr(sqldatet(tct05),1,9) tct05, substr(sqltime(tct06||'00'),1,5) as tct06 " & _
           ",tct07||s2.st02 as tct07, substr(sqldatet(tct08),1,9) tct08, substr(sqltime(tct09||'00'),1,5) as tct09 " & _
           ",tct10||s3.st02 as tct10, substr(sqldatet(tct11),1,9) tct11, substr(sqltime(tct12||'00'),1,5) as tct12 ,trackingcasename.* " & _
           "From caseprogress, patent, transcasetitle, trackingcasename ,staff s1, staff s2, staff s3 " & _
           "where cp01='" & text1 & "' and cp02='" & Text2 & "' and cp03='" & Text3 & "' and cp04='" & Text4 & "' and cp31='Y' and cp09=tct01(+) " & _
           "and cp09=tcn05(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
           "and tct04=s1.st01(+) and tct07=s2.st01(+) and tct10=s3.st01(+)"
   intJ = 1
   Set rsQD = ClsLawReadRstMsg(intJ, strB1)
   If intJ = 0 Then
      strB2 = ""
      MsgBox text1 & "-" & Text2 & IIf(Text3 & Text4 <> "000", "-" & Text3 & "-" & Text4, "") & "本案無新案命名作業！", vbOKOnly + vbInformation
   Else
      strB2 = "本所案號：" & rsQD.Fields("caseno") & String(4, " ") & "追蹤號：" & rsQD.Fields("tcn01")
      If "" & rsQD.Fields("tcn13") <> "" Then
         strB2 = strB2 & vbCrLf & Mid(PUB_GetTCNmTitle(text1, Text2, Text3, Text4, "", rsQD.Fields("tcn13"), ""), 1, InStr(PUB_GetTCNmTitle(text1, Text2, Text3, Text4, "", rsQD.Fields("tcn13"), ""), "〕"))
      End If
      strB2 = strB2 & vbCrLf
      'Added by Lydia 2023/10/12 增加說明; ex.FCP-070257應該要設定有英文本,確沒有設定
      If "" & rsQD.Fields("TCN12") = "Y" Then
        strB2 = strB2 & vbCrLf & "客戶有提供彩圖：" & rsQD.Fields("tcn12")
      End If
      If "" & rsQD.Fields("TCN13") <> "" Then
         strB2 = strB2 & vbCrLf & "外文本的對應英/中說："
         Select Case "" & rsQD.Fields("tcn13")
            Case "0": strB2 = strB2 & "無"
            Case "1": strB2 = strB2 & "有"
            Case "2": strB2 = strB2 & "待確定"
            'Modified by Lydia 2024/11/22 +TCN26英說參考本收件日
            Case "3": strB2 = strB2 & "已收參考本" & IIf(Val("" & rsQD.Fields("tcn26")) > 0, vbCrLf & "參考本收件日：" & ChangeWStringToTDateString("" & rsQD.Fields("tcn26")) & vbCrLf, "")
            Case "4": strB2 = strB2 & "確定無參考本"
         End Select
      End If
      If "" & rsQD.Fields("TCN17") <> "" Then
         strB2 = strB2 & vbCrLf & "相似舊案案號：" & rsQD.Fields("tcn17")
         strB2 = strB2 & String(4, "　") & "指定組別：" & PUB_GetFCPGrpName("" & rsQD.Fields("tcn18"), , False)
      End If
      If "" & rsQD.Fields("TCN12") & rsQD.Fields("TCN13") & rsQD.Fields("TCN17") <> "" Then
         strB2 = strB2 & vbCrLf
      End If
      'end 2023/10/12
      
      If "" & rsQD.Fields("tcn16") = "Y" Then
         strB2 = strB2 & vbCrLf & "新案暫不認領：" & rsQD.Fields("tcn16")
      'Added by Lydia 2023/12/05
      ElseIf "" & rsQD.Fields("tct01") = "" Then
         strB2 = strB2 & vbCrLf & "本案無新案命名作業！"
      'end 2023/12/05
      Else
         If "" & rsQD.Fields("tct04") = "" Then
            If "" & rsQD.Fields("tcn21") = "99999999" Then
               strB2 = strB2 & vbCrLf & "認領狀態：暫不認領"
            Else
               strSql = ""
               Select Case "" & rsQD.Fields("tcn23")
                  'Modified by Lydia 2024/02/16 增加TFA09的歸類strB3
                  Case "0"
                     strSql = "急件認領"
                     strB3 = "0"
                  Case "1"
                     strSql = "主管認領"
                     strB3 = "1"
                  Case "2"
                     strSql = "主管+職代認領"
                     strB3 = "1"
                  Case "3":
                     strSql = "協調認領認領"
                     strB3 = "2"
                  Case "4": strSql = "非英說協調認領"
                     strB3 = "4"
               End Select
               strB2 = strB2 & vbCrLf & "認領狀態：" & strSql & String(4, " ") & "(期限:" & ChangeWStringToTDateString("" & rsQD.Fields("tcn21")) & "  " & Format("" & rsQD.Fields("tcn22"), "00:00") & ")"
               
               'Added by Lydia 2024/01/10 增加顯示主管認領狀態
               'Modified by Lydia 2024/02/16 增加TFA09的歸類; rsQD.Fields("tcn23")=>strB3
               strB1 = "select st16, cst16(st16) grpname,st02||'，'||decode(tfa05,'Y','Y認領','N','N不認領')||'，時間:'||sqldatet(tfa02)||' '||sqltime6(tfa03) stype " & _
                       "from transfeeassign,staff where tfa01='" & rsQD.Fields("tct01") & "' and tfa09='" & strB3 & "' and tfa04=st01(+) order by st16 "
               intJ = 1
               Set rsQD = ClsLawReadRstMsg(intJ, strB1)
               If intJ = 1 Then
                  strSql = ""
                  rsQD.MoveFirst
                  Do While Not rsQD.EOF
                     strSql = strSql & vbCrLf & PUB_StrToStr(rsQD.Fields("grpname"), 12, True) & "：" & rsQD.Fields("stype")
                     rsQD.MoveNext
                  Loop
                  strB2 = strB2 & strSql
               End If
               'end 2024/01/10
            End If
         Else
            strB2 = strB2 & vbCrLf & "命名工程師主管：" & rsQD.Fields("tct04") & IIf("" & rsQD.Fields("tct05") <> "", " (已確認)", "")
            If "" & rsQD.Fields("tct07") <> "" Then
               strB2 = strB2 & vbCrLf & "命名工程師主任：" & rsQD.Fields("tct07") & IIf("" & rsQD.Fields("tct08") <> "", " (已確認)", "")
            End If
            If "" & rsQD.Fields("tct10") <> "" Then
               strB2 = strB2 & vbCrLf & "命名工程師：" & rsQD.Fields("tct10") & IIf("" & rsQD.Fields("tct11") <> "", " (已確認)", "")
            End If
         End If
      End If
      MsgBox strB2, vbOKOnly + vbInformation, "命名作業查詢時間:" & ChangeTStringToTDateString(strSrvDate(2)) & "  " & Format(ServerTime, "00:00:00")
   End If
   Set rsQD = Nothing
End Sub

