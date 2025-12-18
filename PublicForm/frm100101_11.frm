VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_11 
   BackColor       =   &H80000004&
   BorderStyle     =   1  '單線固定
   Caption         =   "申請人資料查詢"
   ClientHeight    =   6384
   ClientLeft      =   1440
   ClientTop       =   2316
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6384
   ScaleWidth      =   9324
   Begin VB.CommandButton CmdOk1 
      Caption         =   "被介紹者"
      Height          =   360
      Index           =   6
      Left            =   4290
      Style           =   1  '圖片外觀
      TabIndex        =   303
      Top             =   45
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "各項指示"
      Height          =   360
      Index           =   4
      Left            =   3390
      TabIndex        =   296
      Top             =   45
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "合約資料"
      Height          =   360
      Index           =   5
      Left            =   5175
      TabIndex        =   290
      Top             =   45
      Width           =   900
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "平台帳號"
      Height          =   360
      Index           =   3
      Left            =   6075
      TabIndex        =   258
      Top             =   45
      Width           =   900
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "減免身份"
      Height          =   360
      Index           =   2
      Left            =   6960
      TabIndex        =   168
      Top             =   45
      Width           =   900
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "回前畫面"
      Height          =   360
      Index           =   0
      Left            =   7845
      TabIndex        =   37
      Top             =   45
      Width           =   900
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束"
      Height          =   360
      Index           =   1
      Left            =   8760
      TabIndex        =   38
      Top             =   45
      Width           =   550
   End
   Begin TabDlg.SSTab tabCustomer 
      Height          =   5856
      Left            =   60
      TabIndex        =   0
      Top             =   504
      Width           =   9228
      _ExtentX        =   16277
      _ExtentY        =   10329
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   10
      TabHeight       =   420
      TabCaption(0)   =   "基本"
      TabPicture(0)   =   "frm100101_11.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(18)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(17)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(16)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(15)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(14)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(13)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(12)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(11)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(10)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(7)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "LABEL"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl1(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl1(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lbl1(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl1(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl1(4)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl1(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lbl1(6)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lbl1(7)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl1(8)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl1(9)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl1(10)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl1(11)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lbl1(12)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lbl1(13)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lbl1(14)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbl1(15)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label2"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label50"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "lbl1(52)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label48"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label1(20)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label1(22)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "lbl1(59)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Label1(23)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "lbl1(60)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label1(24)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "lbl1(66)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Label1(29)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "lbl1(68)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Label1(26)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "lbl1(69)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Label1(30)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "lbl1(73)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Label1(31)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Cu180"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "lstDeveloper"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txt1(0)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txt1(1)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txt1(44)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txt1(45)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txt1(2)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "LblCU144"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Label83"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Label84"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "lbl1(41)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "lbl1(42)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "lbl1(48)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "lbl1(49)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "lbl1(50)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "lbl1(51)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "FrameID"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).ControlCount=   71
      TabCaption(1)   =   "通訊"
      TabPicture(1)   =   "frm100101_11.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(34)"
      Tab(1).Control(1)=   "lbl1(74)"
      Tab(1).Control(2)=   "txt1(49)"
      Tab(1).Control(3)=   "txt1(48)"
      Tab(1).Control(4)=   "txt1(47)"
      Tab(1).Control(5)=   "txt1(46)"
      Tab(1).Control(6)=   "txt1(4)"
      Tab(1).Control(7)=   "txt1(5)"
      Tab(1).Control(8)=   "txt1(6)"
      Tab(1).Control(9)=   "txt1(7)"
      Tab(1).Control(10)=   "txt1(8)"
      Tab(1).Control(11)=   "txt1(9)"
      Tab(1).Control(12)=   "txt1(10)"
      Tab(1).Control(13)=   "txt1(11)"
      Tab(1).Control(14)=   "txt1(12)"
      Tab(1).Control(15)=   "txt1(3)"
      Tab(1).Control(16)=   "Label63(11)"
      Tab(1).Control(17)=   "Label63(17)"
      Tab(1).Control(18)=   "Label63(18)"
      Tab(1).Control(19)=   "Label63(19)"
      Tab(1).Control(20)=   "Label63(20)"
      Tab(1).Control(21)=   "Label63(8)"
      Tab(1).Control(22)=   "Label63(7)"
      Tab(1).Control(23)=   "Label63(6)"
      Tab(1).Control(24)=   "Label63(5)"
      Tab(1).Control(25)=   "Label63(4)"
      Tab(1).Control(26)=   "Label63(3)"
      Tab(1).Control(27)=   "Label63(2)"
      Tab(1).Control(28)=   "Label63(1)"
      Tab(1).Control(29)=   "Label63(0)"
      Tab(1).Control(30)=   "lbl1(21)"
      Tab(1).Control(31)=   "lbl1(16)"
      Tab(1).Control(32)=   "lbl1(20)"
      Tab(1).Control(33)=   "lbl1(19)"
      Tab(1).Control(34)=   "lbl1(18)"
      Tab(1).Control(35)=   "lbl1(17)"
      Tab(1).Control(36)=   "Label63(9)"
      Tab(1).Control(37)=   "Label63(10)"
      Tab(1).Control(38)=   "Label63(12)"
      Tab(1).Control(39)=   "Label63(13)"
      Tab(1).Control(40)=   "Label63(14)"
      Tab(1).Control(41)=   "Label63(15)"
      Tab(1).ControlCount=   42
      TabCaption(2)   =   "地址"
      TabPicture(2)   =   "frm100101_11.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt1(57)"
      Tab(2).Control(1)=   "Label41(39)"
      Tab(2).Control(2)=   "txt1(21)"
      Tab(2).Control(3)=   "txt1(20)"
      Tab(2).Control(4)=   "txt1(19)"
      Tab(2).Control(5)=   "txt1(18)"
      Tab(2).Control(6)=   "txt1(17)"
      Tab(2).Control(7)=   "txt1(16)"
      Tab(2).Control(8)=   "txt1(15)"
      Tab(2).Control(9)=   "txt1(14)"
      Tab(2).Control(10)=   "txt1(13)"
      Tab(2).Control(11)=   "Label41(31)"
      Tab(2).Control(12)=   "lbl1(53)"
      Tab(2).Control(13)=   "lbl1(23)"
      Tab(2).Control(14)=   "lbl1(22)"
      Tab(2).Control(15)=   "Label41(18)"
      Tab(2).Control(16)=   "Label41(19)"
      Tab(2).Control(17)=   "Label41(20)"
      Tab(2).Control(18)=   "Label41(21)"
      Tab(2).Control(19)=   "Label41(22)"
      Tab(2).Control(20)=   "Label41(23)"
      Tab(2).Control(21)=   "Label41(24)"
      Tab(2).Control(22)=   "Label41(25)"
      Tab(2).Control(23)=   "Label41(26)"
      Tab(2).Control(24)=   "Label41(27)"
      Tab(2).Control(25)=   "Label41(28)"
      Tab(2).ControlCount=   26
      TabCaption(3)   =   "代表人"
      TabPicture(3)   =   "frm100101_11.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text1(125)"
      Tab(3).Control(1)=   "Text1(102)"
      Tab(3).Control(2)=   "Text1(103)"
      Tab(3).Control(3)=   "txt1(39)"
      Tab(3).Control(4)=   "txt1(38)"
      Tab(3).Control(5)=   "txt1(37)"
      Tab(3).Control(6)=   "txt1(36)"
      Tab(3).Control(7)=   "txt1(35)"
      Tab(3).Control(8)=   "txt1(34)"
      Tab(3).Control(9)=   "txt1(33)"
      Tab(3).Control(10)=   "txt1(32)"
      Tab(3).Control(11)=   "txt1(31)"
      Tab(3).Control(12)=   "txt1(30)"
      Tab(3).Control(13)=   "txt1(29)"
      Tab(3).Control(14)=   "txt1(28)"
      Tab(3).Control(15)=   "txt1(27)"
      Tab(3).Control(16)=   "txt1(26)"
      Tab(3).Control(17)=   "txt1(25)"
      Tab(3).Control(18)=   "txt1(24)"
      Tab(3).Control(19)=   "txt1(23)"
      Tab(3).Control(20)=   "txt1(22)"
      Tab(3).Control(21)=   "Label41(32)"
      Tab(3).Control(22)=   "Label41(29)"
      Tab(3).Control(23)=   "Label41(30)"
      Tab(3).Control(24)=   "Label41(0)"
      Tab(3).Control(25)=   "Label41(1)"
      Tab(3).Control(26)=   "Label41(2)"
      Tab(3).Control(27)=   "Label41(3)"
      Tab(3).Control(28)=   "Label41(4)"
      Tab(3).Control(29)=   "Label41(5)"
      Tab(3).Control(30)=   "Label41(6)"
      Tab(3).Control(31)=   "Label41(7)"
      Tab(3).Control(32)=   "Label41(8)"
      Tab(3).Control(33)=   "Label41(9)"
      Tab(3).Control(34)=   "Label41(10)"
      Tab(3).Control(35)=   "Label41(11)"
      Tab(3).Control(36)=   "Label41(12)"
      Tab(3).Control(37)=   "Label41(13)"
      Tab(3).Control(38)=   "Label41(14)"
      Tab(3).Control(39)=   "Label41(15)"
      Tab(3).Control(40)=   "Label41(16)"
      Tab(3).Control(41)=   "Label41(17)"
      Tab(3).ControlCount=   42
      TabCaption(4)   =   "專利"
      TabPicture(4)   =   "frm100101_11.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Combo4"
      Tab(4).Control(1)=   "Combo3(0)"
      Tab(4).Control(2)=   "lblCU(202)"
      Tab(4).Control(3)=   "lblCU(189)"
      Tab(4).Control(4)=   "txt1(40)"
      Tab(4).Control(5)=   "lblCU(177)"
      Tab(4).Control(6)=   "lbl1(72)"
      Tab(4).Control(7)=   "Label1(19)"
      Tab(4).Control(8)=   "Label6"
      Tab(4).Control(9)=   "Label80(30)"
      Tab(4).Control(10)=   "Label80(22)"
      Tab(4).Control(11)=   "lblCU(113)"
      Tab(4).Control(12)=   "Label1(27)"
      Tab(4).Control(13)=   "lbl1(24)"
      Tab(4).Control(14)=   "lbl1(37)"
      Tab(4).Control(15)=   "lbl1(36)"
      Tab(4).Control(16)=   "lblCU(131)"
      Tab(4).Control(17)=   "lblCU(130)"
      Tab(4).Control(18)=   "lbl1(56)"
      Tab(4).Control(19)=   "lblCU(133)"
      Tab(4).Control(20)=   "lblCU(135)"
      Tab(4).Control(21)=   "lblCU(137)"
      Tab(4).Control(22)=   "lbl1(44)"
      Tab(4).Control(23)=   "lbl1(43)"
      Tab(4).Control(24)=   "lbl1(33)"
      Tab(4).Control(25)=   "lbl1(32)"
      Tab(4).Control(26)=   "lbl1(29)"
      Tab(4).Control(27)=   "lbl1(55)"
      Tab(4).Control(28)=   "lbl1(28)"
      Tab(4).Control(29)=   "lbl1(54)"
      Tab(4).Control(30)=   "Label80(19)"
      Tab(4).Control(31)=   "Label69"
      Tab(4).Control(32)=   "Label38"
      Tab(4).Control(33)=   "Label70"
      Tab(4).Control(34)=   "Label60(0)"
      Tab(4).Control(35)=   "Label60(1)"
      Tab(4).Control(36)=   "Label67(0)"
      Tab(4).Control(37)=   "Label67(1)"
      Tab(4).Control(38)=   "Label67(4)"
      Tab(4).Control(39)=   "Label1(21)"
      Tab(4).Control(40)=   "lbl1(47)"
      Tab(4).Control(41)=   "Label80(18)"
      Tab(4).Control(42)=   "Label80(17)"
      Tab(4).Control(43)=   "lbl1(39)"
      Tab(4).Control(44)=   "lbl1(27)"
      Tab(4).Control(45)=   "lbl1(26)"
      Tab(4).Control(46)=   "lbl1(25)"
      Tab(4).Control(47)=   "Label80(3)"
      Tab(4).Control(48)=   "Label80(4)"
      Tab(4).Control(49)=   "Label80(7)"
      Tab(4).Control(50)=   "Label80(8)"
      Tab(4).Control(51)=   "Label80(10)"
      Tab(4).Control(52)=   "Label80(11)"
      Tab(4).Control(53)=   "Label80(12)"
      Tab(4).Control(54)=   "Label80(13)"
      Tab(4).Control(55)=   "Label80(15)"
      Tab(4).Control(56)=   "Label80(14)"
      Tab(4).Control(57)=   "Label52"
      Tab(4).Control(58)=   "Label80(33)"
      Tab(4).Control(59)=   "Label1(32)"
      Tab(4).Control(60)=   "Label1(35)"
      Tab(4).ControlCount=   61
      TabCaption(5)   =   "商標"
      TabPicture(5)   =   "frm100101_11.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Combo3(1)"
      Tab(5).Control(1)=   "lbl1(77)"
      Tab(5).Control(2)=   "Label11"
      Tab(5).Control(3)=   "lbl1(76)"
      Tab(5).Control(4)=   "Label9"
      Tab(5).Control(5)=   "lbl1(75)"
      Tab(5).Control(6)=   "lblCU(190)"
      Tab(5).Control(7)=   "txt1(51)"
      Tab(5).Control(8)=   "txt1(50)"
      Tab(5).Control(9)=   "Label7"
      Tab(5).Control(10)=   "lbl1(67)"
      Tab(5).Control(11)=   "Label1(28)"
      Tab(5).Control(12)=   "Label80(29)"
      Tab(5).Control(13)=   "lbl1(65)"
      Tab(5).Control(14)=   "lbl1(64)"
      Tab(5).Control(15)=   "lbl1(63)"
      Tab(5).Control(16)=   "lbl1(62)"
      Tab(5).Control(17)=   "lbl1(61)"
      Tab(5).Control(18)=   "Label80(23)"
      Tab(5).Control(19)=   "Label80(24)"
      Tab(5).Control(20)=   "Label80(25)"
      Tab(5).Control(21)=   "Label80(26)"
      Tab(5).Control(22)=   "Label80(27)"
      Tab(5).Control(23)=   "Label80(28)"
      Tab(5).Control(24)=   "Label1(25)"
      Tab(5).Control(25)=   "lbl1(58)"
      Tab(5).Control(26)=   "lbl1(34)"
      Tab(5).Control(27)=   "lbl1(35)"
      Tab(5).Control(28)=   "lbl1(38)"
      Tab(5).Control(29)=   "lblCU(138)"
      Tab(5).Control(30)=   "lblCU(136)"
      Tab(5).Control(31)=   "lblCU(134)"
      Tab(5).Control(32)=   "lbl1(46)"
      Tab(5).Control(33)=   "lbl1(57)"
      Tab(5).Control(34)=   "lbl1(45)"
      Tab(5).Control(35)=   "lbl1(40)"
      Tab(5).Control(36)=   "Label80(6)"
      Tab(5).Control(37)=   "Label80(5)"
      Tab(5).Control(38)=   "Label80(21)"
      Tab(5).Control(39)=   "Label80(16)"
      Tab(5).Control(40)=   "Label67(5)"
      Tab(5).Control(41)=   "Label67(3)"
      Tab(5).Control(42)=   "Label67(2)"
      Tab(5).Control(43)=   "Label3"
      Tab(5).Control(44)=   "Label5"
      Tab(5).Control(45)=   "Label4"
      Tab(5).Control(46)=   "Label80(20)"
      Tab(5).Control(47)=   "Label1(33)"
      Tab(5).Control(48)=   "Label10"
      Tab(5).ControlCount=   49
      TabCaption(6)   =   "其他"
      TabPicture(6)   =   "frm100101_11.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1"
      Tab(6).Control(1)=   "Frame1K"
      Tab(6).Control(2)=   "lblCU(140)"
      Tab(6).Control(3)=   "Label8"
      Tab(6).Control(4)=   "txt1(41)"
      Tab(6).Control(5)=   "txt1(42)"
      Tab(6).Control(6)=   "lbl1(70)"
      Tab(6).Control(7)=   "lbl1(71)"
      Tab(6).Control(8)=   "Label80(32)"
      Tab(6).Control(9)=   "Label80(31)"
      Tab(6).Control(10)=   "lblCU16X(5)"
      Tab(6).Control(11)=   "lblCU16X(4)"
      Tab(6).Control(12)=   "lblCU16X(3)"
      Tab(6).Control(13)=   "lblCU16X(2)"
      Tab(6).Control(14)=   "lblCU16X(1)"
      Tab(6).Control(15)=   "lblCU16X(0)"
      Tab(6).Control(16)=   "lblComp(0)"
      Tab(6).Control(17)=   "lblComp(1)"
      Tab(6).Control(18)=   "lblComp(2)"
      Tab(6).Control(19)=   "lblComp(3)"
      Tab(6).Control(20)=   "lblComp(4)"
      Tab(6).Control(21)=   "lblComp(5)"
      Tab(6).Control(22)=   "lblCU(141)"
      Tab(6).Control(23)=   "Label67(9)"
      Tab(6).Control(24)=   "Label80(9)"
      Tab(6).Control(25)=   "Label80(2)"
      Tab(6).Control(26)=   "Label80(1)"
      Tab(6).Control(27)=   "Label80(0)"
      Tab(6).Control(28)=   "lbl1(30)"
      Tab(6).Control(29)=   "lbl1(31)"
      Tab(6).ControlCount=   30
      TabCaption(7)   =   "參考備註"
      TabPicture(7)   =   "frm100101_11.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "txt1(43)"
      Tab(7).ControlCount=   1
      Begin VB.Frame FrameID 
         BorderStyle     =   0  '沒有框線
         Enabled         =   0   'False
         Height          =   280
         Left            =   3840
         TabIndex        =   346
         Top             =   3020
         Width           =   1050
         Begin VB.CheckBox ChkID 
            Caption         =   "不提供ID"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   0
            TabIndex        =   347
            Top             =   0
            Width           =   1000
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "全E化客戶"
         Height          =   1668
         Left            =   -74952
         TabIndex        =   327
         Top             =   4152
         Width           =   9108
         Begin VB.CheckBox ChkCU186 
            Caption         =   "不寄官方收據"
            Height          =   180
            Index           =   1
            Left            =   1512
            TabIndex        =   339
            Top             =   1392
            Width           =   1536
         End
         Begin VB.CheckBox ChkCU186 
            Caption         =   "勾選讀取回條"
            Height          =   180
            Index           =   2
            Left            =   3192
            TabIndex        =   338
            Top             =   1392
            Width           =   1632
         End
         Begin MSForms.TextBox txt1 
            Height          =   288
            Index           =   56
            Left            =   8352
            TabIndex        =   337
            Top             =   1344
            Visible         =   0   'False
            Width           =   672
            VariousPropertyBits=   -1467989985
            BackColor       =   16777215
            ScrollBars      =   2
            Size            =   "1185;508"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblECustXSet 
            AutoSize        =   -1  'True
            Caption         =   "特殊設定："
            Height          =   180
            Left            =   588
            TabIndex        =   336
            Top             =   1368
            Width           =   900
         End
         Begin MSForms.TextBox txt1 
            Height          =   312
            Index           =   55
            Left            =   1488
            TabIndex        =   335
            Top             =   1032
            Width           =   7548
            VariousPropertyBits=   -1467989985
            BackColor       =   16777215
            ScrollBars      =   2
            Size            =   "13314;550"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label63 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "(副本)："
            Height          =   180
            Index           =   23
            Left            =   828
            TabIndex        =   334
            Top             =   1104
            Width           =   660
         End
         Begin MSForms.TextBox txt1 
            Height          =   312
            Index           =   54
            Left            =   1488
            TabIndex        =   333
            Top             =   744
            Width           =   7548
            VariousPropertyBits=   -1467989985
            BackColor       =   16777215
            ScrollBars      =   2
            Size            =   "13314;550"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label63 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "商標信箱(正本)："
            Height          =   180
            Index           =   22
            Left            =   108
            TabIndex        =   332
            Top             =   816
            Width           =   1380
         End
         Begin MSForms.TextBox txt1 
            Height          =   312
            Index           =   53
            Left            =   1488
            TabIndex        =   331
            Top             =   456
            Width           =   7548
            VariousPropertyBits=   -1467989985
            BackColor       =   16777215
            ScrollBars      =   2
            Size            =   "13314;550"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label63 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "(副本)："
            Height          =   180
            Index           =   21
            Left            =   828
            TabIndex        =   330
            Top             =   528
            Width           =   660
         End
         Begin MSForms.TextBox txt1 
            Height          =   312
            Index           =   52
            Left            =   1488
            TabIndex        =   329
            Top             =   144
            Width           =   7548
            VariousPropertyBits=   -1467989985
            BackColor       =   16777215
            ScrollBars      =   2
            Size            =   "13314;550"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label63 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "指定信箱(正本)："
            Height          =   180
            Index           =   16
            Left            =   108
            TabIndex        =   328
            Top             =   216
            Width           =   1380
         End
      End
      Begin VB.Frame Frame1K 
         Enabled         =   0   'False
         Height          =   280
         Left            =   -71100
         TabIndex        =   312
         Top             =   1620
         Width           =   4930
         Begin VB.CheckBox Chk1K 
            Caption         =   "帳單另寄"
            Height          =   180
            Index           =   0
            Left            =   1740
            TabIndex        =   315
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "上傳平台"
            Height          =   180
            Index           =   1
            Left            =   2790
            TabIndex        =   314
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "月帳單"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   313
            Top             =   60
            Width           =   1030
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "請款單寄送類型："
            Height          =   180
            Index           =   34
            Left            =   150
            TabIndex        =   316
            Top             =   60
            Width           =   1440
         End
      End
      Begin VB.ComboBox Combo4 
         Height          =   276
         ItemData        =   "frm100101_11.frx":00E0
         Left            =   -69120
         List            =   "frm100101_11.frx":00F3
         TabIndex        =   304
         Top             =   870
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         Index           =   0
         ItemData        =   "frm100101_11.frx":012A
         Left            =   -70860
         List            =   "frm100101_11.frx":013D
         Style           =   2  '單純下拉式
         TabIndex        =   262
         Top             =   1140
         Width           =   1470
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         Index           =   1
         ItemData        =   "frm100101_11.frx":0171
         Left            =   -70710
         List            =   "frm100101_11.frx":0184
         Style           =   2  '單純下拉式
         TabIndex        =   260
         Top             =   2362
         Width           =   1470
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   77
         Left            =   -66900
         TabIndex        =   344
         Top             =   1245
         Width           =   410
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "723;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "延展折扣：           ％"
         Height          =   180
         Left            =   -67830
         TabIndex        =   345
         Top             =   1245
         Width           =   1630
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   76
         Left            =   -68670
         TabIndex        =   342
         Top             =   1245
         Width           =   410
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "723;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "商標全部折扣終止日："
         Height          =   180
         Left            =   -71250
         TabIndex        =   341
         Top             =   1530
         Width           =   1800
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   75
         Left            =   -69430
         TabIndex        =   340
         Top             =   1530
         Width           =   1470
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2593;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   51
         Left            =   7800
         TabIndex        =   326
         Top             =   5544
         Width           =   800
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1411;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   50
         Left            =   6888
         TabIndex        =   325
         Top             =   5544
         Width           =   800
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1411;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   49
         Left            =   3768
         TabIndex        =   324
         Top             =   5544
         Width           =   804
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1418;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   48
         Left            =   2856
         TabIndex        =   323
         Top             =   5544
         Width           =   804
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1411;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   42
         Left            =   5568
         TabIndex        =   322
         Top             =   5544
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   41
         Left            =   1536
         TabIndex        =   321
         Top             =   5544
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "UPDATE："
         Height          =   180
         Left            =   4704
         TabIndex        =   320
         Top             =   5556
         Width           =   852
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "CREATE："
         Height          =   180
         Left            =   672
         TabIndex        =   319
         Top             =   5556
         Width           =   840
      End
      Begin MSForms.Label lblCU 
         Height          =   252
         Index           =   202
         Left            =   -66972
         TabIndex        =   317
         Top             =   4080
         Width           =   408
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "720;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblCU 
         Height          =   230
         Index           =   140
         Left            =   -72780
         TabIndex        =   310
         Top             =   3900
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;406"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "財務處是否寄發催款單：              (1. 每月寄對帳單　2. 客戶要求不寄對帳單　3. 其他)"
         Height          =   180
         Left            =   -74850
         TabIndex        =   311
         Top             =   3930
         Width           =   6690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "顧問專用信箱："
         Height          =   180
         Index           =   34
         Left            =   -70080
         TabIndex        =   309
         Top             =   1680
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   74
         Left            =   -68808
         TabIndex        =   308
         Top             =   1656
         Width           =   2916
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5143;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblCU144 
         AutoSize        =   -1  'True
         Caption         =   "(N:不開發票)"
         Height          =   180
         Left            =   7590
         TabIndex        =   307
         Top             =   3600
         Width           =   1550
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   57
         Left            =   -68580
         TabIndex        =   306
         Top             =   1740
         Width           =   1095
         VariousPropertyBits=   679493663
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "跨所同意主管："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   39
         Left            =   -69870
         TabIndex        =   305
         Top             =   1800
         Width           =   1260
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   190
         Left            =   -72780
         TabIndex        =   301
         Top             =   4560
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   189
         Left            =   -68880
         TabIndex        =   299
         Top             =   330
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox txt1 
         Height          =   5412
         Index           =   43
         Left            =   -74880
         TabIndex        =   269
         Top             =   336
         Width           =   9000
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "15875;9546"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   51
         Left            =   -72810
         TabIndex        =   252
         Top             =   4170
         Width           =   2565
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4524;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   610
         Index           =   50
         Left            =   -73620
         TabIndex        =   246
         Top             =   2670
         Width           =   7755
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13679;1076"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   41
         Left            =   -73320
         TabIndex        =   204
         Top             =   656
         Width           =   7350
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12965;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   42
         Left            =   -73320
         TabIndex        =   203
         Top             =   1323
         Width           =   7350
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12965;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   900
         Index           =   125
         Left            =   -67530
         TabIndex        =   199
         Top             =   3105
         Width           =   1644
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "2900;1587"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   49
         Left            =   -73380
         TabIndex        =   182
         Top             =   1630
         Width           =   3225
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5689;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   48
         Left            =   -69285
         TabIndex        =   181
         Top             =   1290
         Width           =   3225
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5689;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   47
         Left            =   -73380
         TabIndex        =   180
         Top             =   1290
         Width           =   3225
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5689;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   46
         Left            =   -69285
         TabIndex        =   179
         Top             =   950
         Width           =   3225
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5689;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   900
         Index           =   102
         Left            =   -67536
         TabIndex        =   34
         Top             =   528
         Width           =   1644
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "2900;1587"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   900
         Index           =   103
         Left            =   -67536
         TabIndex        =   35
         Top             =   1824
         Width           =   1644
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "2900;1587"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   4
         Left            =   -73125
         TabIndex        =   143
         Top             =   2265
         Width           =   7200
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12700;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   5
         Left            =   -73125
         TabIndex        =   142
         Top             =   2610
         Width           =   7200
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12700;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   6
         Left            =   -73125
         TabIndex        =   141
         Top             =   2940
         Width           =   7200
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12700;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   7
         Left            =   -73125
         TabIndex        =   140
         Top             =   3285
         Width           =   7200
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12700;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   8
         Left            =   -73125
         TabIndex        =   139
         Top             =   3630
         Width           =   7200
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12700;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   9
         Left            =   -73125
         TabIndex        =   138
         Top             =   3960
         Width           =   7200
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12700;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   10
         Left            =   -73125
         TabIndex        =   137
         Top             =   4305
         Width           =   7200
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12700;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   11
         Left            =   -73125
         TabIndex        =   136
         Top             =   4650
         Width           =   7200
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12700;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   12
         Left            =   -73110
         TabIndex        =   135
         Top             =   5010
         Width           =   7200
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12700;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   3
         Left            =   -73380
         TabIndex        =   6
         Top             =   950
         Width           =   3225
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5689;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   380
         Index           =   2
         Left            =   1590
         TabIndex        =   3
         Top             =   1689
         Width           =   7560
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13335;670"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   45
         Left            =   6615
         TabIndex        =   5
         Top             =   4185
         Width           =   2565
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4524;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   44
         Left            =   5520
         TabIndex        =   4
         Top             =   2100
         Width           =   3465
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "6112;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   640
         Index           =   40
         Left            =   -73710
         TabIndex        =   36
         Top             =   1472
         Width           =   7815
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13785;1129"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   39
         Left            =   -73515
         TabIndex        =   33
         Top             =   5460
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   21
         Left            =   -73890
         TabIndex        =   15
         Top             =   4950
         Width           =   7950
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14033;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   38
         Left            =   -73515
         TabIndex        =   32
         Top             =   5166
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   37
         Left            =   -73515
         TabIndex        =   31
         Top             =   4860
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   36
         Left            =   -73515
         TabIndex        =   30
         Top             =   4554
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   35
         Left            =   -73515
         TabIndex        =   29
         Top             =   4248
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   34
         Left            =   -73515
         TabIndex        =   28
         Top             =   3942
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   33
         Left            =   -73515
         TabIndex        =   27
         Top             =   3636
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   32
         Left            =   -73515
         TabIndex        =   26
         Top             =   3330
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   31
         Left            =   -73515
         TabIndex        =   25
         Top             =   3024
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   30
         Left            =   -73515
         TabIndex        =   24
         Top             =   2718
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   29
         Left            =   -73515
         TabIndex        =   23
         Top             =   2412
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   28
         Left            =   -73515
         TabIndex        =   22
         Top             =   2106
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   27
         Left            =   -73515
         TabIndex        =   21
         Top             =   1800
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   26
         Left            =   -73515
         TabIndex        =   20
         Top             =   1494
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   25
         Left            =   -73515
         TabIndex        =   19
         Top             =   1188
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   24
         Left            =   -73515
         TabIndex        =   18
         Top             =   882
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   23
         Left            =   -73515
         TabIndex        =   17
         Top             =   576
         Width           =   5800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10231;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   315
         Index           =   22
         Left            =   -73515
         TabIndex        =   16
         Top             =   270
         Width           =   5805
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "10239;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   20
         Left            =   -73890
         TabIndex        =   14
         Top             =   4635
         Width           =   7956
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14033;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   19
         Left            =   -73890
         TabIndex        =   13
         Top             =   4317
         Width           =   7956
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14033;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   18
         Left            =   -73890
         TabIndex        =   12
         Top             =   3999
         Width           =   7956
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14033;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   17
         Left            =   -73890
         TabIndex        =   11
         Top             =   3681
         Width           =   7956
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14033;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   420
         Index           =   16
         Left            =   -73890
         TabIndex        =   10
         Top             =   3243
         Width           =   7956
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14033;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   1200
         Index           =   15
         Left            =   -73890
         TabIndex        =   9
         Top             =   2025
         Width           =   7956
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14033;2117"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   420
         Index           =   14
         Left            =   -73890
         TabIndex        =   8
         Top             =   1314
         Width           =   7956
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14033;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   420
         Index           =   13
         Left            =   -73890
         TabIndex        =   7
         Top             =   330
         Width           =   7956
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14033;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   640
         Index           =   1
         Left            =   1590
         TabIndex        =   2
         Top             =   1016
         Width           =   7560
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13335;1129"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   380
         Index           =   0
         Left            =   1590
         TabIndex        =   1
         Top             =   600
         Width           =   7560
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13335;670"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstDeveloper 
         Height          =   315
         Left            =   7680
         TabIndex        =   298
         Top             =   2475
         Width           =   1500
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "2646;556"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Cu180 
         Height          =   525
         Left            =   1080
         TabIndex        =   297
         Top             =   4740
         Width           =   7950
         VariousPropertyBits=   -1467989989
         Size            =   "14023;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "狀態備註："
         Height          =   180
         Index           =   31
         Left            =   120
         TabIndex        =   295
         Top             =   4770
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   73
         Left            =   1800
         TabIndex        =   294
         Top             =   3600
         Width           =   630
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1111;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶特殊付款週期："
         Height          =   180
         Index           =   30
         Left            =   120
         TabIndex        =   293
         Top             =   3600
         Width           =   1620
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   177
         Left            =   -73245
         TabIndex        =   291
         Top             =   916
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   72
         Left            =   -67185
         TabIndex        =   289
         Top             =   2413
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否電子送件：         　　   (Y:是)"
         Height          =   255
         Index           =   19
         Left            =   -69120
         TabIndex        =   288
         Top             =   2413
         Width           =   2925
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   70
         Left            =   -73380
         TabIndex        =   287
         Top             =   3620
         Width           =   3620
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6376;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   71
         Left            =   -67800
         TabIndex        =   286
         Top             =   3620
         Width           =   1790
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3149;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "國內副本接洽人："
         Height          =   260
         Index           =   32
         Left            =   -69360
         TabIndex        =   285
         Top             =   3620
         Width           =   1440
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "國內副本收件人："
         Height          =   260
         Index           =   31
         Left            =   -74850
         TabIndex        =   284
         Top             =   3620
         Width           =   1440
      End
      Begin MSForms.Label lblCU16X 
         Height          =   260
         Index           =   5
         Left            =   -72270
         TabIndex        =   283
         Top             =   3320
         Width           =   360
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "635;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblCU16X 
         Height          =   260
         Index           =   4
         Left            =   -72270
         TabIndex        =   282
         Top             =   3050
         Width           =   360
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "635;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblCU16X 
         Height          =   260
         Index           =   3
         Left            =   -72270
         TabIndex        =   281
         Top             =   2760
         Width           =   360
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "635;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblCU16X 
         Height          =   260
         Index           =   2
         Left            =   -72270
         TabIndex        =   280
         Top             =   2490
         Width           =   360
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "635;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblCU16X 
         Height          =   260
         Index           =   1
         Left            =   -72270
         TabIndex        =   279
         Top             =   2230
         Width           =   360
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "635;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblCU16X 
         Height          =   260
         Index           =   0
         Left            =   -72270
         TabIndex        =   278
         Top             =   1960
         Width           =   360
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "635;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "專利案預設收據公司別-台灣：             (1：專利商標 2：專利法律)"
         Height          =   260
         Index           =   0
         Left            =   -74850
         TabIndex        =   277
         Top             =   1990
         Width           =   5130
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "專利案預設收據公司別-非台灣：         (1：專利商標 2：專利法律 J：台一智權)"
         Height          =   260
         Index           =   1
         Left            =   -74850
         TabIndex        =   276
         Top             =   2260
         Width           =   6140
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "商標案預設收據公司別-台灣：             (1：專利商標 2：專利法律)"
         Height          =   260
         Index           =   2
         Left            =   -74850
         TabIndex        =   275
         Top             =   2520
         Width           =   5130
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "商標案預設收據公司別-非台灣：         (1：專利商標 2：專利法律 J：台一智權)"
         Height          =   260
         Index           =   3
         Left            =   -74850
         TabIndex        =   274
         Top             =   2790
         Width           =   6140
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "其他案預設收據公司別-台灣：             (1：專利商標 2：專利法律)"
         Height          =   260
         Index           =   4
         Left            =   -74850
         TabIndex        =   273
         Top             =   3080
         Width           =   5130
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "其他案預設收據公司別-非台灣：         (1：專利商標 2：專利法律 J：台一智權)"
         Height          =   260
         Index           =   5
         Left            =   -74850
         TabIndex        =   272
         Top             =   3380
         Width           =   6140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(僅提醒)"
         Height          =   180
         Left            =   -74400
         TabIndex        =   271
         Top             =   2955
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "(僅提醒)"
         Height          =   180
         Left            =   -74490
         TabIndex        =   270
         Top             =   1740
         Width           =   660
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   69
         Left            =   7290
         TabIndex        =   267
         Top             =   3600
         Width           =   230
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "406;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊發票："
         Height          =   180
         Index           =   26
         Left            =   6390
         TabIndex        =   268
         Top             =   3600
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   68
         Left            =   2004
         TabIndex        =   266
         Top             =   5292
         Visible         =   0   'False
         Width           =   720
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "不顯示"
         Size            =   "1270;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "預定收款日放寬月數："
         Height          =   180
         Index           =   29
         Left            =   120
         TabIndex        =   265
         Top             =   5292
         Visible         =   0   'False
         Width           =   1800
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   67
         Left            =   -68655
         TabIndex        =   263
         Top             =   4208
         Width           =   390
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "688;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "不催延展：              (Y:不催)"
         Height          =   255
         Index           =   28
         Left            =   -69690
         TabIndex        =   264
         Top             =   4208
         Width           =   2175
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利請款單列印幣別格式："
         Height          =   180
         Index           =   30
         Left            =   -73020
         TabIndex        =   261
         Top             =   1194
         Width           =   2160
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單列印幣別格式："
         Height          =   255
         Index           =   29
         Left            =   -72870
         TabIndex        =   259
         Top             =   2385
         Width           =   2160
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "年費聯絡人："
         Height          =   255
         Index           =   22
         Left            =   -74910
         TabIndex        =   257
         Top             =   2969
         Width           =   1080
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   113
         Left            =   -73620
         TabIndex        =   256
         Top             =   2969
         Width           =   7740
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13652;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP提申急件預設組別："
         Height          =   180
         Index           =   27
         Left            =   -71085
         TabIndex        =   255
         Top             =   916
         Visible         =   0   'False
         Width           =   1920
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   66
         Left            =   4815
         TabIndex        =   253
         Top             =   3600
         Width           =   285
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "503;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發顧問電子報：        (Y:寄/N:不寄)"
         Height          =   180
         Index           =   24
         Left            =   3030
         TabIndex        =   254
         Top             =   3600
         Width           =   3195
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   65
         Left            =   -73260
         TabIndex        =   251
         Top             =   3880
         Width           =   7170
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12647;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   64
         Left            =   -72900
         TabIndex        =   250
         Top             =   3595
         Width           =   6840
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12065;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   63
         Left            =   -73200
         TabIndex        =   249
         Top             =   3310
         Width           =   7110
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12541;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   62
         Left            =   -67020
         TabIndex        =   248
         Top             =   2385
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   61
         Left            =   -73560
         TabIndex        =   247
         Top             =   2385
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N是否列印申請人：            ( Y：印)"
         Height          =   255
         Index           =   23
         Left            =   -69150
         TabIndex        =   245
         Top             =   2385
         Width           =   3270
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "延展D/N列印對象："
         Height          =   255
         Index           =   24
         Left            =   -74820
         TabIndex        =   244
         Top             =   3880
         Width           =   1545
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N固定列印對象："
         Height          =   255
         Index           =   25
         Left            =   -74820
         TabIndex        =   243
         Top             =   3595
         Width           =   1905
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標請款幣別："
         Height          =   255
         Index           =   26
         Left            =   -74820
         TabIndex        =   242
         Top             =   2385
         Width           =   1260
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N備註："
         Height          =   180
         Index           =   27
         Left            =   -74820
         TabIndex        =   241
         Top             =   2730
         Width           =   1185
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標固定請款對象："
         Height          =   255
         Index           =   28
         Left            =   -74820
         TabIndex        =   240
         Top             =   3310
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶彼所商標財務編號："
         Height          =   255
         Index           =   25
         Left            =   -74820
         TabIndex        =   239
         Top             =   4208
         Width           =   1980
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   58
         Left            =   -68820
         TabIndex        =   215
         Top             =   390
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   34
         Left            =   -73560
         TabIndex        =   218
         Top             =   675
         Width           =   3600
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6350;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   35
         Left            =   -73560
         TabIndex        =   217
         Top             =   960
         Width           =   7200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12700;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   38
         Left            =   -73560
         TabIndex        =   213
         Top             =   390
         Visible         =   0   'False
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   138
         Left            =   -68730
         TabIndex        =   198
         Top             =   2100
         Width           =   255
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "450;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   136
         Left            =   -69180
         TabIndex        =   197
         Top             =   1815
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   134
         Left            =   -73560
         TabIndex        =   196
         Top             =   1815
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   46
         Left            =   -73560
         TabIndex        =   162
         Top             =   1245
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   57
         Left            =   -73215
         TabIndex        =   191
         Top             =   2100
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   45
         Left            =   -70920
         TabIndex        =   161
         Top             =   1245
         Width           =   410
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "723;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   40
         Left            =   -73005
         TabIndex        =   160
         Top             =   1530
         Width           =   1470
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2593;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   24
         Left            =   -73560
         TabIndex        =   233
         Top             =   4359
         Width           =   345
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "609;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   37
         Left            =   -73080
         TabIndex        =   235
         Top             =   4637
         Width           =   1650
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2910;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   36
         Left            =   -69120
         TabIndex        =   234
         Top             =   4359
         Width           =   345
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "609;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   131
         Left            =   -69570
         TabIndex        =   229
         Top             =   4081
         Width           =   345
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "609;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   130
         Left            =   -73560
         TabIndex        =   228
         Top             =   4081
         Width           =   345
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "609;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   56
         Left            =   -73290
         TabIndex        =   227
         Top             =   4920
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   133
         Left            =   -69510
         TabIndex        =   223
         Top             =   4637
         Width           =   345
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "609;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   135
         Left            =   -67020
         TabIndex        =   222
         Top             =   4637
         Width           =   345
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "609;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblCU 
         Height          =   255
         Index           =   137
         Left            =   -68985
         TabIndex        =   221
         Top             =   4920
         Width           =   315
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "556;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   44
         Left            =   -73500
         TabIndex        =   158
         Top             =   3803
         Width           =   7605
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13414;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   43
         Left            =   -72990
         TabIndex        =   156
         Top             =   3525
         Width           =   7095
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12515;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   33
         Left            =   -73620
         TabIndex        =   130
         Top             =   3247
         Width           =   7740
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13652;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   32
         Left            =   -73620
         TabIndex        =   129
         Top             =   2691
         Width           =   7740
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13652;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   29
         Left            =   -73290
         TabIndex        =   128
         Top             =   2413
         Width           =   3495
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6165;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   55
         Left            =   -70800
         TabIndex        =   178
         Top             =   2135
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   28
         Left            =   -73800
         TabIndex        =   127
         Top             =   2135
         Width           =   315
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "556;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   54
         Left            =   -67185
         TabIndex        =   174
         Top             =   2135
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利以 EMail 通知：          （Y：是  D：僅D/N）"
         Height          =   255
         Index           =   19
         Left            =   -74910
         TabIndex        =   238
         Top             =   4920
         Width           =   3735
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "專利申請/翻譯折扣：            ％"
         Height          =   180
         Left            =   -70890
         TabIndex        =   237
         Top             =   4359
         Width           =   2385
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "專利全部折扣：           ％"
         Height          =   255
         Left            =   -74910
         TabIndex        =   236
         Top             =   4359
         Width           =   1935
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "專利全部折扣起始日："
         Height          =   255
         Left            =   -74910
         TabIndex        =   232
         Top             =   4637
         Width           =   1800
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "專利領證折扣：           ％"
         Height          =   255
         Index           =   0
         Left            =   -74910
         TabIndex        =   231
         Top             =   4081
         Width           =   1935
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "專利年費折扣：          ％"
         Height          =   255
         Index           =   1
         Left            =   -70890
         TabIndex        =   230
         Top             =   4081
         Width           =   1890
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "專利定稿份數： "
         Height          =   255
         Index           =   0
         Left            =   -70890
         TabIndex        =   226
         Top             =   4637
         Width           =   1305
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "專利請款單份數："
         Height          =   255
         Index           =   1
         Left            =   -68460
         TabIndex        =   225
         Top             =   4637
         Width           =   1440
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "專利 Email 同時寄紙本：       (Y:是)"
         Height          =   180
         Index           =   4
         Left            =   -70890
         TabIndex        =   224
         Top             =   4920
         Width           =   2715
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "延展請款對象："
         Height          =   255
         Index           =   6
         Left            =   -74820
         TabIndex        =   220
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "延展代理人："
         Height          =   255
         Index           =   5
         Left            =   -74820
         TabIndex        =   219
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCT註冊費自動代繳：             (Y:自動代繳)"
         Height          =   255
         Index           =   21
         Left            =   -70650
         TabIndex        =   216
         Top             =   390
         Width           =   3345
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "延展單筆不跑：            （Y：單筆不跑）"
         Height          =   255
         Index           =   16
         Left            =   -74820
         TabIndex        =   214
         Top             =   390
         Visible         =   0   'False
         Width           =   3180
      End
      Begin MSForms.Label lblCU 
         Height          =   260
         Index           =   141
         Left            =   -72750
         TabIndex        =   211
         Top             =   1660
         Width           =   360
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "635;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "是否用LEDES電子帳單：              (Y：是)"
         Height          =   260
         Index           =   9
         Left            =   -74850
         TabIndex        =   212
         Top             =   1690
         Width           =   3260
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人："
         Height          =   260
         Index           =   9
         Left            =   -74850
         TabIndex        =   210
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "實體副本聯絡人："
         Height          =   255
         Index           =   2
         Left            =   -74850
         TabIndex        =   209
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "實體副本收受人："
         Height          =   255
         Index           =   1
         Left            =   -74850
         TabIndex        =   208
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   255
         Index           =   0
         Left            =   -74850
         TabIndex        =   207
         Top             =   660
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   30
         Left            =   -73320
         TabIndex        =   206
         Top             =   360
         Width           =   7300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12876;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   31
         Left            =   -73320
         TabIndex        =   205
         Top             =   1027
         Width           =   7300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12876;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   60
         Left            =   8160
         TabIndex        =   201
         Top             =   3315
         Width           =   315
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "556;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發專利雙週報：        (N:不寄)"
         Height          =   180
         Index           =   23
         Left            =   6360
         TabIndex        =   202
         Top             =   3315
         Width           =   2805
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "業務備註："
         Height          =   180
         Index           =   32
         Left            =   -67530
         TabIndex        =   200
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單份數："
         Height          =   255
         Index           =   5
         Left            =   -70650
         TabIndex        =   195
         Top             =   1815
         Width           =   1440
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "商標定稿份數： "
         Height          =   255
         Index           =   3
         Left            =   -74820
         TabIndex        =   194
         Top             =   1815
         Width           =   1305
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "商標 Email 同時寄紙本：       (Y:是)"
         Height          =   255
         Index           =   2
         Left            =   -70650
         TabIndex        =   193
         Top             =   2100
         Width           =   2715
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   59
         Left            =   5115
         TabIndex        =   190
         Top             =   3315
         Width           =   300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "529;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報：        (N:不寄)"
         Height          =   180
         Index           =   22
         Left            =   3870
         TabIndex        =   189
         Top             =   3315
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發人員："
         Height          =   180
         Index           =   20
         Left            =   6810
         TabIndex        =   188
         Top             =   2475
         Width           =   900
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail(代表)："
         Height          =   180
         Index           =   11
         Left            =   -74640
         TabIndex        =   187
         Top             =   1005
         Width           =   1140
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(財務)："
         Height          =   180
         Index           =   17
         Left            =   -70095
         TabIndex        =   186
         Top             =   1005
         Width           =   660
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(其他1)："
         Height          =   180
         Index           =   18
         Left            =   -74145
         TabIndex        =   185
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(其他2)："
         Height          =   180
         Index           =   19
         Left            =   -70095
         TabIndex        =   184
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(其他3)："
         Height          =   180
         Index           =   20
         Left            =   -74145
         TabIndex        =   183
         Top             =   1630
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP重新委任註記：           (N：不辦理)"
         Height          =   255
         Index           =   21
         Left            =   -72360
         TabIndex        =   177
         Top             =   2135
         Width           =   3015
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "中文地址郵遞區號："
         Height          =   180
         Index           =   31
         Left            =   -74820
         TabIndex        =   173
         Top             =   1800
         Width           =   1620
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   53
         Left            =   -73125
         TabIndex        =   172
         Top             =   1752
         Width           =   2800
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4939;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "呆帳紀錄："
         Height          =   180
         Left            =   7380
         TabIndex        =   171
         Top             =   3885
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   52
         Left            =   8295
         TabIndex        =   170
         Top             =   3885
         Width           =   300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "529;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Left            =   8700
         TabIndex        =   169
         Top             =   3885
         Width           =   465
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   47
         Left            =   -67140
         TabIndex        =   166
         Top             =   1194
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "商標全部折扣起始日："
         Height          =   255
         Left            =   -74820
         TabIndex        =   163
         Top             =   1530
         Width           =   1800
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "年費D/N列對象："
         Height          =   255
         Index           =   18
         Left            =   -74910
         TabIndex        =   159
         Top             =   3803
         Width           =   1365
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N固定列印對象："
         Height          =   255
         Index           =   17
         Left            =   -74910
         TabIndex        =   157
         Top             =   3525
         Width           =   1905
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "公司負責人英文名稱："
         Height          =   180
         Index           =   29
         Left            =   -67530
         TabIndex        =   155
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "定稿中文名稱："
         Height          =   180
         Index           =   30
         Left            =   -67530
         TabIndex        =   154
         Top             =   1605
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "客戶狀態除 業務自行處理 外, 其餘都不列印在客戶名冊中 !!"
         Height          =   180
         Left            =   2790
         TabIndex        =   153
         Top             =   4485
         Width           =   5145
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人（日）："
         Height          =   180
         Index           =   8
         Left            =   -74700
         TabIndex        =   152
         Top             =   4650
         Width           =   1620
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人（英）："
         Height          =   180
         Index           =   7
         Left            =   -74700
         TabIndex        =   151
         Top             =   5040
         Width           =   1620
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人（中）："
         Height          =   180
         Index           =   6
         Left            =   -74700
         TabIndex        =   150
         Top             =   4305
         Width           =   1620
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2（日）："
         Height          =   180
         Index           =   5
         Left            =   -74700
         TabIndex        =   149
         Top             =   3960
         Width           =   1350
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2（英）："
         Height          =   180
         Index           =   4
         Left            =   -74700
         TabIndex        =   148
         Top             =   3630
         Width           =   1350
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2（中）："
         Height          =   180
         Index           =   3
         Left            =   -74700
         TabIndex        =   147
         Top             =   3285
         Width           =   1350
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1（日）："
         Height          =   180
         Index           =   2
         Left            =   -74700
         TabIndex        =   146
         Top             =   2940
         Width           =   1350
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1（英）："
         Height          =   180
         Index           =   1
         Left            =   -74700
         TabIndex        =   145
         Top             =   2610
         Width           =   1350
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1（中）："
         Height          =   180
         Index           =   0
         Left            =   -74700
         TabIndex        =   144
         Top             =   2265
         Width           =   1350
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   -69255
         TabIndex        =   134
         Top             =   1995
         Width           =   3195
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5636;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   -73890
         TabIndex        =   133
         Top             =   1041
         Width           =   2800
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4939;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   -73980
         TabIndex        =   132
         Top             =   420
         Width           =   3200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5644;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   39
         Left            =   -69480
         TabIndex        =   131
         Top             =   638
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   27
         Left            =   -73650
         TabIndex        =   126
         Top             =   1194
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   26
         Left            =   -73245
         TabIndex        =   125
         Top             =   638
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   25
         Left            =   -72780
         TabIndex        =   124
         Top             =   360
         Width           =   420
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "741;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   22
         Left            =   -73890
         TabIndex        =   123
         Top             =   768
         Width           =   2800
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4939;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   20
         Left            =   -69510
         TabIndex        =   122
         Top             =   685
         Width           =   3200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5644;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   19
         Left            =   -69510
         TabIndex        =   121
         Top             =   420
         Width           =   3200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5644;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   18
         Left            =   -73230
         TabIndex        =   120
         Top             =   1995
         Width           =   2970
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5239;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   17
         Left            =   -73980
         TabIndex        =   119
         Top             =   685
         Width           =   3200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5644;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   4800
         TabIndex        =   118
         Top             =   3885
         Width           =   1380
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2434;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   2880
         TabIndex        =   117
         Top             =   3315
         Width           =   660
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1164;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   13
         Left            =   6096
         TabIndex        =   116
         Top             =   3036
         Width           =   2676
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4710;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   5010
         TabIndex        =   115
         Top             =   2745
         Width           =   2640
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4657;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   4950
         TabIndex        =   114
         Top             =   2475
         Width           =   1635
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2884;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   4950
         TabIndex        =   113
         Top             =   2102
         Width           =   555
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "979;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   5205
         TabIndex        =   112
         Top             =   315
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   1170
         TabIndex        =   111
         Top             =   4485
         Width           =   1575
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2778;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   1650
         TabIndex        =   110
         Top             =   4185
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   6
         Left            =   1170
         TabIndex        =   109
         Top             =   3885
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   1275
         TabIndex        =   108
         Top             =   3315
         Width           =   660
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1164;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   2010
         TabIndex        =   107
         Top             =   3030
         Width           =   1770
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3122;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   106
         Top             =   2745
         Width           =   1905
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3360;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   2
         Left            =   1065
         TabIndex        =   105
         Top             =   2475
         Width           =   2415
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4260;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   1065
         TabIndex        =   104
         Top             =   2102
         Width           =   2280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4022;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   103
         Top             =   315
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LABEL 
         AutoSize        =   -1  'True
         Caption         =   "個人或公司："
         Height          =   180
         Left            =   120
         TabIndex        =   102
         Top             =   2745
         Width           =   1080
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人1（中）："
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   101
         Top             =   337
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   100
         Top             =   315
         Width           =   900
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "年費代理人："
         Height          =   255
         Index           =   3
         Left            =   -74910
         TabIndex        =   99
         Top             =   2691
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "年費請款對象："
         Height          =   255
         Index           =   4
         Left            =   -74910
         TabIndex        =   98
         Top             =   3247
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱（中）："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   97
         Top             =   615
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱（英）："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   96
         Top             =   1016
         Width           =   1440
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人1（英）："
         Height          =   180
         Index           =   1
         Left            =   -74910
         TabIndex        =   95
         Top             =   640
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人1（日）："
         Height          =   180
         Index           =   2
         Left            =   -74910
         TabIndex        =   94
         Top             =   943
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人2（中）："
         Height          =   180
         Index           =   3
         Left            =   -74910
         TabIndex        =   93
         Top             =   1246
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人2（英）："
         Height          =   180
         Index           =   4
         Left            =   -74910
         TabIndex        =   92
         Top             =   1549
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人2（日）："
         Height          =   180
         Index           =   5
         Left            =   -74910
         TabIndex        =   91
         Top             =   1852
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人3（中）："
         Height          =   180
         Index           =   6
         Left            =   -74910
         TabIndex        =   90
         Top             =   2155
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人3（英）："
         Height          =   180
         Index           =   7
         Left            =   -74910
         TabIndex        =   89
         Top             =   2458
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人3（日）："
         Height          =   180
         Index           =   8
         Left            =   -74940
         TabIndex        =   88
         Top             =   2761
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人4（中）："
         Height          =   180
         Index           =   9
         Left            =   -74940
         TabIndex        =   87
         Top             =   3064
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人4（英）："
         Height          =   180
         Index           =   10
         Left            =   -74940
         TabIndex        =   86
         Top             =   3367
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人4（日）："
         Height          =   180
         Index           =   11
         Left            =   -74925
         TabIndex        =   85
         Top             =   3670
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人5（中）："
         Height          =   180
         Index           =   12
         Left            =   -74925
         TabIndex        =   84
         Top             =   3973
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人5（英）："
         Height          =   180
         Index           =   13
         Left            =   -74925
         TabIndex        =   83
         Top             =   4276
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人5（日）："
         Height          =   180
         Index           =   14
         Left            =   -74925
         TabIndex        =   82
         Top             =   4579
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人6（中）："
         Height          =   180
         Index           =   15
         Left            =   -74925
         TabIndex        =   81
         Top             =   4882
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人6（英）："
         Height          =   180
         Index           =   16
         Left            =   -74925
         TabIndex        =   80
         Top             =   5185
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人6（日）："
         Height          =   240
         Index           =   17
         Left            =   -74925
         TabIndex        =   79
         Top             =   5497
         Width           =   1350
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "聯絡地址："
         Height          =   180
         Index           =   18
         Left            =   -74820
         TabIndex        =   78
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "郵遞區號："
         Height          =   180
         Index           =   19
         Left            =   -74820
         TabIndex        =   77
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "中文地址："
         Height          =   180
         Index           =   20
         Left            =   -74820
         TabIndex        =   76
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "英文地址："
         Height          =   180
         Index           =   21
         Left            =   -74820
         TabIndex        =   75
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "地址國籍："
         Height          =   180
         Index           =   22
         Left            =   -74820
         TabIndex        =   74
         Top             =   1050
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB1："
         Height          =   180
         Index           =   23
         Left            =   -74820
         TabIndex        =   73
         Top             =   3690
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB2："
         Height          =   180
         Index           =   24
         Left            =   -74820
         TabIndex        =   72
         Top             =   4020
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB3："
         Height          =   180
         Index           =   25
         Left            =   -74820
         TabIndex        =   71
         Top             =   4350
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB4："
         Height          =   180
         Index           =   26
         Left            =   -74820
         TabIndex        =   70
         Top             =   4650
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "POB5："
         Height          =   180
         Index           =   27
         Left            =   -74820
         TabIndex        =   69
         Top             =   4953
         Width           =   600
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "日文地址："
         Height          =   240
         Index           =   28
         Left            =   -74820
         TabIndex        =   68
         Top             =   3240
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人編號："
         Height          =   180
         Index           =   3
         Left            =   4065
         TabIndex        =   67
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶名稱（日）："
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   66
         Top             =   1689
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶國籍："
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   65
         Top             =   2102
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   64
         Top             =   2475
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公司負責人："
         Height          =   180
         Index           =   7
         Left            =   3870
         TabIndex        =   63
         Top             =   2745
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶來源："
         Height          =   180
         Index           =   8
         Left            =   3840
         TabIndex        =   62
         Top             =   2102
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "業務區別："
         Height          =   180
         Index           =   9
         Left            =   3810
         TabIndex        =   61
         Top             =   2475
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "身分證字號/統一編號："
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   60
         Top             =   3030
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "預設接洽人："
         Height          =   180
         Index           =   11
         Left            =   4956
         TabIndex        =   59
         Top             =   3036
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "行業別："
         Height          =   180
         Index           =   12
         Left            =   2130
         TabIndex        =   58
         Top             =   3315
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發日期："
         Height          =   180
         Index           =   13
         Left            =   3870
         TabIndex        =   57
         Top             =   3885
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分所編號/客戶彼所專利財務編號："
         Height          =   180
         Index           =   14
         Left            =   3825
         TabIndex        =   56
         Top             =   4185
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "每月收款日："
         Height          =   180
         Index           =   15
         Left            =   120
         TabIndex        =   55
         Top             =   3315
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文：            （1:中文   2:英文   3:日文）"
         Height          =   180
         Index           =   16
         Left            =   120
         TabIndex        =   54
         Top             =   3885
         Width           =   3555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄台一雜誌：           （N:不寄）"
         Height          =   180
         Index           =   17
         Left            =   120
         TabIndex        =   53
         Top             =   4185
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶狀態："
         Height          =   180
         Index           =   18
         Left            =   120
         TabIndex        =   52
         Top             =   4485
         Width           =   900
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "TEL1："
         Height          =   180
         Index           =   9
         Left            =   -74640
         TabIndex        =   51
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "FAX1："
         Height          =   180
         Index           =   10
         Left            =   -74640
         TabIndex        =   50
         Top             =   685
         Width           =   600
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "TEL2："
         Height          =   180
         Index           =   12
         Left            =   -70080
         TabIndex        =   49
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "FAX2："
         Height          =   180
         Index           =   13
         Left            =   -70080
         TabIndex        =   48
         Top             =   685
         Width           =   600
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         Height          =   180
         Index           =   14
         Left            =   -70155
         TabIndex        =   47
         Top             =   1995
         Width           =   795
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "MOBILE PHONE："
         Height          =   180
         Index           =   15
         Left            =   -74700
         TabIndex        =   46
         Top             =   1995
         Width           =   1440
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "收款後辦案：          （Y：先收）"
         Height          =   255
         Index           =   7
         Left            =   -74910
         TabIndex        =   45
         Top             =   2135
         Width           =   2550
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利固定請款對象："
         Height          =   255
         Index           =   8
         Left            =   -74910
         TabIndex        =   44
         Top             =   2413
         Width           =   1620
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N備註："
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   43
         Top             =   1515
         Width           =   1185
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利請款幣別："
         Height          =   180
         Index           =   11
         Left            =   -74880
         TabIndex        =   42
         Top             =   1194
         Width           =   1260
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCP領證自動代繳：          （Y：自動代繳）"
         Height          =   180
         Index           =   12
         Left            =   -74880
         TabIndex        =   41
         Top             =   638
         Width           =   3390
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費通知函單筆不跑：        （Y：單筆不跑）"
         Height          =   180
         Index           =   13
         Left            =   -74880
         TabIndex        =   40
         Top             =   360
         Width           =   3840
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費自動代繳：       （Y：自動代繳 /  N：寄證書後年費不續辦)"
         Height          =   180
         Index           =   15
         Left            =   -71085
         TabIndex        =   39
         Top             =   638
         Width           =   5235
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "商標全部折扣：           ％"
         Height          =   255
         Left            =   -74820
         TabIndex        =   165
         Top             =   1245
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "商標申請/翻譯折扣：           ％"
         Height          =   260
         Left            =   -72630
         TabIndex        =   164
         Top             =   1245
         Width           =   2340
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N是否列印申請人：          （Y：印）"
         Height          =   180
         Index           =   14
         Left            =   -69255
         TabIndex        =   167
         Top             =   1194
         Width           =   3315
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否核對已准專利：          （N：否）"
         Height          =   255
         Left            =   -69120
         TabIndex        =   175
         Top             =   2135
         Width           =   3210
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標以 EMail 通知：        （Y：是  D：僅D/N）"
         Height          =   255
         Index           =   20
         Left            =   -74820
         TabIndex        =   192
         Top             =   2100
         Width           =   3645
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "FCP實審自動代繳：          （Y：自動代繳）"
         Height          =   180
         Index           =   33
         Left            =   -74880
         TabIndex        =   292
         Top             =   916
         Width           =   3390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "台灣案專利證書形式：             (1:電子 2:紙本)"
         Height          =   180
         Index           =   32
         Left            =   -70740
         TabIndex        =   300
         Top             =   360
         Width           =   3540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "台灣案商標註冊證形式：             (1:電子 2:紙本)"
         Height          =   180
         Index           =   33
         Left            =   -74820
         TabIndex        =   302
         Top             =   4560
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利不得請雜費：     　   (Y:是)"
         Height          =   180
         Index           =   35
         Left            =   -68460
         TabIndex        =   318
         Top             =   4080
         Width           =   2472
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "繳註冊費折扣：           ％"
         Height          =   180
         Left            =   -69960
         TabIndex        =   343
         Top             =   1245
         Width           =   1990
      End
   End
   Begin VB.Label SpecCU 
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   30
      TabIndex        =   176
      Top             =   30
      Width           =   3465
   End
End
Attribute VB_Name = "frm100101_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/15 改成Form2.0 ; lbl1(index)、txt1(index)、Cu180、lblCU(index)、Text1(index)、lblCU16X(index)、lstDeveloper
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit

'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'add by nickc 2007/08/24 改成同 cu 維護
Dim strTmp As String
Public m_pub_QL05 As String 'Add By Sindy 2025/8/28 只記錄於此Form


'92.04.16 nick
Public Sub PubShowNextData()
Dim stName As String 'Add by Amy 2022/12/05
    
Select Case cmdState
Case 0
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1
     fnCloseAllFrm100
' add by nick 2004/07/15 加入客戶減免身份查詢
Case 2
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_13.Show
     '傳申請人代號
     frm100101_13.txtFn(0).Text = lbl1(0).Caption
     frm100101_13.txtFn(1).Text = lbl1(0).Caption
     frm100101_13.cmQueryData
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'Add By Sindy 2012/9/27
Case 3 '平台帳號
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_26.Show
     frm100101_26.Tag = Trim(lbl1(0).Caption)
     'frm100101_26.Tag = "X23450000"
     frm100101_26.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'Added by Lydia 2016/11/23
Case 4 '各項指示
     'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
     If PUB_CheckFormExist("frm12040159") Then
         MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
         Exit Sub
     End If
     'end 2020/05/05
     
    cmdState = -1
    Me.Enabled = False
    If fnSaveParentForm(Me) = False Then
       Me.Enabled = True
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    frm12040159.SetParent "Q", Trim(lbl1(0).Caption), Me
    frm12040159.Show
    Screen.MousePointer = vbDefault
    Me.Enabled = True
'Add by Amy 2018/01/16
Case 5 '合約資料查詢
    cmdState = -1
    Me.Enabled = False
    If fnSaveParentForm(Me) = False Then
       Me.Enabled = True
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    frm100101_N.lblCT01 = Left(lbl1(0).Caption, 8)
    frm100101_N.Show
    Call frm100101_N.QueryData(Left(lbl1(0).Caption, 8))
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
Case 6 'Add by Amy 2022/12/05 被介紹者
    If CmdOk1(6).BackColor <> &HFFFF80 Then
        MsgBox "無被介紹者資料"
        Exit Sub
    End If
    If PUB_CheckFormExist("frm050705_1") Then
         MsgBox "請先關閉〔被介紹資料〕的畫面！", vbInformation
         Exit Sub
    End If
    cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
    '英->中->日
    If Trim(txt1(1)) = MsgText(601) Then
        If Trim(txt1(0)) = MsgText(601) Then
            stName = txt1(2)
        Else
            stName = txt1(0)
        End If
    Else
        stName = txt1(1)
    End If
    frm050705_1.txtNo = Left(lbl1(0), 8)
    frm050705_1.lbl1(0) = Mid(lbl1(1), 1, InStr(lbl1(1), " "))
    frm050705_1.lbl1(1) = Mid(lbl1(1), InStr(lbl1(1), " ") + 1)
    frm050705_1.lbl1(3) = stName
    frm050705_1.SetParent Me
    frm050705_1.QueryData
    frm050705_1.Show
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Case Else
End Select
End Sub

Private Sub cmdok1_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
''92.04.16 nick 以下無效
'Select Case Index
'Case 0
'      Me.Hide
'Case 1
'      bolToEndByNick = True
'     Unload Me
'     Exit Sub
'Case Else
'End Select
End Sub

Sub StrMenu()
Dim strSql  As String, i As Integer
Dim Str01 As String  ', str02 As String, str03 As String, str04 As String
'Modify by Morgan 2008/1/17
'Dim strArr() As String, StrOk(54) As String, StrOkTxt(45) As String
'2008/12/9 modify by sonia
'Dim strArr() As String, StrOk(57) As String, StrOkTxt(49) As String
'Modified by Lydia 2017/11/30
'Dim strArr() As String, StrOk(71) As String, StrOkTxt(51) As String
'Modified by Morgan 2018/10/29
'Dim strArr() As String, StrOk(72) As String, StrOkTxt(51) As String
'Modified by Lydia 2019/05/17
'Dim strArr() As String, StrOk(72) As String, StrOkTxt(52) As String
'Modified by Morgan 2021/10/15
'Dim strArr() As String, StrOk(73) As String, StrOkTxt(52) As String
'Modified by Lydia 2024/01/15
'Dim strArr() As String, StrOk(73) As String, StrOkTxt(57) As String 'Modify by Amy 2023/05/10 原:StrOkTxt(56)
Dim strArr() As String, StrOk(77) As String, StrOkTxt(57) As String

ReDim strArr(TF_CU) As String
'end 2007/10/29
Dim CU102 As String
'add by nickc 2007/12/28
Dim Cu121 As String
'Add by Morgan 2008/11/13
Dim oLbl As Object
Dim arrID 'Add By Sindy 2025/1/6

Str01 = ""
Str01 = Me.Tag

'Added by Lydia 2021/12/15
For Each oLbl In lbl1
   oLbl.BackColor = &H8000000F
Next
'end 2021/12/15

'add by sonia 2024/8/13 增加檢查有無申請人查詢權限
If CheckUse("frm100102_1", strExec, False) Then
Else
   i = MsgBox("您沒有查詢申請人資料的權限！", , "沒權限")
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
'end 2024/8/13

'Add By Sindy 2011/01/03 檢查國內外權限
If Len(Str01) <> 9 Then Str01 = Str01 & "0"
If CheckSR12(Str01) = False Then
   Screen.MousePointer = vbDefault
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
pub_QL05 = m_pub_QL05 & ";客戶編號：" & Str01 & "(基本資料)" 'Add By Sindy 2025/8/13

If Len(Str01) = 9 Then
    strSql = "SELECT * FROM customer Where CU01='" & Left(Str01, 8) & "' AND CU02='" & Right(Str01, 1) & "' "
Else
    strSql = "SELECT * FROM CUSTOMER WHERE CU01='" & Str01 & "' AND CU02='0'"
End If
CheckOC

adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/13
   If Not IsNull(adoRecordset.Fields("CU102")) Then
      CU102 = adoRecordset.Fields("CU102")
   Else
      CU102 = ""
   End If
   
   If Not IsNull(adoRecordset.Fields("CU103")) Then
      Text1(102) = adoRecordset.Fields("CU103")
   Else
      Text1(102) = ""
   End If
   
   If Not IsNull(adoRecordset.Fields("CU104")) Then
      Text1(103) = adoRecordset.Fields("CU104")
   Else
      Text1(103) = ""
   End If
   
   'Add By Sindy 2009/10/26
   If Not IsNull(adoRecordset.Fields("CU125")) Then
      Text1(125) = adoRecordset.Fields("CU125")
   Else
      Text1(125) = ""
   End If
   
   'add by nickc 2007/12/28
   Cu121 = CheckStr(adoRecordset.Fields("CU121"))
   'Add by Amy 2019/08/27
   Cu180 = CheckStr("" & adoRecordset.Fields("CU180"))
   
   'For i = 0 To (108 - 1)
   'edit by nickc 2005/12/06
   'For i = 0 To (109 - 1)
   For i = 0 To UBound(strArr) - 1
      Select Case i
      'Modify By Sindy 2013/1/30 +, 155, 156
      '數值
      Case 13, 34, 35, 36, 37, 81, 82, 84, 85, 106, 107, 108, 155, 156
           If IsNull(adoRecordset.Fields(i)) Then
               strArr(i + 1) = " "
           Else
               strArr(i + 1) = str(adoRecordset.Fields(i))
           End If
      '文字
      Case Else
           If IsNull(adoRecordset.Fields(i)) Then
                strArr(i + 1) = " "
           Else
                strArr(i + 1) = adoRecordset.Fields(i)
           End If
      End Select
      DoEvents
   Next i
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
   ShowNoData
   Screen.MousePointer = vbDefault
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
CheckOC
Dim strTemp As String    '暫存
Dim strTemp1 As Variant, strTemp2 As Variant, strTemp3 As Variant
Dim j As Integer
'For i = 0 To 100
'For i = 1 To 108
'edit by nickc 2005/12/06
'For i = 1 To 109
For i = 1 To UBound(strArr)
    Select Case i
    Case 1
         StrOk(0) = strArr(1) + strArr(2)
    Case 10
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(1) = strArr(i) + ""
              Else
                  StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
              'Add by Morgan 2004/1/19
              lbl1(1).ForeColor = vbBlack
         Else
              'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
              'StrOk(1) = ""
              StrOk(1) = strArr(i) + ""
              'Add by Morgan 2004/1/19
              lbl1(1).ForeColor = vbRed
              StrOk(1) = strArr(i)
         End If
         CheckOC
    Case 4
         StrOkTxt(0) = strArr(i)
    Case 5
         'Modify by Amy 2017/07/11 +IIF 避免太多個換行,游標跳入TextBox會跑到最後,不到文字
         StrOkTxt(1) = strArr(i) & IIf(Len(Trim(strArr(88))) = 0, "", Chr(13) & Chr(10) & strArr(88)) & IIf(Len(Trim(strArr(89))) = 0, "", Chr(13) & Chr(10) & strArr(89)) & _
                    IIf(Len(Trim(strArr(90))) = 0, "", Chr(13) & Chr(10) & strArr(90))
    Case 6
         StrOkTxt(2) = strArr(i)
    Case 13
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(2) = strArr(i) + ""
              Else
                  StrOk(2) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
              'Add by Morgan 2004/1/19
              lbl1(2).ForeColor = vbBlack
         Else
              'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
              'StrOk(2) = ""
              StrOk(2) = strArr(i) + ""
              'Add by Morgan 2004/1/19
              lbl1(2).ForeColor = vbRed
              StrOk(2) = strArr(i)
         End If
         CheckOC
    Case 15
'2012/5/29 MODIF BY SONIA
'         If strArr(i) = "0" Then
'              StrOk(3) = "0  個人"
'         Else
'            If strArr(i) = "1" Then
'                StrOk(3) = "1  公司"
'            Else
'                StrOk(3) = "錯誤"
'            End If
'         End If
         Select Case strArr(i)
            Case "0"
               StrOk(3) = "0  個人"
            Case "1"
               StrOk(3) = "1  公司"
            Case "2"
               StrOk(3) = "2  學校"
            Case "3"
               StrOk(3) = "3  特殊機構"
            Case Else
               'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
               'StrOk(3) = "錯誤"
               StrOk(3) = strArr(i) & "  錯誤"
         End Select
'2012/5/29 END
    Case 11
         StrOk(4) = strArr(i)
    Case 35
         StrOk(5) = strArr(i)
    Case 64
         StrOk(6) = strArr(i)
    Case 32
         StrOk(7) = strArr(i)
    Case 80
         StrOk(8) = strArr(i)
    Case 3
         StrOk(9) = strArr(i)
    Case 9
         strSql = "SELECT CSM02 FROM CASESOURCEMAP WHERE CSM01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(10) = strArr(i) + ""
                  StrOkTxt(44) = ""
              Else
                  StrOk(10) = strArr(i)
                  StrOkTxt(44) = adoRecordset.Fields(0)
              End If
         Else
              'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
              'StrOk(10) = ""
              StrOk(10) = strArr(i) + ""
              StrOkTxt(44) = ""
         End If
         CheckOC
    Case 12
         strSql = "SELECT A0902 FROM ACC090 WHERE A0901='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(11) = strArr(i) + ""
              Else
                  StrOk(11) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
              'Add by Morgan 2004/1/19
              lbl1(11).ForeColor = vbBlack
         Else
              'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
              'StrOk(11) = ""
              StrOk(11) = strArr(i) + ""
              'Add by Morgan 2004/1/19
              lbl1(11).ForeColor = vbRed
              StrOk(11) = strArr(i)
         End If
         CheckOC
   Case 7
         StrOk(12) = strArr(i)
   Case 34
         StrOk(14) = strArr(i)
   Case 14
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(15) = ""
         Else
             StrOk(15) = ChangeWStringToTString(strArr(i))
         End If
   Case 33
        StrOkTxt(45) = strArr(i)
   Case 79
        StrOkTxt(43) = strArr(i)
   Case 16
        StrOk(16) = strArr(i)
   Case 17
        StrOk(19) = strArr(i)
   Case 18
        StrOk(17) = strArr(i)
   Case 19
        StrOk(20) = strArr(i)
   Case 20
        StrOkTxt(3) = strArr(i)
   Case 21
        StrOk(21) = strArr(i)
   Case 22
        StrOk(18) = strArr(i)
   Case 58
        StrOkTxt(4) = strArr(i)
   Case 59
        StrOkTxt(5) = strArr(i)
   Case 60
        StrOkTxt(6) = strArr(i)
   Case 61
        StrOkTxt(7) = strArr(i)
   Case 62
        StrOkTxt(8) = strArr(i)
   Case 63
        StrOkTxt(9) = strArr(i)
   Case 91
        StrOkTxt(10) = strArr(i)
   Case 92
        StrOkTxt(11) = strArr(i)
   Case 93
        StrOkTxt(12) = strArr(i)
   Case 31
        StrOkTxt(13) = strArr(i)
   Case 30
        StrOk(22) = strArr(i)
   Case 23
        StrOkTxt(14) = strArr(i)
   Case 24
        'Modify by Amy 2017/07/11 +IIF 避免太多個換行,游標跳入TextBox會跑到最後,不到文字
        StrOkTxt(15) = strArr(i) & IIf(Len(Trim(strArr(25))) = 0, "", Chr(13) & Chr(10) & strArr(25)) & IIf(Len(Trim(strArr(26))) = 0, "", Chr(13) & Chr(10) & strArr(26)) & _
                    IIf(Len(Trim(strArr(27))) = 0, "", Chr(13) & Chr(10) & strArr(27)) & IIf(Len(Trim(strArr(28))) = 0, "", Chr(13) & Chr(10) & strArr(28)) & IIf(Len(Trim(strArr(102))) = 0, "", Chr(13) & Chr(10) & CU102)
   Case 29
        StrOkTxt(16) = strArr(i)
   Case 87
        strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
        CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(23) = strArr(i) + ""
              Else
                  StrOk(23) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
              'Add by Morgan 2004/1/19
              lbl1(23).ForeColor = vbBlack
         Else
              'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
              'StrOk(23) = ""
              StrOk(23) = strArr(i) + ""
              'Add by Morgan 2004/1/19
              lbl1(23).ForeColor = vbRed
              StrOk(23) = strArr(i)
         End If
         CheckOC
    Case 65
         StrOkTxt(17) = strArr(i)
    Case 66
         StrOkTxt(18) = strArr(i)
    Case 67
         StrOkTxt(19) = strArr(i)
    Case 68
         StrOkTxt(20) = strArr(i)
    Case 69
         StrOkTxt(21) = strArr(i)
    Case 39
         StrOkTxt(22) = strArr(i)
    Case 40
         StrOkTxt(23) = strArr(i)
    Case 41
         StrOkTxt(24) = strArr(i)
    Case 42
         StrOkTxt(25) = strArr(i)
    Case 43
         StrOkTxt(26) = strArr(i)
    Case 44
         StrOkTxt(27) = strArr(i)
    Case 45
         StrOkTxt(28) = strArr(i)
    Case 46
         StrOkTxt(29) = strArr(i)
    Case 47
         StrOkTxt(30) = strArr(i)
    Case 48
         StrOkTxt(31) = strArr(i)
    Case 49
         StrOkTxt(32) = strArr(i)
    Case 50
         StrOkTxt(33) = strArr(i)
    Case 51
         StrOkTxt(34) = strArr(i)
    Case 52
         StrOkTxt(35) = strArr(i)
    Case 53
         StrOkTxt(36) = strArr(i)
    Case 54
         StrOkTxt(37) = strArr(i)
    Case 55
         StrOkTxt(38) = strArr(i)
    Case 56
         StrOkTxt(39) = strArr(i)
    Case 36
         StrOk(24) = strArr(i)
    Case 73
         StrOk(25) = strArr(i)
    Case 75
         StrOk(26) = strArr(i)
    Case 76
         StrOk(27) = strArr(i)
    Case 78
         StrOkTxt(40) = strArr(i)
    Case 72
         StrOk(28) = strArr(i)
    Case 57
'edit by nickc 2007/08/24 改成同 cu 維護
'         If Left$(strArr(i), 1) = "X" Then
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
'         Else
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                StrOk(29) = strArr(i) + ""
'            Else
'                StrOk(29) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
          If Trim(strArr(i)) <> "" Then
                If ClsLawLawGetName(strArr(i), strTmp) = True Then
                    StrOk(29) = strArr(i) + "  " + strTmp
                
                    'Add by Morgan 2004/1/19
                    lbl1(29).ForeColor = vbBlack
                Else
                   'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
                   'StrOk(29) = ""
                   StrOk(29) = strArr(i) + ""
                   lbl1(29).ForeColor = vbRed
                   StrOk(29) = strArr(i)
                End If
         Else
            StrOk(29) = ""
            'Add by Morgan 2004/1/19
            lbl1(29).ForeColor = vbRed
            StrOk(29) = strArr(i)
         End If
         CheckOC
    Case 71
'edit by nickc 2007/08/24 改成同 cu 維護
'         If Left$(strArr(i), 1) = "X" Then
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
'         Else
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                StrOk(30) = strArr(i) + ""
'            Else
'                StrOk(30) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
          If Trim(strArr(i)) <> "" Then
             If ClsLawLawGetName(strArr(i), strTmp) = True Then
                StrOk(30) = strArr(i) + "  " + strTmp
                
                'Add by Morgan 2004/1/19
                lbl1(30).ForeColor = vbBlack
            Else
               'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
               'StrOk(30) = ""
               StrOk(30) = strArr(i) + ""
               lbl1(30).ForeColor = vbRed
               StrOk(30) = strArr(i)
            End If
         Else
            StrOk(30) = ""
            'Add by Morgan 2004/1/19
            lbl1(30).ForeColor = vbRed
            StrOk(30) = strArr(i)
         End If
         CheckOC
    Case 70
         StrOkTxt(41) = strArr(i)
    Case 94
'edit by nickc 2007/08/24 改成同 cu 維護
'         If Left$(strArr(i), 1) = "X" Then
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
'         Else
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                StrOk(31) = strArr(i) + ""
'            Else
'                StrOk(31) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
          If Trim(strArr(i)) <> "" Then
            If ClsLawLawGetName(strArr(i), strTmp) = True Then
                StrOk(31) = strArr(i) + "  " + strTmp
            Else
                'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
                'StrOk(31) = ""
                StrOk(31) = strArr(i) + ""
            End If
          Else
              StrOk(31) = ""
          End If
            CheckOC
    Case 95
         StrOkTxt(42) = strArr(i)
    Case 96
'edit by nickc 2007/08/24 改成同 cu 維護
'         If Left$(strArr(i), 1) = "X" Then
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
'         Else
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                StrOk(32) = strArr(i) + ""
'            Else
'                StrOk(32) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
          If Trim(strArr(i)) <> "" Then
              If ClsLawLawGetName(strArr(i), strTmp) = True Then
                  StrOk(32) = strArr(i) + "  " + strTmp
                
                'Add by Morgan 2004/1/19
                lbl1(32).ForeColor = vbBlack
            Else
               'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
               'StrOk(32) = ""
               StrOk(32) = strArr(i) + ""
               lbl1(32).ForeColor = vbRed
               StrOk(32) = strArr(i)
            End If
         Else
            StrOk(32) = ""
            'Add by Morgan 2004/1/19
            lbl1(32).ForeColor = vbRed
            StrOk(32) = strArr(i)
         End If
         CheckOC
    Case 97
'edit by nickc 2007/08/24 改成同 cu 維護
'         If Left$(strArr(i), 1) = "X" Then
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
'         Else
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                StrOk(33) = strArr(i) + ""
'            Else
'                StrOk(33) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
          If Trim(strArr(i)) <> "" Then
             If ClsLawLawGetName(strArr(i), strTmp) = True Then
                    StrOk(33) = strArr(i) + "  " + strTmp
                'Add by Morgan 2004/1/19
                lbl1(33).ForeColor = vbBlack
            Else
               'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
               'StrOk(33) = ""
               StrOk(33) = strArr(i) + ""
               lbl1(33).ForeColor = vbRed
               StrOk(33) = strArr(i)
            End If
         Else
            StrOk(33) = ""
            'Add by Morgan 2004/1/19
            lbl1(33).ForeColor = vbRed
            StrOk(33) = strArr(i)
         End If
         CheckOC
    Case 98
'edit by nickc 2007/08/24 改成同 cu 維護
'         If Left$(strArr(i), 1) = "X" Then
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
'         Else
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                StrOk(34) = strArr(i) + ""
'            Else
'                StrOk(34) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
          If Trim(strArr(i)) <> "" Then
            If ClsLawLawGetName(strArr(i), strTmp) = True Then
                StrOk(34) = strArr(i) + "  " + strTmp
                
                'Add by Morgan 2004/1/19
                lbl1(34).ForeColor = vbBlack
            Else
               'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
               'StrOk(34) = ""
               StrOk(34) = strArr(i) + ""
               lbl1(34).ForeColor = vbRed
               StrOk(34) = strArr(i)
            End If
         Else
            StrOk(34) = ""
            'Add by Morgan 2004/1/19
            lbl1(34).ForeColor = vbRed
            StrOk(34) = strArr(i)
         End If
         CheckOC
    Case 99
'edit by nickc 2007/08/24 改成同 cu 維護
'         If Left$(strArr(i), 1) = "X" Then
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
'         Else
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                StrOk(35) = strArr(i) + ""
'            Else
'                StrOk(35) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
          If Trim(strArr(i)) <> "" Then
              If ClsLawLawGetName(strArr(i), strTmp) = True Then
                    StrOk(35) = strArr(i) + "  " + strTmp
                
                'Add by Morgan 2004/1/19
                lbl1(35).ForeColor = vbBlack
              Else
                  'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
                  'StrOk(35) = ""
                  StrOk(35) = strArr(i) + ""
                  lbl1(35).ForeColor = vbRed
                  StrOk(35) = strArr(i)
              End If
         Else
            StrOk(35) = ""
            'Add by Morgan 2004/1/19
            lbl1(35).ForeColor = vbRed
            StrOk(35) = strArr(i)
         End If
         CheckOC
    Case 37
         StrOk(36) = strArr(i)
    Case 38
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(37) = ""
         Else
             StrOk(37) = ChangeWStringToTString(strArr(i))
         End If
    Case 100
         StrOk(38) = strArr(i)
    Case 74
         StrOk(39) = strArr(i)
    Case 77 'D/N是否列印申請人
         StrOk(47) = strArr(i)
    Case 81
         'edit by nick 2004/10/05
         'StrOk(41) = strArr(i)
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               StrOk(41) = strArr(i) + ""
            Else
               StrOk(41) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
            'StrOk(41) = ""
            StrOk(41) = strArr(i) + ""
         End If
         CheckOC
    Case 84
         'edit by nick 2004/10/05
         'StrOk(42) = strArr(i)
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               StrOk(42) = strArr(i) + ""
            Else
               StrOk(42) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
            'StrOk(42) = ""
            StrOk(42) = strArr(i) + ""
         End If
         CheckOC
    'add by nick 2004/10/05   start
    Case 82
         '2011/8/16 ADD BY SONIA
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
            StrOk(48) = ""
         Else
         '2011/8/16 END
            StrOk(48) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If  '2011/8/16 ADD BY SONIA
    Case 83
          StrOk(49) = Format(strArr(i), "##:##")
    Case 85
         '2011/8/16 ADD BY SONIA
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
            StrOk(50) = ""
         Else
         '2011/8/16 END
            StrOk(50) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If  '2011/8/16 ADD BY SONIA
    Case 86
          StrOk(51) = Format(strArr(i), "##:##")
    'add by nick 2004/10/05 end
    Case 105 'D/N固定列印對象
'edit by nickc 2007/08/24 改成同 cu 維護
'         If Left$(strArr(i), 1) = "X" Then
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
'         Else
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                StrOk(43) = strArr(i) + ""
'            Else
'                StrOk(43) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
          If Trim(strArr(i)) <> "" Then
                If ClsLawLawGetName(strArr(i), strTmp) = True Then
                       StrOk(43) = strArr(i) + "  " + strTmp
                Else
                   'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
                   'StrOk(43) = ""
                   StrOk(43) = strArr(i) + ""
                End If
         Else
            StrOk(43) = ""
         End If
         CheckOC
    Case 106 '年費延展D/N列印對象
'edit by nickc 2007/08/24 改成同 cu 維護
'         If Left$(strArr(i), 1) = "X" Then
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
'         Else
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                StrOk(44) = strArr(i) + ""
'            Else
'                StrOk(44) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
          If Trim(strArr(i)) <> "" Then
                If ClsLawLawGetName(strArr(i), strTmp) = True Then
                       StrOk(44) = strArr(i) + "  " + strTmp
                Else
                   'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
                   'StrOk(44) = ""
                   StrOk(44) = strArr(i) + ""
                End If
         Else
            StrOk(44) = ""
         End If
         CheckOC
    'Add By Cheng 2003/11/18
    Case 107 '商標全部折扣
        StrOk(46) = strArr(i)
    Case 108 '商標申請/翻譯折扣
        StrOk(45) = strArr(i)
    Case 109 '商標全部折扣起始日
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(40) = ""
         Else
             StrOk(40) = ChangeWStringToTString(strArr(i))
         End If
    'End
    'Add By Sindy 2025/3/10
    Case 203 '繳註冊費折扣
        StrOk(76) = strArr(i)
    Case 204 '延展折扣
        StrOk(77) = strArr(i)
    Case 205 '商標全部折扣終止日
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(75) = ""
         Else
             StrOk(75) = ChangeWStringToTString(strArr(i))
         End If
    '2025/3/10 END
    'add by nickc 2005/12/06
    Case 111
        StrOk(52) = strArr(i)
        If StrOk(52) = "Y" Then
            lbl1(0).ForeColor = &HFF&
        Else
            lbl1(0).ForeColor = &H80000012
        End If
    'add by nickc 2005/12/27
    Case 112
        StrOk(53) = strArr(i)
    'Add by Morgan 2007/10/29
    Case 122
      StrOk(54) = strArr(i)
   'Add by Morgan 2008/1/17
    Case 123
      StrOk(55) = strArr(i)
    Case 124
      StrOk(56) = strArr(i)
    'Add by Morgan 2008/5/26
    Case 126
      StrOk(57) = strArr(i)
   'Add by Morgan 2008/1/17
    Case 115, 116, 117, 118
      StrOkTxt(i - 69) = strArr(i)
    'Add by Morgan 2008/8/12
    Case 127
      StrOk(13) = PUB_GetContact(strArr(1), strArr(i))
    'add by Toni 2008/10/21
    Case 128
      StrOk(58) = strArr(i)
    'end 2008/10/21
    '2008/12/9 add by sonia
    Case 132
      StrOk(59) = strArr(i)
    '2008/12/9 end
    'Add By Sindy 2011/1/14
    Case 145
      StrOk(60) = strArr(i)
    '2011/1/14 End
    'Add By Sindy 2011/3/4
    Case 146
      StrOkTxt(51) = strArr(i)
    Case 147
      'StrOk(63) = strArr(i)
      If Trim(strArr(i)) <> "" Then
         If ClsLawLawGetName(strArr(i), strTmp) = True Then
            StrOk(63) = strArr(i) + "  " + strTmp
         Else
            'modify by sonia 2020/7/29 抓不到檔案也要顯示原編號
            'StrOk(63) = ""
            StrOk(63) = strArr(i) + ""
         End If
      Else
         StrOk(63) = ""
      End If
      CheckOC
    Case 148
      StrOk(61) = strArr(i)
    Case 149
      StrOk(62) = strArr(i)
    Case 150
      StrOkTxt(50) = strArr(i)
    Case 151
      'StrOk(64) = strArr(i)
      If Trim(strArr(i)) <> "" Then
         If ClsLawLawGetName(strArr(i), strTmp) = True Then
            StrOk(64) = strArr(i) + "  " + strTmp
         Else
            StrOk(64) = ""
         End If
      Else
         StrOk(64) = ""
      End If
      CheckOC
    Case 152
      'StrOk(65) = strArr(i)
      If Trim(strArr(i)) <> "" Then
         If ClsLawLawGetName(strArr(i), strTmp) = True Then
                StrOk(65) = strArr(i) + "  " + strTmp
         Else
            StrOk(65) = ""
         End If
      Else
         StrOk(65) = ""
      End If
      CheckOC
    '2011/3/4 End
    'Add By Sindy 2011/3/17
    Case 153
      StrOk(66) = strArr(i)
    '2011/3/17 End
    'Add By Sindy 2013/8/26
    Case 139
      StrOk(67) = strArr(i)
    '2013/8/26 End
    'Add By Sindy 2013/11/20
    'Memo by Lydia 2019/05/17 預定收款日放寬月數(隱藏)
    Case 143
      StrOk(68) = strArr(i)
    '2013/11/20 End
    'Add By Sindy 2013/12/17
    Case 144
      StrOk(69) = strArr(i)
      LblCU144 = ShowLblCU144(StrOk(69)) 'Add By Sindy 2023/9/4
    '2013/12/17 End
    'Added by Lydia 2022/12/20  改成「FCP提申急件預設組別」
    Case 154
      Combo4.ListIndex = 0
      If Not IsNull(strArr(i)) And Trim(strArr(i)) <> "" Then
         Combo4.ListIndex = strArr(i)
      End If
    'end 2022/12/20
    'Add By Sindy 2013/1/30
    Case 156
      Combo3(0).ListIndex = 0
      If Not IsNull(strArr(i)) And Trim(strArr(i)) <> "" Then
         Combo3(0).ListIndex = strArr(i)
      End If
    Case 157
      Combo3(1).ListIndex = 0
      If Not IsNull(strArr(i)) And Trim(strArr(i)) <> "" Then
         Combo3(1).ListIndex = strArr(i)
      End If
    '2013/1/30 End
    'Add by Amy 2025/11/03 +不提供ID
    Case 182
      ChkID.Value = vbUnchecked
      If strArr(i) = "Y" Then
         ChkID.Value = vbChecked
      End If
   'end 2025/11/03
    'Added by Morgan 2016/12/8
    Case 166 '國內副本收件人
      StrOk(70) = strArr(i)
      If Trim(strArr(i)) <> "" Then
         If ClsLawLawGetName(strArr(i), strTmp) = True Then
            StrOk(70) = StrOk(70) + "  " + strTmp
         End If
      End If
      CheckOC

    Case 167 '國內副本接洽人
      If strArr(166) <> "" And strArr(i) <> "" Then
         StrOk(71) = PUB_GetContact(strArr(166), strArr(i))
      Else
         StrOk(71) = ""
      End If
    'end 2016/12/8
    'Added by Lydia 2019/05/17 客戶特殊付款週期
    Case 175
         StrOk(73) = strArr(i)
    'Added by Lydia 2017/11/30 FCP是否電子送件
    Case 174
         StrOk(72) = strArr(i)
    'Added by Morgan 2018/10/29
    Case 176 'E化客戶指定信箱-正本
         StrOkTxt(52) = strArr(i)
   'Added by Morgan 2021/10/15
    Case 185 'E化客戶指定信箱-副本
         StrOkTxt(53) = strArr(i)
    Case 186 'E化客戶特殊設定
         StrOkTxt(56) = strArr(i)
         'Added by Moran 2025/2/27
         If Trim(strArr(i)) <> "" Then
            arrID = Split(strArr(i), ",")
            For intI = LBound(arrID) To UBound(arrID)
               If arrID(intI) <> "" Then
                  ChkCU186(Val(arrID(intI))).Value = 1
               End If
            Next intI
         End If
         'end 2025/2/27
    Case 187 'E化客戶商標指定信箱-正本
         StrOkTxt(54) = strArr(i)
    Case 188 'E化客戶商標指定信箱-副本
         StrOkTxt(55) = strArr(i)
    'end 2021/10/15
    Case 191 'Add by Amy 2023/05/10 跨所同意主管
        StrOkTxt(57) = strArr(i)
    'Added by Lydia 2024/01/11
    Case 199
        StrOk(74) = strArr(i)
    'end 2024/01/11
    'Add by Sindy 2025/1/6
    Case 201
         If Trim(strArr(i)) <> "" Then
            arrID = Split(strArr(i), ",")
            For intI = UBound(arrID) To LBound(arrID) Step -1
               Chk1K(Val(arrID(intI)) - 1).Value = 1
            Next intI
         End If
    '2025/1/6 END
    Case Else
    End Select
Next i
For i = 0 To UBound(StrOk)
    lbl1(i) = StrOk(i)
Next i
For i = 0 To UBound(StrOkTxt)
    txt1(i) = StrOkTxt(i)
Next i
'Add by Amy 2015/08/24 +cu160~165
If 案件預設收據公司別啟用日 <= Val(strSrvDate(1)) Then
    For i = 0 To 5
        lblCU16X(i).Caption = strArr(i + 160)
        lblCU16X(i).BackColor = &H8000000F
    Next i
Else
    For i = 0 To 5
        lblCU16X(i).Visible = False
    Next i
End If
'end 2015/08/24

'Add by Morgan 2008/11/13 改index與欄位序次相同的陣列，將來再新增欄位時只需加畫面的物件並指定相同的index就好
For Each oLbl In lblCU
   oLbl = strArr(oLbl.Index)
   oLbl.BackColor = &H8000000F
Next
'Modified by Lydia 2021/12/13 改為Form 2.0元件
PUB_SetUserList lstDeveloper, strArr(129), True
'end 2008/11/13

'add by nickc 2007/12/28 加入是否為特殊客戶
If Cu121 = "Y" Then
    strSql = "select * from custspeciallog where cl01='" & Str01 & "' and (cl02,cl03)  in (select cl02,max(cl03) from custspeciallog where cl01='" & Str01 & "' and cl02 in ( select max(cl02) from custspeciallog where cl01='" & Str01 & "'  ) group by cl02) "
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 Then
        SpecCU.Caption = "此為特殊客戶，原因為：" & CheckStr(adoRecordset.Fields("cl04"))
    End If
End If

'Added by Lydia 2023/01/19 往來紀錄中有「A14客戶名稱資訊不得宣傳」者，在申請人/代理人資料查詢首頁提示
strSql = "SELECT ac03 as memo FROM allcode where AC01='11' and ac02='A14' and exists (select * from contactrecord where instr(cr05,'A14')>0 and substr(cr03,1,8)='" & Mid(Str01, 1, 8) & "' and substr(cr03,9,1)='" & Mid(Str01, 9, 1) & "') "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
    SpecCU.Caption = SpecCU.Caption & IIf(Trim(SpecCU.Caption) <> "", "；", "") & adoRecordset.Fields("memo")
    SpecCU.Font.Size = 14
    SpecCU.AutoSize = True
End If
'end 2023/01/19

'Add By Sindy 2012/10/2
If PUB_ChkCustWebIDUserRights(lbl1(0), strUserNum) = True Then
   CmdOk1(3).Visible = True
Else
   CmdOk1(3).Visible = False
   'Modified by Lydia 2020/09/18
   'CmdOk1(5).Left = 4630 'Add by Amy 2017/01/17
   'Modify by Amy 2022/12/05
   'Modify by Amy 2022/12/08 各項指示放最左邊-外專(Morgan通知)
   CmdOk1(5).Left = 6075 '原:4980
   CmdOk1(4).Left = 4292 '原:4020
   CmdOk1(6).Left = 5175 '被介紹者
   'end 2022/12/05
   'end 2020/09/18
End If
'2012/10/2 End
'Add by Amy 2022/12/05
If strSrvDate(1) >= 代理人來源啟用日 Then
    CmdOk1(6).Visible = True
    CmdOk1(6).BackColor = &H8000000F
    If Pub_GetXYSource(2, Left(lbl1(0), 8)) = True Then
        CmdOk1(6).BackColor = &HFFFF80
    End If
End If
End Sub

Private Sub Form_Load()
   Dim objTxt As Object 'Add by Amy 2015/07/24
   
   bolToEndByNick = False
   MoveFormToCenter Me
   If bolFNation = False Then
      'tabCustomer.TabVisible(4) = False 'Mark by Lydia 2021/12/15 已在外層控制權限，所以不用限制顯示
      Label1(3).Visible = False
      lbl1(9).Visible = False
   End If
   '92.04.16 nick
   cmdState = -1
   
   'Added by Morgan 2025/10/16
   If Pub_StrUserSt03 <> "M51" Then
      lblECustXSet.Visible = False
      ChkCU186(1).Visible = False
      ChkCU186(2).Visible = False
   End If
   'end 2025/10/16
   
   'Add By Sindy 2013/12/18
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M31" Then
      Label1(26).Visible = True
      lbl1(69).Visible = True
   Else
      Label1(26).Visible = False
      lbl1(69).Visible = False
   End If
   '2013/12/18 END
   'Add by Amy 2015/07/24
   If 案件預設收據公司別啟用日 >= Val(strSrvDate(1)) Then
        For Each objTxt In Me.lblComp
            objTxt.Visible = False
        Next
        For Each objTxt In Me.lblCU16X
            objTxt.Visible = False
        Next
   End If
   'end 2015/07/24
   
   'Added by Lydia 2020/03/31 事務所合併日起台灣案取消(1:專利商標 2:專利法律) 的標題，非台灣案改標題為(J:智權公司 空白:系統預設)。
   If strSrvDate(1) >= 事務所合併日 Then
       For intI = 0 To 5
          Select Case intI
              Case 0, 2, 4  '台灣案:CU160,CU162,CU164
                  'Modifed by Lydia 2021/07/13 debug-統一改標題為(J:智權公司 空白:系統預設)
                  'lblComp(intI).Visible = False
                  'lblCU16X(intI).Visible = False
                  lblComp(intI).Caption = Replace(lblComp(intI).Caption, "1：專利商標 2：專利法律", "J：智權公司 空白:系統預設")
                  'end 2021/07/13
              Case 1, 3, 5  '非台灣案:CU161,CU163,CU165
                  lblComp(intI).Caption = Replace(lblComp(intI).Caption, "1：專利商標 2：專利法律 J：台一智權", "J：智權公司 空白:系統預設")
          End Select

       Next
   End If
   'end 2020/03/30
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      CmdOk1(4).Visible = True
   Else
      CmdOk1(4).Visible = False
      'Mark by Lydia 2020/09/18 按鈕移到最上方
      'txt1(43).Top = 360
      'txt1(43).Height = 4740
      'end 2020/09/18
   End If
   'end 2020/05/05
   
   'Added by Lydia 2021/12/15 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstDeveloper.Height = 720
   lstDeveloper.Width = 1500
   
   'Added by Lydia 2023/03/03 外專新案認領
   If strSrvDate(1) >= 外專新案認領啟用日 Then
      Label1(27).Visible = True
      Combo4.Visible = True
   End If
   'Memo by Amy 2025/02/11 FCT註冊費自動代繳移位,切其他頁籤會有殘影搬至按頁籤
   
   Frame1K.BorderStyle = 0 'Add By Sindy 2025/1/6
   
   tabCustomer.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100101_11 = Nothing
End Sub

Private Sub tabCustomer_Click(PreviousTab As Integer)
   'Add by Amy 2025/02/11 從Form_Load搬過來,否則切其他頁籤會有殘影
   If tabCustomer.Tab = 5 Then
      'Add by Amy 2024/03/08 隱藏延展單筆不跑,將FCT註冊費自動代繳移位
      'tabCustomer.Tab = 5 'Add By Sindy 2024/12/12 下列欄位移動位置才不會亂掉(在別的Tab會看到)
      Label80(21).Left = 180
      lbl1(58).Left = 2000
      '2024/03/08 END
   End If
End Sub

'Added by Lydia 2016/10/29 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
End Sub
