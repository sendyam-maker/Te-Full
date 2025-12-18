VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利案件基本資料查詢"
   ClientHeight    =   6480
   ClientLeft      =   216
   ClientTop       =   1464
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9312
   Begin VB.CommandButton cmdOK 
      Caption         =   "各項指示"
      Height          =   330
      Index           =   11
      Left            =   30
      TabIndex        =   349
      Top             =   30
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度"
      Height          =   330
      Index           =   12
      Left            =   5439
      TabIndex        =   8
      Top             =   30
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申5"
      Height          =   330
      Index           =   10
      Left            =   3438
      TabIndex        =   5
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申4"
      Height          =   330
      Index           =   9
      Left            =   3021
      TabIndex        =   4
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申3"
      Height          =   330
      Index           =   8
      Left            =   2604
      TabIndex        =   3
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申2"
      Height          =   330
      Index           =   7
      Left            =   2187
      TabIndex        =   2
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "已設定代表圖"
      Height          =   330
      Index           =   6
      Left            =   930
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   30
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "分割案"
      Height          =   330
      Index           =   5
      Left            =   3855
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "相關卷號"
      Height          =   330
      Index           =   4
      Left            =   4542
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人"
      Height          =   330
      Index           =   3
      Left            =   6216
      TabIndex        =   9
      Top             =   30
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人"
      Height          =   330
      Index           =   2
      Left            =   6993
      TabIndex        =   10
      Top             =   30
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      Height          =   330
      Index           =   1
      Left            =   8550
      TabIndex        =   11
      Top             =   30
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      Height          =   330
      Index           =   0
      Left            =   7770
      TabIndex        =   0
      Top             =   30
      Width           =   765
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6045
      Left            =   30
      TabIndex        =   61
      Top             =   390
      Width           =   9255
      _ExtentX        =   16341
      _ExtentY        =   10668
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      Tab             =   2
      TabsPerRow      =   11
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm100101_3.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(1)=   "Label30"
      Tab(0).Control(2)=   "Label29"
      Tab(0).Control(3)=   "Label8(0)"
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(5)=   "Label1(15)"
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(7)=   "Label1(10)"
      Tab(0).Control(8)=   "Label11(0)"
      Tab(0).Control(9)=   "Label19(0)"
      Tab(0).Control(10)=   "Label20(0)"
      Tab(0).Control(11)=   "Label21(1)"
      Tab(0).Control(12)=   "Label31(1)"
      Tab(0).Control(13)=   "Label33(0)"
      Tab(0).Control(14)=   "lbl1(16)"
      Tab(0).Control(15)=   "lbl1(20)"
      Tab(0).Control(16)=   "lbl1(21)"
      Tab(0).Control(17)=   "lbl1(22)"
      Tab(0).Control(18)=   "lbl1(23)"
      Tab(0).Control(19)=   "lbl1(24)"
      Tab(0).Control(20)=   "lbl1(25)"
      Tab(0).Control(21)=   "Label3"
      Tab(0).Control(22)=   "lbl1(26)"
      Tab(0).Control(23)=   "lbl1(14)"
      Tab(0).Control(24)=   "lbl1(13)"
      Tab(0).Control(25)=   "lbl1(11)"
      Tab(0).Control(26)=   "lbl1(10)"
      Tab(0).Control(27)=   "lbl1(9)"
      Tab(0).Control(28)=   "lbl1(8)"
      Tab(0).Control(29)=   "lbl1(7)"
      Tab(0).Control(30)=   "lbl1(1)"
      Tab(0).Control(31)=   "Label35"
      Tab(0).Control(32)=   "Label32"
      Tab(0).Control(33)=   "Label31(0)"
      Tab(0).Control(34)=   "Label21(0)"
      Tab(0).Control(35)=   "Label18(0)"
      Tab(0).Control(36)=   "Label9(0)"
      Tab(0).Control(37)=   "Label1(8)"
      Tab(0).Control(38)=   "Label1(2)"
      Tab(0).Control(39)=   "Label1(9)"
      Tab(0).Control(40)=   "Label1(6)"
      Tab(0).Control(41)=   "Label1(3)"
      Tab(0).Control(42)=   "Label1(13)"
      Tab(0).Control(43)=   "Label8(6)"
      Tab(0).Control(44)=   "Label1(7)"
      Tab(0).Control(45)=   "Label8(5)"
      Tab(0).Control(46)=   "Label22"
      Tab(0).Control(47)=   "Label23"
      Tab(0).Control(48)=   "lbl1(57)"
      Tab(0).Control(49)=   "Label33(1)"
      Tab(0).Control(50)=   "lbl1(70)"
      Tab(0).Control(51)=   "Label1(1)"
      Tab(0).Control(52)=   "lbl1(15)"
      Tab(0).Control(53)=   "Label2(0)"
      Tab(0).Control(54)=   "Label5"
      Tab(0).Control(55)=   "lbl1(71)"
      Tab(0).Control(56)=   "lbl1(12)"
      Tab(0).Control(57)=   "Label28"
      Tab(0).Control(58)=   "Label26(1)"
      Tab(0).Control(59)=   "lbl1(86)"
      Tab(0).Control(60)=   "lblFilingDate(0)"
      Tab(0).Control(61)=   "lblFilingDate(1)"
      Tab(0).Control(62)=   "Label19(1)"
      Tab(0).Control(63)=   "lbl1(90)"
      Tab(0).Control(64)=   "lbl1(91)"
      Tab(0).Control(65)=   "Label40"
      Tab(0).Control(66)=   "lbl1(160)"
      Tab(0).Control(67)=   "Label1(11)"
      Tab(0).Control(68)=   "lbl1(140)"
      Tab(0).Control(69)=   "Label1(12)"
      Tab(0).Control(70)=   "Label1(14)"
      Tab(0).Control(71)=   "lbl1(164)"
      Tab(0).Control(72)=   "lblCaseMap"
      Tab(0).Control(73)=   "lblCMboth"
      Tab(0).Control(74)=   "lblCaseMap2"
      Tab(0).Control(75)=   "lblPA174"
      Tab(0).Control(76)=   "lblPA176"
      Tab(0).Control(77)=   "lbl1(176)"
      Tab(0).Control(78)=   "txt1(81)"
      Tab(0).Control(79)=   "txt1(82)"
      Tab(0).Control(80)=   "txt1(83)"
      Tab(0).Control(81)=   "txt1(84)"
      Tab(0).Control(82)=   "txt1(85)"
      Tab(0).Control(83)=   "txt1(86)"
      Tab(0).Control(84)=   "txt1(23)"
      Tab(0).Control(85)=   "txt1(24)"
      Tab(0).Control(86)=   "txt1(25)"
      Tab(0).Control(87)=   "Label45"
      Tab(0).Control(88)=   "Label46"
      Tab(0).Control(89)=   "lblPA(178)"
      Tab(0).Control(90)=   "Label43"
      Tab(0).Control(91)=   "lbl1(179)"
      Tab(0).Control(92)=   "CmdPA174"
      Tab(0).ControlCount=   93
      TabCaption(1)   =   "申請人/代理人"
      TabPicture(1)   =   "frm100101_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(0)"
      Tab(1).Control(1)=   "txt1(1)"
      Tab(1).Control(2)=   "txt1(2)"
      Tab(1).Control(3)=   "txt1(3)"
      Tab(1).Control(4)=   "txt1(7)"
      Tab(1).Control(5)=   "txt1(8)"
      Tab(1).Control(6)=   "txt1(9)"
      Tab(1).Control(7)=   "txt1(10)"
      Tab(1).Control(8)=   "txt1(11)"
      Tab(1).Control(9)=   "txt1(14)"
      Tab(1).Control(10)=   "txt1(15)"
      Tab(1).Control(11)=   "txt1(16)"
      Tab(1).Control(12)=   "txt1(17)"
      Tab(1).Control(13)=   "txt1(18)"
      Tab(1).Control(14)=   "txt1(4)"
      Tab(1).Control(15)=   "Label11(13)"
      Tab(1).Control(16)=   "lbl1(169)"
      Tab(1).Control(17)=   "Label11(2)"
      Tab(1).Control(18)=   "lbl1(168)"
      Tab(1).Control(19)=   "lbl1(85)"
      Tab(1).Control(20)=   "Label11(12)"
      Tab(1).Control(21)=   "lbl1(74)"
      Tab(1).Control(22)=   "Label24"
      Tab(1).Control(23)=   "lbl1(73)"
      Tab(1).Control(24)=   "Label10"
      Tab(1).Control(25)=   "lbl1(72)"
      Tab(1).Control(26)=   "Label7"
      Tab(1).Control(27)=   "lbl1(54)"
      Tab(1).Control(28)=   "lbl1(53)"
      Tab(1).Control(29)=   "Label11(5)"
      Tab(1).Control(30)=   "Label44"
      Tab(1).Control(31)=   "lbl1(40)"
      Tab(1).Control(32)=   "Label11(3)"
      Tab(1).Control(33)=   "lbl1(36)"
      Tab(1).Control(34)=   "lbl1(35)"
      Tab(1).Control(35)=   "lbl1(34)"
      Tab(1).Control(36)=   "lbl1(33)"
      Tab(1).Control(37)=   "lbl1(32)"
      Tab(1).Control(38)=   "lbl1(31)"
      Tab(1).Control(39)=   "lbl1(30)"
      Tab(1).Control(40)=   "lbl1(29)"
      Tab(1).Control(41)=   "lbl1(28)"
      Tab(1).Control(42)=   "Label42"
      Tab(1).Control(43)=   "Label37"
      Tab(1).Control(44)=   "Label12(1)"
      Tab(1).Control(45)=   "Label12(0)"
      Tab(1).Control(46)=   "Label13(0)"
      Tab(1).Control(47)=   "Label13(1)"
      Tab(1).Control(48)=   "Label16(1)"
      Tab(1).Control(49)=   "Label17"
      Tab(1).Control(50)=   "Label12(7)"
      Tab(1).Control(51)=   "Label12(8)"
      Tab(1).Control(52)=   "Label12(9)"
      Tab(1).Control(53)=   "Label11(1)"
      Tab(1).Control(54)=   "Label13(2)"
      Tab(1).Control(55)=   "Label13(3)"
      Tab(1).Control(56)=   "Label13(4)"
      Tab(1).ControlCount=   57
      TabCaption(2)   =   "FC資料"
      TabPicture(2)   =   "frm100101_3.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label49(4)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lbl1(181)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label11(14)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label41"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label18(3)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label18(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label11(9)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lbl1(83)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label49(1)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label50"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label49(0)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label11(6)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label18(1)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label11(4)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label48(0)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label21(2)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label20(1)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "lbl1(38)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "lbl1(42)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "lbl1(44)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "lbl1(55)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Label11(7)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Label48(1)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "lbl1(59)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "lbl1(60)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Label51(4)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Label49(2)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Label1(156)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "lbl1(80)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "lbl1(56)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "lbl1(41)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "lbl1(46)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "lbl1(58)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Label11(8)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "lbl1(82)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "lblPA(151)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "lblPA(152)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "lbl1(159)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "lbl1(167)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "lblPA(69)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Label11(15)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "lblPA(156)"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Label49(3)"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "lbl1(177)"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "txt1(21)"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "txt1(22)"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).ControlCount=   47
      TabCaption(3)   =   "聯絡人"
      TabPicture(3)   =   "frm100101_3.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txt1(55)"
      Tab(3).Control(1)=   "txt1(54)"
      Tab(3).Control(2)=   "Label12(11)"
      Tab(3).Control(3)=   "lbl1(92)"
      Tab(3).Control(4)=   "lbl1(43)"
      Tab(3).Control(5)=   "lbl1(37)"
      Tab(3).Control(6)=   "lbl1(61)"
      Tab(3).Control(7)=   "lbl1(62)"
      Tab(3).Control(8)=   "lbl1(63)"
      Tab(3).Control(9)=   "lbl1(64)"
      Tab(3).Control(10)=   "Label6(7)"
      Tab(3).Control(11)=   "Label6(6)"
      Tab(3).Control(12)=   "lbl1(69)"
      Tab(3).Control(13)=   "Label6(5)"
      Tab(3).Control(14)=   "lbl1(68)"
      Tab(3).Control(15)=   "Label6(4)"
      Tab(3).Control(16)=   "lbl1(67)"
      Tab(3).Control(17)=   "lbl1(66)"
      Tab(3).Control(18)=   "lbl1(65)"
      Tab(3).Control(19)=   "Label6(3)"
      Tab(3).Control(20)=   "Label6(2)"
      Tab(3).Control(21)=   "Label6(1)"
      Tab(3).Control(22)=   "Label12(10)"
      Tab(3).Control(23)=   "Label12(6)"
      Tab(3).Control(24)=   "Label12(3)"
      Tab(3).Control(25)=   "Label12(2)"
      Tab(3).Control(26)=   "lbl1(39)"
      Tab(3).Control(27)=   "lbl1(45)"
      Tab(3).Control(28)=   "Label12(5)"
      Tab(3).Control(29)=   "Label12(4)"
      Tab(3).Control(30)=   "Label6(0)"
      Tab(3).Control(31)=   "Label15"
      Tab(3).ControlCount=   32
      TabCaption(4)   =   "繳年費"
      TabPicture(4)   =   "frm100101_3.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdDataList1"
      Tab(4).Control(1)=   "Label34"
      Tab(4).Control(2)=   "lbl1(76)"
      Tab(4).Control(3)=   "lbl1(75)"
      Tab(4).Control(4)=   "Label27"
      Tab(4).Control(5)=   "Label25"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "優先權"
      TabPicture(5)   =   "frm100101_3.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grdDataList2"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "發明人"
      TabPicture(6)   =   "frm100101_3.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "GRD1"
      Tab(6).Control(1)=   "lblInventor(32)"
      Tab(6).Control(2)=   "lblInventor(31)"
      Tab(6).Control(3)=   "lblInventor(30)"
      Tab(6).Control(4)=   "lblInventor(29)"
      Tab(6).Control(5)=   "lblInventor(28)"
      Tab(6).Control(6)=   "lblInventor(27)"
      Tab(6).Control(7)=   "lblInventor(26)"
      Tab(6).Control(8)=   "lblInventor(25)"
      Tab(6).Control(9)=   "lblInventor(24)"
      Tab(6).Control(10)=   "Label14(4)"
      Tab(6).Control(11)=   "Label52(0)"
      Tab(6).Control(12)=   "Label52(1)"
      Tab(6).Control(13)=   "lbl1(47)"
      Tab(6).Control(14)=   "lbl1(48)"
      Tab(6).Control(15)=   "lbl1(49)"
      Tab(6).Control(16)=   "lbl1(50)"
      Tab(6).Control(17)=   "lbl1(51)"
      Tab(6).Control(18)=   "lbl1(52)"
      Tab(6).ControlCount=   19
      TabCaption(7)   =   "代表人"
      TabPicture(7)   =   "frm100101_3.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "txt1(56)"
      Tab(7).Control(1)=   "txt1(79)"
      Tab(7).Control(2)=   "txt1(78)"
      Tab(7).Control(3)=   "txt1(77)"
      Tab(7).Control(4)=   "txt1(76)"
      Tab(7).Control(5)=   "txt1(75)"
      Tab(7).Control(6)=   "txt1(74)"
      Tab(7).Control(7)=   "txt1(73)"
      Tab(7).Control(8)=   "txt1(72)"
      Tab(7).Control(9)=   "txt1(71)"
      Tab(7).Control(10)=   "txt1(70)"
      Tab(7).Control(11)=   "txt1(69)"
      Tab(7).Control(12)=   "txt1(68)"
      Tab(7).Control(13)=   "txt1(67)"
      Tab(7).Control(14)=   "txt1(66)"
      Tab(7).Control(15)=   "txt1(65)"
      Tab(7).Control(16)=   "txt1(64)"
      Tab(7).Control(17)=   "txt1(63)"
      Tab(7).Control(18)=   "txt1(62)"
      Tab(7).Control(19)=   "txt1(61)"
      Tab(7).Control(20)=   "txt1(60)"
      Tab(7).Control(21)=   "txt1(59)"
      Tab(7).Control(22)=   "txt1(58)"
      Tab(7).Control(23)=   "txt1(57)"
      Tab(7).Control(24)=   "txt1(5)"
      Tab(7).Control(25)=   "txt1(6)"
      Tab(7).Control(26)=   "txt1(12)"
      Tab(7).Control(27)=   "txt1(13)"
      Tab(7).Control(28)=   "txt1(19)"
      Tab(7).Control(29)=   "txt1(20)"
      Tab(7).Control(30)=   "Label13(14)"
      Tab(7).Control(31)=   "Label13(13)"
      Tab(7).Control(32)=   "Label13(12)"
      Tab(7).Control(33)=   "Label13(11)"
      Tab(7).Control(34)=   "Label13(10)"
      Tab(7).Control(35)=   "Label13(9)"
      Tab(7).Control(36)=   "Label13(8)"
      Tab(7).Control(37)=   "Label13(6)"
      Tab(7).Control(38)=   "Label13(5)"
      Tab(7).Control(39)=   "Label13(7)"
      Tab(7).ControlCount=   40
      TabCaption(8)   =   "銷卷"
      TabPicture(8)   =   "frm100101_3.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label39"
      Tab(8).Control(1)=   "lbl1(79)"
      Tab(8).Control(2)=   "Label38"
      Tab(8).Control(3)=   "lbl1(78)"
      Tab(8).Control(4)=   "Label36"
      Tab(8).Control(5)=   "lbl1(77)"
      Tab(8).Control(6)=   "Label4"
      Tab(8).Control(7)=   "lbl1(27)"
      Tab(8).ControlCount=   8
      TabCaption(9)   =   "其他"
      TabPicture(9)   =   "frm100101_3.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame1K"
      Tab(9).Control(1)=   "cmdDivSug"
      Tab(9).Control(2)=   "Combo3(1)"
      Tab(9).Control(3)=   "lstPA166"
      Tab(9).Control(4)=   "txt1(80)"
      Tab(9).Control(5)=   "lblPA(61)"
      Tab(9).Control(6)=   "lblPA(60)"
      Tab(9).Control(7)=   "lblPA(163)"
      Tab(9).Control(8)=   "lblPA(162)"
      Tab(9).Control(9)=   "Label1(175)"
      Tab(9).Control(10)=   "Label1(174)"
      Tab(9).Control(11)=   "Label1(67)"
      Tab(9).Control(12)=   "Label1(68)"
      Tab(9).Control(13)=   "Label11(11)"
      Tab(9).Control(14)=   "lbl1(84)"
      Tab(9).Control(15)=   "Label1(165)"
      Tab(9).Control(16)=   "Label1(164)"
      Tab(9).Control(17)=   "lbl1(87)"
      Tab(9).Control(18)=   "lbl1(88)"
      Tab(9).Control(19)=   "lbl1(81)"
      Tab(9).Control(20)=   "lblPA(161)"
      Tab(9).Control(21)=   "lbl1(89)"
      Tab(9).Control(22)=   "Label1(177)"
      Tab(9).Control(23)=   "Label80(29)"
      Tab(9).Control(24)=   "lbl1(170)"
      Tab(9).Control(25)=   "Label80(26)"
      Tab(9).Control(26)=   "lblPA(64)"
      Tab(9).Control(27)=   "Label80(0)"
      Tab(9).Control(28)=   "Label1(73)"
      Tab(9).Control(29)=   "Label1(72)"
      Tab(9).Control(30)=   "Label1(71)"
      Tab(9).Control(31)=   "Label1(70)"
      Tab(9).Control(32)=   "Label1(69)"
      Tab(9).Control(33)=   "lblPA(65)"
      Tab(9).Control(34)=   "lblPA(66)"
      Tab(9).Control(35)=   "lblPA(67)"
      Tab(9).Control(36)=   "lblPA(68)"
      Tab(9).Control(37)=   "Label1(16)"
      Tab(9).Control(38)=   "lblTot6"
      Tab(9).Control(39)=   "Label1(75)"
      Tab(9).Control(40)=   "lblPA(172)"
      Tab(9).Control(41)=   "Label1(17)"
      Tab(9).Control(42)=   "lblPA(173)"
      Tab(9).Control(43)=   "Label1(18)"
      Tab(9).Control(44)=   "Label1(27)"
      Tab(9).Control(45)=   "Label68"
      Tab(9).Control(46)=   "Label1(166)"
      Tab(9).Control(47)=   "Label11(10)"
      Tab(9).ControlCount=   48
      TabCaption(10)  =   "參考備註"
      TabPicture(10)  =   "frm100101_3.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "txt1(53)"
      Tab(10).ControlCount=   1
      Begin VB.Frame Frame1K 
         Enabled         =   0   'False
         Height          =   280
         Left            =   -75060
         TabIndex        =   363
         Top             =   4650
         Width           =   4870
         Begin VB.CheckBox Chk1K 
            Caption         =   "帳單另寄"
            Height          =   180
            Index           =   0
            Left            =   1740
            TabIndex        =   366
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "上傳平台"
            Height          =   180
            Index           =   1
            Left            =   2790
            TabIndex        =   365
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "月帳單"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   364
            Top             =   60
            Width           =   910
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "請款單寄送類型："
            Height          =   180
            Index           =   34
            Left            =   150
            TabIndex        =   367
            Top             =   60
            Width           =   1440
         End
      End
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   -75000
         Style           =   1  '圖片外觀
         TabIndex        =   348
         Top             =   2010
         Width           =   800
      End
      Begin VB.CommandButton cmdDivSug 
         Caption         =   "分割建議"
         Height          =   315
         Left            =   -71250
         TabIndex        =   343
         Top             =   2790
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         Index           =   1
         ItemData        =   "frm100101_3.frx":0134
         Left            =   -67920
         List            =   "frm100101_3.frx":0147
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   298
         Top             =   1080
         Width           =   1470
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList2 
         Height          =   5565
         Left            =   -74610
         TabIndex        =   63
         Top             =   360
         Width           =   8505
         _ExtentX        =   15007
         _ExtentY        =   9800
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   3
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList1 
         Height          =   5445
         Left            =   -71520
         TabIndex        =   62
         Top             =   300
         Width           =   5655
         _ExtentX        =   9991
         _ExtentY        =   9589
         _Version        =   393216
         Rows            =   21
         Cols            =   5
         FocusRect       =   2
         AllowUserResizing=   1
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   3795
         Left            =   -74850
         TabIndex        =   286
         Top             =   600
         Width           =   8745
         _ExtentX        =   15409
         _ExtentY        =   6710
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "V|發明人編號|中文名稱|英文名稱|日文名稱"
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
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   179
         Left            =   -67440
         TabIndex        =   362
         Top             =   4425
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
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "個體身份:"
         Height          =   180
         Left            =   -68280
         TabIndex        =   361
         Top             =   4440
         Width           =   765
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   178
         Left            =   -67440
         TabIndex        =   360
         Top             =   4155
         Width           =   345
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "609;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   " (1:電子 2:紙本)"
         Height          =   180
         Left            =   -67080
         TabIndex        =   359
         Top             =   4155
         Width           =   1200
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "證書形式:"
         Height          =   180
         Left            =   -68280
         TabIndex        =   358
         Top             =   4155
         Width           =   765
      End
      Begin MSForms.ListBox lstPA166 
         Height          =   600
         Left            =   -73080
         TabIndex        =   357
         TabStop         =   0   'False
         Top             =   4020
         Width           =   4365
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "7699;1057"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   25
         Left            =   -74010
         TabIndex        =   352
         Top             =   2641
         Width           =   1185
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2090;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   24
         Left            =   -70470
         TabIndex        =   351
         Top             =   2975
         Width           =   2085
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3678;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   23
         Left            =   -70470
         TabIndex        =   350
         Top             =   2641
         Width           =   2085
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3678;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   640
         Index           =   80
         Left            =   -73980
         TabIndex        =   299
         Top             =   390
         Width           =   7995
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14102;1129"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   5445
         Index           =   53
         Left            =   -74940
         TabIndex        =   297
         Top             =   360
         Width           =   9045
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "15954;9604"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   86
         Left            =   -74010
         TabIndex        =   281
         Top             =   2307
         Width           =   1185
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2090;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   85
         Left            =   -71100
         TabIndex        =   280
         Top             =   2310
         Width           =   2295
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "4048;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   420
         Index           =   84
         Left            =   -73740
         TabIndex        =   279
         Top             =   1853
         Width           =   7875
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13891;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   420
         Index           =   83
         Left            =   -73740
         TabIndex        =   278
         Top             =   1399
         Width           =   7875
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13891;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   420
         Index           =   82
         Left            =   -73740
         TabIndex        =   277
         Top             =   945
         Width           =   7875
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13891;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   81
         Left            =   -74040
         TabIndex        =   276
         Top             =   322
         Width           =   1815
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3201;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   56
         Left            =   -72912
         TabIndex        =   37
         Top             =   1446
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   525
         Index           =   79
         Left            =   -68205
         TabIndex        =   60
         Top             =   5415
         Width           =   2355
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   525
         Index           =   78
         Left            =   -70560
         TabIndex        =   59
         Top             =   5415
         Width           =   2355
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4154;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   525
         Index           =   77
         Left            =   -72912
         TabIndex        =   58
         Top             =   5415
         Width           =   2355
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4154;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   76
         Left            =   -68208
         TabIndex        =   57
         Top             =   4848
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   75
         Left            =   -70560
         TabIndex        =   56
         Top             =   4848
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   74
         Left            =   -72912
         TabIndex        =   55
         Top             =   4848
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   73
         Left            =   -68208
         TabIndex        =   54
         Top             =   4281
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   72
         Left            =   -70560
         TabIndex        =   53
         Top             =   4281
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   71
         Left            =   -72912
         TabIndex        =   52
         Top             =   4281
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   70
         Left            =   -68208
         TabIndex        =   51
         Top             =   3714
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   69
         Left            =   -70560
         TabIndex        =   50
         Top             =   3714
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   68
         Left            =   -72912
         TabIndex        =   49
         Top             =   3714
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   67
         Left            =   -68208
         TabIndex        =   48
         Top             =   3147
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   66
         Left            =   -70560
         TabIndex        =   47
         Top             =   3147
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   65
         Left            =   -72912
         TabIndex        =   46
         Top             =   3147
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   64
         Left            =   -68208
         TabIndex        =   45
         Top             =   2580
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   63
         Left            =   -70560
         TabIndex        =   44
         Top             =   2580
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   62
         Left            =   -72912
         TabIndex        =   43
         Top             =   2580
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   61
         Left            =   -68208
         TabIndex        =   42
         Top             =   2013
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   60
         Left            =   -70560
         TabIndex        =   41
         Top             =   2013
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   59
         Left            =   -72912
         TabIndex        =   40
         Top             =   2013
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   58
         Left            =   -68208
         TabIndex        =   39
         Top             =   1446
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   57
         Left            =   -70560
         TabIndex        =   38
         Top             =   1446
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   5
         Left            =   -72912
         TabIndex        =   31
         Top             =   312
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   6
         Left            =   -72912
         TabIndex        =   34
         Top             =   879
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   12
         Left            =   -70560
         TabIndex        =   32
         Top             =   312
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   13
         Left            =   -70560
         TabIndex        =   35
         Top             =   879
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   19
         Left            =   -68208
         TabIndex        =   33
         Top             =   312
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   520
         Index           =   20
         Left            =   -68208
         TabIndex        =   36
         Top             =   879
         Width           =   2352
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4149;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   525
         Index           =   55
         Left            =   -72495
         TabIndex        =   30
         Top             =   4830
         Width           =   6630
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11695;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   525
         Index           =   54
         Left            =   -72495
         TabIndex        =   29
         Top             =   4260
         Width           =   6630
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11695;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   0
         Left            =   -72510
         TabIndex        =   12
         Top             =   1685
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   1
         Left            =   -72510
         TabIndex        =   15
         Top             =   2047
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   2
         Left            =   -72510
         TabIndex        =   18
         Top             =   2409
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   345
         Index           =   3
         Left            =   -72510
         TabIndex        =   21
         Top             =   2760
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   7
         Left            =   -70275
         TabIndex        =   13
         Top             =   1685
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   8
         Left            =   -70275
         TabIndex        =   16
         Top             =   2047
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   9
         Left            =   -70275
         TabIndex        =   19
         Top             =   2409
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   10
         Left            =   -70275
         TabIndex        =   22
         Top             =   2762
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   11
         Left            =   -70275
         TabIndex        =   25
         Top             =   3122
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   14
         Left            =   -68040
         TabIndex        =   14
         Top             =   1685
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   15
         Left            =   -68040
         TabIndex        =   17
         Top             =   2047
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   16
         Left            =   -68040
         TabIndex        =   20
         Top             =   2409
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   17
         Left            =   -68040
         TabIndex        =   23
         Top             =   2762
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   340
         Index           =   18
         Left            =   -68040
         TabIndex        =   26
         Top             =   3122
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;600"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   345
         Index           =   4
         Left            =   -72510
         TabIndex        =   24
         Top             =   3120
         Width           =   2160
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "3810;609"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   1092
         Index           =   22
         Left            =   1200
         TabIndex        =   28
         Top             =   3996
         Width           =   7860
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13864;1940"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   780
         Index           =   21
         Left            =   1200
         TabIndex        =   27
         Top             =   2856
         Width           =   7860
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13864;1376"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   177
         Left            =   5112
         TabIndex        =   355
         Top             =   1032
         Width           =   264
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
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "專利連結通知：        (Y:是)"
         Height          =   180
         Index           =   3
         Left            =   3828
         TabIndex        =   356
         Top             =   1032
         Width           =   2088
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   176
         Left            =   -66240
         TabIndex        =   354
         Top             =   660
         Width           =   315
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "556;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPA176 
         Caption         =   "專利權期間延長相關:"
         Height          =   420
         Left            =   -67440
         TabIndex        =   353
         Top             =   600
         Width           =   1000
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -75000
         TabIndex        =   347
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblCaseMap2 
         AutoSize        =   -1  'True
         Caption         =   "lblCaseMap2"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -67260
         TabIndex        =   346
         Top             =   5340
         Width           =   1410
      End
      Begin MSForms.Label lblPA 
         Height          =   252
         Index           =   156
         Left            =   1752
         TabIndex        =   344
         Top             =   1944
         Width           =   264
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "459;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費特殊管制：       (Y:年費續辦:有別於Y / X設定  N:寄證書/二核後年費不續辦  空白:視Y / X設定)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   15
         Left            =   180
         TabIndex        =   345
         Top             =   1944
         Width           =   7896
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   61
         Left            =   -73350
         TabIndex        =   342
         Top             =   3690
         Width           =   280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   60
         Left            =   -72870
         TabIndex        =   341
         Top             =   3420
         Width           =   280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblPA 
         Height          =   252
         Index           =   163
         Left            =   -72672
         TabIndex        =   340
         Top             =   3156
         Width           =   276
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   162
         Left            =   -72900
         TabIndex        =   339
         Top             =   2850
         Width           =   280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否加註核准分割建議：       ( Y：是 N：否)"
         Height          =   180
         Index           =   175
         Left            =   -74895
         TabIndex        =   338
         Top             =   2850
         Width           =   3465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否初審階段提分割/改請：       ( Y：是 N：否)"
         Height          =   180
         Index           =   174
         Left            =   -74892
         TabIndex        =   337
         Top             =   3156
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "一案兩請是否放棄新型：       ( Y：是 N：否)"
         Height          =   180
         Index           =   67
         Left            =   -74895
         TabIndex        =   336
         Top             =   3420
         Width           =   3465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFP有無關聯P案：         (  N：無)"
         Height          =   180
         Index           =   68
         Left            =   -74865
         TabIndex        =   335
         Top             =   3690
         Width           =   2565
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "帳單備註："
         Height          =   180
         Index           =   11
         Left            =   -74895
         TabIndex        =   331
         Top             =   390
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   84
         Left            =   -73320
         TabIndex        =   330
         Top             =   1140
         Width           =   280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Caption         =   "請款單份數："
         Height          =   180
         Index           =   165
         Left            =   -74385
         TabIndex        =   328
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "定稿份數："
         Height          =   180
         Index           =   164
         Left            =   -74205
         TabIndex        =   327
         Top             =   1410
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   87
         Left            =   -73320
         TabIndex        =   326
         Top             =   1410
         Width           =   280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   88
         Left            =   -73320
         TabIndex        =   325
         Top             =   1680
         Width           =   280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   81
         Left            =   -73320
         TabIndex        =   324
         Top             =   1980
         Width           =   280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   161
         Left            =   -73313
         TabIndex        =   323
         Top             =   2580
         Width           =   280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   89
         Left            =   -73320
         TabIndex        =   322
         Top             =   2280
         Width           =   280
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊請款單列印對象："
         Height          =   180
         Index           =   177
         Left            =   -74895
         TabIndex        =   321
         Top             =   4035
         Width           =   1800
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "請款單列印幣別格式："
         Height          =   180
         Index           =   29
         Left            =   -69780
         TabIndex        =   320
         Top             =   1140
         Width           =   1800
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   170
         Left            =   -70710
         TabIndex        =   319
         Top             =   1140
         Width           =   435
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "767;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "請款幣別："
         Height          =   180
         Index           =   26
         Left            =   -71640
         TabIndex        =   318
         Top             =   1140
         Width           =   900
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   64
         Left            =   -67470
         TabIndex        =   317
         Top             =   1410
         Width           =   500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "中文本資訊："
         Height          =   180
         Index           =   0
         Left            =   -69780
         TabIndex        =   316
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "圖式頁數："
         Height          =   180
         Index           =   73
         Left            =   -68670
         TabIndex        =   315
         Top             =   2535
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "申請專利範圍頁數："
         Height          =   180
         Index           =   72
         Left            =   -68670
         TabIndex        =   314
         Top             =   2250
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "序列表："
         Height          =   180
         Index           =   71
         Left            =   -68670
         TabIndex        =   313
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "說明書頁數："
         Height          =   180
         Index           =   70
         Left            =   -68670
         TabIndex        =   312
         Top             =   1695
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "摘要頁數："
         Height          =   180
         Index           =   69
         Left            =   -68670
         TabIndex        =   311
         Top             =   1410
         Width           =   900
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   65
         Left            =   -67470
         TabIndex        =   310
         Top             =   1695
         Width           =   500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   66
         Left            =   -66870
         TabIndex        =   309
         Top             =   1980
         Width           =   500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   67
         Left            =   -66870
         TabIndex        =   308
         Top             =   2250
         Width           =   500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   68
         Left            =   -67470
         TabIndex        =   307
         Top             =   2535
         Width           =   500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "頁數總計："
         Height          =   180
         Index           =   16
         Left            =   -68670
         TabIndex        =   306
         Top             =   2820
         Width           =   900
      End
      Begin VB.Label lblTot6 
         BackColor       =   &H8000000E&
         Caption         =   " "
         Height          =   255
         Left            =   -67470
         TabIndex        =   305
         Top             =   2820
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "不算超頁費"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   75
         Left            =   -67890
         TabIndex        =   304
         Top             =   1980
         Width           =   900
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   172
         Left            =   -66870
         TabIndex        =   303
         Top             =   3105
         Width           =   500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "申請專利範圍項數："
         Height          =   180
         Index           =   17
         Left            =   -68670
         TabIndex        =   302
         Top             =   3105
         Width           =   1620
      End
      Begin MSForms.Label lblPA 
         Height          =   255
         Index           =   173
         Left            =   -66870
         TabIndex        =   301
         Top             =   3360
         Width           =   500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "圖式圖數："
         Height          =   180
         Index           =   18
         Left            =   -68670
         TabIndex        =   300
         Top             =   3360
         Width           =   1620
      End
      Begin MSForms.Label lblPA 
         Height          =   252
         Index           =   69
         Left            =   1752
         TabIndex        =   295
         Top             =   2244
         Width           =   264
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "459;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "國內副本接洽人:"
         Height          =   180
         Index           =   13
         Left            =   -69780
         TabIndex        =   294
         Top             =   5730
         Width           =   1365
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   169
         Left            =   -68295
         TabIndex        =   293
         Top             =   5730
         Width           =   2490
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4392;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "國內副本收件人:"
         Height          =   180
         Index           =   2
         Left            =   -74910
         TabIndex        =   292
         Top             =   5730
         Width           =   1305
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   168
         Left            =   -73470
         TabIndex        =   291
         Top             =   5730
         Width           =   3570
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6297;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCMboth 
         Caption         =   "lblCMboth"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -67260
         TabIndex        =   290
         Top             =   5580
         Width           =   945
      End
      Begin VB.Label lblCaseMap 
         AutoSize        =   -1  'True
         Caption         =   "lblCaseMap"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -67260
         TabIndex        =   289
         Top             =   5100
         Width           =   1320
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   167
         Left            =   2640
         TabIndex        =   287
         Top             =   1032
         Width           =   264
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "450;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日)："
         Height          =   180
         Index           =   11
         Left            =   -74790
         TabIndex        =   285
         Top             =   2142
         Width           =   1380
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   92
         Left            =   -73380
         TabIndex        =   284
         Top             =   2142
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   164
         Left            =   -67620
         TabIndex        =   283
         Top             =   2975
         Width           =   1605
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "存取碼:"
         Height          =   180
         Index           =   14
         Left            =   -68325
         TabIndex        =   282
         Top             =   2975
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新穎性優惠日期:"
         Height          =   180
         Index           =   12
         Left            =   -68325
         TabIndex        =   275
         Top             =   3264
         Width           =   1305
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   140
         Left            =   -66840
         TabIndex        =   274
         Top             =   3264
         Width           =   975
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1720;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "國際分類:"
         Height          =   180
         Index           =   11
         Left            =   -68325
         TabIndex        =   273
         Top             =   2641
         Width           =   855
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   160
         Left            =   -67470
         TabIndex        =   272
         Top             =   2641
         Width           =   1455
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2566;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   159
         Left            =   2160
         TabIndex        =   270
         Top             =   5280
         Width           =   3168
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5583;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "案件屬性: "
         Height          =   180
         Left            =   -67920
         TabIndex        =   269
         Top             =   322
         Width           =   810
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   91
         Left            =   -66990
         TabIndex        =   268
         Top             =   322
         Width           =   1065
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1879;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   90
         Left            =   -66690
         TabIndex        =   267
         Top             =   3842
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請人是否同發明人:        (Y/ N)"
         Height          =   180
         Index           =   1
         Left            =   -68415
         TabIndex        =   266
         Top             =   3842
         Width           =   2475
      End
      Begin VB.Label lblFilingDate 
         AutoSize        =   -1  'True
         Caption         =   "lblFilingDate"
         Height          =   180
         Index           =   1
         Left            =   -67650
         TabIndex        =   265
         Top             =   2367
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblFilingDate 
         AutoSize        =   -1  'True
         Caption         =   "提交日:"
         Height          =   180
         Index           =   0
         Left            =   -68325
         TabIndex        =   264
         Top             =   2367
         Visible         =   0   'False
         Width           =   585
      End
      Begin MSForms.Label lblPA 
         Height          =   252
         Index           =   152
         Left            =   6252
         TabIndex        =   262
         Top             =   400
         Width           =   396
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblPA 
         Height          =   252
         Index           =   151
         Left            =   4644
         TabIndex        =   260
         Top             =   400
         Width           =   396
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   86
         Left            =   -68820
         TabIndex        =   259
         Top             =   656
         Width           =   1275
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2249;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label26 
         Caption         =   "工程師組別:"
         Height          =   165
         Index           =   1
         Left            =   -69780
         TabIndex        =   258
         Top             =   656
         Width           =   1035
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   85
         Left            =   -73710
         TabIndex        =   257
         Top             =   5439
         Width           =   2490
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4392;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "接洽人:"
         Height          =   180
         Index           =   12
         Left            =   -74910
         TabIndex        =   256
         Top             =   5439
         Width           =   585
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   82
         Left            =   7896
         TabIndex        =   253
         Top             =   2244
         Width           =   264
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "年費本所是否出名：       (N:不出名)"
         Height          =   180
         Index           =   8
         Left            =   6264
         TabIndex        =   252
         Top             =   2244
         Width           =   2784
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   58
         Left            =   4296
         TabIndex        =   173
         Top             =   2556
         Width           =   264
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "450;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   46
         Left            =   8232
         TabIndex        =   121
         Top             =   1644
         Width           =   264
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
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   41
         Left            =   8232
         TabIndex        =   114
         Top             =   1344
         Width           =   264
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
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   56
         Left            =   1752
         TabIndex        =   170
         Top             =   1644
         Width           =   264
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "459;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   80
         Left            =   5376
         TabIndex        =   250
         Top             =   1332
         Width           =   252
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否核對已准專利：         (N:不核對)"
         Height          =   180
         Index           =   156
         Left            =   3336
         TabIndex        =   251
         Top             =   1344
         Width           =   3156
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   249
         Top             =   1230
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   79
         Left            =   -73620
         TabIndex        =   248
         Top             =   1230
         Width           =   6165
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10874;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   180
         Left            =   -74880
         TabIndex        =   247
         Top             =   930
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   78
         Left            =   -73800
         TabIndex        =   246
         Top             =   930
         Width           =   1000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1773;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   180
         Left            =   -74880
         TabIndex        =   245
         Top             =   660
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   77
         Left            =   -73800
         TabIndex        =   244
         Top             =   660
         Width           =   1000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1764;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   180
         Left            =   -74880
         TabIndex        =   243
         Top             =   390
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   27
         Left            =   -73800
         TabIndex        =   242
         Top             =   390
         Width           =   1000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1764;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "尚未公告，年費為預估值"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   -74910
         TabIndex        =   241
         Top             =   1530
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "尚未公告，下次繳費日為預估值"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   -74910
         TabIndex        =   240
         Top             =   5370
         Visible         =   0   'False
         Width           =   3990
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   76
         Left            =   -73740
         TabIndex        =   239
         Top             =   1260
         Width           =   1545
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2725;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   75
         Left            =   -73740
         TabIndex        =   238
         Top             =   540
         Width           =   1545
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2725;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "法定期限:"
         Height          =   180
         Left            =   -74730
         TabIndex        =   237
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "本所期限:"
         Height          =   180
         Index           =   0
         Left            =   -74730
         TabIndex        =   236
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日:"
         Height          =   180
         Left            =   -74775
         TabIndex        =   235
         Top             =   540
         Width           =   945
      End
      Begin VB.Label lblInventor 
         Height          =   180
         Index           =   32
         Left            =   -74040
         TabIndex        =   234
         Top             =   3630
         Width           =   1200
      End
      Begin VB.Label lblInventor 
         Height          =   180
         Index           =   31
         Left            =   -74040
         TabIndex        =   233
         Top             =   3270
         Width           =   1200
      End
      Begin VB.Label lblInventor 
         Height          =   180
         Index           =   30
         Left            =   -74040
         TabIndex        =   232
         Top             =   2895
         Width           =   1200
      End
      Begin VB.Label lblInventor 
         Height          =   180
         Index           =   29
         Left            =   -74040
         TabIndex        =   231
         Top             =   2550
         Width           =   1200
      End
      Begin VB.Label lblInventor 
         Height          =   180
         Index           =   28
         Left            =   -74040
         TabIndex        =   230
         Top             =   2190
         Width           =   1200
      End
      Begin VB.Label lblInventor 
         Height          =   180
         Index           =   27
         Left            =   -74040
         TabIndex        =   229
         Top             =   1830
         Width           =   1200
      End
      Begin VB.Label lblInventor 
         Height          =   180
         Index           =   26
         Left            =   -74040
         TabIndex        =   228
         Top             =   1470
         Width           =   1200
      End
      Begin VB.Label lblInventor 
         Height          =   180
         Index           =   25
         Left            =   -74040
         TabIndex        =   227
         Top             =   1110
         Width           =   1200
      End
      Begin VB.Label lblInventor 
         Height          =   180
         Index           =   24
         Left            =   -74040
         TabIndex        =   226
         Top             =   750
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "發明人:"
         Height          =   180
         Index           =   4
         Left            =   -74820
         TabIndex        =   225
         Top             =   375
         Width           =   585
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   74
         Left            =   -73710
         TabIndex        =   224
         Top             =   4885
         Width           =   3825
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6747;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "年費聯絡人:"
         Height          =   180
         Left            =   -74910
         TabIndex        =   223
         Top             =   4885
         Width           =   945
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   73
         Left            =   -68295
         TabIndex        =   222
         Top             =   4608
         Width           =   2490
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4392;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "年費D/N列印對象:"
         Height          =   180
         Left            =   -69825
         TabIndex        =   221
         Top             =   4608
         Width           =   1410
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   72
         Left            =   -73470
         TabIndex        =   220
         Top             =   4050
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象:"
         Height          =   180
         Left            =   -74910
         TabIndex        =   219
         Top             =   4050
         Width           =   1410
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   -73860
         TabIndex        =   142
         Top             =   4420
         Width           =   525
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "926;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   71
         Left            =   -69000
         TabIndex        =   218
         Top             =   322
         Width           =   1000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1764;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "是否PCT案: "
         Height          =   180
         Left            =   -70080
         TabIndex        =   217
         Top             =   322
         Width           =   945
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   14
         Left            =   -74970
         TabIndex        =   216
         Top             =   5587
         Width           =   2010
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   13
         Left            =   -74910
         TabIndex        =   215
         Top             =   5018
         Width           =   1920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   12
         Left            =   -74910
         TabIndex        =   214
         Top             =   4451
         Width           =   1920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   11
         Left            =   -74910
         TabIndex        =   213
         Top             =   3884
         Width           =   1920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   10
         Left            =   -74910
         TabIndex        =   212
         Top             =   3317
         Width           =   1920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   9
         Left            =   -74910
         TabIndex        =   211
         Top             =   2750
         Width           =   1920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   8
         Left            =   -74910
         TabIndex        =   210
         Top             =   2183
         Width           =   1920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   6
         Left            =   -74910
         TabIndex        =   209
         Top             =   1616
         Width           =   1920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   5
         Left            =   -74910
         TabIndex        =   208
         Top             =   482
         Width           =   1920
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   7
         Left            =   -74910
         TabIndex        =   207
         Top             =   1049
         Width           =   1920
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   43
         Left            =   -73380
         TabIndex        =   206
         Top             =   1236
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   37
         Left            =   -73380
         TabIndex        =   205
         Top             =   330
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   61
         Left            =   -73380
         TabIndex        =   204
         Top             =   632
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   62
         Left            =   -73380
         TabIndex        =   203
         Top             =   934
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   63
         Left            =   -73380
         TabIndex        =   202
         Top             =   1538
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   64
         Left            =   -73380
         TabIndex        =   201
         Top             =   1840
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "卷宗性質: "
         Height          =   180
         Index           =   0
         Left            =   -72135
         TabIndex        =   200
         Top             =   322
         Width           =   816
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   -71295
         TabIndex        =   199
         Top             =   322
         Width           =   1095
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1931;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   1
         Left            =   -74925
         TabIndex        =   198
         Top             =   322
         Width           =   765
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   70
         Left            =   -70170
         TabIndex        =   197
         Top             =   5010
         Width           =   2505
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4419;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "分所案號:"
         Height          =   180
         Index           =   1
         Left            =   -71145
         TabIndex        =   196
         Top             =   5010
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "實體副本收受人彼所案號2："
         Height          =   180
         Index           =   7
         Left            =   -74835
         TabIndex        =   195
         Top             =   4830
         Width           =   2250
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "實體副本收受人彼所案號1："
         Height          =   180
         Index           =   6
         Left            =   -74835
         TabIndex        =   194
         Top             =   4260
         Width           =   2250
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   69
         Left            =   -73380
         TabIndex        =   193
         Top             =   3960
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "實體副本聯絡人:"
         Height          =   180
         Index           =   5
         Left            =   -74715
         TabIndex        =   192
         Top             =   3960
         Width           =   1305
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   68
         Left            =   -73380
         TabIndex        =   191
         Top             =   3652
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "實體副本收受人:"
         Height          =   180
         Index           =   4
         Left            =   -74715
         TabIndex        =   190
         Top             =   3652
         Width           =   1305
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   67
         Left            =   -73380
         TabIndex        =   189
         Top             =   3048
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   66
         Left            =   -73380
         TabIndex        =   188
         Top             =   2746
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   65
         Left            =   -73380
         TabIndex        =   187
         Top             =   2444
         Width           =   7450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13141;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人(英)："
         Height          =   180
         Index           =   3
         Left            =   -74790
         TabIndex        =   186
         Top             =   2746
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人(日)："
         Height          =   180
         Index           =   2
         Left            =   -74790
         TabIndex        =   185
         Top             =   3048
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人(中)："
         Height          =   180
         Index           =   1
         Left            =   -74790
         TabIndex        =   184
         Top             =   2444
         Width           =   1380
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(日) 2："
         Height          =   180
         Index           =   10
         Left            =   -74565
         TabIndex        =   183
         Top             =   1840
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(英) 2："
         Height          =   180
         Index           =   6
         Left            =   -74565
         TabIndex        =   182
         Top             =   1538
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(日) 1："
         Height          =   180
         Index           =   3
         Left            =   -74565
         TabIndex        =   181
         Top             =   934
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(英) 1："
         Height          =   180
         Index           =   2
         Left            =   -74565
         TabIndex        =   180
         Top             =   632
         Width           =   1155
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "客戶收款後辦案："
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   179
         Top             =   3672
         Width           =   1440
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "(Y:先收)"
         Height          =   180
         Index           =   4
         Left            =   2220
         TabIndex        =   178
         Top             =   3696
         Width           =   648
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   60
         Left            =   1716
         TabIndex        =   177
         Top             =   3672
         Width           =   468
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "820;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   59
         Left            =   1752
         TabIndex        =   175
         Top             =   2556
         Width           =   264
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "459;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "年費單筆不跑：        (Y:不跑)"
         Height          =   180
         Index           =   1
         Left            =   2988
         TabIndex        =   174
         Top             =   2556
         Width           =   2268
      End
      Begin MSForms.Label lbl1 
         Height          =   375
         Index           =   57
         Left            =   -70140
         TabIndex        =   172
         Top             =   5310
         Width           =   2800
         ForeColor       =   255
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "紅14-lblFM2"
         Size            =   "4939;661"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   285
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費自動代繳：       (Y:自動代繳)"
         Height          =   180
         Index           =   7
         Left            =   180
         TabIndex        =   171
         Top             =   1644
         Width           =   2880
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   55
         Left            =   1752
         TabIndex        =   168
         Top             =   1344
         Width           =   264
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
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   39
         Left            =   -73380
         TabIndex        =   167
         Top             =   3350
         Width           =   3480
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6138;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   45
         Left            =   -68790
         TabIndex        =   166
         Top             =   3350
         Width           =   2880
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5080;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(1.中文 2.英文 3.日文)"
         Height          =   180
         Left            =   -73230
         TabIndex        =   165
         Top             =   4420
         Width           =   1695
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文:"
         Height          =   180
         Left            =   -74925
         TabIndex        =   164
         Top             =   4420
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "專利種類:"
         Height          =   180
         Index           =   5
         Left            =   -74925
         TabIndex        =   163
         Top             =   656
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利名稱(中):"
         Height          =   180
         Index           =   7
         Left            =   -74925
         TabIndex        =   162
         Top             =   1065
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "申請日期:"
         Height          =   180
         Index           =   6
         Left            =   -74925
         TabIndex        =   161
         Top             =   2367
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公開日:"
         Height          =   180
         Index           =   13
         Left            =   -74925
         TabIndex        =   160
         Top             =   2641
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公告日:"
         Height          =   180
         Index           =   3
         Left            =   -74925
         TabIndex        =   159
         Top             =   2975
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發證日:"
         Height          =   180
         Index           =   6
         Left            =   -74895
         TabIndex        =   158
         Top             =   3264
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利號數:"
         Height          =   180
         Index           =   9
         Left            =   -74925
         TabIndex        =   157
         Top             =   3553
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(外):"
         Height          =   180
         Index           =   2
         Left            =   -74175
         TabIndex        =   156
         Top             =   1875
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   8
         Left            =   -74175
         TabIndex        =   155
         Top             =   1519
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "目前准/駁:"
         Height          =   180
         Index           =   0
         Left            =   -74925
         TabIndex        =   154
         Top             =   3842
         Width           =   810
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "(1.准2.駁)"
         Height          =   180
         Index           =   0
         Left            =   -73230
         TabIndex        =   153
         Top             =   3842
         Width           =   750
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "是否有救濟程序:"
         Height          =   180
         Index           =   0
         Left            =   -74955
         TabIndex        =   152
         Top             =   4155
         Width           =   1305
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "(Y:有)"
         Height          =   180
         Index           =   0
         Left            =   -73020
         TabIndex        =   151
         Top             =   4155
         Width           =   465
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期:"
         Height          =   180
         Left            =   -74925
         TabIndex        =   150
         Top             =   4709
         Width           =   765
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日:"
         Height          =   180
         Left            =   -74925
         TabIndex        =   149
         Top             =   5010
         Width           =   945
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   -74070
         TabIndex        =   148
         Top             =   656
         Width           =   1275
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2249;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   -74040
         TabIndex        =   147
         Top             =   2975
         Width           =   1425
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2514;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   -74040
         TabIndex        =   146
         Top             =   3264
         Width           =   1455
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2566;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   -74040
         TabIndex        =   145
         Top             =   3553
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
         Index           =   10
         Left            =   -73950
         TabIndex        =   144
         Top             =   3842
         Width           =   585
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1032;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   -73560
         TabIndex        =   143
         Top             =   4155
         Width           =   465
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "820;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   -73860
         TabIndex        =   141
         Top             =   4709
         Width           =   1725
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3043;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   14
         Left            =   -73860
         TabIndex        =   140
         Top             =   5016
         Width           =   2640
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4657;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(中) 2："
         Height          =   180
         Index           =   5
         Left            =   -74565
         TabIndex        =   139
         Top             =   1236
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(中) 1："
         Height          =   180
         Index           =   4
         Left            =   -74565
         TabIndex        =   138
         Top             =   330
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人:"
         Height          =   180
         Index           =   0
         Left            =   -74355
         TabIndex        =   137
         Top             =   3350
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人:"
         Height          =   180
         Left            =   -69780
         TabIndex        =   136
         Top             =   3350
         Width           =   945
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "CREATE ID:"
         Height          =   180
         Index           =   0
         Left            =   -74835
         TabIndex        =   135
         Top             =   4635
         Width           =   945
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "UPDATE ID:"
         Height          =   180
         Index           =   1
         Left            =   -74835
         TabIndex        =   134
         Top             =   4935
         Width           =   930
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   47
         Left            =   -73605
         TabIndex        =   133
         Top             =   4635
         Width           =   1700
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2990;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   48
         Left            =   -73605
         TabIndex        =   132
         Top             =   4935
         Width           =   1700
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2990;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   49
         Left            =   -71325
         TabIndex        =   131
         Top             =   4635
         Width           =   1700
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2990;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   50
         Left            =   -71325
         TabIndex        =   130
         Top             =   4935
         Width           =   1700
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2990;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   51
         Left            =   -69165
         TabIndex        =   129
         Top             =   4635
         Width           =   1700
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2990;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   52
         Left            =   -69165
         TabIndex        =   128
         Top             =   4935
         Width           =   1700
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2990;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   54
         Left            =   -68295
         TabIndex        =   127
         Top             =   4331
         Width           =   2490
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4392;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   53
         Left            =   -68295
         TabIndex        =   126
         Top             =   4065
         Width           =   2490
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4392;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "年費請款對象:"
         Height          =   180
         Index           =   5
         Left            =   -69840
         TabIndex        =   125
         Top             =   4331
         Width           =   1125
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象:"
         Height          =   180
         Left            =   -74910
         TabIndex        =   124
         Top             =   4608
         Width           =   1125
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   40
         Left            =   -73710
         TabIndex        =   123
         Top             =   4608
         Width           =   3825
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6747;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "年費彼所案號:"
         Height          =   180
         Index           =   3
         Left            =   -69855
         TabIndex        =   122
         Top             =   4095
         Width           =   1125
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   36
         Left            =   -73710
         TabIndex        =   120
         Top             =   4331
         Width           =   3825
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6747;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   35
         Left            =   -73710
         TabIndex        =   119
         Top             =   3777
         Width           =   7830
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13811;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   26
         Left            =   -70170
         TabIndex        =   118
         Top             =   4709
         Width           =   2505
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4419;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷:"
         Height          =   180
         Left            =   -71115
         TabIndex        =   117
         Top             =   4420
         Width           =   765
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   44
         Left            =   3024
         TabIndex        =   116
         Top             =   400
         Width           =   396
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   42
         Left            =   1800
         TabIndex        =   115
         Top             =   720
         Width           =   264
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
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   38
         Left            =   990
         TabIndex        =   113
         Top             =   400
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   315
         Index           =   34
         Left            =   -73710
         TabIndex        =   112
         Top             =   5130
         Width           =   7845
         BackColor       =   16777215
         Caption         =   "lblFM2"
         Size            =   "13838;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   33
         Left            =   -73710
         TabIndex        =   111
         Top             =   3500
         Width           =   7815
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13785;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   32
         Left            =   -74070
         TabIndex        =   110
         Top             =   1408
         Width           =   8190
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "14446;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   31
         Left            =   -74070
         TabIndex        =   109
         Top             =   1131
         Width           =   8190
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "14446;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   30
         Left            =   -74070
         TabIndex        =   108
         Top             =   854
         Width           =   8190
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "14446;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   29
         Left            =   -74070
         TabIndex        =   107
         Top             =   577
         Width           =   8190
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "14446;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   28
         Left            =   -74070
         TabIndex        =   106
         Top             =   300
         Width           =   8190
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "14446;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   25
         Left            =   -70170
         TabIndex        =   105
         Top             =   4425
         Width           =   645
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1138;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   24
         Left            =   -69765
         TabIndex        =   104
         Top             =   4155
         Width           =   705
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1244;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   -69750
         TabIndex        =   103
         Top             =   3842
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
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   22
         Left            =   -68712
         TabIndex        =   102
         Top             =   3552
         Width           =   2760
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4868;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   -70155
         TabIndex        =   101
         Top             =   3553
         Width           =   1005
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1773;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   20
         Left            =   -69960
         TabIndex        =   100
         Top             =   3270
         Width           =   1035
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1826;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   -71715
         TabIndex        =   99
         Top             =   656
         Width           =   1695
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2990;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "代理人備註:"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   98
         Top             =   2856
         Width           =   948
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "客戶備註:"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   97
         Top             =   3630
         Width           =   765
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人：       (Y:印)"
         Height          =   180
         Index           =   0
         Left            =   6528
         TabIndex        =   94
         Top             =   1644
         Width           =   2508
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "全部折扣:           %"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   93
         Top             =   400
         Width           =   1395
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "申請/翻譯折扣:           %"
         Height          =   180
         Index           =   1
         Left            =   1764
         TabIndex        =   92
         Top             =   400
         Width           =   1848
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "年費代理人:"
         Height          =   180
         Left            =   -74910
         TabIndex        =   91
         Top             =   4331
         Width           =   945
      End
      Begin VB.Label Label37 
         Caption         =   "Label37"
         Height          =   15
         Left            =   -71085
         TabIndex        =   90
         Top             =   660
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "申請人5:"
         Height          =   180
         Index           =   1
         Left            =   -74910
         TabIndex        =   89
         Top             =   1408
         Width           =   675
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "閉卷原因:"
         Height          =   180
         Index           =   0
         Left            =   -71145
         TabIndex        =   88
         Top             =   4709
         Width           =   765
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "(Y:有)"
         Height          =   180
         Index           =   1
         Left            =   -68970
         TabIndex        =   87
         Top             =   4155
         Width           =   615
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "是否有爭議程序:"
         Height          =   180
         Index           =   1
         Left            =   -71175
         TabIndex        =   86
         Top             =   4155
         Width           =   1305
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "(Y/N)"
         Height          =   180
         Index           =   0
         Left            =   -69345
         TabIndex        =   85
         Top             =   3842
         Width           =   405
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "專利權是否存在:"
         Height          =   180
         Index           =   0
         Left            =   -71145
         TabIndex        =   84
         Top             =   3842
         Width           =   1305
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "申請國家:"
         Height          =   180
         Index           =   0
         Left            =   -72555
         TabIndex        =   83
         Top             =   656
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專用期限:"
         Height          =   180
         Index           =   10
         Left            =   -71115
         TabIndex        =   82
         Top             =   3553
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公告號:"
         Height          =   180
         Index           =   5
         Left            =   -71115
         TabIndex        =   81
         Top             =   2975
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "公開號:"
         Height          =   180
         Index           =   15
         Left            =   -71115
         TabIndex        =   80
         Top             =   2641
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "准駁通知日:"
         Height          =   180
         Index           =   4
         Left            =   -71115
         TabIndex        =   79
         Top             =   3264
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "申請案號:"
         Height          =   180
         Index           =   0
         Left            =   -71970
         TabIndex        =   78
         Top             =   2370
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "申請人1:"
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   77
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "申請人1地址(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   76
         Top             =   1765
         Width           =   2280
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "申請人2地址(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   1
         Left            =   -74910
         TabIndex        =   75
         Top             =   2127
         Width           =   2280
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "FC代理人:"
         Height          =   180
         Index           =   1
         Left            =   -74910
         TabIndex        =   74
         Top             =   3500
         Width           =   795
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號:"
         Height          =   180
         Left            =   -74910
         TabIndex        =   73
         Top             =   3777
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "申請人4:"
         Height          =   180
         Index           =   7
         Left            =   -74910
         TabIndex        =   72
         Top             =   1131
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "申請人3:"
         Height          =   180
         Index           =   8
         Left            =   -74910
         TabIndex        =   71
         Top             =   854
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "申請人2:"
         Height          =   180
         Index           =   9
         Left            =   -74910
         TabIndex        =   70
         Top             =   577
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號:"
         Height          =   180
         Index           =   1
         Left            =   -74910
         TabIndex        =   69
         Top             =   5162
         Width           =   1125
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "申請人3地址(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   2
         Left            =   -74910
         TabIndex        =   68
         Top             =   2489
         Width           =   2280
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "申請人4地址(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   3
         Left            =   -74910
         TabIndex        =   67
         Top             =   2842
         Width           =   2280
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "申請人5地址(1:中, 2:英 ,3:日):"
         Height          =   180
         Index           =   4
         Left            =   -74910
         TabIndex        =   66
         Top             =   3202
         Width           =   2280
      End
      Begin VB.Label Label29 
         Caption         =   "(N:不是)"
         Height          =   255
         Left            =   -65700
         TabIndex        =   65
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "(Y:閉卷)"
         Height          =   180
         Left            =   -69360
         TabIndex        =   64
         Top             =   4425
         Width           =   645
      End
      Begin VB.Line Line1 
         X1              =   -69120
         X2              =   -68880
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "FCP領證自動代繳：       (Y:自動代繳)"
         Height          =   180
         Index           =   6
         Left            =   180
         TabIndex        =   169
         Top             =   1344
         Width           =   2916
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "信函是否列印Title：        (Y:印)"
         Height          =   180
         Index           =   0
         Left            =   6636
         TabIndex        =   95
         Top             =   1344
         Width           =   2400
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "後續准駁簡單報告：         (Y: 核准以及C類來函簡單報告)"
         Height          =   180
         Left            =   180
         TabIndex        =   96
         Top             =   720
         Width           =   4476
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "代理人收款後辦案：       (Y:先收)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   176
         Top             =   2556
         Width           =   2580
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   83
         Left            =   4968
         TabIndex        =   254
         Top             =   2244
         Width           =   264
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "C類收文是否請款：       (N:否)"
         Height          =   180
         Index           =   9
         Left            =   3396
         TabIndex        =   255
         Top             =   2244
         Width           =   2340
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "領證折扣:           %"
         Height          =   180
         Index           =   2
         Left            =   3780
         TabIndex        =   261
         Top             =   400
         Width           =   1440
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "年費折扣:           %"
         Height          =   180
         Index           =   3
         Left            =   5400
         TabIndex        =   263
         Top             =   400
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID："
         Height          =   180
         Index           =   0
         Left            =   228
         TabIndex        =   271
         Top             =   5280
         Width           =   1860
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "年費逾期補繳通知函是否寄發:        (N:不寄)"
         Height          =   180
         Left            =   180
         TabIndex        =   288
         Top             =   1032
         Width           =   3396
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "FCP實審自動代繳：       (Y:自動代繳)"
         Height          =   180
         Index           =   14
         Left            =   180
         TabIndex        =   296
         Top             =   2244
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司：        ( T：專利商標 J：智權公司 空白:系統預設)"
         Height          =   180
         Index           =   27
         Left            =   -74550
         TabIndex        =   334
         Top             =   2580
         Width           =   4965
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "以Email通知：        (Y：是   D：僅D/N）"
         Height          =   180
         Left            =   -74430
         TabIndex        =   333
         Top             =   1980
         Width           =   3105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email 同時寄紙本：        (Y：是)"
         Height          =   180
         Index           =   166
         Left            =   -74820
         TabIndex        =   329
         Top             =   2280
         Width           =   2490
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "帳單備註是否提醒：       (N：否)"
         Height          =   180
         Index           =   10
         Left            =   -74895
         TabIndex        =   332
         Top             =   1140
         Width           =   2535
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   181
         Left            =   8448
         TabIndex        =   368
         Top             =   400
         Width           =   240
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "423;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "專利不得請雜費：     (Y:是)"
         Height          =   180
         Index           =   4
         Left            =   7032
         TabIndex        =   369
         Top             =   400
         Width           =   2148
      End
   End
End
Attribute VB_Name = "frm100101_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/13 改成Form2.0 ;  lbl1(index)、txt1(index)、lblPA(index)、grdDataList1改字型=新細明體-ExtB、grdDataList2改字型=新細明體-ExtB、GRD1改字型=新細明體-ExtB、lstPA166
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit

Dim StrTag As String, StrTag1 As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add By Sindy 2010/02/04
Dim StrTag2 As String, StrTag3 As String, StrTag4 As String, StrTag5 As String
Dim m_FixNo As Integer   '2010/2/22 add by sonia 修法次數
Dim pa(1 To 4) As String 'Added by Lydia 2019/11/04 本所案號
Dim m_bolFMP As Boolean 'Added by Lydia 2023/03/09
Dim m_pub_QL05 As String 'Add By Sindy 2025/8/28 只記錄於此Form


'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1
     fnCloseAllFrm100
Case 2
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag1 ' StrTag    傳申請人代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 3
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_10.Show
     frm100101_10.Tag = StrTag ' StrTag  傳代理人代號
     frm100101_10.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_10.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'add by nick 2004/09/15
'相關卷號
Case 4
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100108_3.Show
     frm100108_3.Tag = txt1(81).Text
     frm100108_3.Caption = "相關卷號"
     frm100108_3.StrMenu2
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'分割案
Case 5
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100108_4.Show
     frm100108_4.frm100108_txt_7 = "3"
     frm100108_4.SetDataListWidth
     frm100108_4.Tag = txt1(81).Text
     frm100108_4.Caption = "分割案"
     frm100108_4.StrMenu1
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'add by nickc 2005/12/12
Case 6
    frmPic001.oCP01 = SystemNumber(txt1(81), 1)
    frmPic001.oCP02 = SystemNumber(txt1(81), 2)
    frmPic001.oCP03 = SystemNumber(txt1(81), 3)
    frmPic001.oCP04 = SystemNumber(txt1(81), 4)
    frmPic001.StrMenu
    frmPic001.CanScan
    frmPic001.SetSeekCmdok 'Add by Amy 2018/07/16
    frmPic001.Show vbModal
    'add by nickc 2005/12/15 檢查有無代表圖
    'Modify by Amy 2018/07/16  改寫至function
'    strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(txt1(81), 1) & "' and ibf02='" & SystemNumber(txt1(81), 2) & "' and ibf03='" & SystemNumber(txt1(81), 3) & "' and ibf04='" & SystemNumber(txt1(81), 4) & "' and ibf05='1'"
'    CheckOC2
'    adoRecordset1.CursorLocation = adUseClient
'    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    If ChkImgByteFile(SystemNumber(txt1(81), 1), SystemNumber(txt1(81), 2), SystemNumber(txt1(81), 3), SystemNumber(txt1(81), 4)) = True Then
        'Modified by Lydia 2021/12/16 拿掉快速鍵(&I)
        cmdOK(6).Caption = "已設定代表圖"
        cmdOK(6).BackColor = &HC0FFC0
    Else
        'Modified by Lydia 2021/12/16 拿掉快速鍵(&I)
        cmdOK(6).Caption = "未設定代表圖"
        cmdOK(6).BackColor = &HC0C0FF
    End If
'    CheckOC2
    'end 2018/07/16
'Add By Sindy 2010/02/04
Case 7
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag2 ' StrTag2    傳申請人2代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 8
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag3 ' StrTag3    傳申請人3代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 9
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag4 ' StrTag4    傳申請人4代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 10
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag5 ' StrTag5    傳申請人5代號
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'2010/02/04 End
'Added by Lydia 2016/11/23
Case 11 '各項指示
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
     frm12040159.SetParent "Q", Trim(Replace(txt1(81), "-", "")), Me
     frm12040159.Show
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'Add By Sindy 2020/7/15
Case 12 '進度
   cmdState = -1
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_2.Show
   frm100101_2.Tag = txt1(81)
   frm100101_2.StrMenu
   Screen.MousePointer = vbDefault
   Me.Enabled = True
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/28 還原此Form的查詢條件記錄
End Sub

Private Sub Form_Load()
Dim Lbl As Object

For Each Lbl In Me.lbl1
   Lbl.BackColor = &H8000000F
Next

cmdDivSug.Visible = False 'Added by Lydia 2019/11/04
lblTot6.BackColor = &H8000000F 'Added by Lydia 2018/12/27

bolToEndByNick = False
MoveFormToCenter Me
SSTab1.Tab = 0 'Add by Amy 2014/04/10
GRIDHEAND
Grid

'Removed by Morgan 2021/3/12 FC資料可不必限制--秀玲
'If bolFNation = False Then
'    SSTab1.TabVisible(2) = False
'    Label16(1).Visible = False
'    lbl1(33).Visible = False
'    Label17.Visible = False
'    lbl1(35).Visible = False
'    Label42.Visible = False
'    lbl1(36).Visible = False
'    Label11(3).Visible = False
'    lbl1(53).Visible = False
'    Label44.Visible = False
'    lbl1(40).Visible = False
'    Label11(5).Visible = False
'    lbl1(54).Visible = False
'    cmdok(3).Visible = False
'End If
'end 2021/3/12

'92.04.16 nick
cmdState = -1
   'Add By Sindy 2013/12/17
   If strSrvDate(1) >= InvoiceStartDate Then
      lblPA(161).Left = 1620
   Else
      Label1(27).Caption = "是否以專利商標出名：      ( Y：是 )"
      lblPA(161).Left = 1890
   End If
   'Added by Lydia 2020/03/30 事務所合併日起取消(T:專利商標 J:智權公司 空白:系統預設)的標題改為(J:智權公司 空白:系統預設)
   If strSrvDate(1) >= 事務所合併日 Then
       Label1(27).Caption = "特殊出名公司：             ( J：智權公司 空白:系統預設)"
   End If
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdOK(11).Visible = True
   Else
      cmdOK(11).Visible = False
      'Mark by Lydia 2020/09/18 按鈕移到最上方
      'txt1(53).Top = 330
      'txt1(53).Height = 4770
      'end 2020/09/18
   End If
   'end 2020/05/05
   
   Frame1K.BorderStyle = 0 'Add By Sindy 2025/1/7
End Sub

Sub StrMenu()
Dim strSql  As String
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
'Dim strArr(132) As String, i As Integer, StrOk(70) As String, StrOkTxt(79) As String
Dim i As Integer, StrOk(75 - 1) As String, StrOkTxt(86) As String
Dim strArr() As String
ReDim strArr(TF_PA) As String
'Add By Cheng 2002/07/08
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSK03 As String
'add by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
Dim tmp01 As String, tmp02 As String, tmp03 As String, tmp04 As String, tmp05 As String
'add by Toni 20080926 控制跨部門權限訊息
Dim strTit As String
Dim strMsg As String
Dim nResponse
'end by Toni 20080926
Dim strFeeType As String, strYF15 As String 'Add By Sindy 2009/06/29
Dim oLbl As Object 'Add by Morgan 2008/11/13
'Dim pa(5) As String '2010/2/22 add by sonia 'Remove by Lydia 2019/11/04
Dim strPA09 As String 'Added by Lydia 2016/06/14
Dim strPA08 As String 'Added by Morgan 2024/10/4
Dim arrID 'Add By Sindy 2025/1/7

Str01 = ""
Str02 = ""
Str03 = ""
Str04 = ""
If Left(Me.Tag, 1) = "N" Then
   strSql = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   strSql = Me.Tag
End If
Str01 = SystemNumber(strSql, 1)
Str02 = SystemNumber(strSql, 2)
Str03 = SystemNumber(strSql, 3)
Str04 = SystemNumber(strSql, 4)
'Added by Lydia 2019/11/04
pa(1) = Str01
pa(2) = Str02
pa(3) = Str03
pa(4) = Str04
'end 2019/11/04

'add by Toni 20080926 控制跨部門權限訊息 for 專利基本資料用
'2008/10/2 modify by sonia
'If IsUserHasRightOfSystem(strUserNum, Str01) = False Then
'   If IsUserHasRightOfFunction("frm100101_1", strCrossDept, False) = False Then
'      strTit = "檢核資料"
'      strMsg = "您沒有使用該系統類別的權限"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   End If
'End If
If CheckSR09(strUserNum, Str01, "Y", , Str01, Str02, Str03, Str04) = False Then
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
'2008/10/2 end
'End by 20080926

pub_QL05 = ";本所案號：" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & _
           "(基本資料)" 'Add By Sindy 2025/8/7

'Add By Cheng 2002/07/08
strSK03 = ""
StrSQLa = "Select SK03 From SystemKind Where SK01='" & Str01 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic
If rsA.RecordCount > 0 Then
   strSK03 = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close

'欲搜尋的SQL字串
strSql = "SELECT * FROM PATENT WHERE PA01='" & Str01 & "' AND PA02='" & Str02 & "' AND PA03='" & Str03 & "' AND PA04='" & Str04 & "'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28 記錄此Form的查詢條件
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/7
'    For i = 0 To 131
    For i = 0 To (TF_PA - 1) ' edit by nickc 2006/07/12 (T_PA - 1)
        Select Case i
            Case 9, 11, 13, 19, 20, 22, 23, 24, 48, 49, 57, 92, 93, 95, 96
                If IsNull(adoRecordset.Fields(i)) Then
                    strArr(i + 1) = ""
                Else
                    strArr(i + 1) = Trim(str(adoRecordset.Fields(i)))
                End If
            Case Else
                If IsNull(adoRecordset.Fields(i)) Then
                    strArr(i + 1) = ""
                Else
                    strArr(i + 1) = adoRecordset.Fields(i)
                End If
        End Select
'        DoEvents
    Next i
    
   strPA08 = "" & adoRecordset.Fields("PA08") 'Added by Morgan 2024/10/4
   strPA09 = "" & adoRecordset.Fields("PA09") 'Added by Lydia 2016/06/14
   
   'Added by Lydia 2018/12/27 中文本-頁數總計
   '-----與各式申請書的總計不同,有加上序列表
   i = Val(strArr(64)) + Val(strArr(65)) + Val(strArr(66)) + Val(strArr(67)) + Val(strArr(68))
   If i = 0 Then
       lblTot6 = Empty
   Else
       lblTot6 = i
   End If
   'end 2018/12/27
   
   'add by Toni 2008/10/20
   If strArr(150) = "" Then
      lbl1(86) = ""
   Else
      lbl1(86) = strArr(150) + "." + PUB_GetFCPGrpName(strArr(150))
   End If
   'end 2008/10/20
   
   'Add By Sindy 2010/10/27
   If strArr(158) = "" Then
      lbl1(91) = ""
   Else
      'Modify By Sindy 2014/7/8 +設計案的案件屬性
      'lbl1(91) = strArr(158) + "." + PUB_GetCaseAttributeName(strArr(158))
      lbl1(91) = strArr(158) + "." + PUB_GetCaseAttributeName(strArr(158), strArr(8))
   End If
   '2010/10/27 End
   
   'Add By Sindy 2014/10/27
   If strArr(139) = "" Then
      lbl1(92) = ""
   Else
      lbl1(92) = strArr(139)
   End If
   '2014/10/27 End
   
   'Add by Amy 2014/04/07 +pa164
    If strArr(164) = "" Then
      lbl1(164) = ""
   Else
      lbl1(164) = strArr(164)
   End If
   'end2014/04/07
   
   'Added by Lydia 2020/02/17 預設「名稱有特殊字」
   lblPA174.Visible = False
   CmdPA174.Visible = False
   If strArr(174) = "Y" Then
        lblPA174.Visible = True
        CmdPA174.Visible = True
   End If
   'end 2020/02/17
   
   'Add by Morgan 2009/12/24 PCT案,香港標準專利案,分割案要帶出提交日
   lblFilingDate(0).Visible = False
   lblFilingDate(1).Visible = False
   If (adoRecordset("pa09") = "013" And adoRecordset("pa08") = "1") Or (adoRecordset("pa46") = "Y") Then
      lblFilingDate(0).Visible = True
      lblFilingDate(1).Visible = True
      lblFilingDate(1) = ""
      strExc(0) = "select cp47 from caseprogress where cp01='" & Str01 & "' and cp02='" & Str02 & "' and cp03='" & Str03 & "' and cp04='" & Str04 & "' and cp10 in (" & NewCasePtyList & ") and cp57 is null and cp27>0 and cp47>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         lblFilingDate(1) = TransDate(RsTemp(0), 1)
      End If
   Else
      If adoRecordset("pa09") = "000" Then
         strExc(0) = "select cp27 from caseprogress where cp01='" & Str01 & "' and cp02='" & Str02 & "' and cp03='" & Str03 & "' and cp04='" & Str04 & "' and cp10='307' and cp57 is null"
      Else
         strExc(0) = "select cp47 from caseprogress where cp01='" & Str01 & "' and cp02='" & Str02 & "' and cp03='" & Str03 & "' and cp04='" & Str04 & "' and cp10='307' and cp57 is null"
         'Added by Morgan 2019/11/28  +CFP接續案(母案非本所案件)
         If Str01 = "CFP" And Str03 = "0" Then
            strExc(0) = strExc(0) & " union select cp47 from caseprogress where cp01='" & Str01 & "' and cp02='" & Str02 & "' and cp03='" & Str03 & "' and cp04='" & Str04 & "' and cp10 in ('122','113') and cp57 is null"
         End If
         'end 2019/11/28
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         lblFilingDate(0).Visible = True
         lblFilingDate(1).Visible = True
         If IsNull(RsTemp(0)) Then
            lblFilingDate(1) = ""
         Else
            lblFilingDate(1) = TransDate(RsTemp(0), 1)
         End If
      End If
   End If
   'end 2009/12/24
   '2010/2/22 add by sonia 抓修法次數
   'Remove by Lydia 2019/11/04
   'pa(1) = Str01
   'pa(2) = Str02
   'pa(3) = Str03
   'pa(4) = Str04
   'end 2019/11/04
   GetMoneyDate adoRecordset("pa08"), adoRecordset("pa09"), pa, strExc(0), strExc(1), , , m_FixNo
   '2010/2/22 end
   
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
   ShowNoData
   '920416 nick
   'Me.Hide
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
CheckOC
Dim strTemp As String    '暫存
Dim strTemp1 As Variant, strTemp2 As Variant, strTemp3 As Variant
Dim j As Integer
intK = 107
'Modify By Cheng 2002/07/29
'For i = 0 To 107
'For i = 0 To 132
For i = 1 To TF_PA 'edit by nickc 2006/07/12 T_PA
    Select Case i
    Case 1
         StrOk(0) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4)
         StrOkTxt(81) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4) 'Add By Sindy 2013/1/31
        'Modify By Cheng 2002/12/04
        '與基本檔維護的控制一樣
'         strSQL = "SELECT " & SQLDate("NP09") & " FROM NEXTPROGRESS WHERE NP02='" & StrArr(1) & "' AND NP03='" & StrArr(2) & "' AND NP04='" & StrArr(3) & "' AND NP05='" & StrArr(4) & "' AND NP07>=605 AND NP07<=607 AND NP06 IS NULL"
         '92.9.10 MODIFY BY SONIA
         'strSQL = "SELECT " & SQLDate("NP09") & " FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NULL AND NP09 Is Not Null Order By NP09 Desc  "
         
         'Modify by Morgan 2004/3/11
         '下次繳費日全部都顯示法定期限
'         If strArr(1) = "P" Then
'            '92.10.21 MODIFY BY SONIA
'            'strSQL = "SELECT MAX(" & SQLDate("NP08") & ") FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NULL AND NP01<'B' AND NP09 Is Not Null Order By NP09 Desc  "
'            strSQL = "SELECT MAX(" & SQLDate("NP08") & ") FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NULL AND NP09 Is Not Null Order By NP09 Desc  "
'            '92.10.21 END
'         Else
'            '92.10.21 MODIFY BY SONIA
'            'strSQL = "SELECT MAX(" & SQLDate("NP09") & ") FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NULL AND NP01<'B' AND NP09 Is Not Null Order By NP09 Desc  "
'            strSQL = "SELECT MAX(" & SQLDate("NP09") & ") FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NULL AND NP09 Is Not Null Order By NP09 Desc  "
'            '92.10.21 END
'         End If
'         '92.9.10 END
          '2011/1/27 MODIFY BY SONIA 百年蟲
          'strSql = "SELECT MAX(" & SQLDate("NP09") & "),MAX(" & SQLDate("NP08") & "), MAX(NP09) FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NULL AND NP09 Is Not Null Order By NP09 Desc  "
          'Modified by Morgan 2022/6/13 可能同時會有1個以上的期限 Ex:CFP-023718 年費&延展費, 秀玲:有未過期的帶出最小期限,都過期則帶最大期限(不同時間點可能看到不同期限??)
          'strSql = "SELECT " & SQLDate("NP09") & "," & SQLDate("NP08") & ",NP09 FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND (NP09||NP22) IN " & _
                   "(SELECT MAX(NP09||NP22) FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NULL AND NP09 Is Not NulL) "
          strSql = "SELECT " & SQLDate("NP09") & "," & SQLDate("NP08") & ",NP09 FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "'" & _
                   " AND NP07>=605 AND NP07<=607 And NP06 IS NULL AND NP09 Is Not NulL" & _
                   " order by sign(to_date(np09,'yyyymmdd')-sysdate)*np09 asc"
          'end 2022/6/13
          'Modify end 2004/3/11
         
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If adoRecordset.Fields(0) <> "" Then
               StrOk(14) = CheckStr(adoRecordset.Fields(0))
               'Add by Morgan 2004/3/11
               '下次繳費日本所期限，法定期限
               lbl1(75) = "" & adoRecordset.Fields(1)
               lbl1(76) = "" & adoRecordset.Fields(0)
               'Add by Morgan 2004/8/6
               '若繳費期限有延期過加6個月逾繳期限
               If PUB_IfCtrlDateExtended(strArr, Format(adoRecordset.Fields(2))) = True Then
                  StrOk(14) = StrOk(14) & " (6個月逾繳期限)"
               End If
            End If
         Else
            '92.9.10 MODIFY BY SONIA
            'StrOk(14) = ""
            
            'Modify by Morgan 2004/3/11
            '下次繳費日全部都顯示法定期限
'            If strArr(1) = "P" Then
'               strSQL = "SELECT MAX(" & SQLDate("NP08") & "||' '||DECODE(NP06,'Y',' (已收文)',' (不續辦)')) FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NOT NULL AND NP09 Is Not Null Order By NP09 Desc  "
'            Else
'               strSQL = "SELECT MAX(" & SQLDate("NP09") & "||' '||DECODE(NP06,'Y',' (已收文)',' (不續辦)')) FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NOT NULL AND NP09 Is Not Null Order By NP09 Desc  "
'            End If
            '2011/1/27 MODIFY BY SONIA 百年蟲 CFP-008574會抓到96年資料
            'strSql = "SELECT MAX(" & SQLDate("NP09") & "||' '||DECODE(NP06,'Y',' (已收文)',' (不續辦)')),MAX(" & SQLDate("NP09") & "),MAX(" & SQLDate("NP08") & ") FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NOT NULL AND NP09 Is Not Null Order By NP09 Desc  "
            strSql = "SELECT " & SQLDate("NP09") & "||' '||DECODE(NP06,'Y',' (已收文)',' (不續辦)')," & SQLDate("NP09") & "," & SQLDate("NP08") & " FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP09||NP22 IN " & _
                     "(SELECT MAX(NP09||NP22) FROM NEXTPROGRESS WHERE NP02='" & strArr(1) & "' AND NP03='" & strArr(2) & "' AND NP04='" & strArr(3) & "' AND NP05='" & strArr(4) & "' AND NP07>=605 AND NP07<=607 And NP06 IS NOT NULL AND NP09 Is Not Null) "
            'Modify end 2004/3/11
            
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If adoRecordset.Fields(0) <> "" Then
               StrOk(14) = CheckStr(adoRecordset.Fields(0))
                'Add by Morgan 2004/3/11
                '下次繳費日本所期限，法定期限
'edit by nickc 2005/08/25 放顛倒了
'                lbl1(75) = "" & adoRecordset.Fields(1)
'                lbl1(76) = "" & adoRecordset.Fields(2)
                lbl1(75) = "" & adoRecordset.Fields(2)
                lbl1(76) = "" & adoRecordset.Fields(1)
               'Add by Morgan 2004/2/9
               '檢查若為已收文狀態，若案件進度檔中案件性質為605~607者若都有發文日時則下次繳費日改空白
               If InStr(1, StrOk(14), "已收文") > 0 Then
                    strSql = "SELECT 1 FROM CASEPROGRESS WHERE CP10 BETWEEN '605' AND '607' AND CP01='" & strArr(1) & "' AND CP02='" & strArr(2) & "' AND CP03='" & strArr(3) & "' AND CP04='" & strArr(4) & "' AND CP27 IS NULL AND CP57 IS NULL"
                    CheckOC
                    adoRecordset.CursorLocation = adUseClient
                    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If adoRecordset.RecordCount = 0 Then
                        StrOk(14) = ""
                        'add by nickc 2005/05/24
                        '下次繳費日本所期限，法定期限
                        lbl1(75) = ""
                        lbl1(76) = ""
                    End If
               End If
              End If
            Else
               StrOk(14) = ""
                'Add by Morgan 2004/3/11
                '下次繳費日本所期限，法定期限
                lbl1(75) = ""
                lbl1(76) = ""
            End If
            '92.9.10 END
         End If
         CheckOC
               'add by nick 2004/07/13  fcp 和 P 605 時要檢查(P & FCP 只有 605 沒有 606 或 607)，且沒公告日
               'edit by nick 2004/08/13 避免沒有期限的也秀
               'If (UCase(strArr(1)) = "FCP" Or UCase(strArr(1)) = "P") And strArr(14) = "" Then
               'If (UCase(strArr(1)) = "FCP" Or UCase(strArr(1)) = "P") And strArr(14) = "" And lbl1(76).Caption <> "" Then
               If (UCase(strArr(1)) = "FCP" Or UCase(strArr(1)) = "P") And strArr(14) = "" And lbl1(76).Caption <> "" And strArr(9) = "000" Then
                     Label28.Visible = True
                     Label34.Visible = True
                     lbl1(14).ForeColor = &HC000&
               Else
                     Label28.Visible = False
                     Label34.Visible = False
                     lbl1(14).ForeColor = &H80000012
               End If
               'Modify by Amy 2014/04/10 +存取碼
               'Modified by Lydia 2016/10/19 +本所案號
               'strSql = "select PD05 AS  優先權日,PD06 AS 優先權號,NA03 AS 優先權國家,PD09 as 優先權存取碼 from PRIDATE,NATION WHERE PD01='" & strArr(1) & "' AND PD02='" & strArr(2) & "' AND PD03='" & strArr(3) & "' AND PD04 ='" & strArr(4) & "' AND PD07=NA01(+) ORDER BY PD01,PD02,PD03,PD04"
               strSql = "select PD05 AS  優先權日,PD06 AS 優先權號,NA03 AS 優先權國家,PD09 as 優先權存取碼,PA01||PA02||PA03||PA04 AS 本所案號 " & _
                        "From PRIDATE, Nation, PATENT " & _
                        "WHERE PD01='" & strArr(1) & "' AND PD02='" & strArr(2) & "' AND PD03='" & strArr(3) & "' AND PD04 ='" & strArr(4) & "' AND PD07=NA01(+) " & _
                        "AND PD06=PA11(+) AND PD05=PA10(+) AND PD07=PA09(+) ORDER BY PD01,PD02,PD03,PD04 "
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         Set grdDataList2.Recordset = adoRecordset
         CheckOC
         strSql = "SELECT CR05,CR06,CR07,CR08 FROM CASERELATION WHERE CR01='" & Str01 & "' AND CR02='" & Str02 & "' AND CR03='" & Str03 & "' AND CR04='" & Str04 & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 Then
              StrOk(57) = "有相關卷號資料"
         Else
              StrOk(57) = ""
         End If
         CheckOC
        'add by nickc 2005/12/15 檢查有無代表圖
        'Modify by Amy 2018/07/16  改寫至function
'        strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & strArr(1) & "' and ibf02='" & strArr(2) & "' and ibf03='" & strArr(3) & "' and ibf04='" & strArr(4) & "' and ibf05='1'"
'        CheckOC2
'        adoRecordset1.CursorLocation = adUseClient
'        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
        If ChkImgByteFile(strArr(1), strArr(2), strArr(3), strArr(4)) = True Then
            'Modified by Lydia 2020/09/18 拿掉(&I)
            cmdOK(6).Caption = "已設定代表圖"
            cmdOK(6).BackColor = &HC0FFC0
        Else
            'Modified by Lydia 2020/09/18 拿掉(&I)
            cmdOK(6).Caption = "未設定代表圖"
            cmdOK(6).BackColor = &HC0C0FF
        End If
'        CheckOC2
       'end 2018/07/16
    Case 8
         strSql = "SELECT SK02 FROM SYSTEMKIND WHERE SK01='" & strArr(1) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               strTemp = ""
            Else
               strTemp = str(adoRecordset.Fields(0))
            End If
            CheckOC
            strSql = "SELECT PTM03,PTM04 FROM PATENTTRADEMARKMAP WHERE PTM01='" & Val(strTemp) & "' AND PTM02='" & strArr(i) & "'"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                '91.08.19  nick   cfp 時只抓  ptm03
                If UCase(Str01) = "CFP" Then
                    If IsNull(adoRecordset.Fields(0)) Then
                         StrOk(1) = strArr(i) + ""
                    Else
                         StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(0)
                    End If
                Else
                    If strArr(9) = "000" Then
                        If IsNull(adoRecordset.Fields(0)) Then
                             StrOk(1) = strArr(i) + ""
                        Else
                             StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(0)
                        End If
                    Else
                        If IsNull(adoRecordset.Fields(1)) Then
                             StrOk(1) = strArr(i) + ""
                        Else
                             StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(1)
                        End If
                     End If
                End If
               'Add by Morgan 2004/1/2
               lbl1(1).ForeColor = vbBlack
            Else
               'Modify by Morgan 2004/1/2
               'StrOk(1) = ""
               lbl1(1).ForeColor = vbRed
               StrOk(1) = strArr(i)
            End If
            CheckOC
         Else
            CheckOC
            StrOk(1) = ""
         End If
    Case 5
         StrOk(2) = strArr(i)
         StrOkTxt(82) = strArr(i) 'Add By Sindy 2013/1/31
    Case 6
         StrOk(3) = strArr(i)
         StrOkTxt(83) = strArr(i) 'Add By Sindy 2013/1/31
    Case 7
         StrOk(4) = strArr(i)
         StrOkTxt(84) = strArr(i) 'Add By Sindy 2013/1/31
    Case 10
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(5) = ""
             StrOkTxt(86) = "" 'Add By Sindy 2013/1/31
         Else
             StrOk(5) = ChangeWStringToTString(strArr(i))
             StrOkTxt(86) = ChangeWStringToTString(strArr(i)) 'Add By Sindy 2013/1/31
         End If
    Case 12
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(6) = ""
         Else
             StrOk(6) = ChangeWStringToTString(strArr(i))
         End If
    Case 14
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(7) = ""
         Else
             StrOk(7) = ChangeWStringToTString(strArr(i))
         End If
    Case 21
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(8) = ""
         Else
             StrOk(8) = ChangeWStringToTString(strArr(i))
         End If
    Case 22
         StrOk(9) = strArr(i)
    Case 16
         StrOk(10) = strArr(i)
    Case 18
         StrOk(11) = strArr(i)
    Case 85
         StrOk(12) = strArr(i)
    Case 58
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(13) = ""
         Else
             StrOk(13) = ChangeWStringToTString(strArr(i))
         End If
    Case 23
         If strArr(i) = "2" Then
             StrOk(15) = " 2  異議"
         Else
             If strArr(i) = "3" Then
                 StrOk(15) = " 3  舉發"
             Else
                 If strArr(i) = "1" Then
                     StrOk(15) = " 1  申請"
                 Else
                     If Len(Trim(strArr(i))) = 0 Then
                         StrOk(15) = ""
                     Else
                         StrOk(15) = strArr(i) + "  錯誤代號"
                     End If
                 End If
             End If
         End If
    Case 9
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(16) = strArr(i) + ""
              Else
                  StrOk(16) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
            'Add by Morgan 2004/1/2
            lbl1(16).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/2
            'StrOk(16) = ""
            lbl1(16).ForeColor = vbRed
            StrOk(16) = strArr(i)
         End If
         CheckOC
    Case 11
         StrOk(17) = strArr(i)
         StrOkTxt(85) = strArr(i) 'Add By Sindy 2013/1/31
    Case 13
         StrOk(18) = strArr(i)
    Case 15
         StrOk(19) = strArr(i)
    Case 20
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(20) = ""
         Else
             StrOk(20) = ChangeWStringToTString(strArr(i))
         End If
    Case 24
         StrOk(21) = strArr(i)
    Case 25
         StrOk(22) = strArr(i)
    Case 17
         StrOk(23) = strArr(i)
    Case 19
         StrOk(24) = strArr(i)
    Case 57
         StrOk(25) = strArr(i)
    Case 59
         strSql = "SELECT ROR02 FROM REASONOFRELIEF WHERE ROR01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             If IsNull(adoRecordset.Fields(0)) Then
                     StrOk(26) = strArr(i) + ""
             Else
                     StrOk(26) = strArr(i) + "  " + adoRecordset.Fields(0)
             End If
         Else
             StrOk(26) = ""
         End If
         CheckOC
    Case 108
         'edit by nickc 2006/07/12
         'StrOk(27) = strArr(i)
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(27) = ""
         Else
             StrOk(27) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 26
         If Len(strArr(i)) = 9 Then
              strSql = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79,CU72 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              strSql = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79,CU72 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(28) = strArr(i) + ""
'                     Else
'                          StrOk(28) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(28) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(28) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
             ClsPDGetCustomerNameAndAddress Trim(strArr(i)), tmp02, tmp03, tmp04, tmp05
             StrOk(28) = strArr(i) + "  " + tmp02
             
             If IsNull(adoRecordset.Fields(3)) Then
                  StrOkTxt(22) = ""
             Else
                  StrOkTxt(22) = adoRecordset.Fields(3)
             End If
             If IsNull(adoRecordset.Fields(4)) Then
                 StrOk(60) = ""
             Else
                 StrOk(60) = adoRecordset.Fields(4)
             End If
            'Add by Morgan 2004/1/2
            lbl1(28).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/2
            'StrOk(28) = ""
            lbl1(28).ForeColor = vbRed
            StrOk(28) = strArr(i)
             
             StrOkTxt(22) = ""
             StrOk(60) = ""
         End If
         CheckOC
    Case 27
         If Len(strArr(i)) = 9 Then
              strSql = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              strSql = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(29) = strArr(i) + ""
'                     Else
'                          StrOk(29) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(29) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(29) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
             ClsPDGetCustomerNameAndAddress Trim(strArr(i)), tmp02, tmp03, tmp04, tmp05
             StrOk(29) = strArr(i) + "  " + tmp02
             
            'Add by Morgan 2004/1/2
            lbl1(29).ForeColor = vbBlack
         Else
            'Add by Morgan 2004/1/2
            'StrOk(29) = ""
            lbl1(29).ForeColor = vbRed
            StrOk(29) = strArr(i)
         End If
         CheckOC
    Case 28
         If Len(strArr(i)) = 9 Then
              strSql = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              strSql = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(30) = strArr(i) + ""
'                     Else
'                          StrOk(30) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(30) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(30) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
             ClsPDGetCustomerNameAndAddress Trim(strArr(i)), tmp02, tmp03, tmp04, tmp05
             StrOk(30) = strArr(i) + "  " + tmp02
             
            'Add by Morgan 2004/1/2
            lbl1(30).ForeColor = vbBlack
         Else
            'Add by Morgan 2004/1/2
            'StrOk(30) = ""
            lbl1(30).ForeColor = vbRed
            StrOk(30) = strArr(i)
         End If
         CheckOC
    Case 29
         If Len(strArr(i)) = 9 Then
              strSql = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              strSql = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(31) = strArr(i) + ""
'                     Else
'                          StrOk(31) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(31) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(31) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
             ClsPDGetCustomerNameAndAddress Trim(strArr(i)), tmp02, tmp03, tmp04, tmp05
             StrOk(31) = strArr(i) + "  " + tmp02
             
            'Add by Morgan 2004/1/2
            lbl1(31).ForeColor = vbBlack
         Else
            'Add by Morgan 2004/1/2
            'StrOk(31) = ""
            lbl1(31).ForeColor = vbRed
            StrOk(31) = strArr(i)
         End If
         CheckOC
    Case 30
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Len(strArr(i)) = 9 Then
'              strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'         Else
'              strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(32) = strArr(i) + ""
'                     Else
'                          StrOk(32) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(32) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(32) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
             tmp02 = ""
             If Trim(strArr(i)) <> "" Then
                ClsPDGetCustomerNameAndAddress Trim(strArr(i)), tmp02, tmp03, tmp04, tmp05
             End If
        If tmp02 <> "" Then
             StrOk(32) = strArr(i) + "  " + tmp02
             
            'Add by Morgan 2004/1/2
            lbl1(32).ForeColor = vbBlack
         Else
            'Add by Morgan 2004/1/2
            'StrOk(32) = ""
            lbl1(32).ForeColor = vbRed
            StrOk(32) = strArr(i)
         End If
         CheckOC
    Case 75
         If Len(strArr(i)) = 9 Then
              strSql = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29,FA39 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
         Else
              strSql = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29,FA39 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
         
'            '2005/9/14 MODIFY BY SONIA
'            'If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))) Then
'            If CheckStr(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))) = "" Then
'            '2005/9/14 END
'               'Modify By Cheng 2002/07/08
''               If IsNull(adoRecordset.Fields(1)) Then
'               If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))) Then
'                   If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(33) = strArr(i) + ""
'                   Else
'                         StrOk(33) = strArr(i) + "  " + CheckStr(adoRecordset.Fields(2))
'                   End If
'               Else
'                  'Modify By Cheng 2002/07/08
''                   StrOk(33) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                   StrOk(33) = strArr(i) + "  " + CheckStr(adoRecordset.Fields(IIf(strSK03 = "0", 0, 1)))
'               End If
'            Else
'               StrOk(33) = strArr(i) + "  " + CheckStr(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0)))
'            End If
            tmp02 = ""
            If Trim(strArr(i)) <> "" Then
                PUB_GetAgentName Str01, Trim(strArr(i)), tmp02
            End If
            StrOk(33) = strArr(i) + "  " + tmp02
            
            If IsNull(adoRecordset.Fields(3)) Then
                StrOkTxt(21) = ""
            Else
                StrOkTxt(21) = adoRecordset.Fields(3)
            End If
            If IsNull(adoRecordset.Fields(4)) Then
                 StrOk(59) = ""
            Else
                 StrOk(59) = adoRecordset.Fields(4)
            End If
            'Add by Morgan 2004/1/2
            lbl1(33).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/2
            'StrOk(33) = ""
            lbl1(33).ForeColor = vbRed
            StrOk(33) = strArr(i)
            
            StrOkTxt(21) = ""
            StrOk(59) = ""
         End If
         CheckOC
    Case 48
         StrOk(34) = strArr(i)
    Case 77
         StrOk(35) = strArr(i)
    Case 76
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Left$(strArr(i), 1) = "X" Then
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
'         Else
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            'Modify By Cheng 2002/07/08
'            If IsNull(adoRecordset.Fields(0)) Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(36) = strArr(i) + ""
'                    Else
'                        StrOk(36) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                     'Modify by Cheng 2002/07/08
''                    StrOk(36) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(36) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(36) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(36) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
        tmp02 = ""
        If Trim(strArr(i)) <> "" Then
            ClsLawLawGetName Trim(strArr(i)), tmp02
        End If
        If tmp02 <> "" Then
            StrOk(36) = strArr(i) + "  " + tmp02
            
            'Add by Morgan 2004/1/2
            lbl1(36).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/2
            'StrOk(36) = ""
            lbl1(36).ForeColor = vbRed
            StrOk(36) = strArr(i)
         End If
         CheckOC
    Case 51
         StrOk(37) = strArr(i)
    Case 49
         StrOk(38) = strArr(i)
    Case 86
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Left$(strArr(i), 1) = "X" Then
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
'         Else
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(39) = strArr(i) + ""
'                    Else
'                        StrOk(39) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                     'Modify By Cheng 2002/07/08
''                    StrOk(39) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(39) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(39) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(39) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
        tmp02 = ""
        If Trim(strArr(i)) <> "" Then
            ClsLawLawGetName Trim(strArr(i)), tmp02
        End If
        If tmp02 <> "" Then
            StrOk(39) = strArr(i) + "  " + tmp02
            
            'Add by Morgan 2004/1/5
            lbl1(39).ForeColor = vbBlack
         Else
            'Add by Morgan 2004/1/5
            'StrOk(39) = ""
            lbl1(39).ForeColor = vbRed
            StrOk(39) = strArr(i)
         End If
         CheckOC
    Case 88
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Left$(strArr(i), 1) = "X" Then
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
'         Else
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Modify by Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(40) = strArr(i) + ""
'                    Else
'                        StrOk(40) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                     'Modify By Cheng 2002/07/08
''                    StrOk(40) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(40) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(40) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(40) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsLawLawGetName Trim(strArr(i)), tmp02
         End If
         If tmp02 <> "" Then
            StrOk(40) = strArr(i) + "  " + tmp02
            
            'Add by Morgan 2004/1/5
            lbl1(40).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/5
            'StrOk(40) = ""
            lbl1(40).ForeColor = vbRed
            StrOk(40) = strArr(i)
         End If
         CheckOC
    Case 90
         StrOk(41) = strArr(i)
    Case 89
         StrOk(42) = strArr(i)
    Case 54
         StrOk(43) = strArr(i)
    Case 50
         StrOk(44) = strArr(i)
    Case 87
         StrOk(45) = strArr(i)
    Case 78
         StrOk(46) = strArr(i)
    Case 92
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               StrOk(47) = strArr(i) + ""
            Else
               StrOk(47) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            StrOk(47) = ""
         End If
         CheckOC
    Case 95
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             If IsNull(adoRecordset.Fields(0)) Then
                 StrOk(48) = strArr(i) + ""
             Else
                 StrOk(48) = strArr(i) + "  " + adoRecordset.Fields(0)
             End If
         Else
             StrOk(48) = ""
         End If
         CheckOC
    Case 93
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(49) = ""
         Else
             StrOk(49) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 96
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(50) = ""
         Else
             StrOk(50) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 94
         StrOk(51) = Format(strArr(i), "##:##")
    Case 97
         StrOk(52) = Format(strArr(i), "##:##")
    Case 31
         StrOkTxt(0) = strArr(i)
    Case 32
         StrOkTxt(1) = strArr(i)
    Case 33
         StrOkTxt(2) = strArr(i)
    Case 34
         StrOkTxt(3) = strArr(i)
    Case 35
         StrOkTxt(4) = strArr(i)
    Case 79
         StrOkTxt(5) = strArr(i)
    Case 82
         StrOkTxt(6) = strArr(i)
    Case 36
         StrOkTxt(7) = strArr(i)
    Case 37
         StrOkTxt(8) = strArr(i)
    Case 38
         StrOkTxt(9) = strArr(i)
    Case 39
         StrOkTxt(10) = strArr(i)
    Case 40
         StrOkTxt(11) = strArr(i)
    Case 80
         StrOkTxt(12) = strArr(i)
    Case 83
         StrOkTxt(13) = strArr(i)
    Case 41
         StrOkTxt(14) = strArr(i)
    Case 42
         StrOkTxt(15) = strArr(i)
    Case 43
         StrOkTxt(16) = strArr(i)
    Case 44
         StrOkTxt(17) = strArr(i)
    Case 45
         StrOkTxt(18) = strArr(i)
    Case 81
         StrOkTxt(19) = strArr(i)
    Case 84
         StrOkTxt(20) = strArr(i)
   'Modify By Sindy 2014/11/6 Mark 取消發明人欄位,另開新Table不限人數
'    Case 60
'         'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'         strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            For j = 0 To 20 Step 10
'                If IsNull(adoRecordset.Fields(j / 10)) Then
'                   StrOkTxt(23 + j) = ""
'                Else
'                   StrOkTxt(23 + j) = adoRecordset.Fields(j / 10)
'                End If
'             Next j
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbBlack
'         Else
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbRed
'
'             StrOkTxt(23) = ""
'             StrOkTxt(33) = ""
'             StrOkTxt(43) = ""
'         End If
'         CheckOC
'    Case 61
'         'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'         strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            For j = 0 To 20 Step 10
'               If IsNull(adoRecordset.Fields(j / 10)) Then
'                   StrOkTxt(24 + j) = ""
'               Else
'                   StrOkTxt(24 + j) = adoRecordset.Fields(j / 10)
'               End If
'            Next j
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbBlack
'         Else
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbRed
'
'            StrOkTxt(24) = ""
'            StrOkTxt(34) = ""
'            StrOkTxt(44) = ""
'         End If
'         CheckOC
'    Case 62
'         'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'         strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            For j = 0 To 20 Step 10
'               If IsNull(adoRecordset.Fields(j / 10)) Then
'                  StrOkTxt(25 + j) = ""
'               Else
'                  StrOkTxt(25 + j) = adoRecordset.Fields(j / 10)
'               End If
'            Next j
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbBlack
'         Else
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbRed
'
'            StrOkTxt(25) = ""
'            StrOkTxt(35) = ""
'            StrOkTxt(45) = ""
'         End If
'         CheckOC
'    Case 63
'         'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'         strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            For j = 0 To 20 Step 10
'               If IsNull(adoRecordset.Fields(j / 10)) Then
'                   StrOkTxt(26 + j) = ""
'               Else
'                   StrOkTxt(26 + j) = adoRecordset.Fields(j / 10)
'               End If
'            Next j
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbBlack
'         Else
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbRed
'
'            StrOkTxt(26) = ""
'            StrOkTxt(36) = ""
'            StrOkTxt(46) = ""
'         End If
'         CheckOC
'    Case 64
'         'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'         strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            For j = 0 To 20 Step 10
'               If IsNull(adoRecordset.Fields(j / 10)) Then
'                  StrOkTxt(27 + j) = ""
'               Else
'                  StrOkTxt(27 + j) = adoRecordset.Fields(j / 10)
'               End If
'            Next j
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbBlack
'         Else
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbRed
'
'            StrOkTxt(27) = ""
'            StrOkTxt(37) = ""
'            StrOkTxt(47) = ""
'         End If
'         CheckOC
'    Case 65
'         'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'         strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            For j = 0 To 20 Step 10
'               If IsNull(adoRecordset.Fields(j / 10)) Then
'                   StrOkTxt(28 + j) = ""
'               Else
'                   StrOkTxt(28 + j) = adoRecordset.Fields(j / 10)
'               End If
'            Next j
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbBlack
'         Else
'            'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbRed
'
'            StrOkTxt(28) = ""
'            StrOkTxt(38) = ""
'            StrOkTxt(48) = ""
'         End If
'         CheckOC
'    Case 66
'         'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'         strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            For j = 0 To 20 Step 10
'                If IsNull(adoRecordset.Fields(j / 10)) Then
'                    StrOkTxt(29 + j) = ""
'                Else
'                    StrOkTxt(29 + j) = adoRecordset.Fields(j / 10)
'                End If
'             Next j
'             'Add by Morgan 2004/1/2
'            lblInventor(i - 37).ForeColor = vbBlack
'          Else
'             'Add by Morgan 2004/1/2
'             lblInventor(i - 37).ForeColor = vbRed
'
'             StrOkTxt(29) = ""
'             StrOkTxt(39) = ""
'             StrOkTxt(49) = ""
'          End If
'          CheckOC
'    Case 67
'          'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'          strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'          CheckOC
'          adoRecordset.CursorLocation = adUseClient
'          adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'          If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'              For j = 0 To 20 Step 10
'                 If IsNull(adoRecordset.Fields(j / 10)) Then
'                    StrOkTxt(30 + j) = ""
'                 Else
'                    StrOkTxt(30 + j) = adoRecordset.Fields(j / 10)
'                 End If
'              Next j
'              'Add by Morgan 2004/1/2
'              lblInventor(i - 37).ForeColor = vbBlack
'          Else
'              'Add by Morgan 2004/1/2
'              lblInventor(i - 37).ForeColor = vbRed
'
'              StrOkTxt(30) = ""
'              StrOkTxt(40) = ""
'              StrOkTxt(50) = ""
'          End If
'          CheckOC
'    Case 68
'         'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'          strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'          CheckOC
'          adoRecordset.CursorLocation = adUseClient
'          adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'          If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'             For j = 0 To 20 Step 10
'                If IsNull(adoRecordset.Fields(j / 10)) Then
'                    StrOkTxt(31 + j) = ""
'                Else
'                    StrOkTxt(31 + j) = adoRecordset.Fields(j / 10)
'                End If
'             Next j
'             'Add by Morgan 2004/1/2
'              lblInventor(i - 37).ForeColor = vbBlack
'          Else
'             'Add by Morgan 2004/1/2
'             lblInventor(i - 37).ForeColor = vbRed
'
'             StrOkTxt(31) = ""
'             StrOkTxt(41) = ""
'             StrOkTxt(51) = ""
'          End If
'          CheckOC
'    Case 69
'         'Add by Morgan 2004/1/5
'         lblInventor(i - 37) = strArr(i)
'
'          strSql = "SELECT IN04,IN05,IN06 FROM INVENTOR WHERE IN01='" & Left$(strArr(i), 8) & "' AND IN02='" & Right$(Left$(strArr(i), 10), 2) & "'"
'          CheckOC
'          adoRecordset.CursorLocation = adUseClient
'          adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'          If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'             For j = 0 To 20 Step 10
'                If IsNull(adoRecordset.Fields(j / 10)) Then
'                   StrOkTxt(32 + j) = ""
'                Else
'                   StrOkTxt(32 + j) = adoRecordset.Fields(j / 10)
'                End If
'             Next j
'             'Add by Morgan 2004/1/2
'              lblInventor(i - 37).ForeColor = vbBlack
'           Else
'             'Add by Morgan 2004/1/2
'             lblInventor(i - 37).ForeColor = vbRed
'
'             StrOkTxt(32) = ""
'             StrOkTxt(42) = ""
'             StrOkTxt(52) = ""
'           End If
'           CheckOC
     Case 72     '繳費年度
           strTemp1 = Split(strArr(i), ",")
           'grdDataList1.Rows = UBound(StrTemp1) + 1
           'Modify By Sindy 2009/06/29 改顯示繳費次數及年度說明
           '繳費次數
           grdDataList1.col = 1
           For j = 0 To UBound(strTemp1)
               grdDataList1.row = j + 1
               grdDataList1.Text = j + 1
           Next j
           '年度說明
           grdDataList1.col = 2
           'Modified by Morgan 2022/6/13 +strArr(10)
           strFeeType = PUB_GetNa20Na22Na24(strArr(9), strArr(8), strArr(10))
           For j = 0 To UBound(strTemp1)
               'strYF15 = PUB_GetYF15(strArr(9), strArr(8), "Y0000000", strFeeType, j + 1)
               '2010/2/22 modify by sonia
               'strYF15 = PUB_GetYF15(strArr(9), strArr(8), "Y0000000", strFeeType, CDbl(strTemp1(j)))
               strYF15 = PUB_GetYF15(strArr(9), strArr(8), "Y000000" & m_FixNo, strFeeType, CDbl(strTemp1(j)))
               grdDataList1.row = j + 1
               grdDataList1.Text = strYF15
           Next j
    Case 73
          strTemp1 = Split(strArr(i), ",")
          'grdDataList1.Rows = UBound(StrTemp1) + 1
          'Modify By Sindy 2009/06/29
          'grdDataList1.col = 2
          grdDataList1.col = 3
          For j = 0 To UBound(strTemp1)
              grdDataList1.row = j + 1
              'Modify by Morgan 2005/1/7 改民國年
              'grdDataList1.Text = strTemp1(j)
               If strTemp1(j) <> "" Then
                  grdDataList1.Text = ChangeTStringToTDateString(strTemp1(j) - 19110000)
               End If
          Next j
    Case 74
          strTemp1 = Split(strArr(i), ",")
          'grdDataList1.Rows = UBound(StrTemp1) + 1
          'Modify By Sindy 2009/06/29
          'grdDataList1.col = 3
          grdDataList1.col = 4
          For j = 0 To UBound(strTemp1)
              grdDataList1.row = j + 1
              grdDataList1.Text = strTemp1(j)
          Next j
    Case 91
         StrOkTxt(53) = strArr(i)
    Case 106
         StrOk(53) = strArr(i)
    Case 105
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Left$(strArr(i), 1) = "X" Then
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
'         Else
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                  'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(54) = strArr(i) + ""
'                    Else
'                        StrOk(54) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                     'Modify By Cheng 2002/07/08
''                    StrOk(54) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(54) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(54) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(54) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsLawLawGetName Trim(strArr(i)), tmp02
         End If
         If tmp02 <> "" Then
            StrOk(54) = strArr(i) + "  " + tmp02
            
            'Add by Morgan 2004/1/5
            lbl1(54).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/5
            'StrOk(54) = ""
            lbl1(54).ForeColor = vbRed
            StrOk(54) = strArr(i)
         End If
         CheckOC
    Case 70
          StrOk(56) = strArr(i)
    Case 71
         StrOk(55) = strArr(i)
    Case 107
         StrOk(58) = strArr(i)
    Case 52
         StrOk(61) = strArr(i)
    Case 53
         StrOk(62) = strArr(i)
    Case 55
         StrOk(63) = strArr(i)
    Case 56
         StrOk(64) = strArr(i)
    Case 98
         StrOk(65) = strArr(i)
    Case 99
         StrOk(66) = strArr(i)
    Case 100
         StrOk(67) = strArr(i)
    Case 101
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Left$(strArr(i), 1) = "X" Then
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
'         Else
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                  'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(68) = strArr(i) + ""
'                    Else
'                        StrOk(68) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                     'Modify By Cheng 2002/07/08
''                    StrOk(68) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(68) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(68) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(68) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
        tmp02 = ""
        If Trim(strArr(i)) <> "" Then
            ClsLawLawGetName Trim(strArr(i)), tmp02
        End If
        If tmp02 <> "" Then
            StrOk(68) = strArr(i) + "  " + tmp02
            
            'Add by Morgan 2004/1/5
            lbl1(68).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/5
            'StrOk(68) = ""
            lbl1(68).ForeColor = vbRed
            StrOk(68) = strArr(i)
         End If
         CheckOC
    Case 102
         StrOk(69) = strArr(i)
    Case 103
         StrOkTxt(54) = strArr(i)
    Case 104
         StrOkTxt(55) = strArr(i)
    Case 47
         StrOk(70) = strArr(i)
    'Add By Cheng 2003/05/30
    Case 46 '是否PCT案
         StrOk(71) = IIf(strArr(i) = "Y", "是", "否")
    Case 109
         StrOkTxt(56) = strArr(i)
    Case 110
         StrOkTxt(57) = strArr(i)
    Case 111
         StrOkTxt(58) = strArr(i)
    Case 112
         StrOkTxt(59) = strArr(i)
    Case 113
         StrOkTxt(60) = strArr(i)
    Case 114
         StrOkTxt(61) = strArr(i)
    Case 115
         StrOkTxt(62) = strArr(i)
    Case 116
         StrOkTxt(63) = strArr(i)
    Case 117
         StrOkTxt(64) = strArr(i)
    Case 118
         StrOkTxt(65) = strArr(i)
    Case 119
         StrOkTxt(66) = strArr(i)
    Case 120
         StrOkTxt(67) = strArr(i)
    Case 121
         StrOkTxt(68) = strArr(i)
    Case 122
         StrOkTxt(69) = strArr(i)
    Case 123
         StrOkTxt(70) = strArr(i)
    Case 124
         StrOkTxt(71) = strArr(i)
    Case 125
         StrOkTxt(72) = strArr(i)
    Case 126
         StrOkTxt(73) = strArr(i)
    Case 127
         StrOkTxt(74) = strArr(i)
    Case 128
         StrOkTxt(75) = strArr(i)
    Case 129
         StrOkTxt(76) = strArr(i)
    Case 130
         StrOkTxt(77) = strArr(i)
    Case 131
         StrOkTxt(78) = strArr(i)
    Case 132
         StrOkTxt(79) = strArr(i)
    Case 133 'D/N固定列印對象
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Left$(strArr(i), 1) = "X" Then
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
'         Else
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(72) = strArr(i) + ""
'                    Else
'                        StrOk(72) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(72) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'                StrOk(72) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsLawLawGetName Trim(strArr(i)), tmp02
         End If
         If tmp02 <> "" Then
            StrOk(72) = strArr(i) + "  " + tmp02
            
            'Add by Morgan 2004/1/5
            lbl1(72).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/5
            'StrOk(72) = ""
            lbl1(72).ForeColor = vbRed
            StrOk(72) = strArr(i)
         End If
         CheckOC
    Case 134 '年費D/N列印對象
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Left$(strArr(i), 1) = "X" Then
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
'         Else
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(73) = strArr(i) + ""
'                    Else
'                        StrOk(73) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(73) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'                StrOk(73) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsLawLawGetName Trim(strArr(i)), tmp02
        End If
        If tmp02 <> "" Then
            StrOk(73) = strArr(i) + "  " + tmp02
            
            'Add by Morgan 2004/1/5
            lbl1(73).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/5
            'StrOk(73) = ""
            lbl1(73).ForeColor = vbRed
            StrOk(73) = strArr(i)
         End If
         CheckOC
    Case 135 '年費聯絡人
        StrOk(74) = strArr(i)
    'add by nickc 2006/07/12
    Case 136
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             lbl1(77) = ""
         Else
             lbl1(77) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 137
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               lbl1(78) = strArr(i) + ""
            Else
               lbl1(78) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            lbl1(78) = ""
         End If
         CheckOC
    Case 138
         lbl1(79) = strArr(i)
   
    'Add by Morgan 2012/5/9
    Case 140
         lbl1(140) = ChangeWStringToTString(strArr(i))
         
    'Add by Morgan 2007/10/29
    Case 141
         lbl1(80) = strArr(i)
   'Add by Morgan 2008/1/17
    Case 142
         lbl1(81) = strArr(i)
         
   'Add by Morgan 2008/2/27
    Case 143
         lbl1(82) = strArr(i)
   'Add by Morgan 2008/4/10
    Case 146
         lbl1(83) = strArr(i)
   'Add by Morgan 2008/6/3
    Case 147
         lbl1(84) = strArr(i)
   'Add by Morgan 2008/6/10
    Case 148
         txt1(80) = strArr(i)
    'Add by Morgan 2008/8/4
    Case 149
         lbl1(85) = PUB_GetContact(strArr(26), strArr(i))
    'Add by Morgan 2009/10/16
    Case 153
      lbl1(87) = strArr(i)
    Case 154
      lbl1(88) = strArr(i)
    Case 155
      lbl1(89) = strArr(i)
    'end 2009/10/16
    Case 157 'Add by Morgan 2010/6/18
      lbl1(90) = strArr(i)
    Case 159 'Add by Morgan 2010/11/5
      lbl1(159) = strArr(i)
    Case 160 'Add by Sindy 2012/3/2
      lbl1(160) = strArr(i)
    
    'Added by Morgan 2014/8/29
    Case 166
      'Modified by Morgan 2015/6/11
      'tmp02 = ""
      'If Trim(strArr(i)) <> "" Then
      '   ClsLawLawGetName Trim(strArr(i)), tmp02
      'End If
      'If tmp02 <> "" Then
      '   strArr(i) = strArr(i) + "  " + tmp02
      'End If
      If strArr(i) <> "" Then
         strExc(1) = ""
         If Len(strArr(i)) = 9 Then
            strExc(1) = strArr(i) & " " & GetName(strArr(i))
         Else
            strExc(0) = "select fa01||fa02||' '||nvl(fa06, nvl(rtrim(fa05||' '||fa63||' '||fa64||' '||fa65),fa04)),instr('" & strArr(i) & "',fa01||fa02) srt from fagent where instr('" & strArr(i) & "',fa01||fa02)>0" & _
               " union all select cu01||cu02||' '||nvl(cu06, nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu04)),instr('" & strArr(i) & "',cu01||cu02) srt from customer where instr('" & strArr(i) & "',cu01||cu02)>0" & _
               " order by srt"
        
            intI = 1
            Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = Replace("" & adoRecordset(0), vbCrLf, "")
               adoRecordset.MoveNext
               Do While Not adoRecordset.EOF
                  strExc(1) = strExc(1) & vbCrLf & Replace("" & adoRecordset(0), vbCrLf, "")
                  adoRecordset.MoveNext
               Loop
            End If
         End If
         SetList lstPA166, strExc(1)
      Else
         lstPA166.Clear
      End If
      'end 2015/6/11
   
    Case 167 'Added by Morgan 2015/5/27
      lbl1(167) = strArr(i)
    
    'Added by Morgan 2016/12/8
    Case 168
      lbl1(168) = strArr(i)
      If Trim(strArr(i)) <> "" Then
         If ClsLawLawGetName(strArr(i), strExc(9)) = True Then
            lbl1(168) = lbl1(168) + "  " + strExc(9)
         End If
      End If
    Case 169
      If strArr(168) <> "" And strArr(i) <> "" Then
         lbl1(169) = PUB_GetContact(strArr(168), strArr(i))
      Else
         lbl1(169) = ""
      End If
    'end 2016/12/8
    
    'Add By Sindy 2016/11/24
    Case 170
      lbl1(170) = strArr(i)
    Case 171
      If strArr(i) <> "" Then
         Combo3(1).ListIndex = Val(strArr(i))
      End If
    '2016/11/24 END
    
    'Added by Morgan 2021/6/29
    Case 176   '是否新藥專利
      lbl1(176) = strArr(i)
      
    'Added by Lydia 2021/08/16
    Case 177   'FCP專利連結通知
      lbl1(177) = strArr(i)
      
    'Add by Sindy 2025/1/7
    Case 180
      If Trim(strArr(i)) <> "" Then
         arrID = Split(strArr(i), ",")
         For intI = UBound(arrID) To LBound(arrID) Step -1
            Chk1K(Val(arrID(intI)) - 1).Value = 1
         Next intI
      End If
      '2025/1/7 END
      
   'Added by Morgan 2025/2/10
    Case 181   '專利不可請雜費
      lbl1(181) = strArr(i)
      
    Case Else
    
End Select
    'DoEvents
Next i

'Add by Morgan 2008/11/13 改index與欄位序次相同的陣列，將來再新增欄位時只需加畫面的物件並指定相同的index就好
For Each oLbl In lblPA
   oLbl = strArr(oLbl.Index)
   oLbl.BackColor = &H8000000F
Next

'Added by Lydia 2019/11/04
If pa(1) = "FCP" And lblPA(162) = "Y" Then
   cmdDivSug.Visible = True
Else
   cmdDivSug.Visible = False
End If
'end 2019/11/04

'Modify By Cheng 2003/05/30
'For i = 0 To 70
For i = 0 To 75 - 1 '2006/07/12 加備註，以後新增欄位，直接在上面修改，此2段迴圈
   If i <> 0 And i <> 2 And i <> 3 And i <> 4 And i <> 5 And i <> 6 And _
      i <> 17 And i <> 18 And i <> 19 Then  'Add By Sindy 2013/1/31 +if
      lbl1(i) = StrOk(i) '不可修改，不然會影響資料顯現，而且陣列的宣告也不用一直的修改
   End If
Next i
For i = 0 To 86 '79
   'Modify By Sindy 2014/11/6 +And (i < 23 Or i > 52)
   If i <> 80 And (i < 23 Or i > 52) Then 'Add By Sindy 2013/1/31 +if
      txt1(i) = StrOkTxt(i)
   End If
Next i
'Add By Sindy 2021/2/5
txt1(23) = StrOk(18) '公開號
txt1(24) = StrOk(19) '公告號
txt1(25) = StrOk(6) '公開日
'2021/2/5 END
StrTag = strArr(75)
StrTag1 = strArr(26)
'Add By Sindy 2010/02/04
cmdOK(7).Visible = False
cmdOK(8).Visible = False
cmdOK(9).Visible = False
cmdOK(10).Visible = False
StrTag2 = strArr(27)
StrTag3 = strArr(28)
StrTag4 = strArr(29)
StrTag5 = strArr(30)
If Trim(StrTag2) <> "" Then cmdOK(7).Visible = True
If Trim(StrTag3) <> "" Then cmdOK(8).Visible = True
If Trim(StrTag4) <> "" Then cmdOK(9).Visible = True
If Trim(StrTag5) <> "" Then cmdOK(10).Visible = True
'2010/02/04 End
'Add By Cheng 2003/06/16
'避免&符號無法正常顯示
'Me.lbl1(33).Caption = Replace(Me.lbl1(33).Caption, "&", "&&") 'Removed by Morgan 2020/6/9 已改在下面統一轉換
'add by nickc 2005/05/30  檢查有無分割或相關卷號
     cmdOK(5).Visible = ChkDataBy308(txt1(81).Text)
     cmdOK(4).Visible = ChkDataByCR(txt1(81).Text)
    'Added by Lydia 2015/11/03　顯示一案兩請，擬制喪失新穎性案件
    lblCaseMap.Caption = ""
    lblCaseMap2.Caption = "" 'Added by Lydia 2019/11/28
    If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "3") = True Then
       lblCaseMap.Caption = "一案兩請"
    End If
    If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "6") = True Then
       'Modified by Lydia 2019/11/28 P-123733有一案兩請和擬制喪失新穎性案件
       'lblCaseMap.Caption = "擬制喪失新穎性案件"
       lblCaseMap2.Caption = "擬制喪失新穎性案件"
    End If
    'end 2015/11/03
    'Added by Lydia 2016/06/14 +台灣大陸案件提示
    lblCMboth.Caption = ""
    If (Str01 = "P" Or Str01 = "FCP") And strPA09 = 台灣國家代號 Then
       If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "0", "A", 大陸國家代號) Then
          lblCMboth.Caption = "有大陸案"
       End If
    ElseIf Str01 = "P" And strPA09 = 大陸國家代號 Then
       If PUB_GetRefCaseChk(Str01, Str02, Str03, Str04, "CASEMAP", "0", "A", 台灣國家代號) Then
          lblCMboth.Caption = "有台灣案"
       End If
    End If
    'end 2016/06/14

   'Add by Sindy 2014/11/6
   'Modify by Sindy 2019/4/18 + 國籍
   GRD1.Clear
   SetGrd
   StrSQLa = "SELECT pi05 as 序號,pi06 as 發明人編號,in04 as 中文名稱,in05 as 英文名稱,in06 as 日文名稱,na03 國籍 from PatentInventor,Inventor,nation where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
             " and substr(pi06,1,8)=in01(+) and substr(pi06,9,2)=in02(+) and in11=na01(+)" & _
             " order by pi05 asc"
   If rsA.State <> adStateClosed Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      Set GRD1.Recordset = rsA
   End If
   '2014/11/6 END
   
   
   
    'Modified by Morgan 2018/4/11 LABEL 若有 "&" 時要改為 "&&" 否則會變 "_"
    For Each oLbl In lbl1
      oLbl.Caption = Replace(oLbl.Caption, "&", "&&")
    Next
    'end 2018/4/11
    
   'Added by Morgan 2023/3/30
   If pa(1) = "CFP" Then
      lbl1(179) = PUB_GetP_PA91Individual(pa(1), pa(2), pa(3), pa(4))
   Else
      lbl1(179) = ""
   End If
   'end 2023/3/30
    
   'Added by Lydia 2023/03/09 PA176改說明
   m_bolFMP = PUB_ChkIsFMP(pa(1), pa(2), pa(3), pa(4))
   If m_bolFMP = True Then
       lblPA176.Caption = "專利權期間延長相關:"
       lblPA176.Width = 1000
   Else
       lblPA176.Caption = "是否新藥專利:"
       lblPA176.Width = 1200
   End If
   'end 2023/03/09
   
   'Added by Morgan 2024/10/28
   '大陸發明案的專用期止日增加判斷是否需帶補償天數及補償年費期限
   If pa(1) = "P" And strPA09 = "020" And strPA08 = "1" And lbl1(22) <> "" Then
      If lbl1(14) = "" Then lbl1(14) = PUB_GetCN615DueDate(pa())
      If PUB_GetCNExtDays(pa(), DBDATE(lbl1(22)), i) = True Then
         If i > 0 Then
            lbl1(22) = lbl1(22) & " (含補償 " & i & " 天)"
         End If
      End If
   End If
   'end 2024/10/28
   
   Set rsA = Nothing
End Sub

'Add By Sindy 2014/11/6
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify by Sindy 2019/4/18 + 國籍
   arrGridHeadText = Array("序號", "發明人編號", "中文名稱", "英文名稱", "日文名稱", "國籍")
   arrGridHeadWidth = Array(400, 1000, 2000, 2000, 2000, 1000)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Function GRIDHEAND()
With grdDataList1
.row = 0
.col = 0
.ColWidth(0) = 1000
'Modify By Sindy 2009/06/29
'.col = 1
'.ColWidth(1) = 1200
'.Text = "繳費年度"
'.col = 2
'.ColWidth(2) = 1200
'.Text = "繳費日期"
'.col = 3
'.ColWidth(3) = 1800
'.Text = "費用是否雙倍"
.col = 1
.ColWidth(1) = 500
.Text = "次數"
.ColAlignment = flexAlignCenterCenter
.col = 2
.ColWidth(2) = 2500
.Text = "年度說明"
.ColAlignment = flexAlignLeftCenter
.col = 3
.ColWidth(3) = 900
.Text = "繳費日期"
.ColAlignment = flexAlignRightCenter
.col = 4
.ColWidth(4) = 900
'Modified by Morgan 2021/3/24
'.Text = "費用雙倍"
.Text = "逾期補繳"
.ColAlignment = flexAlignLeftCenter
'2009/06/29 End
End With
With grdDataList2
.row = 0
.col = 0
.ColWidth(0) = 1000
.Text = "優先權日"
.col = 1
.ColWidth(1) = 3000
.Text = "優先權號"
.col = 2
.ColWidth(2) = 1000
.Text = "優先權國家"
'Add by Amy 2014/04/10
.col = 3
.ColWidth(3) = 1300
.Text = "優先權存取碼"
'end2014/04/10
'Added by Lydia 2016/10/19
.col = 4
.ColWidth(4) = 1300
.Text = "本所案號"
'end 2016/10/19
End With
End Function

Private Function Grid()
'Modify by Morgan 2004/3/11
'取消第0列內容
'Dim i As Long
'With grdDataList1
'.Col = 0
'For i = 1 To 20
'.Row = i
'.Text = "第" & i & "次"
'Next
'Modify end 2004/3/11
'End With
grdDataList1.ColWidth(0) = 0
End Function

Private Sub Form_Unload(Cancel As Integer)
pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
Set frm100101_3 = Nothing
End Sub

Private Function GetName(pNo As String) As String
   
   If Left(pNo, 1) = "Y" Then
      strExc(0) = "select nvl(fa06, nvl(rtrim(fa05||' '||fa63||' '||fa64||' '||fa65),fa04)) from fagent where fa01='" & Left(pNo, 8) & "' and fa02='" & Mid(pNo, 9) & "'"
   Else
      strExc(0) = "select nvl(cu06, nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu04)) from customer where cu01='" & Left(pNo, 8) & "' and cu02='" & Mid(pNo, 9) & "'"
   End If
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetName = "" & adoRecordset(0)
   End If
End Function

'Modified by Lydia 2021/12/23 As ListBox=> As Object
Private Sub SetList(pList As Object, pText As String)
   Dim arrList() As String, ii As Integer
   
   pList.Clear
   If pText <> "" Then
      arrList = Split(pText, vbCrLf)
      For ii = LBound(arrList) To UBound(arrList)
         If arrList(ii) <> "" Then
            pList.AddItem arrList(ii)
         End If
      Next
   End If
End Sub

'Added by Lydia 2016/10/27 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
   'Added by Lydia 2021/12/15 Focus在Form 2.0元件，關閉Default_Button的功能
   If cmdOK(0).Default = True Then
       cmdOK(0).Default = False
   End If
End Sub

'Added by Lydia 2019/11/04
Private Sub cmdDivSug_Click()
   strExc(0) = "select dst05 from divsugtext where dst01='" & pa(1) & "' and dst02='" & pa(2) & "' and dst03='" & pa(3) & "' and dst04='" & pa(4) & "' and dst05 is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Mofieid by Morgan 2022/10/11
      'MsgBox RsTemp(0), , "分割建議"
      MsgBoxU RsTemp(0), , "分割建議"
   Else
      MsgBox "未輸入分割建議!!"
   End If
End Sub

'Added by Lydia 2020/02/17 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
