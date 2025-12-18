VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050701 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利案件基本資料維護"
   ClientHeight    =   7668
   ClientLeft      =   1764
   ClientTop       =   1860
   ClientWidth     =   8952
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7668
   ScaleWidth      =   8952
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   30
      TabIndex        =   165
      Top             =   930
      Width           =   8895
      _ExtentX        =   15706
      _ExtentY        =   11896
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      Tab             =   2
      TabsPerRow      =   11
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm050701.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(28)"
      Tab(0).Control(1)=   "Label1(27)"
      Tab(0).Control(2)=   "Label1(26)"
      Tab(0).Control(3)=   "Label1(25)"
      Tab(0).Control(4)=   "Label1(24)"
      Tab(0).Control(5)=   "Label1(23)"
      Tab(0).Control(6)=   "Label1(22)"
      Tab(0).Control(7)=   "Label1(21)"
      Tab(0).Control(8)=   "Label1(19)"
      Tab(0).Control(9)=   "Label1(18)"
      Tab(0).Control(10)=   "Label1(17)"
      Tab(0).Control(11)=   "Label1(16)"
      Tab(0).Control(12)=   "Label1(15)"
      Tab(0).Control(13)=   "Label1(14)"
      Tab(0).Control(14)=   "Label1(13)"
      Tab(0).Control(15)=   "Label1(12)"
      Tab(0).Control(16)=   "Label1(11)"
      Tab(0).Control(17)=   "Label1(10)"
      Tab(0).Control(18)=   "Label1(9)"
      Tab(0).Control(19)=   "Label1(8)"
      Tab(0).Control(20)=   "Label1(7)"
      Tab(0).Control(21)=   "Label1(6)"
      Tab(0).Control(22)=   "Label1(5)"
      Tab(0).Control(23)=   "Label1(4)"
      Tab(0).Control(24)=   "Label1(3)"
      Tab(0).Control(25)=   "Label1(2)"
      Tab(0).Control(26)=   "Label1(1)"
      Tab(0).Control(27)=   "Label1(0)"
      Tab(0).Control(28)=   "Label2(0)"
      Tab(0).Control(29)=   "Label2(1)"
      Tab(0).Control(30)=   "Label2(5)"
      Tab(0).Control(31)=   "Label2(50)"
      Tab(0).Control(32)=   "Label2(44)"
      Tab(0).Control(33)=   "Label1(161)"
      Tab(0).Control(34)=   "lblFilingDate(1)"
      Tab(0).Control(35)=   "lblFilingDate(0)"
      Tab(0).Control(36)=   "Label1(167)"
      Tab(0).Control(37)=   "Label1(168)"
      Tab(0).Control(38)=   "Label1(170)"
      Tab(0).Control(39)=   "Label1(171)"
      Tab(0).Control(40)=   "Label1(176)"
      Tab(0).Control(41)=   "lblCaseMap"
      Tab(0).Control(42)=   "lblCMboth"
      Tab(0).Control(43)=   "lblCaseMap2"
      Tab(0).Control(44)=   "Label1(130)"
      Tab(0).Control(45)=   "Text1(46)"
      Tab(0).Control(46)=   "Text1(17)"
      Tab(0).Control(47)=   "Text1(9)"
      Tab(0).Control(48)=   "Text1(85)"
      Tab(0).Control(49)=   "Text1(10)"
      Tab(0).Control(50)=   "Text1(12)"
      Tab(0).Control(51)=   "Text1(14)"
      Tab(0).Control(52)=   "Text1(21)"
      Tab(0).Control(53)=   "Text1(22)"
      Tab(0).Control(54)=   "Text1(25)"
      Tab(0).Control(55)=   "Text1(5)"
      Tab(0).Control(56)=   "Text1(6)"
      Tab(0).Control(57)=   "Text1(7)"
      Tab(0).Control(58)=   "Text1(3)"
      Tab(0).Control(59)=   "Text1(4)"
      Tab(0).Control(60)=   "Text1(16)"
      Tab(0).Control(61)=   "Text1(18)"
      Tab(0).Control(62)=   "Text1(58)"
      Tab(0).Control(63)=   "Text1(11)"
      Tab(0).Control(64)=   "Text1(13)"
      Tab(0).Control(65)=   "Text1(15)"
      Tab(0).Control(66)=   "Text1(20)"
      Tab(0).Control(67)=   "Text1(24)"
      Tab(0).Control(68)=   "Text1(19)"
      Tab(0).Control(69)=   "Text1(57)"
      Tab(0).Control(70)=   "Text1(59)"
      Tab(0).Control(71)=   "Text1(47)"
      Tab(0).Control(72)=   "Text1(8)"
      Tab(0).Control(73)=   "Text1(2)"
      Tab(0).Control(74)=   "Text1(1)"
      Tab(0).Control(75)=   "Text1(157)"
      Tab(0).Control(76)=   "Text1(160)"
      Tab(0).Control(77)=   "Text1(140)"
      Tab(0).Control(78)=   "Text1(164)"
      Tab(0).Control(79)=   "Text1(176)"
      Tab(0).Control(80)=   "Text1(23)"
      Tab(0).Control(81)=   "Label1(132)"
      Tab(0).Control(82)=   "Text1(178)"
      Tab(0).Control(83)=   "lblPA176"
      Tab(0).Control(84)=   "Label1(133)"
      Tab(0).Control(85)=   "Combo2"
      Tab(0).Control(86)=   "Combo3"
      Tab(0).Control(87)=   "FraPA174"
      Tab(0).Control(88)=   "CmdPA174"
      Tab(0).Control(89)=   "Combo6"
      Tab(0).ControlCount=   90
      TabCaption(1)   =   "申請人"
      TabPicture(1)   =   "frm050701.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(48)"
      Tab(1).Control(1)=   "Text1(30)"
      Tab(1).Control(2)=   "Text1(29)"
      Tab(1).Control(3)=   "Text1(28)"
      Tab(1).Control(4)=   "Text1(27)"
      Tab(1).Control(5)=   "Text1(26)"
      Tab(1).Control(6)=   "Text1(45)"
      Tab(1).Control(7)=   "Text1(40)"
      Tab(1).Control(8)=   "Text1(35)"
      Tab(1).Control(9)=   "Text1(44)"
      Tab(1).Control(10)=   "Text1(39)"
      Tab(1).Control(11)=   "Text1(34)"
      Tab(1).Control(12)=   "Text1(43)"
      Tab(1).Control(13)=   "Text1(38)"
      Tab(1).Control(14)=   "Text1(33)"
      Tab(1).Control(15)=   "Text1(42)"
      Tab(1).Control(16)=   "Text1(37)"
      Tab(1).Control(17)=   "Text1(32)"
      Tab(1).Control(18)=   "Text1(41)"
      Tab(1).Control(19)=   "Text1(36)"
      Tab(1).Control(20)=   "Text1(31)"
      Tab(1).Control(21)=   "Label1(160)"
      Tab(1).Control(22)=   "Label1(110)"
      Tab(1).Control(23)=   "Label1(109)"
      Tab(1).Control(24)=   "Label1(108)"
      Tab(1).Control(25)=   "Label1(107)"
      Tab(1).Control(26)=   "Label1(106)"
      Tab(1).Control(27)=   "Label1(105)"
      Tab(1).Control(28)=   "Label1(104)"
      Tab(1).Control(29)=   "Label1(103)"
      Tab(1).Control(30)=   "Label1(102)"
      Tab(1).Control(31)=   "Label1(101)"
      Tab(1).Control(32)=   "Label1(95)"
      Tab(1).Control(33)=   "Label1(94)"
      Tab(1).Control(34)=   "Label1(93)"
      Tab(1).Control(35)=   "Label1(92)"
      Tab(1).Control(36)=   "Label1(20)"
      Tab(1).Control(37)=   "Label2(10)"
      Tab(1).Control(38)=   "Label2(9)"
      Tab(1).Control(39)=   "Label2(8)"
      Tab(1).Control(40)=   "Label2(7)"
      Tab(1).Control(41)=   "Label2(6)"
      Tab(1).Control(42)=   "Label1(62)"
      Tab(1).Control(43)=   "Label1(49)"
      Tab(1).Control(44)=   "Label1(50)"
      Tab(1).Control(45)=   "Label1(51)"
      Tab(1).Control(46)=   "Label1(52)"
      Tab(1).Control(47)=   "Label1(53)"
      Tab(1).Control(48)=   "cboContact"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).ControlCount=   49
      TabCaption(2)   =   "FC資料"
      TabPicture(2)   =   "frm050701.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(134)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(48)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(45)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(37)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(36)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label1(35)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(34)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label1(32)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label2(13)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label2(12)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label2(2)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label1(65)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label1(64)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label1(59)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label1(154)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label2(49)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label1(55)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label2(48)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label1(155)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label1(156)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Label1(47)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Label1(157)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Label1(158)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Label1(162)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Label1(163)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Label1(169)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Label1(61)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Label1(128)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Label1(129)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Label1(131)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Text1(70)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Text1(89)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Text1(78)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Text1(90)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Text1(75)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Text1(77)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Text1(76)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Text1(134)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Text1(133)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Text1(135)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Text1(141)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Text1(143)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Text1(146)"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Text1(49)"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Text1(50)"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Text1(88)"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Text1(71)"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Text1(151)"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "Text1(152)"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "Text1(159)"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "Text1(167)"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "Text1(69)"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "Text1(156)"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "Text1(177)"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "Text1(181)"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).ControlCount=   55
      TabCaption(3)   =   "聯絡人"
      TabPicture(3)   =   "frm050701.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(76)"
      Tab(3).Control(1)=   "Label2(11)"
      Tab(3).Control(2)=   "Label1(29)"
      Tab(3).Control(3)=   "Label1(30)"
      Tab(3).Control(4)=   "Label1(31)"
      Tab(3).Control(5)=   "Label1(33)"
      Tab(3).Control(6)=   "Label1(42)"
      Tab(3).Control(7)=   "Label1(43)"
      Tab(3).Control(8)=   "Label1(44)"
      Tab(3).Control(9)=   "Label1(46)"
      Tab(3).Control(10)=   "Label2(46)"
      Tab(3).Control(11)=   "Label1(83)"
      Tab(3).Control(12)=   "Label1(82)"
      Tab(3).Control(13)=   "Label1(81)"
      Tab(3).Control(14)=   "Label1(80)"
      Tab(3).Control(15)=   "Label1(79)"
      Tab(3).Control(16)=   "Label1(78)"
      Tab(3).Control(17)=   "Label1(77)"
      Tab(3).Control(18)=   "Text1(101)"
      Tab(3).Control(19)=   "Text1(99)"
      Tab(3).Control(20)=   "Text1(98)"
      Tab(3).Control(21)=   "Text1(104)"
      Tab(3).Control(22)=   "Text1(103)"
      Tab(3).Control(23)=   "Text1(102)"
      Tab(3).Control(24)=   "Text1(100)"
      Tab(3).Control(25)=   "Text1(139)"
      Tab(3).Control(26)=   "Text1(56)"
      Tab(3).Control(27)=   "Text1(53)"
      Tab(3).Control(28)=   "Text1(87)"
      Tab(3).Control(29)=   "Text1(86)"
      Tab(3).Control(30)=   "Text1(55)"
      Tab(3).Control(31)=   "Text1(54)"
      Tab(3).Control(32)=   "Text1(52)"
      Tab(3).Control(33)=   "Text1(51)"
      Tab(3).ControlCount=   34
      TabCaption(4)   =   "繳費/代理人備註"
      TabPicture(4)   =   "frm050701.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(41)"
      Tab(4).Control(1)=   "Label1(40)"
      Tab(4).Control(2)=   "Label1(39)"
      Tab(4).Control(3)=   "Label1(38)"
      Tab(4).Control(4)=   "Label1(88)"
      Tab(4).Control(5)=   "Label1(87)"
      Tab(4).Control(6)=   "Label1(86)"
      Tab(4).Control(7)=   "Label1(85)"
      Tab(4).Control(8)=   "Label1(84)"
      Tab(4).Control(9)=   "Label2(45)"
      Tab(4).Control(10)=   "Text1(105)"
      Tab(4).Control(11)=   "Text1(106)"
      Tab(4).Control(12)=   "Text2(0)"
      Tab(4).Control(13)=   "Text2(1)"
      Tab(4).Control(14)=   "Text1(72)"
      Tab(4).Control(15)=   "Text1(73)"
      Tab(4).Control(16)=   "Text1(74)"
      Tab(4).Control(17)=   "Text1(107)"
      Tab(4).Control(18)=   "Label2(3)"
      Tab(4).Control(19)=   "Label2(4)"
      Tab(4).Control(20)=   "Label1(54)"
      Tab(4).Control(21)=   "Label2(51)"
      Tab(4).Control(22)=   "Label2(47)"
      Tab(4).Control(23)=   "MSHFlexGrid1"
      Tab(4).Control(24)=   "Command1(3)"
      Tab(4).Control(25)=   "Text3(0)"
      Tab(4).Control(26)=   "Command1(2)"
      Tab(4).Control(27)=   "Command1(1)"
      Tab(4).Control(28)=   "Text3(1)"
      Tab(4).Control(29)=   "Command1(0)"
      Tab(4).ControlCount=   30
      TabCaption(5)   =   "發明人"
      TabPicture(5)   =   "frm050701.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Lb_IN11"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Lb_IN11N"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label1(126)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label1(90)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label1(89)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label1(66)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "GRDtmp"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "GRD1"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "txtIN11"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "cmdAddRow"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "cmdDelRow"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "Combo1"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Frame3"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "cmdUpdRow"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "Frame4"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).ControlCount=   15
      TabCaption(6)   =   "代表人"
      TabPicture(6)   =   "frm050701.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1(125)"
      Tab(6).Control(1)=   "Label1(124)"
      Tab(6).Control(2)=   "Label1(123)"
      Tab(6).Control(3)=   "Label1(122)"
      Tab(6).Control(4)=   "Label1(121)"
      Tab(6).Control(5)=   "Label1(120)"
      Tab(6).Control(6)=   "Label1(119)"
      Tab(6).Control(7)=   "Label1(118)"
      Tab(6).Control(8)=   "Label1(117)"
      Tab(6).Control(9)=   "Label1(116)"
      Tab(6).Control(10)=   "Label1(115)"
      Tab(6).Control(11)=   "Label1(114)"
      Tab(6).Control(12)=   "Label1(113)"
      Tab(6).Control(13)=   "Label1(112)"
      Tab(6).Control(14)=   "Label1(111)"
      Tab(6).Control(15)=   "Label1(100)"
      Tab(6).Control(16)=   "Label1(99)"
      Tab(6).Control(17)=   "Label1(98)"
      Tab(6).Control(18)=   "Label1(97)"
      Tab(6).Control(19)=   "Label1(96)"
      Tab(6).Control(20)=   "Label1(58)"
      Tab(6).Control(21)=   "Label1(57)"
      Tab(6).Control(22)=   "Label1(56)"
      Tab(6).Control(23)=   "Label1(60)"
      Tab(6).Control(24)=   "Label1(63)"
      Tab(6).Control(25)=   "Text1(132)"
      Tab(6).Control(26)=   "Text1(131)"
      Tab(6).Control(27)=   "Text1(130)"
      Tab(6).Control(28)=   "Text1(129)"
      Tab(6).Control(29)=   "Text1(128)"
      Tab(6).Control(30)=   "Text1(127)"
      Tab(6).Control(31)=   "Text1(126)"
      Tab(6).Control(32)=   "Text1(125)"
      Tab(6).Control(33)=   "Text1(124)"
      Tab(6).Control(34)=   "Text1(123)"
      Tab(6).Control(35)=   "Text1(122)"
      Tab(6).Control(36)=   "Text1(121)"
      Tab(6).Control(37)=   "Text1(120)"
      Tab(6).Control(38)=   "Text1(119)"
      Tab(6).Control(39)=   "Text1(118)"
      Tab(6).Control(40)=   "Text1(117)"
      Tab(6).Control(41)=   "Text1(116)"
      Tab(6).Control(42)=   "Text1(115)"
      Tab(6).Control(43)=   "Text1(114)"
      Tab(6).Control(44)=   "Text1(113)"
      Tab(6).Control(45)=   "Text1(112)"
      Tab(6).Control(46)=   "Text1(111)"
      Tab(6).Control(47)=   "Text1(110)"
      Tab(6).Control(48)=   "Text1(109)"
      Tab(6).Control(49)=   "Text1(84)"
      Tab(6).Control(50)=   "Text1(83)"
      Tab(6).Control(51)=   "Text1(81)"
      Tab(6).Control(52)=   "Text1(80)"
      Tab(6).Control(53)=   "Text1(79)"
      Tab(6).Control(54)=   "Text1(82)"
      Tab(6).ControlCount=   55
      TabCaption(7)   =   "銷卷資料"
      TabPicture(7)   =   "frm050701.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label39"
      Tab(7).Control(1)=   "Label38"
      Tab(7).Control(2)=   "Label36"
      Tab(7).Control(3)=   "Label4"
      Tab(7).Control(4)=   "Text1(108)"
      Tab(7).Control(5)=   "Text1(136)"
      Tab(7).Control(6)=   "Text1(137)"
      Tab(7).Control(7)=   "Text1(138)"
      Tab(7).ControlCount=   8
      TabCaption(8)   =   "其他"
      TabPicture(8)   =   "frm050701.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame1K"
      Tab(8).Control(1)=   "cmdDivSug"
      Tab(8).Control(2)=   "lstPA166"
      Tab(8).Control(3)=   "Frame2"
      Tab(8).Control(4)=   "Combo4"
      Tab(8).Control(5)=   "Combo5"
      Tab(8).Control(6)=   "Frame1"
      Tab(8).Control(7)=   "Text1(166)"
      Tab(8).Control(8)=   "Text1(163)"
      Tab(8).Control(9)=   "Text1(162)"
      Tab(8).Control(10)=   "Text1(161)"
      Tab(8).Control(11)=   "Text1(142)"
      Tab(8).Control(12)=   "Text1(155)"
      Tab(8).Control(13)=   "Text1(154)"
      Tab(8).Control(14)=   "Text1(153)"
      Tab(8).Control(15)=   "Text1(147)"
      Tab(8).Control(16)=   "Text1(148)"
      Tab(8).Control(17)=   "Text1(60)"
      Tab(8).Control(18)=   "Text1(61)"
      Tab(8).Control(19)=   "Label1(177)"
      Tab(8).Control(20)=   "Label1(175)"
      Tab(8).Control(21)=   "Label1(174)"
      Tab(8).Control(22)=   "Label1(173)"
      Tab(8).Control(23)=   "Label68"
      Tab(8).Control(24)=   "Label1(166)"
      Tab(8).Control(25)=   "Label1(165)"
      Tab(8).Control(26)=   "Label1(164)"
      Tab(8).Control(27)=   "Label1(159)"
      Tab(8).Control(28)=   "Label70"
      Tab(8).Control(29)=   "Label1(67)"
      Tab(8).Control(30)=   "Label49"
      Tab(8).Control(31)=   "Label11(0)"
      Tab(8).Control(32)=   "Label1(68)"
      Tab(8).ControlCount=   33
      TabCaption(9)   =   "參考備註"
      TabPicture(9)   =   "frm050701.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label1(172)"
      Tab(9).Control(1)=   "Text1(91)"
      Tab(9).Control(2)=   "cmdIns"
      Tab(9).ControlCount=   3
      Begin VB.Frame Frame1K 
         Height          =   280
         Left            =   -74970
         TabIndex        =   406
         Top             =   4890
         Width           =   4930
         Begin VB.CheckBox Chk1K 
            Caption         =   "月帳單"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   156
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "上傳平台"
            Height          =   180
            Index           =   1
            Left            =   2790
            TabIndex        =   155
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "帳單另寄"
            Height          =   180
            Index           =   0
            Left            =   1740
            TabIndex        =   154
            Top             =   60
            Width           =   1030
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "請款單寄送類型："
            Height          =   180
            Index           =   26
            Left            =   150
            TabIndex        =   407
            Top             =   60
            Width           =   1440
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "移動順序:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -70560
         TabIndex        =   403
         Top             =   1704
         Width           =   2025
         Begin VB.CommandButton cmdDown 
            Caption         =   "▼"
            Height          =   255
            Left            =   1410
            TabIndex        =   405
            Top             =   90
            Width           =   375
         End
         Begin VB.CommandButton cmdUp 
            Caption         =   "▲"
            Height          =   255
            Left            =   960
            TabIndex        =   404
            Top             =   90
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdUpdRow 
         Caption         =   "修改"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72168
         TabIndex        =   98
         Top             =   1776
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   948
         Left            =   -73920
         TabIndex        =   401
         Top             =   720
         Width           =   6228
         Begin MSForms.TextBox txtInvField 
            Height          =   288
            Index           =   0
            Left            =   24
            TabIndex        =   94
            Top             =   324
            Width           =   6132
            VariousPropertyBits=   671105051
            BackColor       =   -2147483644
            MaxLength       =   70
            Size            =   "10816;508"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtInvField 
            Height          =   288
            Index           =   1
            Left            =   24
            TabIndex        =   93
            Top             =   24
            Width           =   6132
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
            Height          =   288
            Index           =   2
            Left            =   24
            TabIndex        =   95
            Top             =   624
            Width           =   6132
            VariousPropertyBits=   671105051
            BackColor       =   -2147483644
            MaxLength       =   40
            Size            =   "10816;508"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.ComboBox Combo6 
         Height          =   260
         ItemData        =   "frm050701.frx":0118
         Left            =   -67890
         List            =   "frm050701.frx":011A
         Style           =   2  '單純下拉式
         TabIndex        =   400
         Top             =   4350
         Width           =   1590
      End
      Begin VB.ComboBox cboContact 
         BeginProperty Font 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   -68205
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   212
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1800
      End
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   -74700
         Style           =   1  '圖片外觀
         TabIndex        =   362
         Top             =   1770
         Width           =   840
      End
      Begin VB.Frame FraPA174 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame3"
         Height          =   240
         Left            =   -74910
         TabIndex        =   363
         Top             =   1515
         Width           =   1035
         Begin VB.CheckBox ChkPA174 
            Caption         =   "有特殊字"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   30
            TabIndex        =   364
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "加入(&A)"
         Height          =   300
         Index           =   0
         Left            =   -73125
         TabIndex        =   81
         Top             =   390
         Width           =   765
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   1
         Left            =   -73620
         MaxLength       =   1
         TabIndex        =   85
         Top             =   690
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "刪除(&D)"
         Height          =   300
         Index           =   1
         Left            =   -71580
         TabIndex        =   83
         Top             =   390
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         Caption         =   "清除(&C)"
         Height          =   300
         Index           =   2
         Left            =   -70800
         TabIndex        =   84
         Top             =   390
         Width           =   765
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Index           =   0
         Left            =   -73980
         MaxLength       =   7
         TabIndex        =   80
         Top             =   375
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "修改(&E)"
         Height          =   300
         Index           =   3
         Left            =   -72360
         TabIndex        =   82
         Top             =   390
         Width           =   765
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         ItemData        =   "frm050701.frx":011C
         Left            =   -73896
         List            =   "frm050701.frx":011E
         Style           =   2  '單純下拉式
         TabIndex        =   92
         Top             =   390
         Width           =   7425
      End
      Begin VB.CommandButton cmdDelRow 
         Caption         =   "刪除"
         Height          =   285
         Left            =   -73020
         TabIndex        =   97
         Top             =   1776
         Width           =   735
      End
      Begin VB.CommandButton cmdAddRow 
         Caption         =   "加入"
         Height          =   285
         Left            =   -73845
         TabIndex        =   96
         Top             =   1776
         Width           =   735
      End
      Begin VB.TextBox txtIN11 
         Height          =   270
         Left            =   -67656
         MaxLength       =   3
         TabIndex        =   99
         Top             =   1752
         Width           =   400
      End
      Begin VB.CommandButton cmdDivSug 
         Caption         =   "分割建議"
         Height          =   315
         Left            =   -70680
         TabIndex        =   139
         Top             =   2940
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ListBox lstPA166 
         Height          =   228
         Left            =   -73005
         TabIndex        =   143
         Top             =   4110
         Width           =   4155
      End
      Begin VB.Frame Frame2 
         Height          =   1065
         Left            =   -68820
         TabIndex        =   270
         Top             =   4020
         Width           =   2535
         Begin VB.TextBox txtNo 
            Height          =   300
            Left            =   45
            MaxLength       =   9
            TabIndex        =   151
            Top             =   120
            Width           =   945
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<- 新增"
            Height          =   285
            Left            =   45
            TabIndex        =   152
            Top             =   450
            Width           =   735
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除 ->"
            Height          =   285
            Left            =   45
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   750
            Width           =   735
         End
         Begin VB.Label lblName 
            Caption         =   "lblName"
            Height          =   840
            Left            =   1080
            TabIndex        =   271
            Top             =   150
            Width           =   1395
            WordWrap        =   -1  'True
         End
      End
      Begin VB.ComboBox Combo4 
         Height          =   276
         ItemData        =   "frm050701.frx":0120
         Left            =   -71070
         List            =   "frm050701.frx":0122
         Style           =   2  '單純下拉式
         TabIndex        =   131
         Top             =   1110
         Width           =   990
      End
      Begin VB.ComboBox Combo5 
         Height          =   276
         ItemData        =   "frm050701.frx":0124
         Left            =   -67950
         List            =   "frm050701.frx":0137
         Style           =   2  '單純下拉式
         TabIndex        =   132
         Top             =   1110
         Width           =   1470
      End
      Begin VB.Frame Frame1 
         Caption         =   "中文本資訊"
         Height          =   2600
         Left            =   -69030
         TabIndex        =   265
         Top             =   1440
         Width           =   2535
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   68
            Left            =   1815
            TabIndex        =   148
            Top             =   1365
            Width           =   630
            VariousPropertyBits=   671105051
            MaxLength       =   4
            Size            =   "1111;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   67
            Left            =   1815
            TabIndex        =   147
            Top             =   1077
            Width           =   630
            VariousPropertyBits=   671105051
            MaxLength       =   4
            Size            =   "1111;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   66
            Left            =   1815
            TabIndex        =   146
            Top             =   788
            Width           =   630
            VariousPropertyBits=   671105051
            MaxLength       =   4
            Size            =   "1111;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   65
            Left            =   1815
            TabIndex        =   145
            Top             =   499
            Width           =   630
            VariousPropertyBits=   671105051
            MaxLength       =   4
            Size            =   "1111;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   64
            Left            =   1815
            TabIndex        =   144
            Top             =   210
            Width           =   630
            VariousPropertyBits=   671105051
            MaxLength       =   4
            Size            =   "1111;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   172
            Left            =   1815
            TabIndex        =   149
            Top             =   1920
            Width           =   630
            VariousPropertyBits=   671105051
            MaxLength       =   4
            Size            =   "1111;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox Text1 
            Height          =   300
            Index           =   173
            Left            =   1815
            TabIndex        =   150
            Top             =   2220
            Width           =   630
            VariousPropertyBits=   671105051
            MaxLength       =   4
            Size            =   "1111;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Caption         =   "摘要頁數："
            Height          =   180
            Index           =   69
            Left            =   200
            TabIndex        =   269
            Top             =   315
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "說明書頁數："
            Height          =   180
            Index           =   70
            Left            =   195
            TabIndex        =   268
            Top             =   600
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "序列表："
            Height          =   180
            Index           =   71
            Left            =   195
            TabIndex        =   267
            Top             =   885
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "申請專利範圍頁數："
            Height          =   180
            Index           =   72
            Left            =   200
            TabIndex        =   266
            Top             =   1155
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "圖式頁數："
            Height          =   180
            Index           =   73
            Left            =   200
            TabIndex        =   259
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "頁數總計："
            Height          =   180
            Index           =   74
            Left            =   200
            TabIndex        =   264
            Top             =   1725
            Width           =   900
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
            Left            =   960
            TabIndex        =   263
            Top             =   885
            Width           =   900
         End
         Begin VB.Label lblTot6 
            Caption         =   "Label3"
            Height          =   255
            Left            =   1860
            TabIndex        =   262
            Top             =   1688
            Width           =   560
         End
         Begin VB.Label Label1 
            Caption         =   "申請專利範圍項數："
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   91
            Left            =   200
            TabIndex        =   261
            Top             =   1965
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "圖式圖數："
            Height          =   180
            Index           =   127
            Left            =   200
            TabIndex        =   260
            Top             =   2235
            Width           =   1620
         End
      End
      Begin VB.CommandButton cmdIns 
         Caption         =   "各項指示"
         Height          =   300
         Left            =   -74910
         TabIndex        =   163
         Top             =   420
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         ItemData        =   "frm050701.frx":016B
         Left            =   -68010
         List            =   "frm050701.frx":0178
         TabIndex        =   7
         Top             =   660
         Width           =   1785
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   260
         ItemData        =   "frm050701.frx":019C
         Left            =   -68370
         List            =   "frm050701.frx":01AC
         TabIndex        =   10
         Top             =   975
         Width           =   2145
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   3912
         Left            =   -74940
         TabIndex        =   100
         Top             =   2136
         Width           =   8748
         _ExtentX        =   15409
         _ExtentY        =   6922
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3495
         Left            =   -74910
         TabIndex        =   86
         Top             =   1020
         Width           =   4425
         _ExtentX        =   7789
         _ExtentY        =   6181
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
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
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRDtmp 
         Height          =   828
         Left            =   -74952
         TabIndex        =   402
         Top             =   1752
         Visible         =   0   'False
         Width           =   1032
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
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   181
         Left            =   7848
         TabIndex        =   50
         Top             =   1968
         Width           =   288
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "個體身份:"
         Height          =   180
         Index           =   133
         Left            =   -68700
         TabIndex        =   399
         Top             =   4410
         Width           =   765
      End
      Begin VB.Label lblPA176 
         Caption         =   "專利權期間延長相關:"
         Height          =   360
         Left            =   -71670
         TabIndex        =   398
         Top             =   660
         Width           =   1000
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   178
         Left            =   -67890
         TabIndex        =   33
         Top             =   4035
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "證書形式:         (1:電子 2:紙本)"
         Height          =   180
         Index           =   132
         Left            =   -68700
         TabIndex        =   397
         Top             =   4065
         Width           =   2325
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   47
         Left            =   -73365
         TabIndex        =   396
         Top             =   5475
         Width           =   7125
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "12568;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   51
         Left            =   -72225
         TabIndex        =   329
         Top             =   5160
         Width           =   1995
         ForeColor       =   49152
         VariousPropertyBits=   276824091
         Caption         =   "尚未公告　年費為預估值"
         Size            =   "3519;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日:"
         Height          =   180
         Index           =   54
         Left            =   -74820
         TabIndex        =   395
         Top             =   5475
         Width           =   1065
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   -68970
         TabIndex        =   331
         Top             =   2813
         Width           =   600
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1058;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   -68850
         TabIndex        =   330
         Top             =   983
         Width           =   600
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1058;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人5:"
         Height          =   180
         Index           =   53
         Left            =   -74700
         TabIndex        =   394
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人4:"
         Height          =   180
         Index           =   52
         Left            =   -74700
         TabIndex        =   393
         Top             =   1260
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人3:"
         Height          =   180
         Index           =   51
         Left            =   -74700
         TabIndex        =   392
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人2:"
         Height          =   180
         Index           =   50
         Left            =   -74700
         TabIndex        =   391
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人1:"
         Height          =   180
         Index           =   49
         Left            =   -74700
         TabIndex        =   390
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號:"
         Height          =   180
         Index           =   62
         Left            =   -74700
         TabIndex        =   389
         Top             =   1860
         Width           =   1125
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   6
         Left            =   -72900
         TabIndex        =   388
         Top             =   360
         Width           =   6630
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11695;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   7
         Left            =   -72900
         TabIndex        =   387
         Top             =   660
         Width           =   6630
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11695;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   8
         Left            =   -72900
         TabIndex        =   386
         Top             =   960
         Width           =   6630
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11695;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   9
         Left            =   -72900
         TabIndex        =   385
         Top             =   1260
         Width           =   6630
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11695;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   10
         Left            =   -72900
         TabIndex        =   384
         Top             =   1560
         Width           =   6630
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11695;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申  1:中"
         Height          =   180
         Index           =   20
         Left            =   -74850
         TabIndex        =   383
         Top             =   2190
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "請  2:中"
         Height          =   180
         Index           =   92
         Left            =   -74850
         TabIndex        =   382
         Top             =   3090
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "人  3:中"
         Height          =   180
         Index           =   93
         Left            =   -74850
         TabIndex        =   381
         Top             =   3990
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "地  4:中"
         Height          =   180
         Index           =   94
         Left            =   -74850
         TabIndex        =   380
         Top             =   4890
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "址  5:中"
         Height          =   180
         Index           =   95
         Left            =   -74850
         TabIndex        =   379
         Top             =   5790
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   101
         Left            =   -74445
         TabIndex        =   378
         Top             =   2505
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   102
         Left            =   -74445
         TabIndex        =   377
         Top             =   2790
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   103
         Left            =   -74445
         TabIndex        =   376
         Top             =   3405
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   104
         Left            =   -74445
         TabIndex        =   375
         Top             =   4305
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   105
         Left            =   -74445
         TabIndex        =   374
         Top             =   5205
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   106
         Left            =   -74445
         TabIndex        =   373
         Top             =   6105
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   107
         Left            =   -74445
         TabIndex        =   372
         Top             =   3690
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   108
         Left            =   -74445
         TabIndex        =   371
         Top             =   4590
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   109
         Left            =   -74445
         TabIndex        =   370
         Top             =   5490
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   110
         Left            =   -74445
         TabIndex        =   369
         Top             =   6390
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人:"
         Height          =   180
         Index           =   160
         Left            =   -68925
         TabIndex        =   368
         Top             =   1860
         Width           =   585
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   31
         Left            =   -74235
         TabIndex        =   213
         Top             =   2160
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   36
         Left            =   -74235
         TabIndex        =   214
         Top             =   2460
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   41
         Left            =   -74235
         TabIndex        =   215
         Top             =   2760
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   32
         Left            =   -74235
         TabIndex        =   216
         Top             =   3060
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   37
         Left            =   -74235
         TabIndex        =   217
         Top             =   3360
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   42
         Left            =   -74235
         TabIndex        =   218
         Top             =   3660
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   33
         Left            =   -74235
         TabIndex        =   219
         Top             =   3960
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   38
         Left            =   -74235
         TabIndex        =   220
         Top             =   4260
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   43
         Left            =   -74235
         TabIndex        =   221
         Top             =   4560
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   34
         Left            =   -74235
         TabIndex        =   222
         Top             =   4860
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   39
         Left            =   -74235
         TabIndex        =   223
         Top             =   5160
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   44
         Left            =   -74235
         TabIndex        =   224
         Top             =   5460
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   35
         Left            =   -74235
         TabIndex        =   225
         Top             =   5760
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   40
         Left            =   -74235
         TabIndex        =   226
         Top             =   6060
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   185
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   45
         Left            =   -74235
         TabIndex        =   227
         Top             =   6360
         Width           =   7905
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13944;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   26
         Left            =   -73980
         TabIndex        =   206
         Top             =   360
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   27
         Left            =   -73980
         TabIndex        =   207
         Top             =   660
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   28
         Left            =   -73980
         TabIndex        =   208
         Top             =   960
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   29
         Left            =   -73980
         TabIndex        =   209
         Top             =   1260
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   30
         Left            =   -73980
         TabIndex        =   210
         Top             =   1560
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   48
         Left            =   -73440
         TabIndex        =   211
         Top             =   1860
         Width           =   4485
         VariousPropertyBits=   671105051
         Size            =   "7911;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   177
         Left            =   7740
         TabIndex        =   42
         Top             =   1050
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   23
         Left            =   -70560
         TabIndex        =   4
         Top             =   345
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   176
         Left            =   -70392
         TabIndex        =   6
         Top             =   660
         Width           =   252
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   156
         Left            =   1740
         TabIndex        =   56
         Top             =   2895
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   51
         Left            =   -73770
         TabIndex        =   64
         Top             =   450
         Width           =   2805
         VariousPropertyBits=   671105051
         Size            =   "4948;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   52
         Left            =   -73770
         TabIndex        =   66
         Top             =   765
         Width           =   7365
         VariousPropertyBits=   671105051
         Size            =   "12991;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   54
         Left            =   -73770
         TabIndex        =   67
         Top             =   1080
         Width           =   2805
         VariousPropertyBits=   671105051
         Size            =   "4948;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   55
         Left            =   -73770
         TabIndex        =   69
         Top             =   1380
         Width           =   7365
         VariousPropertyBits=   671105051
         Size            =   "12991;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   86
         Left            =   -73770
         TabIndex        =   70
         Top             =   1695
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   87
         Left            =   -73770
         TabIndex        =   72
         Top             =   1995
         Width           =   7365
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "12991;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   53
         Left            =   -69810
         TabIndex        =   65
         Top             =   450
         Width           =   3405
         VariousPropertyBits=   671105051
         Size            =   "6006;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   56
         Left            =   -69810
         TabIndex        =   68
         Top             =   1080
         Width           =   3405
         VariousPropertyBits=   671105051
         Size            =   "6006;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   139
         Left            =   -69495
         TabIndex        =   71
         Top             =   1695
         Width           =   3090
         VariousPropertyBits=   671105051
         Size            =   "5450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   100
         Left            =   -69600
         TabIndex        =   78
         Top             =   3240
         Width           =   3165
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5583;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   102
         Left            =   -69315
         TabIndex        =   74
         Top             =   2310
         Width           =   2880
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "5080;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   103
         Left            =   -72660
         TabIndex        =   75
         Top             =   2640
         Width           =   6225
         VariousPropertyBits=   -1466941413
         MaxLength       =   264
         ScrollBars      =   2
         Size            =   "10980;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   104
         Left            =   -72660
         TabIndex        =   76
         Top             =   2940
         Width           =   6225
         VariousPropertyBits=   -1466941413
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "10980;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   98
         Left            =   -73560
         TabIndex        =   77
         Top             =   3240
         Width           =   2565
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "4524;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   99
         Left            =   -73560
         TabIndex        =   79
         Top             =   3540
         Width           =   7125
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "12568;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   101
         Left            =   -73500
         TabIndex        =   73
         Top             =   2310
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   107
         Left            =   -73380
         TabIndex        =   91
         Top             =   5130
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   74
         Left            =   -66900
         TabIndex        =   328
         Top             =   4380
         Visible         =   0   'False
         Width           =   1575
         VariousPropertyBits=   671105051
         MaxLength       =   39
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   73
         Left            =   -66900
         TabIndex        =   327
         Top             =   4110
         Visible         =   0   'False
         Width           =   1575
         VariousPropertyBits=   671105051
         MaxLength       =   179
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   72
         Left            =   -66900
         TabIndex        =   326
         Top             =   3840
         Visible         =   0   'False
         Width           =   1575
         VariousPropertyBits=   671105051
         MaxLength       =   99
         Size            =   "7223;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   1215
         Index           =   1
         Left            =   -70410
         TabIndex        =   89
         Top             =   3315
         Width           =   4155
         VariousPropertyBits=   -1466941409
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "7329;2143"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   1215
         Index           =   0
         Left            =   -70410
         TabIndex        =   87
         Top             =   1530
         Width           =   4155
         VariousPropertyBits=   -1466941409
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "7329;2143"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   106
         Left            =   -73380
         TabIndex        =   90
         Top             =   4830
         Width           =   2325
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "4101;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   105
         Left            =   -73380
         TabIndex        =   88
         Top             =   4530
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   82
         Left            =   -69705
         TabIndex        =   104
         Top             =   420
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   79
         Left            =   -73950
         TabIndex        =   101
         Top             =   420
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   80
         Left            =   -73950
         TabIndex        =   102
         Top             =   720
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   81
         Left            =   -73950
         TabIndex        =   103
         Top             =   1020
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   83
         Left            =   -69705
         TabIndex        =   105
         Top             =   720
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   84
         Left            =   -69705
         TabIndex        =   106
         Top             =   1020
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   109
         Left            =   -73950
         TabIndex        =   107
         Top             =   1380
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   110
         Left            =   -73950
         TabIndex        =   108
         Top             =   1680
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   111
         Left            =   -73950
         TabIndex        =   109
         Top             =   1980
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   112
         Left            =   -69705
         TabIndex        =   110
         Top             =   1380
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   113
         Left            =   -69705
         TabIndex        =   111
         Top             =   1680
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   114
         Left            =   -69705
         TabIndex        =   112
         Top             =   1980
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   115
         Left            =   -73950
         TabIndex        =   113
         Top             =   2355
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   116
         Left            =   -73950
         TabIndex        =   114
         Top             =   2655
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   117
         Left            =   -73950
         TabIndex        =   115
         Top             =   2955
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   118
         Left            =   -69705
         TabIndex        =   116
         Top             =   2355
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   119
         Left            =   -69705
         TabIndex        =   117
         Top             =   2655
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   120
         Left            =   -69705
         TabIndex        =   118
         Top             =   2955
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   121
         Left            =   -73950
         TabIndex        =   119
         Top             =   3330
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   122
         Left            =   -73950
         TabIndex        =   120
         Top             =   3630
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   123
         Left            =   -73950
         TabIndex        =   121
         Top             =   3930
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   124
         Left            =   -69705
         TabIndex        =   122
         Top             =   3330
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   125
         Left            =   -69705
         TabIndex        =   123
         Top             =   3630
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   126
         Left            =   -69705
         TabIndex        =   124
         Top             =   3930
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   127
         Left            =   -73950
         TabIndex        =   125
         Top             =   4305
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   128
         Left            =   -73950
         TabIndex        =   126
         Top             =   4605
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   129
         Left            =   -73950
         TabIndex        =   127
         Top             =   4905
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   130
         Left            =   -69705
         TabIndex        =   128
         Top             =   4305
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   131
         Left            =   -69705
         TabIndex        =   157
         Top             =   4605
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   132
         Left            =   -69705
         TabIndex        =   158
         Top             =   4905
         Width           =   3375
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "5953;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   138
         Left            =   -73500
         TabIndex        =   290
         Top             =   1410
         Width           =   7245
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         Size            =   "12779;503"
         Value           =   "text1"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   137
         Left            =   -73500
         TabIndex        =   289
         Top             =   1080
         Width           =   1515
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         Size            =   "2672;503"
         Value           =   "text1"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   136
         Left            =   -73500
         TabIndex        =   288
         Top             =   750
         Width           =   1515
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         Size            =   "2672;503"
         Value           =   "text1"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   108
         Left            =   -73500
         TabIndex        =   287
         Top             =   420
         Width           =   1515
         VariousPropertyBits=   671105055
         BackColor       =   -2147483644
         Size            =   "2672;503"
         Value           =   "text1"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   166
         Left            =   -74805
         TabIndex        =   272
         Top             =   4350
         Visible         =   0   'False
         Width           =   1695
         VariousPropertyBits=   671105051
         Size            =   "2990;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   163
         Left            =   -72588
         TabIndex        =   140
         Top             =   3240
         Width           =   252
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   162
         Left            =   -72840
         TabIndex        =   138
         Top             =   2940
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   161
         Left            =   -73200
         TabIndex        =   137
         Top             =   2670
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   142
         Left            =   -73200
         TabIndex        =   135
         Top             =   2055
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   155
         Left            =   -73200
         TabIndex        =   136
         Top             =   2355
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   154
         Left            =   -73200
         TabIndex        =   134
         Top             =   1755
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   153
         Left            =   -73200
         TabIndex        =   133
         Top             =   1425
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   147
         Left            =   -73200
         TabIndex        =   130
         Top             =   1110
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   615
         Index           =   148
         Left            =   -73920
         TabIndex        =   129
         Top             =   450
         Width           =   7635
         VariousPropertyBits=   -1466941413
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "13467;1085"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   60
         Left            =   -72840
         TabIndex        =   141
         Top             =   3540
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   61
         Left            =   -73200
         TabIndex        =   142
         Top             =   3810
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   4170
         Index           =   91
         Left            =   -74970
         TabIndex        =   164
         Top             =   780
         Width           =   8775
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "15478;7355"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   69
         Left            =   1740
         TabIndex        =   57
         Top             =   3210
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   167
         Left            =   7740
         TabIndex        =   44
         Top             =   1380
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   164
         Left            =   -67305
         TabIndex        =   21
         Top             =   2805
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   4
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   140
         Left            =   -67305
         TabIndex        =   24
         Top             =   3105
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   160
         Left            =   -67800
         TabIndex        =   18
         Top             =   2505
         Width           =   855
         VariousPropertyBits=   671105051
         MaxLength       =   5
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   159
         Left            =   5850
         TabIndex        =   63
         Top             =   3840
         Width           =   2730
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "4815;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   157
         Left            =   -66990
         TabIndex        =   30
         Top             =   3735
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   152
         Left            =   6675
         TabIndex        =   48
         Top             =   1680
         Width           =   330
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   151
         Left            =   4965
         TabIndex        =   47
         Top             =   1680
         Width           =   330
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   71
         Left            =   1740
         TabIndex        =   51
         Top             =   2310
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   88
         Left            =   1290
         TabIndex        =   49
         Top             =   1995
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   50
         Left            =   3270
         TabIndex        =   46
         Top             =   1680
         Width           =   330
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   49
         Left            =   1290
         TabIndex        =   45
         Top             =   1680
         Width           =   330
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   146
         Left            =   4755
         TabIndex        =   58
         Top             =   3210
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   143
         Left            =   7485
         TabIndex        =   59
         Top             =   3210
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   141
         Left            =   5250
         TabIndex        =   52
         Top             =   2310
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   135
         Left            =   1620
         TabIndex        =   62
         Top             =   3840
         Width           =   2385
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "4207;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   133
         Left            =   1620
         TabIndex        =   60
         Top             =   3540
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   134
         Left            =   5580
         TabIndex        =   61
         Top             =   3540
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   76
         Left            =   1290
         TabIndex        =   43
         Top             =   1380
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   77
         Left            =   1290
         TabIndex        =   40
         Top             =   720
         Width           =   7470
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "13176;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   75
         Left            =   1290
         TabIndex        =   39
         Top             =   405
         Width           =   1005
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   -73680
         TabIndex        =   0
         Top             =   345
         Width           =   525
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   90
         Left            =   7845
         TabIndex        =   53
         Top             =   2310
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   -73170
         TabIndex        =   1
         Top             =   345
         Width           =   795
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1402;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   78
         Left            =   7845
         TabIndex        =   55
         Top             =   2610
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   8
         Left            =   -73680
         TabIndex        =   8
         Top             =   975
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   47
         Left            =   -70440
         TabIndex        =   38
         Top             =   4935
         Width           =   4140
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "7302;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   59
         Left            =   -70080
         TabIndex        =   37
         Top             =   4620
         Width           =   525
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   57
         Left            =   -70080
         TabIndex        =   35
         Top             =   4335
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   19
         Left            =   -70080
         TabIndex        =   32
         Top             =   4035
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   24
         Left            =   -70680
         TabIndex        =   26
         Top             =   3408
         Width           =   888
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   20
         Left            =   -70440
         TabIndex        =   23
         Top             =   3105
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   15
         Left            =   -71070
         TabIndex        =   20
         Top             =   2805
         Width           =   2235
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "3942;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   13
         Left            =   -71070
         TabIndex        =   17
         Top             =   2505
         Width           =   2235
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "3942;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   11
         Left            =   -70920
         TabIndex        =   15
         Top             =   2205
         Width           =   2085
         VariousPropertyBits=   671105051
         MaxLength       =   25
         Size            =   "3678;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   58
         Left            =   -73920
         TabIndex        =   36
         Top             =   4635
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   18
         Left            =   -73440
         TabIndex        =   31
         Top             =   4035
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   16
         Left            =   -73440
         TabIndex        =   28
         Top             =   3735
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   89
         Left            =   1740
         TabIndex        =   41
         Top             =   1056
         Width           =   288
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   -72150
         TabIndex        =   3
         Top             =   345
         Width           =   420
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "741;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   -72390
         TabIndex        =   2
         Top             =   345
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   -73440
         TabIndex        =   13
         Top             =   1875
         Width           =   7230
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "12753;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   6
         Left            =   -73440
         TabIndex        =   12
         Top             =   1575
         Width           =   7230
         VariousPropertyBits=   671105051
         MaxLength       =   250
         Size            =   "12753;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   -73440
         TabIndex        =   11
         Top             =   1275
         Width           =   7230
         VariousPropertyBits=   536887323
         MaxLength       =   160
         Size            =   "12753;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   25
         Left            =   -69528
         TabIndex        =   27
         Top             =   3408
         Width           =   888
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   22
         Left            =   -73680
         TabIndex        =   25
         Top             =   3405
         Width           =   2145
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "3784;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   21
         Left            =   -73680
         TabIndex        =   22
         Top             =   3105
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   14
         Left            =   -73680
         TabIndex        =   19
         Top             =   2805
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   12
         Left            =   -73680
         TabIndex        =   16
         Top             =   2505
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   10
         Left            =   -73680
         TabIndex        =   14
         Top             =   2205
         Width           =   885
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   85
         Left            =   -73920
         TabIndex        =   34
         Top             =   4335
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   70
         Left            =   1740
         TabIndex        =   54
         Top             =   2610
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   9
         Left            =   -73680
         TabIndex        =   5
         Top             =   660
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   17
         Left            =   -70080
         TabIndex        =   29
         Top             =   3735
         Width           =   285
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   46
         Left            =   -70560
         TabIndex        =   9
         Top             =   975
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利連結通知:        (Y:是)"
         Height          =   180
         Index           =   131
         Left            =   6585
         TabIndex        =   366
         Top             =   1095
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " (Y/N)"
         Height          =   180
         Index           =   130
         Left            =   -70056
         TabIndex        =   365
         Top             =   720
         Width           =   456
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
         Left            =   -67680
         TabIndex        =   361
         Top             =   5595
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP 年費特殊管制:          (Y:年費續辦:有別於Y / X設定  N:寄證書/二核後年費不續辦  空白:視Y / X設定)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   129
         Left            =   150
         TabIndex        =   360
         Top             =   2940
         Width           =   7935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人(中):"
         Height          =   180
         Index           =   77
         Left            =   -74850
         TabIndex        =   359
         Top             =   3285
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人(英):"
         Height          =   180
         Index           =   78
         Left            =   -74850
         TabIndex        =   358
         Top             =   3585
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人(日):"
         Height          =   180
         Index           =   79
         Left            =   -70920
         TabIndex        =   357
         Top             =   3285
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體副本收受人:"
         Height          =   180
         Index           =   80
         Left            =   -74850
         TabIndex        =   356
         Top             =   2355
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體副本聯絡人:"
         Height          =   180
         Index           =   81
         Left            =   -70665
         TabIndex        =   355
         Top             =   2355
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體副本收受人彼此案號1:"
         Height          =   180
         Index           =   82
         Left            =   -74850
         TabIndex        =   354
         Top             =   2685
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體副本收受人彼此案號2:"
         Height          =   180
         Index           =   83
         Left            =   -74850
         TabIndex        =   353
         Top             =   2985
         Width           =   2115
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   46
         Left            =   -72540
         TabIndex        =   352
         Top             =   2340
         Width           =   1710
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3016;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人:"
         Height          =   180
         Index           =   46
         Left            =   -74850
         TabIndex        =   351
         Top             =   2025
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(日)2:"
         Height          =   180
         Index           =   44
         Left            =   -70890
         TabIndex        =   350
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(英)2:"
         Height          =   180
         Index           =   43
         Left            =   -74850
         TabIndex        =   349
         Top             =   1425
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(中)2:"
         Height          =   180
         Index           =   42
         Left            =   -74850
         TabIndex        =   348
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人:"
         Height          =   180
         Index           =   33
         Left            =   -74850
         TabIndex        =   347
         Top             =   1740
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(日)1:"
         Height          =   180
         Index           =   31
         Left            =   -70890
         TabIndex        =   346
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(英)1:"
         Height          =   180
         Index           =   30
         Left            =   -74850
         TabIndex        =   345
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人(中)1:"
         Height          =   180
         Index           =   29
         Left            =   -74850
         TabIndex        =   344
         Top             =   495
         Width           =   975
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   11
         Left            =   -72690
         TabIndex        =   343
         Top             =   1725
         Width           =   1620
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2857;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日):"
         Height          =   180
         Index           =   76
         Left            =   -70905
         TabIndex        =   342
         Top             =   1740
         Width           =   1245
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   45
         Left            =   -72450
         TabIndex        =   341
         Top             =   4590
         Width           =   4140
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "7302;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費請款對象:"
         Height          =   180
         Index           =   84
         Left            =   -74820
         TabIndex        =   340
         Top             =   4575
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費單筆不跑               (Y:不跑)"
         Height          =   180
         Index           =   85
         Left            =   -74820
         TabIndex        =   339
         Top             =   5160
         Width           =   2445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費彼所案號:"
         Height          =   180
         Index           =   86
         Left            =   -74820
         TabIndex        =   338
         Top             =   4845
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "繳費日期:"
         Height          =   180
         Index           =   87
         Left            =   -74820
         TabIndex        =   337
         Top             =   405
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否逾期補繳:         (Y:是)"
         Height          =   180
         Index           =   88
         Left            =   -74820
         TabIndex        =   336
         Top             =   750
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人收款後辦案:                 (Y:先收)"
         Height          =   180
         Index           =   38
         Left            =   -70395
         TabIndex        =   335
         Top             =   1020
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人備註:"
         Height          =   180
         Index           =   39
         Left            =   -70410
         TabIndex        =   334
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶收款後辦案:                    (Y:先收)"
         Height          =   180
         Index           =   40
         Left            =   -70410
         TabIndex        =   333
         Top             =   2850
         Width           =   3330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶備註:"
         Height          =   180
         Index           =   41
         Left            =   -70410
         TabIndex        =   332
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發明人:"
         Height          =   180
         Index           =   66
         Left            =   -74610
         TabIndex        =   325
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   89
         Left            =   -74364
         TabIndex        =   324
         Top             =   768
         Width           =   348
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   90
         Left            =   -74364
         TabIndex        =   323
         Top             =   1080
         Width           =   348
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   126
         Left            =   -74364
         TabIndex        =   322
         Top             =   1392
         Width           =   348
      End
      Begin VB.Label Lb_IN11N 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "Lb_IN11N"
         Height          =   180
         Left            =   -67236
         TabIndex        =   321
         Top             =   1776
         Width           =   972
      End
      Begin VB.Label Lb_IN11 
         AutoSize        =   -1  'True
         Caption         =   "國籍:"
         Height          =   180
         Left            =   -68136
         TabIndex        =   320
         Top             =   1776
         Width           =   408
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人2:"
         Height          =   180
         Index           =   63
         Left            =   -70470
         TabIndex        =   319
         Top             =   765
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人1:"
         Height          =   180
         Index           =   60
         Left            =   -74880
         TabIndex        =   318
         Top             =   765
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人4:"
         Height          =   180
         Index           =   56
         Left            =   -70470
         TabIndex        =   317
         Top             =   1725
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人3:"
         Height          =   180
         Index           =   57
         Left            =   -74880
         TabIndex        =   316
         Top             =   1725
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人6:"
         Height          =   180
         Index           =   58
         Left            =   -70470
         TabIndex        =   315
         Top             =   2700
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人5:"
         Height          =   180
         Index           =   96
         Left            =   -74880
         TabIndex        =   314
         Top             =   2700
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人8:"
         Height          =   180
         Index           =   97
         Left            =   -70470
         TabIndex        =   313
         Top             =   3675
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人7:"
         Height          =   180
         Index           =   98
         Left            =   -74880
         TabIndex        =   312
         Top             =   3675
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人10:"
         Height          =   180
         Index           =   99
         Left            =   -70470
         TabIndex        =   311
         Top             =   4680
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代表人9:"
         Height          =   180
         Index           =   100
         Left            =   -74880
         TabIndex        =   310
         Top             =   4680
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   111
         Left            =   -74130
         TabIndex        =   309
         Top             =   765
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   112
         Left            =   -74130
         TabIndex        =   308
         Top             =   1725
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   113
         Left            =   -74130
         TabIndex        =   307
         Top             =   2700
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   114
         Left            =   -74130
         TabIndex        =   306
         Top             =   3675
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "英"
         Height          =   180
         Index           =   115
         Left            =   -74130
         TabIndex        =   305
         Top             =   4680
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "中"
         Height          =   180
         Index           =   116
         Left            =   -74130
         TabIndex        =   304
         Top             =   465
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "中"
         Height          =   180
         Index           =   117
         Left            =   -74130
         TabIndex        =   303
         Top             =   1425
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "中"
         Height          =   180
         Index           =   118
         Left            =   -74130
         TabIndex        =   302
         Top             =   2400
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "中"
         Height          =   180
         Index           =   119
         Left            =   -74130
         TabIndex        =   301
         Top             =   3375
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "中"
         Height          =   180
         Index           =   120
         Left            =   -74130
         TabIndex        =   300
         Top             =   4350
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   121
         Left            =   -74130
         TabIndex        =   299
         Top             =   1065
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   122
         Left            =   -74130
         TabIndex        =   298
         Top             =   2025
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   123
         Left            =   -74130
         TabIndex        =   297
         Top             =   3000
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   124
         Left            =   -74130
         TabIndex        =   296
         Top             =   3975
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Index           =   125
         Left            =   -74130
         TabIndex        =   295
         Top             =   4950
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   180
         Left            =   -74745
         TabIndex        =   294
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   180
         Left            =   -74745
         TabIndex        =   293
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   180
         Left            =   -74745
         TabIndex        =   292
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   180
         Left            =   -74745
         TabIndex        =   291
         Top             =   1470
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊請款單列印對象："
         Height          =   180
         Index           =   177
         Left            =   -74820
         TabIndex        =   286
         Top             =   4110
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否加註核准分割建議：       ( Y：是 N：否)"
         Height          =   180
         Index           =   175
         Left            =   -74820
         TabIndex        =   285
         Top             =   3000
         Width           =   3465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否初審階段提分割/改請：       ( Y：是 N：否)"
         Height          =   180
         Index           =   174
         Left            =   -74820
         TabIndex        =   284
         Top             =   3300
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司：             ( T：專利商標 J：智權公司 空白:系統預設)"
         Height          =   180
         Index           =   173
         Left            =   -74580
         TabIndex        =   283
         Top             =   2715
         Width           =   5190
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "以EMail通知：             （Y:是   D:僅D/N）"
         Height          =   180
         Left            =   -74580
         TabIndex        =   282
         Top             =   2100
         Width           =   3195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email 同時寄紙本：       (Y:是)"
         Height          =   180
         Index           =   166
         Left            =   -74580
         TabIndex        =   281
         Top             =   2400
         Width           =   2310
      End
      Begin VB.Label Label1 
         Caption         =   "請款單份數："
         Height          =   180
         Index           =   165
         Left            =   -74580
         TabIndex        =   280
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "定稿份數："
         Height          =   180
         Index           =   164
         Left            =   -74580
         TabIndex        =   279
         Top             =   1470
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "帳單備註是否提醒：          (N:否)"
         Height          =   180
         Index           =   159
         Left            =   -74820
         TabIndex        =   278
         Top             =   1140
         Width           =   2535
      End
      Begin VB.Label Label70 
         Caption         =   "帳單備註："
         Height          =   255
         Left            =   -74820
         TabIndex        =   277
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "一案兩請是否放棄新型：       ( Y：是 N：否)"
         Height          =   180
         Index           =   67
         Left            =   -74820
         TabIndex        =   276
         Top             =   3570
         Width           =   3465
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "請款單列印幣別格式："
         Height          =   180
         Left            =   -69780
         TabIndex        =   275
         Top             =   1140
         Width           =   1800
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "請款幣別："
         Height          =   180
         Index           =   0
         Left            =   -71970
         TabIndex        =   274
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFP有無關聯P案：         (  N：無)"
         Height          =   180
         Index           =   68
         Left            =   -74715
         TabIndex        =   273
         Top             =   3840
         Width           =   2565
      End
      Begin VB.Label Label1 
         Caption         =   "不可銷卷案請加註 ""不銷卷"" 字樣！  與他案合併計算結餘，請於案件備註欄註明""與某案號合併計算結餘""！"
         ForeColor       =   &H000000FF&
         Height          =   540
         Index           =   172
         Left            =   -74970
         TabIndex        =   258
         Top             =   5055
         Width           =   8745
      End
      Begin VB.Label Label1 
         Caption         =   "FCP 實審自動代繳:          (Y:自動代繳)"
         Height          =   180
         Index           =   128
         Left            =   150
         TabIndex        =   257
         Top             =   3255
         Width           =   2925
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
         Left            =   -67680
         TabIndex        =   256
         Top             =   5850
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
         Left            =   -67680
         TabIndex        =   255
         Top             =   5340
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費逾期補繳通知函是否寄發:        (N:不寄)"
         Height          =   180
         Index           =   61
         Left            =   5325
         TabIndex        =   254
         Top             =   1425
         Width           =   3390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "優先權存取碼:"
         Height          =   180
         Index           =   176
         Left            =   -68655
         TabIndex        =   253
         Top             =   2850
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新穎性優惠日期:"
         Height          =   180
         Index           =   171
         Left            =   -68655
         TabIndex        =   252
         Top             =   3150
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國際分類:"
         Height          =   180
         Index           =   170
         Left            =   -68655
         TabIndex        =   251
         Top             =   2550
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID:"
         Height          =   180
         Index           =   169
         Left            =   4095
         TabIndex        =   250
         Top             =   3885
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件屬性:"
         Height          =   180
         Index           =   168
         Left            =   -68805
         TabIndex        =   249
         Top             =   690
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人是否同發明人:        (Y/ N)"
         Height          =   180
         Index           =   167
         Left            =   -68700
         TabIndex        =   248
         Top             =   3750
         Width           =   2475
      End
      Begin VB.Label lblFilingDate 
         AutoSize        =   -1  'True
         Caption         =   "提交日:"
         Height          =   180
         Index           =   0
         Left            =   -68655
         TabIndex        =   247
         Top             =   2235
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblFilingDate 
         AutoSize        =   -1  'True
         Caption         =   "lblFilingDate"
         Height          =   180
         Index           =   1
         Left            =   -67800
         TabIndex        =   246
         Top             =   2235
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費折扣:          %"
         Height          =   180
         Index           =   163
         Left            =   5835
         TabIndex        =   245
         Top             =   1710
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "領證折扣:          %"
         Height          =   180
         Index           =   162
         Left            =   4125
         TabIndex        =   244
         Top             =   1710
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP工程師組別:"
         Height          =   180
         Index           =   161
         Left            =   -69645
         TabIndex        =   243
         Top             =   1005
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C類收文是否請款:          (N:否)"
         Height          =   180
         Index           =   158
         Left            =   3225
         TabIndex        =   242
         Top             =   3255
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費本所是否出名：       (N：不出名)"
         Height          =   180
         Index           =   157
         Left            =   5868
         TabIndex        =   241
         Top             =   3252
         Width           =   2916
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人：         (Y:是)"
         Height          =   180
         Index           =   47
         Left            =   6105
         TabIndex        =   240
         Top             =   2655
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否核對已准專利:            (N:否)"
         Height          =   180
         Index           =   156
         Left            =   3270
         TabIndex        =   239
         Top             =   2355
         Width           =   2790
      End
      Begin MSForms.Label Label2 
         Height          =   288
         Index           =   44
         Left            =   -73740
         TabIndex        =   205
         Top             =   4968
         Width           =   1884
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3323;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   480
         Index           =   50
         Left            =   -73776
         TabIndex        =   238
         Top             =   5256
         Width           =   1272
         ForeColor       =   49152
         Caption         =   "尚未公告，下次繳費日為預估值"
         Size            =   "2249;847"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費聯絡人:"
         Height          =   180
         Index           =   155
         Left            =   150
         TabIndex        =   237
         Top             =   3885
         Width           =   945
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   48
         Left            =   2670
         TabIndex        =   234
         Top             =   3570
         Width           =   1290
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2275;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象:"
         Height          =   180
         Index           =   55
         Left            =   150
         TabIndex        =   233
         Top             =   3585
         Width           =   1410
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   49
         Left            =   6630
         TabIndex        =   236
         Top             =   3585
         Width           =   1920
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3387;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費D/N列印對象:"
         Height          =   180
         Index           =   154
         Left            =   4110
         TabIndex        =   235
         Top             =   3585
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FC代理人:"
         Height          =   180
         Index           =   59
         Left            =   150
         TabIndex        =   232
         Top             =   450
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費代理人:"
         Height          =   180
         Index           =   64
         Left            =   150
         TabIndex        =   231
         Top             =   1425
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號:"
         Height          =   180
         Index           =   65
         Left            =   150
         TabIndex        =   230
         Top             =   765
         Width           =   765
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   2
         Left            =   2370
         TabIndex        =   229
         Top             =   405
         Width           =   6330
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "11165;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   12
         Left            =   2370
         TabIndex        =   228
         Top             =   1380
         Width           =   2910
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5133;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   13
         Left            =   2370
         TabIndex        =   204
         Top             =   2040
         Width           =   480
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "847;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   5
         Left            =   -69480
         TabIndex        =   203
         Top             =   4665
         Width           =   3300
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   1
         Left            =   -73200
         TabIndex        =   202
         Top             =   990
         Width           =   1410
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2487;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   285
         Index           =   0
         Left            =   -73200
         TabIndex        =   201
         Top             =   720
         Width           =   1380
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2434;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   200
         Top             =   408
         Width           =   768
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利名稱　(中):"
         Height          =   180
         Index           =   1
         Left            =   -74700
         TabIndex        =   199
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "　　　　(英):"
         Height          =   180
         Index           =   2
         Left            =   -74520
         TabIndex        =   198
         Top             =   1605
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "　　　　(外):"
         Height          =   180
         Index           =   3
         Left            =   -74520
         TabIndex        =   197
         Top             =   1905
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利種類:"
         Height          =   180
         Index           =   4
         Left            =   -74760
         TabIndex        =   196
         Top             =   1005
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請國家:"
         Height          =   180
         Index           =   5
         Left            =   -74760
         TabIndex        =   195
         Top             =   690
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請日期:"
         Height          =   180
         Index           =   6
         Left            =   -74760
         TabIndex        =   194
         Top             =   2205
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公開日:"
         Height          =   180
         Index           =   7
         Left            =   -74760
         TabIndex        =   193
         Top             =   2505
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公告日:"
         Height          =   180
         Index           =   8
         Left            =   -74760
         TabIndex        =   192
         Top             =   2805
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發證日:"
         Height          =   180
         Index           =   9
         Left            =   -74760
         TabIndex        =   191
         Top             =   3105
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利號數:"
         Height          =   180
         Index           =   10
         Left            =   -74760
         TabIndex        =   190
         Top             =   3405
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "目前准/駁:                    (1.准2.駁)"
         Height          =   180
         Index           =   11
         Left            =   -74760
         TabIndex        =   189
         Top             =   3735
         Width           =   2760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否有救濟程序          (Y:有)"
         Height          =   180
         Index           =   12
         Left            =   -74760
         TabIndex        =   188
         Top             =   4035
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文:            (1.中文 2.英文 3.日文)"
         Height          =   180
         Index           =   13
         Left            =   -74760
         TabIndex        =   187
         Top             =   4365
         Width           =   3060
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期:"
         Height          =   180
         Index           =   14
         Left            =   -74760
         TabIndex        =   186
         Top             =   4695
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日:"
         Height          =   180
         Index           =   15
         Left            =   -74760
         TabIndex        =   185
         Top             =   4965
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否PCT案件:        (Y:是)"
         Height          =   180
         Index           =   16
         Left            =   -71670
         TabIndex        =   184
         Top             =   1005
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請案號:"
         Height          =   180
         Index           =   17
         Left            =   -71700
         TabIndex        =   183
         Top             =   2250
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公開號:"
         Height          =   180
         Index           =   18
         Left            =   -71700
         TabIndex        =   182
         Top             =   2550
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公告號:"
         Height          =   180
         Index           =   19
         Left            =   -71700
         TabIndex        =   181
         Top             =   2850
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "准駁通知日:"
         Height          =   180
         Index           =   21
         Left            =   -71520
         TabIndex        =   180
         Top             =   3150
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專用期限:                     －                    (西元年月日)"
         Height          =   180
         Index           =   22
         Left            =   -71520
         TabIndex        =   179
         Top             =   3456
         Width           =   3936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利權是否存在:            (Y/ N)"
         Height          =   180
         Index           =   23
         Left            =   -71520
         TabIndex        =   178
         Top             =   3735
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否有爭議程序:            (Y:有)"
         Height          =   180
         Index           =   24
         Left            =   -71520
         TabIndex        =   177
         Top             =   4035
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷:                        (Y:閉卷)"
         Height          =   180
         Index           =   25
         Left            =   -71520
         TabIndex        =   176
         Top             =   4365
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "閉卷原因:"
         Height          =   180
         Index           =   26
         Left            =   -71520
         TabIndex        =   175
         Top             =   4695
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分所案號:"
         Height          =   180
         Index           =   27
         Left            =   -71520
         TabIndex        =   174
         Top             =   4965
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "卷宗性質:        (1.申請 2.異議 3.舉發)"
         Height          =   180
         Index           =   28
         Left            =   -71370
         TabIndex        =   173
         Top             =   408
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "全部折扣:                %"
         Height          =   180
         Index           =   32
         Left            =   150
         TabIndex        =   172
         Top             =   1710
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象:"
         Height          =   180
         Index           =   34
         Left            =   150
         TabIndex        =   171
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "FCP 領證自動代繳:          (Y:自動代繳)"
         Height          =   180
         Index           =   35
         Left            =   150
         TabIndex        =   170
         Top             =   2355
         Width           =   2925
      End
      Begin VB.Label Label1 
         Caption         =   "信函是否列印Title：         (Y:印)"
         Height          =   180
         Index           =   36
         Left            =   6255
         TabIndex        =   169
         Top             =   2355
         Width           =   2505
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "後續准駁簡單報告:          (Y:核准以及C類來函簡單報告)"
         Height          =   180
         Index           =   37
         Left            =   132
         TabIndex        =   168
         Top             =   1092
         Width           =   4344
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請/翻譯折扣:          %"
         Height          =   180
         Index           =   45
         Left            =   2025
         TabIndex        =   167
         Top             =   1710
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP 年費自動代繳:          (Y:自動代繳)"
         Height          =   180
         Index           =   48
         Left            =   150
         TabIndex        =   166
         Top             =   2655
         Width           =   2925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利不得請雜費：         (Y:是)"
         Height          =   180
         Index           =   134
         Left            =   6384
         TabIndex        =   408
         Top             =   2016
         Width           =   2340
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "相關卷號(&F)"
      Height          =   300
      Left            =   7515
      TabIndex        =   161
      Top             =   615
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "優先權資料(P)"
      Height          =   300
      Left            =   6180
      TabIndex        =   160
      Top             =   615
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "分割案資料(D)"
      Height          =   300
      Left            =   4830
      Style           =   1  '圖片外觀
      TabIndex        =   159
      Top             =   615
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8130
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":01E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":04FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":0819
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":09F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":0D11
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":102D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":1349
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":1665
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":1981
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":1C9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050701.frx":1FB9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   162
      Top             =   0
      Width           =   8952
      _ExtentX        =   15790
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox textCUID 
      Height          =   285
      Left            =   30
      TabIndex        =   367
      TabStop         =   0   'False
      Top             =   645
      Width           =   4725
      VariousPropertyBits=   16415
      BackColor       =   16777215
      Size            =   "8334;503"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   165
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm050701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2024/12/03 直接變更物件Label2(15)->txtInvField(0)、Label2(14)->txtInvField(1)、Label2(16)->txtInvField(2)
                           '參考外專 -新案建檔作業的發明人維護:
                           '1.增加可以改變順序的按鈕
                           '2.輸入發明人的中文、英文或日文名稱可以自動帶出符合的發明人編號。
'end 2024/12/03
'Memo By Sindy 2021/12/07 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

'edit by nickc 2006/07/12
'Dim pA(0 To T_PA) As String
Dim pa() As String
Dim strPriority(1 To 5) As String 'Modify by Amy 2014/03/24 +存取碼
Dim varYear As Variant '繳費年度陣列
Dim strRsStart As String, strRsEnd As String, rsDefineSize As New ADODB.Recordset
Dim intWhere As Integer
'edit by nickc 2006/12/06 改可以外部傳
'Dim ActionEdit As Integer
Public ActionEdit As Integer  'Memo by Lydia 2020/02/21 0:新增,1:修改, 2:查詢, 3:無動作
Dim intRow As Integer
' 90.07.31 modify by louis
Dim m_SysKind As String
' 目前正在顯示的本所案號
Dim m_CurrPA(4) As String
' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'Add By Cheng 2003/08/07
Dim m_CustNo(1 To 5) As String '原申請人
Dim m_FixNo As Integer   '2010/2/12 add by sonia 修法次數
Dim cmd As CommandButton
Dim m_bolFmpAuth As Boolean '是否FMP權限(只能改FC特定欄位) Added by Morgan 2012/5/17
Dim m_form As Form 'Add By Morgan 2012/6/18
Dim pPrevRow As Integer 'Add By Sindy 2014/11/6
Dim pPA12 As String 'Added by Lydia 2015/09/09 記錄修改前的公開日
Dim strChkCuAreaMail As String, strChkCuAreaMailTo As String 'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
'Added by Lydia 2020/02/21
Public bolAskPA174 As Boolean '存檔前檢查有修改案件名稱，將原始檔之維護word檔自動打開，是否有上傳
Public intOpen090801 As Integer  'Add by Amy 2022/12/26 已開接洽單
Dim m_bolFMP As Boolean 'Added by Lydia 2023/03/09
'Added by Lydia 2024/12/03 發明人輸入比對兼自動代入(模糊比對)
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
Dim obj As Object


'Add By Cheng 2002/01/03
Public Sub SelectToolbarButtom()
Dim btn
'設定為按下查詢鈕扭
Set btn = Me.TBar1.Buttons(4)
Tbar1_ButtonClick btn
End Sub

'Modified by Morgan 2012/6/18 + pCallForm
Public Sub SetCurrKey(Optional ByVal strKEY01 As String = Empty, Optional ByVal strKEY02 As String = Empty, Optional ByVal strKEY03 As String = Empty, Optional ByVal strKEY04 As String = Empty, Optional ByRef pCallForm As Form)
   If IsEmptyText(strKEY01) Or IsEmptyText(strKEY02) Then
      m_CurrPA(0) = Empty
      m_CurrPA(1) = Empty
      m_CurrPA(2) = Empty
      m_CurrPA(3) = Empty
      Exit Sub
   End If
   m_CurrPA(0) = strKEY01
   m_CurrPA(1) = strKEY02
   m_CurrPA(2) = strKEY03
   If IsEmptyText(m_CurrPA(2)) Then
      m_CurrPA(2) = "0"
   End If
   m_CurrPA(3) = strKEY04
   If IsEmptyText(m_CurrPA(3)) Then
      m_CurrPA(3) = "00"
   End If
   
   'Added by Morgan 2012/6/18
   If pCallForm Is Nothing = False Then
      Set m_form = pCallForm
      m_form.Enabled = False
   End If
End Sub

'Added by Morgan 2015/6/11
Private Sub cmdAdd_Click()
   If txtNo <> "" And lblName <> "" Then
      If lstPA166.ListCount = 5 Then
         MsgBox "目前只開放 5 個列印對象!!"
         Exit Sub
      End If
      If AddList(lstPA166, txtNo & " " & lblName) = True Then
         Text1(166) = ComposeList(lstPA166)
         txtNo = ""
         lblName = ""
         txtNo.SetFocus
      Else
         txtNo.SetFocus
         txtNo_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2014/11/6
Private Sub cmdAddRow_Click()
Dim bolChk As Boolean
Dim ii As Integer
Dim Cancel As Boolean
Dim rsTmp  As New ADODB.Recordset
Dim strNo As String

   'Added by Lydia 2024/12/03
   If Trim(Text1(26)) = "" Then
        MsgBox "請輸入申請人1編號！", vbCritical, "資料檢核"
        Exit Sub
   End If
   'end 2024/12/03
    
   '檢查發明人
   'Modified by Lydia 2024/12/03
'   bolChk = True
'   strExc(1) = Replace(Right(Combo1.Text, 11), ")", "")
'   If strExc(1) = "" Then Exit Sub
'   For ii = 1 To GRD1.Rows - 1
'      If GRD1.TextMatrix(ii, 1) = strExc(1) Then
'         bolChk = False
'         Exit For
'      End If
'   Next ii
'   If Not bolChk Then
'      MsgBox "發明人不可重覆 !", vbCritical
'      Combo1.SetFocus
'      Exit Sub
'   End If
   strExc(1) = Replace(Right(Combo1.Text, 11), ")", "")
   If strExc(1) = "" Then
      If txtInvField(0) = "" And txtInvField(1) = "" And txtInvField(2) = "" Then
         Exit Sub
      Else
         '判斷國籍是否有輸入
         If txtIN11.Visible = True Then
            If txtIN11 = "" Then
                MsgBox "請輸入國藉！", vbExclamation
                SSTab1.Tab = 5
                txtIN11.SetFocus
                Exit Sub
            Else
                Cancel = False
                txtIN11_Validate Cancel
                If Cancel = True Then
                  SSTab1.Tab = 5
                  txtIN11.SetFocus
                  TextInverse txtIN11
                  Exit Sub
                End If
            End If
         End If
         '判斷客戶發明人檔是否有重覆資料:發明人會有造字無法存檔時會加空白,所以改在語法內trim
         If Len(Text1(26)) < 8 Then
            strNo = Text1(26) & String(8 - Len(Text1(26)), "0")
         Else
            strNo = Left(Text1(26), 8)
         End If
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
               Combo1.SetFocus
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
               Combo1.SetFocus
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
               Combo1.SetFocus
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
         Combo1.SetFocus
         Exit Sub
      End If
   End If
   'end 2024/12/03
   
   If Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) <> "" Then
      GRD1.AddItem ""
   End If
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 1) = strExc(1)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 2) = txtInvField(0)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 3) = txtInvField(1)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 4) = txtInvField(2)
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 5) = Lb_IN11N 'Add By Sindy 2019/4/9
   Me.GRD1.TextMatrix(Me.GRD1.Rows - 1, 7) = txtIN11 'Added by Lydia 2024/12/03
   cmdAddRow.Tag = "I" '記錄有異動資料
   '清空欄位
   Combo1.ListIndex = 0
   txtInvField(0) = ""
   txtInvField(1) = ""
   txtInvField(2) = ""
   txtIN11.Text = "" 'Add By Sindy 2019/4/9
   Lb_IN11N.Caption = "" 'Add By Sindy 2019/4/9
End Sub

'Add By Sindy 2014/11/6
Private Sub cmdDelRow_Click()
   'Added by Lydia 2024/12/03
   If pPrevRow <= 0 Then Exit Sub
   GRD1.col = 0
   GRD1.row = pPrevRow
   If GRD1.CellBackColor <> &HFFC0C0 Then Exit Sub
   'end 2024/12/03
   If pPrevRow = 1 And GRD1.Rows = 2 Then
      GRD1.TextMatrix(pPrevRow, 0) = ""
      GRD1.TextMatrix(pPrevRow, 1) = ""
      GRD1.TextMatrix(pPrevRow, 2) = ""
      GRD1.TextMatrix(pPrevRow, 3) = ""
      GRD1.TextMatrix(pPrevRow, 4) = ""
      'Added by Lydia 2024/12/03
      GRD1.TextMatrix(pPrevRow, 5) = ""
      GRD1.TextMatrix(pPrevRow, 6) = ""
      'end 2024/12/03
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
   Combo1.ListIndex = 0
   txtInvField(0) = ""
   txtInvField(1) = ""
   txtInvField(2) = ""
   'Added by Lydia 2024/12/03
   txtIN11.Text = ""
   Lb_IN11N = ""
   'end 2024/12/03
End Sub

Private Sub cmdDivSug_Click()
   strExc(0) = "select dst05 from divsugtext where dst01='" & Text1(1) & "' and dst02='" & Text1(2) & "' and dst03='" & Text1(3) & "' and dst04='" & Text1(4) & "' and dst05 is not null"
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

'Added by Lydia 2016/11/24 各項指示
Private Sub cmdIns_Click()
   If Me.Text1(1).Text = "" Or Me.Text1(2).Text = "" Then
      MsgBox "請輸入本所案號", vbInformation
      Exit Sub
   End If
   'Added by Lydia 2020/05/05
   If ActionEdit <> 2 And ActionEdit <> 3 Then
      MsgBox IIf(ActionEdit = 0, "新增中", "修改中") & "不可執行！", vbInformation
      Exit Sub
   End If
   'end 2020/05/05
   
   'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
   If PUB_CheckFormExist("frm12040159") Then
       MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
       Exit Sub
   End If
   'end 2020/05/05
   
   frm12040159.SetParent "E", Trim(Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text & Me.Text1(4).Text), Me
   frm12040159.Show
End Sub

Private Sub cmdRemove_Click()
   If lstPA166.ListCount > 0 Then
      RemoveList lstPA166
      Text1(166) = ComposeList(lstPA166)
   End If
End Sub

'Added by Lydia 2024/12/03
Private Sub cmdUpdRow_Click()
   Me.GRD1.TextMatrix(pPrevRow, 2) = txtInvField(0)
   Me.GRD1.TextMatrix(pPrevRow, 3) = txtInvField(1)
   Me.GRD1.TextMatrix(pPrevRow, 4) = txtInvField(2)
   Me.GRD1.TextMatrix(pPrevRow, 5) = Lb_IN11N
   Me.GRD1.TextMatrix(pPrevRow, 7) = txtIN11
   cmdUpdRow.Enabled = False
End Sub

Private Sub Combo1_Click()
Dim strMain As String, i As Integer
   
   'Modify By Cheng 2002/07/02
'   strMain = Combo1(Index).Text
   strMain = Replace(Right(Combo1.Text, 11), ")", "")
   'Modified by Lydia 2024/12/03
   'For i = 0 To 2
   '   Label2(14 + i).Caption = ""
   'Next
   For i = 0 To 2
      txtInvField(i).Text = ""
      txtInvField(i).Tag = "" 'Add By Sindy 2015/3/5
   Next
   txtIN11 = ""
   Lb_IN11N = ""
   'end 2024/12/03
   If strMain = "" Then Exit Sub
   
   'Added by Lydia 2024/12/03
   If ActionEdit = 0 Or ActionEdit = 1 Then
      cmdUpdRow.Enabled = False
      cmdAddRow.Enabled = True
      Frame3.Enabled = True
   Else
      Frame3.Enabled = False
   End If
   'end 2024/12/03
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If objLawDll.GetInventor(strMain, strExc) Then
   If ClsLawGetInventor(strMain, strExc) Then
      'Modified by Lydia 2024/12/03
      'For i = 0 To 2
      '   Label2(14 + i).Caption = strExc(i + 1)
      'Next
      For i = 0 To 2
         txtInvField(i).Text = strExc(i + 1)
         txtInvField(i).Tag = txtInvField(i).Text
      Next
      'end 2024/12/03
      
      'Add By Sindy 2019/4/9
      txtIN11 = strExc(6)
      Call txtIN11_Validate(False)
      '2019/4/9 END
      Frame3.Enabled = False   'Added by Lydia 2024/12/03 控制發明人欄位是否可點選
   End If
  
End Sub

'2010/1/8 ADD BY SONIA
Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 <> "" Then
      Combo2 = Left(Combo2, 1) + "." + PUB_GetFCPGrpName(Left(Combo2, 1))
      If Combo2 = Left(Combo2, 1) + "." Then
         Combo2 = Left(Combo2, 1)
         Cancel = True
         Combo2.SetFocus
      End If
   End If
End Sub
'2010/1/8 end

'Add By Sindy 2010/10/27
Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3 <> "" Then
      'Modify By Sindy 2014/7/8
      'Combo3 = Left(Combo3, 1) + "." + PUB_GetCaseAttributeName(Left(Combo3, 1))
      Combo3 = Left(Combo3, 1) + "." + PUB_GetCaseAttributeName(Left(Combo3, 1), Text1(8))
      '2014/7/8 END
      If Combo3 = Left(Combo3, 1) + "." Then
         Combo3 = Left(Combo3, 1)
         Cancel = True
         Combo3.SetFocus
      End If
   End If
End Sub
'2010/10/27 End

'Add By Sindy 2016/11/23
Private Sub Combo4_Click()
   Call GetCurrType
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo4_Validate(Cancel As Boolean)
   If Combo4 = MsgText(601) Then
      Combo4.Tag = Combo4.Text
      Combo5.ListIndex = 0
      Combo5.Enabled = False
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo4, Label11(0)) = False Then
      Cancel = True
      Combo4.SetFocus
   End If
   If Combo4 <> "USD" Then
      If ExistCheck("DebitNoteRate", "DNR01", Combo4, Label11(0) & "匯率") = False Then
         Cancel = True
         Combo4.SetFocus
         Exit Sub
      End If
   End If
   Call GetCurrType
End Sub
Private Sub GetCurrType()
Dim intType As Integer
   
   If Combo4 = MsgText(601) Then
      Combo4.Tag = Combo4.Text
      Combo5.ListIndex = 0
      Combo5.Enabled = False
      Exit Sub
   End If
   '若更改請款幣別
   If Me.Combo4.Text <> Me.Combo4.Tag Then
      Me.Combo4.Tag = Me.Combo4.Text
      '請款幣別變更要重新預設列印幣別
      '台幣
      If Me.Combo4.Text = "NTD" Then
         intType = 1 '純台幣
      '人民幣
      ElseIf Me.Combo4.Text = "RMB" Then
         intType = 4 '外幣+美金合計
      '其他幣別
      Else
         intType = 2 '台幣+外幣合計
      End If
      Combo5.ListIndex = intType
      '若為台幣時則格式欄位鎖住不可修改
      If Me.Combo4.Text = "NTD" Then
         Combo5.Enabled = False
      Else
         Combo5.Enabled = True
      End If
   End If
End Sub
'2016/11/23 END

Private Sub Command1_Click(Index As Integer)
Dim strTmp As String, i As Integer
'Add By Sindy 2009/06/29
Dim strFeeType As String, strYF15 As String

'On Error GoTo ErrHand
   Select Case Index
      Case 0
         If MSHFlexGrid1.Rows - 2 >= UBound(varYear) Then MsgBox "無繳費年度，無法新增資料 !", vbCritical: Exit Sub
         
         'Add by Morgan 2010/3/10
         If Text1(9) = "000" And Text3(0) = "" Then
            MsgBox "台灣案繳費日期不可空白！"
            Text3(0).SetFocus
            Exit Sub
         End If
         
         'Modify By Sindy 2009/06/29
         'strTmp = varYear(MSHFlexGrid1.Rows - 1)
         '繳費次數
         strTmp = MSHFlexGrid1.Rows
         '年度說明
         'Modified by Morgan 2022/6/13 +pa(10)
         strFeeType = PUB_GetNa20Na22Na24(pa(9), pa(8), pa(10))
         'strYF15 = PUB_GetYF15(pa(9), pa(8), "Y0000000", strFeeType, MSHFlexGrid1.Rows)
         '2010/2/12 modify by sonia
         'strYF15 = PUB_GetYF15(pa(9), pa(8), "Y0000000", strFeeType, CDbl(varYear(MSHFlexGrid1.Rows - 1)))
         strYF15 = PUB_GetYF15(pa(9), pa(8), "Y000000" & m_FixNo, strFeeType, CDbl(varYear(MSHFlexGrid1.Rows - 1)))
         strTmp = strTmp & vbTab & strYF15
         '2009/06/29 End
         '繳費日期
         If Text3(0) = "" Then
            strTmp = strTmp & vbTab & ""
         Else
            strTmp = strTmp & vbTab & Text3(0)
         End If
         '費用是否雙倍
         If Text3(1) <> "" Then strTmp = strTmp & vbTab & Text3(1)
         MSHFlexGrid1.AddItem strTmp
        'Add By Cheng 2002/12/16
         FixGrid MSHFlexGrid1
         'Modify By Sindy 2009/06/29
         'If Me.MSHFlexGrid1.Rows > 1 Then GridClick MSHFlexGrid1, 1, 4
         If Me.MSHFlexGrid1.Rows > 1 Then GridClick MSHFlexGrid1, 1, 5
         
         Text3(0) = ""
         Text3(1) = ""
         Text3(0).SetFocus
      Case 1
         For i = MSHFlexGrid1.row To MSHFlexGrid1.Rows - 1
            MSHFlexGrid1.RemoveItem MSHFlexGrid1.row
         Next
         'Modify By Sindy 2009/06/29
         'GridClick MSHFlexGrid1, MSHFlexGrid1.Rows, 4
         GridClick MSHFlexGrid1, MSHFlexGrid1.Rows, 5
      Case 2
         GridHead
         FixGrid MSHFlexGrid1
         
      'Add by Morgan 2010/3/2
      Case 3
         'Add by Morgan 2010/3/10
         If Text1(9) = "000" And Text3(0) = "" Then
            MsgBox "台灣案繳費日期不可空白！"
            Text3(0).SetFocus
            Exit Sub
         End If
         
         MSHFlexGrid1.TextMatrix(intRow, 2) = Text3(0)
         MSHFlexGrid1.TextMatrix(intRow, 3) = Text3(1)
         
   End Select
   Exit Sub
ErrHand:
   If Err.Number = 30015 Then
      GridHead
      FixGrid MSHFlexGrid1
   End If
End Sub

Private Sub Command2_Click()
   Where1103ComeFrom Me, pa(1), pa(2), pa(3), pa(4)
End Sub

Private Sub Command3_Click()
   'Modify by Morgan 2007/4/24
   'ModifyPriority strPriority(1), strPriority(2), strPriority(3)
   'Modify by Amy 2014/03/24 +strPriority(5)
   ModifyPriority strPriority(1), strPriority(2), strPriority(3), pa(8), , pa(1) & pa(2) & pa(3) & pa(4), pa(9), , strPriority(4), strPriority(5)
End Sub

Private Sub Command4_Click()
    frm02010604_3.m_CP01 = Me.Text1(1).Text
    frm02010604_3.m_CP02 = Me.Text1(2).Text
    frm02010604_3.m_CP03 = IIf(Me.Text1(3).Text = "", "0", Me.Text1(3).Text)
    frm02010604_3.m_CP04 = IIf(Me.Text1(4).Text = "", "00", Me.Text1(4).Text)
    frm02010604_3.intWhereToGo = 1
    frm02010604_3.Show
End Sub

Private Sub Form_Initialize()
ReDim pa(0 To TF_PA) As String
'Added by Lydia 2015/11/03
lblCaseMap.Caption = ""
lblCaseMap2.Caption = "" 'Added by Lydia 2019/11/28
Command4.BackColor = &H8000000F
End Sub

'Modify By Cheng 2003/08/18
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5
         If ActionEdit = 3 Then
            Select Case KeyCode
               Case vbKeyF2
                  RsSitu 0
               Case vbKeyF3
                  RsSitu 1
               Case vbKeyF5
                  RsSitu 2
               Case vbKeyF4
                  RsSitu 5
            End Select
            KeyCode = 0
         End If
      'Modify By Cheng 2001/12/25
'      Case vbKeyF9, vbKeyF10
      'edit by nickc 2006/11/10
      'Case vbKeyF9, vbKeyF10, vbKeyReturn
      Case vbKeyF9, vbKeyF10
         If ActionEdit <> 3 Then
            Select Case KeyCode
               'Modify By Cheng 2001/12/25
'               Case vbKeyF9
               Case vbKeyF9, vbKeyReturn
                  RsSitu 3
               Case vbKeyF10
                  RsSitu 4
            End Select
            KeyCode = 0
         End If
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd
         If ActionEdit = 3 Then
            Select Case KeyCode
               Case vbKeyHome
                  RsAction 0
               Case vbKeyPageUp
                  RsAction 1
               Case vbKeyPageDown
                  RsAction 2
               Case vbKeyEnd
                  RsAction 3
            End Select
            KeyCode = 0
         End If
    Case vbKeyEscape
        If MsgBox("是否確定結束?", vbYesNo + vbCritical) = vbYes Then Unload Me
    End Select
   
'Remove by Morgn 2009/9/14 集中到 CmdSitu 處理
'   ' Ken 90.07.19 -- Start
'   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
'         If m_bInsert Then
'             TBar1.Buttons(1).Enabled = True
'         Else
'             TBar1.Buttons(1).Enabled = False
'         End If
'         If m_bUpdate Then
'             TBar1.Buttons(2).Enabled = True
'         Else
'             TBar1.Buttons(2).Enabled = False
'         End If
'         If m_bDelete Then
'             TBar1.Buttons(3).Enabled = True
'         Else
'             TBar1.Buttons(3).Enabled = False
'         End If
'   End If
'   ' Ken 90.07.19 -- End
End Sub

'Add by Morgan 2004/4/20
'改以使用者權限控管
Private Sub AuthCheck(ByVal stSys As String)

   Dim arrSys, stTmp
   
   arrSys = Split(Systemkind_g, ",")
   For Each stTmp In arrSys
      If stTmp = stSys Then
         m_SysKind = stSys
         Select Case m_SysKind
            Case "P"
               intWhere = 國內
            Case "CFP"
               intWhere = 國外_CF
            Case "FCP"
               intWhere = 國外_FC
            Case "ALL"
               intWhere = 國外_FC
         End Select
         Exit For
      End If
   Next
   
   'Modified by Morgan 2012/5/17 +F2部門可以改FMP案(查詢確定時才檢查)
   'Modified by Morgan 2023/5/12 合併寰華案控制
   'm_bolFmpAuth = False
   'If m_SysKind <> stSys Then
   '   If Left(Pub_StrUserSt03, 2) = "F2" Then
   '      m_SysKind = stSys
   '      m_bolFmpAuth = True
   '   End If
   'End If
   If m_SysKind <> stSys And FMP2open Then m_SysKind = stSys
   'end 2023/5/12
End Sub

'add by nickc 2006/11/10 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
   'Add By Sindy 2014/8/29 當focus在備註欄時按enter鍵維持換行功能而不是存檔功能
   If KeyAscii = 13 And UCase(Me.ActiveControl.Name) = UCase("Text1") Then
      If Me.ActiveControl.Index = 91 Then
         Exit Sub
      End If
   End If
   '2014/8/29 END
   Select Case KeyAscii
      Case 13:
         If ActionEdit <> 3 Then
            KeyAscii = 0
            RsSitu 3
         End If
   End Select
End Sub

Private Sub Form_Load()
Dim i As Integer
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
    
   ' 90.10.18 modify by louis
   Command3.Enabled = False
 
   ' 90.07.16 modify by Ken (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm050701", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm050701", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm050701", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm050701", strFind, False)
   ' Ken 90.07.16 -- End
   
   textCUID.BackColor = &H8000000F
   Lb_IN11N.Caption = ""
   
   ' 90.07.31 modify by louis
   m_SysKind = strSysKind
   
   MoveFormToCenter Me
   
   Frame3.BackColor = &H8000000F 'Added by Lydia 2024/12/03
   
   Label1(22).Tag = Label1(22)  'Added by Morgan 2024/10/15
   
   ' 90.06.29 modify by louis 設定顯示在第一頁
   SSTab1.Tab = 0
   strExc(0) = "SELECT * FROM PATENT WHERE ROWNUM<1"
   intI = 1
   Set rsDefineSize = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      
   Select Case m_SysKind
      Case "P"
         intWhere = 國內
      Case "CFP"
         intWhere = 國外_CF
      Case "FCP"
         intWhere = 國外_FC
      Case "ALL"
         intWhere = 國外_FC
   End Select
      
   Text1(10).MaxLength = 7
   Text1(12).MaxLength = 7
   Text1(14).MaxLength = 7
   Text1(20).MaxLength = 7
   Text1(21).MaxLength = 7
   Text1(24).MaxLength = 8
   Text1(25).MaxLength = 8
   Text1(58).MaxLength = 7
   
   'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    Text1(79).MaxLength = Pub_MaxCEL10
    Text1(80).MaxLength = Pub_MaxCEL11
    Text1(82).MaxLength = Pub_MaxCEL10
    Text1(83).MaxLength = Pub_MaxCEL11
    Text1(109).MaxLength = Pub_MaxCEL10
    Text1(110).MaxLength = Pub_MaxCEL11
    Text1(112).MaxLength = Pub_MaxCEL10
    Text1(113).MaxLength = Pub_MaxCEL11
    Text1(115).MaxLength = Pub_MaxCEL10
    Text1(116).MaxLength = Pub_MaxCEL11
    Text1(118).MaxLength = Pub_MaxCEL10
    Text1(119).MaxLength = Pub_MaxCEL11
    Text1(121).MaxLength = Pub_MaxCEL10
    Text1(122).MaxLength = Pub_MaxCEL11
    Text1(124).MaxLength = Pub_MaxCEL10
    Text1(125).MaxLength = Pub_MaxCEL11
    Text1(127).MaxLength = Pub_MaxCEL10
    Text1(128).MaxLength = Pub_MaxCEL11
    Text1(130).MaxLength = Pub_MaxCEL10
    Text1(131).MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   lblTot6.Caption = Empty 'Added by Lydia 2018/12/27
   
   ' 90.10.18 modify by louis (不去抓第一筆及最後一筆)
   strRsStart = Empty
   strRsEnd = Empty
   'strExc(0) = "SELECT MIN(TO_NUMBER(PA02)),MAX(TO_NUMBER(PA02)) FROM PATENT WHERE PA01=" & CNULL(m_SysKind)
   'intI = 0
   'Set rsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
   'If intI = 1 Then
   '   strRsStart = Format(rsTemp.Fields(0), "000000")
   '   strRsEnd = Format(rsTemp.Fields(1), "000000")
   '   RsAction 0
   'End If
   ActionEdit = 3
   CmdSitu True

'Remove by Morgn 2009/9/14 集中到 CmdSitu 處理
'   ' Ken 90.07.16 -- start
'   If m_bInsert Then
'       TBar1.Buttons(1).Enabled = True
'   Else
'       TBar1.Buttons(1).Enabled = False
'   End If
'   If m_bUpdate Then
'       TBar1.Buttons(2).Enabled = True
'   Else
'       TBar1.Buttons(2).Enabled = False
'   End If
'   If m_bDelete Then
'       TBar1.Buttons(3).Enabled = True
'   Else
'       TBar1.Buttons(3).Enabled = False
'   End If
'   ' Ken 90.07.16 -- End
   
   ' 90.10.18 modify by louis
   If Not IsEmptyText(m_CurrPA(0)) And Not IsEmptyText(m_CurrPA(1)) And Not IsEmptyText(m_CurrPA(2)) And Not IsEmptyText(m_CurrPA(3)) Then
      ActionEdit = 3
      CmdSitu True
   Else
      ActionEdit = 2 'Added by Lydia 2020/02/21
      CmdSitu False
      TxtLock 2
      'Modify by Morgan 2010/3/2
      'For i = 0 To 2
      '   Command1(i).Enabled = False
      'Next
      For Each cmd In Command1
         cmd.Enabled = False
      Next
      'end 2010/3/2
      
      'ActionEdit = 2 'Mark by Lydia 2020/02/21
   End If
   
   'Add By Sindy 2016/11/23
   '抓有輸入過匯率的請款幣別
   Combo4.Clear
   Combo4.AddItem ""
   Combo4.AddItem "USD"
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open "select distinct DNR01 from DebitNoteRate order by DNR01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While RsTemp.EOF = False
      Combo4.AddItem RsTemp.Fields("DNR01").Value
      RsTemp.MoveNext
   Loop
   RsTemp.Close
   '2016/11/23 End
   
   'Add By Cheng 2002/01/04
   Me.Text1(1).Text = m_SysKind
   If Len(Me.Text1(1).Text) > 0 Then SendKeys "{Tab}"
   
   'Added by Lydia 2020/02/21
   FraPA174.BackColor = &H8000000F
   FraPA174.Visible = False
   CmdPA174.Visible = False
   
   'Added by Lydia 2020/03/30 事務所合併日起取消(T:專利商標 J:智權公司 空白:系統預設)的標題改為(J:智權公司 空白:系統預設)
   If strSrvDate(1) >= 事務所合併日 Then
        Label1(173).Caption = "特殊出名公司：             ( J：智權公司 空白:系統預設)"
   End If
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
      Text1(91).Top = 435
      Text1(91).Height = 4515
   End If
   
   Frame1K.BorderStyle = 0 'Add By Sindy 2025/1/7
End Sub

'Modified by Lydia 2024/12/03 ChgCombo改成GetCombo1Data
Private Sub GetCombo1Data(ByVal strTmp As String)
   Combo1.Clear
   Combo1.AddItem ""
   If strTmp = "" Then Exit Sub
   'Modify By Cheng 2002/07/02
'   strExc(0) = "SELECT IN01||IN02 FROM INVENTOR WHERE IN01 IN (" & strTmp & ")"
   'Modify by Morgan 2010/12/16
   'strExc(0) = "SELECT " & IIf(strSysKind <> "P", " NVL(IN05,NVL(IN04,IN06))||'('||IN01||IN02||')' ", " NVL(IN04,NVL(IN05,IN06))||'('||IN01||IN02||')' ") & " FROM INVENTOR WHERE IN01 IN (" & strTmp & ")"
   'Modify By Sindy 2016/12/8 + decode(in03,null,'',' '||in03)||
   'Modified by Lydia 2024/12/03
   'strExc(0) = "SELECT " & IIf(Text1(1) <> "P", " NVL(IN05,NVL(IN04,IN06))||decode(in03,null,'',' '||in03)||' ('||IN01||IN02||')' ", " NVL(IN04,NVL(IN05,IN06))||decode(in03,null,'',' '||in03)||' ('||IN01||IN02||')' ") & " FROM INVENTOR WHERE IN01 IN (" & strTmp & ")"
   strExc(0) = "SELECT " & IIf(Text1(1) <> "P", " NVL(IN05,NVL(IN04,IN06))||decode(in03,null,'',' '||in03)||' ('||IN01||IN02||')' ", " NVL(IN04,NVL(IN05,IN06))||decode(in03,null,'',' '||in03)||' ('||IN01||IN02||')' ") & " AS PINAME " & _
               ", IN01, IN02, IN04, IN05, IN06 FROM INVENTOR WHERE IN01 IN (" & strTmp & ")"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         Do While Not .EOF
            Combo1.AddItem .Fields(0)
            'Added by Lydia 2024/12/03
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
            'end 2024/12/03
            .MoveNext
         Loop
      End If
   End With
End Sub

'end 2024/12/03
' 90.08.22 modify by louis
Private Sub SetComboData(ByVal strData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To Combo1.ListCount - 1
      'Modified by Lydia 2024/12/03 +And strData <> ""
      If InStr(Combo1.List(nPos), strData) > 0 And strData <> "" Then
         bFind = True
         Exit For
      End If
   Next nPos
   
   'Modified by Lydia 2024/12/03
'   If bFind Then
'      'Modify By Sindy 2014/11/6 Mark
''      Combo1.AddItem strData
''      Combo1.Refresh
''      Combo1.ListIndex = Combo1.ListCount - 1
''   Else
'      Combo1.ListIndex = nPos
'   End If
   If Not bFind Then
      Combo1.AddItem strData
      Combo1.Refresh
      Combo1.ListIndex = Combo1.ListCount - 1
      Frame3.Enabled = True
   Else
      Combo1.ListIndex = nPos
      Frame3.Enabled = False
   End If
   'end 2024/12/03

End Sub

'Add By Sindy 2014/11/5
'Mark by Lydia 2024/12/03
'Private Sub SetGrd()
'   Dim arrGridHeadText, arrGridHeadWidth
'   Dim iRow As Integer
'   '                        0    1             2           3           4           5
'   arrGridHeadText = Array("V", "發明人編號", "中文名稱", "英文名稱", "日文名稱", "國籍")
'   arrGridHeadWidth = Array(200, 1000, 2100, 2100, 2100, 1000)
'   GRD1.Visible = False
'   GRD1.Cols = UBound(arrGridHeadText) + 1
'   GRD1.Rows = 2
'   For iRow = 0 To GRD1.Cols - 1
'      GRD1.row = 0
'      GRD1.col = iRow
'      GRD1.Text = arrGridHeadText(iRow)
'      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
'      GRD1.CellAlignment = flexAlignCenterCenter
'   Next
'   GRD1.Visible = True
'End Sub
'end 2024/12/03

'Added by Lydia 2024/12/03
Private Sub SetGrd(tmpGrd As MSHFlexGrid)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
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
'end 2024/12/03

'Add By Sindy 2014/11/5
Private Sub Grd1_Click()
Dim nCol As Integer, nRow As Integer
Dim iCol As Integer
   
   With GRD1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   'Modified by Lydia 2024/12/03
   'If nRow > 0 And .TextMatrix(nRow, 1) <> "" Then
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
      .col = nCol
      pPrevRow = nRow
      Call SetComboData(.TextMatrix(nRow, 1))
      'Added by Lydia 2024/12/03
      If .TextMatrix(nRow, 1) = "" Then
         txtInvField(0) = .TextMatrix(nRow, 2)
         txtInvField(1) = .TextMatrix(nRow, 3)
         txtInvField(2) = .TextMatrix(nRow, 4)
         Lb_IN11N = .TextMatrix(nRow, 5)
         txtIN11 = .TextMatrix(nRow, 7)
         cmdUpdRow.Enabled = True
         cmdAddRow.Enabled = False
      End If
      'end 2024/12/03
   End If
   .Visible = True
   End With
   
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID()
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(pa(92)) = False Then
      If IsEmptyText(pa(92)) = False Then
         strCName = GetStaffName(pa(92), True)
      End If
   End If
   If IsNull(pa(93)) = False Then
      If IsEmptyText(pa(93)) = False Then
         strTemp = TAIWANDATE(pa(93))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(pa(94)) = False Then
      If IsEmptyText(pa(94)) = False Then
         strTemp = pa(94)
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(pa(95)) = False Then
      If IsEmptyText(pa(95)) = False Then
         strUName = GetStaffName(pa(95), True)
      End If
   End If
   If IsNull(pa(96)) = False Then
      If IsEmptyText(pa(96)) = False Then
         strTemp = TAIWANDATE(pa(96))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(pa(97)) = False Then
      If IsEmptyText(pa(97)) = False Then
         strTemp = pa(97)
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   'Modified by Morgan 2014/6/4 內容太長無法完全顯示,去掉欄位間的冒號
   textCUID = "CREATE: " & strCName & " " & _
              strCDate & " " & _
              strCTime & String(4, " ") & _
              "UPDATE: " & strUName & " " & _
              strUDate & " " & _
              strUTime
End Sub

Private Function ReadPatent(ByRef paTmp() As String) As Boolean
Dim i As Integer, j As Integer, Lbl As Object, txt As Object, strTmp As String
Dim strTxt(0 To 4) As String
'Add By Cheng 2002/12/10
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim arrID 'Add By Sindy 2025/1/7
   
   'Add By Sindy 2014/11/5
   cmdAddRow.Tag = ""
   cmdDelRow.Tag = ""
   GRD1.Clear
   'Modified by Lydia 2024/12/03
   'SetGrd
   ''2014/11/5 END
   Call SetGrd(GRD1)
   pPrevRow = 0
   For Each txt In txtInvField
      txt.Text = ""
   Next
   'end 2024/12/13
   
   For Each Lbl In Label2
      Lbl.Caption = ""
   Next
   For Each txt In Text2
      txt.Text = ""
   Next
   
   'Add by Morgan 2010/3/2
   For Each txt In Text3
      txt.Text = ""
   Next
   
   'Added by Lydia 2020/02/21
   ChkPA174.Value = vbUnchecked
   ChkPA174.Tag = ""
   bolAskPA174 = False
   'end 2020/02/21
   
   'Add By Sindy 2016/11/23
   Me.Combo4.ListIndex = 0
   Me.Combo5.ListIndex = 0
   '2016/11/23 End
   
   For i = 1 To 4
      strTxt(i) = paTmp(i)
   Next
   For i = 1 To 4
      pa(i) = strTxt(i)
   Next
   If pa(1) = "" Or pa(2) = "" Then Exit Function
   
   'Add By Sindy 2014/11/5 讀取發明人資料
   'Modified by Lydia 2024/12/03 + '' as 申請人1, IN11
   StrSQLa = "SELECT '' as V,pi06 as 發明人編號,in04 as 中文名稱,in05 as 英文名稱,in06 as 日文名稱,na03 國籍,'' as 申請人1,IN11 " & _
             " from PatentInventor,Inventor,nation where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
             " and substr(pi06,1,8)=in01(+) and substr(pi06,9,2)=in02(+)" & _
             " and in11=na01(+)" & _
             " order by pi05 asc"
   If rsA.State <> adStateClosed Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      Set GRD1.Recordset = rsA
   End If
   '2014/11/5 END
   
   'Modify by Morgan 2006/10/24
   'If Not objPublicData.ReadPatentDatabase(pA(), intWhere, False) Then Exit Function
   If Not PUB_ReadPatentDatabase(pa(), intWhere, False) Then Exit Function
   strTmp = ""
   For i = 26 To 30
      If pa(i) <> "" Then
         If Len(pa(i)) < 9 Then
            strTmp = strTmp & "'" & Left(pa(i), 8) & String(8 - Len(pa(i)), "0") & "',"
         'Added by Morgan 2019/12/5 修正更名前編號未帶出發明人問題
         Else
            strTmp = strTmp & "'" & Left(pa(i), 8) & "',"
         'end 2019/12/5
         End If
      End If
   Next
   If strTmp <> "" Then strTmp = Left(strTmp, Len(strTmp) - 1)
   'Modified by Lydia 2024/12/03 改成GetCombo1Data
   'ChgCombo strTmp
   GetCombo1Data strTmp
   
   'Modify by Morgan 2006/10/18 避免畫面物件未加欄位已新增時發生錯誤
   'For i = 1 To TF_PA 'edit by nickc 2006/07/12  T_PA
   'Modify by Morgan 2007/10/26 改有物件的才放,否則畫面有欄位沒放時會錯
   'For i = 1 To Me.Text1.Count - 1
   '   Text1(i) = pa(i)
   'Next
   For Each txt In Text1
      i = txt.Index
      If i > 0 Then
         Text1(i) = pa(i)
         Text1(i).Tag = pa(i) 'Added by Lydia 2019/11/27 記錄修改前
         
         'Added by Lydia 2015/12/15 X,Y編號縮短
         Select Case i
             Case 26, 27, 28, 29, 30, 75, 76, 86, 88, 101, 105, 133, 134
                 If Left(Text1(i), 1) = "X" Or Left(Text1(i), 1) = "Y" Then
                    Text1(i) = ChangeCustomerS(pa(i))
                 End If
         End Select
         'end 2015/12/15
      End If
   Next
   'end 2007/10/26
   'Add By Sindy 2025/1/7
   If pa(180) <> "" Then
      arrID = Split(pa(180), ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         Chk1K(Val(arrID(intI)) - 1).Value = 1
      Next intI
   End If
   '2025/1/7 END
   
   pPA12 = Text1(12).Text 'Added by Lydia 2015/09/09
   
   'Added by Lydia 2020/02/21 設定「名稱有特殊字」
   FraPA174.Visible = False
   CmdPA174.Visible = False
   If Text1(1) = "P" Or Text1(1) = "FCP" Then
       FraPA174.Visible = True
       CmdPA174.Visible = True
       If pa(174) = "Y" Then
           ChkPA174.Value = vbChecked
           ChkPA174.Tag = pa(174)
       End If
   End If
   'end 2020/02/21
   
    'Add By Cheng 2003/08/07
    '記錄原申請人
    For i = 26 To 30
        If Me.Text1(i).Text <> "" Then
            m_CustNo(i - 25) = Left(Me.Text1(i).Text & "000000000", 9)
        Else
            m_CustNo(i - 25) = Me.Text1(i).Text
        End If
    Next i
   'edit by Toni 2008/10/13
   If pa(150) = "" Then
      Combo2 = ""
   Else
      Combo2 = pa(150) + "." + PUB_GetFCPGrpName(pa(150))
   End If
   'end 2008/10/13
   Combo2.Tag = Combo2.Text 'Added by Lydia 2017/11/30 FCP案件命名電子化
   
   Call Text1_Validate(8, False) 'Add By Sindy 2014/7/8 案件屬性
   'Add By Sindy 2010/10/27
   If pa(158) = "" Then
      Combo3 = ""
   Else
      'Modify By Sindy 2014/7/8
      'Combo3 = pa(158) + "." + PUB_GetCaseAttributeName(pa(158))
      Combo3 = pa(158) + "." + PUB_GetCaseAttributeName(pa(158), pa(8))
      '2014/7/8 END
   End If
   '2010/10/27 End
    
   Text1(24) = TransDate(pa(24), 2)
   Text1(25) = TransDate(pa(25), 2)
   'Modify by Morgan 2006/10/18 避免畫面物件未加欄位已新增時發生錯誤
   'For i = 5 To TF_PA 'edit by nickc 2006/07/12 T_PA
   'Modify by Morgan 2007/10/26 改有物件的才放,否則畫面有欄位沒放時會錯
   'For i = 5 To Me.Text1.Count - 1
   '   Select Case i
   '      Case 26, 27, 28, 29, 30:
   '      Case Else
   '         If Text1(i) <> "" Then ChkKeyIn i
   '   End Select
   'Next
   For Each txt In Text1
      i = txt.Index
      If i > 5 And i <> 26 And i <> 27 And i <> 28 And i <> 29 And i <> 30 And Text1(i) <> "" Then
         ChkKeyIn i
      End If
   Next
   'end 2007/10/26
   
   ' 更新申請人的中文名稱
   UpdateCustomerName
   'edit by nickc 2007/02/02 不用 dll 了
   'If Not objPublicData.ReadPriority(pA, strPriority(1), strPriority(2), strPriority(3)) Then
   'Modify by Morgan 2007/4/24 加 strPriority(4)
   'Modify by Amy 2014/03/24 +strPriority(5)
   If Not ClsPDReadPriority(pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)) Then
      
   End If
   
   'Added by Lydia 2018/12/27 中文本-頁數總計
   '-----與各式申請書的總計不同,有加上序列表
   i = Val(Text1(64)) + Val(Text1(65)) + Val(Text1(66)) + Val(Text1(67)) + Val(Text1(68))
   If i = 0 Then
       lblTot6.Caption = Empty
   Else
       lblTot6.Caption = i
   End If
   'end 2018/12/27
   
'   If GetMoneyDate(Val(pa(8)), pa(9), strTxt, strExc(0), strExc(1)) Then varYear = Split(strExc(1), ",")
   InitFeeData
    'Add By Cheng 2002/12/10
    '下次繳費日內專抓本所期限, 外專抓法定期限(照舊)
    '92.9.10 MODIFY BY SONIA
    'If strSysKind = "P" Then
    '    strSQLA = "SELECT MAX(NP08) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
    '                " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NULL"
    '    rsA.CursorLocation = adUseClient
    '    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
    '    If rsA.RecordCount > 0 Then
    '        Label2(44).Caption = TransDate("" & rsA.Fields(0).Value, 1)
    '        Label2(47).Caption = Label2(44).Caption
    '    End If
    '    If rsA.State <> adStateClosed Then rsA.Close
    '    Set rsA = Nothing
    'Else
    '    If objLawDll.GetNextPayDate(pa, strExc(0)) Then
    '       If strExc(0) <> "" Then
    '          Label2(44).Caption = TransDate(strExc(0), 1)
    '       End If
    '       Label2(47).Caption = Label2(44).Caption
    '    End If
    'End If
    'Modify by Morgan 2004/2/17
    '不必控制為'A'類收文
'    If strSysKind = "P" Then
'        strSQLA = "SELECT MAX(" & SQLDate("NP08") & ") FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                    " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NULL AND NP01<'B'"
'    Else
'        strSQLA = "SELECT MAX(" & SQLDate("NP09") & ") FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                    " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NULL AND NP01<'B'"
'    End If

    'Modify by Morgan 2004/3/11
    '下次繳費日全部都顯示法定期限
'    If strSysKind = "P" Then
'        strSQLA = "SELECT MAX(" & SQLDate("NP08") & ") FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                    " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NULL"
'    Else
'        strSQLA = "SELECT MAX(" & SQLDate("NP09") & ") FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                    " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NULL"
'    End If
    '2011/1/27 MODIFY BY SONIA百年蟲
    'StrSQLa = "SELECT MAX(" & SQLDate("NP09") & "),max(np09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NULL"
    'Modified by Morgan 2022/6/13 可能同時會有1個以上的期限 Ex:CFP-023718 年費&延展費, 秀玲:有未過期的帶出最小期限,都過期則帶最大期限(不同時間點可能看到不同期限??)
    'StrSQLa = "SELECT " & SQLDate("NP09") & ",np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND (NP09||NP22) IN " & _
              "(SELECT MAX(NP09||NP22) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP09 IS NOT NULL AND NP06 IS NULL)"
    StrSQLa = "SELECT " & SQLDate("NP09") & ",np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP09 IS NOT NULL AND NP06 IS NULL" & _
                " order by sign(to_date(np09,'yyyymmdd')-sysdate)*np09 asc"
    'end 2022/6/13
    'Modify end 2004/3/11
    If rsA.State <> adStateClosed Then rsA.Close
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
      If rsA.Fields(0).Value <> "" Then
        Label2(44).Caption = CheckStr(rsA.Fields(0))
         'Add by Morgan 2004/8/16
         '若繳費期限有延期過加說明
         If PUB_IfCtrlDateExtended(pa, Format(rsA.Fields(1))) = True Then
           Label2(44).Caption = Label2(44).Caption & " (6個月逾繳期限)"
         End If
      End If
    Else
        'Modify by Morgan 2004/3/11
        '下次繳費日全部都顯示法定期限
'       If strSysKind = "P" Then
'           strSQLA = "SELECT MAX(" & SQLDate("NP08") & "||' '||DECODE(NP06,'Y',' (已收文)',' (不續辦)')) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                       " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NOT NULL"
'       Else
'           strSQLA = "SELECT MAX(" & SQLDate("NP09") & "||' '||DECODE(NP06,'Y',' (已收文)',' (不續辦)')) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                       " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NOT NULL"
'       End If
        '2011/1/27 MODIFY BY SONIA 百年蟲 CFP-008574會抓到96年資料
        'StrSQLa = "SELECT MAX(" & SQLDate("NP09") & "||' '||DECODE(NP06,'Y',' (已收文)',' (不續辦)')) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                    " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NOT NULL"
        StrSQLa = "SELECT " & SQLDate("NP09") & "||' '||DECODE(NP06,'Y',' (已收文)',' (不續辦)') FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP09||NP22 IN " & _
                  "(SELECT MAX(NP09||NP22) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP09 IS NOT NULL AND NP06 IS NOT NULL)"
       'Modify end 2004/3/11
       
       If rsA.State <> adStateClosed Then rsA.Close
       rsA.CursorLocation = adUseClient
       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
       If rsA.RecordCount > 0 Then
         If rsA.Fields(0).Value <> "" Then
           Label2(44).Caption = CheckStr(rsA.Fields(0))
            'Add by Morgan 2004/2/9
            '檢查若為已收文狀態，若案件進度檔中案件性質為605~607者若都有發文日時則下次繳費日改空白
            If InStr(1, Label2(44).Caption, "已收文") > 0 Then
                 StrSQLa = "SELECT 1 FROM CASEPROGRESS WHERE CP10 BETWEEN '605' AND '607' AND CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP27 IS NULL AND CP57 IS NULL"
                 If rsA.State <> adStateClosed Then rsA.Close
                 rsA.CursorLocation = adUseClient
                 rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                 If rsA.RecordCount = 0 Then
                     Label2(44).Caption = ""
                 End If
            End If
         End If
       Else
           Label2(44) = ""
       End If
    End If
    
   'Add by Morgan 2004/8/16 若無公告日加說明並變色
   If Label2(44).Caption <> "" And pa(14) = "" And Text1(9).Text = "000" Then
      Label2(44).ForeColor = &HC000&
      Label2(50).Caption = "尚未公告，下次繳費日為預估值"
   Else
      Label2(44).ForeColor = &H80000012
   End If
   Label2(47).ForeColor = Label2(44).ForeColor
   Label2(51).Caption = Label2(50).Caption
   'Add end
   
    Label2(47).Caption = Label2(44).Caption
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '92.9.10 END

'On Error Resume Next
'   For i = 60 To 69
'      If pa(i) <> "" Then
'         ' 90.08.22 modify by louis
'         'Combo1(i - 60).Text = pa(i)
'         'If Err.number = 383 Then
'         '   Combo1(i - 60).AddItem pa(i)
'         '   Combo1(i - 60).ListIndex = Combo1(i - 60).ListCount - 1
'         'End If
'         'Modify By Cheng 2002/07/02
''         SetComboData i - 60, GetInventorName(pa(i))
'         SetComboData i - 60, GetInventorName(pa(i))
'      Else
'         Combo1(i - 60).ListIndex = 0
'      End If
'   Next
   
   'Add By Sindy 2021/12/7
   ' 更新CUID
   UpdateCUID
   '2021/12/7 END
   
   'edit by nickc 2007/02/02 不用 dll 了
   'If pA(92) <> "" Then If objPublicData.GetStaffN(pA(92), strTmp) Then Text1(92).Text = strTmp
   'If pA(95) <> "" Then If objPublicData.GetStaffN(pA(95), strTmp) Then Text1(95).Text = strTmp
'   If pa(92) <> "" Then If ClsPDGetStaffN(pa(92), strTmp) Then Text1(92).Text = strTmp
'   If pa(95) <> "" Then If ClsPDGetStaffN(pa(95), strTmp) Then Text1(95).Text = strTmp
'   'Modified by Lydia 2020/04/14
'   'Text1(93).Text = Left(pa(93), 4) - 1911 & "/" & Mid(pa(93), 5, 2) & "/" & Mid(pa(93), 7)
'   'Text1(96).Text = Left(pa(96), 4) - 1911 & "/" & Mid(pa(96), 5, 2) & "/" & Mid(pa(96), 7)
'   Text1(93).Text = ChangeWStringToTDateString(pa(93))
'   Text1(96).Text = ChangeWStringToTDateString(pa(96))
'   'end 2020/04/14
'   Text1(94).Text = Format(pa(94), "##:##")
'   Text1(97).Text = Format(pa(97), "##:##")
   'add by nickc 2006/07/12
   'edit by nickc 2007/02/02 不用 dll 了
   'If pA(137) <> "" Then If objPublicData.GetStaffN(pA(137), strTmp) Then Text1(137).Text = strTmp
   If pa(137) <> "" Then If ClsPDGetStaffN(pa(137), strTmp) Then Text1(137).Text = strTmp
   'Modified by Lydia 2020/04/14
   'Text1(108).Text = Left(pa(108), 4) - 1911 & "/" & Mid(pa(108), 5, 2) & "/" & Mid(pa(108), 7)
   'Text1(136).Text = Left(pa(136), 4) - 1911 & "/" & Mid(pa(136), 5, 2) & "/" & Mid(pa(136), 7)
   Text1(108).Text = ChangeWStringToTDateString(pa(108))
   Text1(136).Text = ChangeWStringToTDateString(pa(136))
   'end 2020/04/14
   PUB_AddContact pa(26), cboContact, pa(149), , True 'Add by Morgan 2008/8/4
   
   'Add by Morgan 2009/12/24 PCT案,香港標準專利案,分割案要帶出提交日
   lblFilingDate(0).Visible = False
   lblFilingDate(1).Visible = False
   If (pa(9) = "013" And pa(8) = "1") Or (pa(46) = "Y") Then
      lblFilingDate(0).Visible = True
      lblFilingDate(1).Visible = True
      lblFilingDate(1) = ""
      strExc(0) = "select cp47 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in (" & NewCasePtyList & ") and cp57 is null and cp27>0 and cp47>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         lblFilingDate(1) = TransDate(RsTemp(0), 1)
      End If
   Else
      If pa(9) = "000" Then
         strExc(0) = "select cp27 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='307' and cp57 is null"
      Else
         strExc(0) = "select cp47 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='307' and cp57 is null"
         'Added by Morgan 2019/11/28  +CFP接續案
         If pa(1) = "CFP" And pa(3) = "0" Then
            strExc(0) = strExc(0) & " union select cp47 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('122','113') and cp57 is null"
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
   
   'Added by Morgan 2012/12/26
   If Text1(1) = "FCP" And Text1(162) = "Y" Then
      cmdDivSug.Visible = True
   Else
      cmdDivSug.Visible = False
   End If
   
   'Added by Morgan 2015/6/11
   If pa(166) <> "" Then
      strExc(1) = ""
      If Len(pa(166)) = 9 Then
         strExc(1) = pa(166) & " " & GetName(pa(166))
      Else
         strExc(0) = "select fa01||fa02||' '||nvl(fa06, nvl(rtrim(fa05||' '||fa63||' '||fa64||' '||fa65),fa04)),instr('" & pa(166) & "',fa01||fa02) srt from fagent where instr('" & pa(166) & "',fa01||fa02)>0" & _
            " union all select cu01||cu02||' '||nvl(cu06, nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu04)),instr('" & pa(166) & "',cu01||cu02) srt from customer where instr('" & pa(166) & "',cu01||cu02)>0" & _
            " order by srt"
     
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = Replace("" & RsTemp(0), vbCrLf, "")
            RsTemp.MoveNext
            Do While Not RsTemp.EOF
               strExc(1) = strExc(1) & vbCrLf & Replace("" & RsTemp(0), vbCrLf, "")
               RsTemp.MoveNext
            Loop
         End If
      End If
      SetList lstPA166, strExc(1)
   Else
      lstPA166.Clear
   End If
   'end 2015/6/11
   
    'Added by Lydia 2015/11/03　顯示一案兩請，擬制喪失新穎性案件
    lblCaseMap.Caption = ""
    lblCaseMap2.Caption = "" 'Added by Lydia 2019/11/28
    If PUB_GetRefCaseChk(pa(1), pa(2), pa(3), pa(4), "CASEMAP", "3") = True Then
       lblCaseMap.Caption = "一案兩請"
    End If
    If PUB_GetRefCaseChk(pa(1), pa(2), pa(3), pa(4), "CASEMAP", "6") = True Then
       'Modified by Lydia 2019/11/28 P-123733有一案兩請和擬制喪失新穎性案件
       'lblCaseMap.Caption = "擬制喪失新穎性案件"
       lblCaseMap2.Caption = "擬制喪失新穎性案件"
    End If
    Command4.BackColor = &H8000000F
    If PUB_GetRefCaseChk(pa(1), pa(2), pa(3), pa(4), "DIVISIONCASE") = True Then
       Command4.BackColor = &H8080FF
    End If
    'end 2015/11/03
    
    'Added by Lydia 2016/06/14 +台灣大陸案件提示
    lblCMboth.Caption = ""
    If (pa(1) = "P" Or pa(1) = "FCP") And pa(9) = 台灣國家代號 Then
       If PUB_GetRefCaseChk(pa(1), pa(2), pa(3), pa(4), "CASEMAP", "0", "A", 大陸國家代號) Then
          lblCMboth.Caption = "有大陸案"
       End If
    ElseIf pa(1) = "P" And pa(9) = 大陸國家代號 Then
       If PUB_GetRefCaseChk(pa(1), pa(2), pa(3), pa(4), "CASEMAP", "0", "A", 台灣國家代號) Then
          lblCMboth.Caption = "有台灣案"
       End If
    End If
    'end 2016/06/14
    
   'Add By Sindy 2016/11/23
   If pa(170) <> "" Then
      For i = 0 To Combo4.ListCount - 1
         Combo4.ListIndex = i
         If InStr(Combo4.Text, pa(170)) > 0 Then
            Exit For
         End If
      Next
   Else
      Combo4.ListIndex = 0
   End If
   If pa(171) <> "" Then
      Combo5.ListIndex = pa(171)
   Else
      Combo5.ListIndex = 0
   End If
   '2016/11/23 End
      
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
   
   'Remove by Lydia 2019/11/27 已預設
   'Text1(161).Tag = pa(161)
   'Text1(11).Tag = Text1(11) 'Added by Morgan 2019/2/19
   
   'Added by Morgan 2023/3/30
   If pa(1) = "CFP" Then
      PUB_SetEntityOpt pa(1), Text1(9), Text1(8), Combo6, pa(179)
   Else
      Combo6.Clear
   End If
   'end 2023/3/30
   
   'Added by Morgan 2024/10/15
   '大陸發明案的專用期止日增加判斷是否需帶補償天數及補償年費期限
   Label1(22) = Label1(22).Tag
   If pa(1) = "P" And pa(9) = "020" And pa(8) = "1" And pa(25) <> "" Then
      If Label2(44) = "" Then Label2(44) = PUB_GetCN615DueDate(pa())
      If PUB_GetCNExtDays(pa(), pa(25), i) = True Then
         If i > 0 Then
            Label1(22) = Label1(22).Tag & " (含補償 " & i & " 天)"
         End If
      End If
   End If
   'end 2024/10/15
   
End Function

Private Sub Form_Unload(Cancel As Integer)

   'Added by Lydia 2024/12/03 發明人輸入比對兼自動代入(模糊比對)
   ' 刪除串列結構
   If m_InventorListCount > 0 Then
      Erase m_InventorList
   End If
   m_InventorListCount = 0
   'end 2024/12/03
   
   'Added By Morgan 2012/6/18
   If m_form Is Nothing = False Then
      m_form.Enabled = True
      If m_form.Name = "frm040101_1" Or m_form.Name = "frm060101_1" Then
         'Modify by Amy 2022/12/26 +if 及intOpen090801(接洽單開啟狀態)
         If m_form.Name = "frm040101_1" Then
            '從專利分案開啟接洽單和此支並存時,關閉接洽單,回分案畫面時不需再重開
            If PUB_CheckFormExist("frm090801_Q") = True Then
                intOpen090801 = 1 'Modify by Amy 2023/03/31 已開啟為1
            End If
            m_form.QueryMainFile intOpen090801
         Else
            m_form.QueryMainFile
         End If
      'Add by Amy 2022/12/26
      ElseIf m_form.Name = "frm050101_2" Then
        If PUB_CheckFormExist("frm090801_Q") = True Then
            intOpen090801 = 1
        End If
      End If
      Set m_form = Nothing
   End If
   'end 2012/6/18
   intOpen090801 = Empty 'Add by Amy 2022/12/26
   Set frm050701 = Nothing
End Sub

Private Sub RsSitu(ByVal Situ As Integer)
 Dim i As Integer, St1 As String, St2 As String
 Dim TBmk As Variant
 Dim txt As Object
 
 '911106 nick
 'On Error GoTo CheckingErr
 
 cmdDivSug.Visible = False 'Added by Morgan 2012/12/26
 'Add By Sindy 2014/11/6
 cmdAddRow.Enabled = False
 cmdDelRow.Enabled = False
 '2014/11/6 END
 'Added by Lydia 2024/12/03
 cmdUp.Enabled = False
 cmdDown.Enabled = False
 Frame3.Enabled = False
 'end 2024/12/03
         
 'Added by Morgan 2015/6/11
 If Situ = 0 Or Situ = 1 Then
   If Frame2.Visible = False Then
      lstPA166.Width = lstPA166.Width - Frame2.Width
      txtNo = ""
      lblName = ""
      Frame2.Visible = True
   End If
 Else
   If Frame2.Visible = True Then
      lstPA166.Width = lstPA166.Width + Frame2.Width
      Frame2.Visible = False
   End If
 End If
 'end 2015/6/11
 
 'Added by Lydia 2015/11/03
 If Situ = 0 Or Situ = 5 Then
    lblCaseMap.Caption = ""
    lblCaseMap2.Caption = "" 'Added by Lydia 2019/11/28
    Command4.BackColor = &H8000000F
    lblCMboth.Caption = "" 'Added by Lydia 2016/06/14
 End If
 'end 2015/11/03
 
 Static TmpPatent(1 To 4) As String
   Select Case Situ
      
      Case 0 'add
         For i = 1 To 4
            TmpPatent(i) = Text1(i).Text
         Next
         Label2(44).Caption = ""
         Label2(47).Caption = ""
         ActionEdit = 0 'Added by Lydia 2020/02/21
         CmdSitu False
         TxtLock 2
         'ActionEdit = 0 'Mark by Lydia 2020/02/21
         GridHead
         Text1(8).Enabled = True
         Text1(23).Enabled = True
         Text1(1).SetFocus
         'Add By Sindy 2014/11/6
         cmdAddRow.Enabled = True
         cmdDelRow.Enabled = True
         '2014/11/6 END
         'Added by Lydia 2024/12/03
         cmdUp.Enabled = True
         cmdDown.Enabled = True
         Frame3.Enabled = True
         'end 2024/12/03
      Case 1 'modi
         ' 90.10.18 modify by louis
         Command3.Enabled = True
         ActionEdit = 1 'Added by Lydia 2020/02/21
         CmdSitu False
         'ActionEdit = 1 'Mark by Lydia 2020/02/21
         For i = 1 To 4
            Text1(i).Locked = True
         Next
         
         'Added by Morgan 2022/11/30
         '若FCP案第1道收文程序為年費時可修改專利種類-- 陳亭妙
         '中間來所案件會一律先收發明，因年費需先補資料才可分案，故此處須開放以免申請號會檢查不過無法輸入
         If IsFCP605Case() = True Then
            Text1(8).Enabled = True
         Else
            Text1(8).Enabled = False
         End If
         'end 2022/11/30
         
         Text1(23).Enabled = False
         For i = 1 To 4
            TmpPatent(i) = Text1(i).Text
         Next
         'Add By Sindy 2014/11/6
         cmdAddRow.Enabled = True
         cmdDelRow.Enabled = True
         '2014/11/6 END
         'Added by Lydia 2024/12/03
         cmdUp.Enabled = True
         cmdDown.Enabled = True
         Frame3.Enabled = True
         'end 2024/12/03
         FraPA174.Enabled = True 'Added by Lydia 2020/02/21
      Case 2 'delete
         If DelMsg Then
            '911106 nick transation
            cnnConnection.BeginTrans
            'If ChkCaseCode("CP", pa(1), pa(2), pa(3), pa(4)) = False Then Exit Sub
            If ChkCaseCode("CP", pa(1), pa(2), pa(3), pa(4)) = False Then cnnConnection.RollbackTrans: Exit Sub
            'Add By Sindy 2010/7/1
            If ChkCaseCode("NP", pa(1), pa(2), pa(3), pa(4)) = False Then cnnConnection.RollbackTrans: Exit Sub
            
            Select Case OnDataDeleteRecord(0, pa(1) & pa(2) & pa(3) & pa(4))
               Case 0
                  PUB_DelPatentRefData pa(1), pa(2), pa(3), pa(4), Me 'Added by Morgan 2025/6/25 以共用函數刪除相關資料
                  
                  strExc(1) = "DELETE FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                  '911106 nick transation
                  'add by nickc 2006/06/07 紀錄分析語法
                  'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
                  'Pub_SeekTbLog strExc(1)
                  Pub_SeekTbLog strExc(1), , , , Me.Caption & "(" & Me.Name & ")"
                  cnnConnection.Execute strExc(1)
                  
'Removed by Morgan 2025/6/25 改在前面以公用函數刪除
'                  strExc(2) = "DELETE FROM PRIDATE WHERE " & ChgPriDate(pa(1) & pa(2) & pa(3) & pa(4))
'                  'add by nickc 2006/06/07 紀錄分析語法
'                  'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
'                  'Pub_SeekTbLog strExc(2)
'                  Pub_SeekTbLog strExc(2), , , , Me.Caption & "(" & Me.Name & ")"
'                  cnnConnection.Execute strExc(2)
'
'                  'Added by Lydia 2016/11/24 一併刪除各項指示
'                  strSql = "DELETE FROM INSTRUCTIONS WHERE ITS01=" & CNULL(Pub_GetITS01Type(pa(1))) & " AND ITS02=" & CNULL(pa(1) & pa(2) & pa(3) & pa(4))
'                  'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
'                  'Pub_SeekTbLog strSql
'                  Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
'                  cnnConnection.Execute strSql
'                  'end 2016/11/24
'
'                  'Add By Sindy 2014/11/5 增加刪除發明人資料(逐筆刪除)
'                  'Add By Sindy 2014/11/6 檢查專利發明人檔是否有資料,若有,逐筆記錄刪除
'                  strExc(0) = "SELECT * FROM PATENTInventor WHERE Pi01='" & pa(1) & "' AND Pi02='" & pa(2) & "' AND Pi03='" & pa(3) & "' AND Pi04='" & pa(4) & "' order by pi05 asc"
'                  intI = 1
'                  'edit by nickc 2007/02/05 不用 dll 了
'                  'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
'                  If intI = 1 Then
'                     RsTemp.MoveFirst
'                     Do While Not RsTemp.EOF
'                        strSql = "delete from patentInventor where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) + " and pi05=" & RsTemp.Fields("pi05")
'                        'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
'                        'Pub_SeekTbLog strSql
'                        Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
'                        cnnConnection.Execute strSql
'                        RsTemp.MoveNext
'                     Loop
'                  End If
'                  '2014/11/6 END
'                  '2014/11/5 END
'end 2025/6/25
                  
                  'If Not objLawDll.ExecSQL(2, strExc) Then Exit Sub
               
                  strExc(0) = "SELECT PA03,PA04 FROM PATENT WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03||PA04>'" & pa(3) & pa(4) & "'"
                  intI = 1
                  'edit by nickc 2007/02/05 不用 dll 了
                  'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
                  If intI = 1 Then
                     strExc(1) = pa(1)
                     strExc(2) = pa(2)
                     'Modify by Morgan 2008/4/18
                     'strExc(3) = pa(3)
                     'strExc(4) = pa(4)
                     strExc(3) = RsTemp.Fields("PA03")
                     strExc(4) = RsTemp.Fields("PA04")
                     ReadPatent strExc
                  Else
                     'Modify By Sindy 2012/3/2
                     strExc(0) = "SELECT count(*) FROM PATENT WHERE PA01='" & pa(1) & _
                        "' AND PA02>'" & pa(2) & "'"
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If RsTemp.Fields(0) > 0 Then
                           strExc(0) = "SELECT MIN(PA01||PA02||PA03||PA04) FROM PATENT WHERE PA01='" & pa(1) & _
                                       "' AND PA02>'" & pa(2) & "'"
                        Else
                           strExc(0) = "SELECT MAX(PA01||PA02||PA03||PA04) FROM PATENT WHERE PA01='" & pa(1) & _
                                       "' AND PA02<'" & pa(2) & "'"
                        End If
                     Else
                        strExc(0) = "SELECT MAX(PA01||PA02||PA03||PA04) FROM PATENT WHERE PA01='" & pa(1) & _
                                    "' AND PA02<'" & pa(2) & "'"
                     End If
                     '2012/3/2 End
'                     strExc(0) = "SELECT MIN(PA01||PA02||PA03||PA04) FROM PATENT WHERE PA01='" & pa(1) & _
'                        "' AND PA02>'" & pa(2) & "'"
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        ChgCaseNo RsTemp.Fields(0), strExc
                        ReadPatent strExc
                     Else
                        RsAction 0
                     End If
                  End If
               Case -3
                  
               Case Else
                  MsgBox "新增資料至案件刪除記錄檔失敗，請洽系統管理員 !", vbCritical
                  
            End Select
            '911108 nick transation
            cnnConnection.CommitTrans
         End If
      Case 3 'update
         If ActionEdit = 0 Then
            'Add By Cheng 2002/01/11
            Text1_Change 27
            Text1_Change 28
            Text1_Change 29
            Text1_Change 30
            'If Not GetData Then Exit Sub 'Removed by Morgan 2012/7/17 移到下面否則欄位若未跳離會沒有更新到
'            For i = 1 To 132
'               pa(i) = ChgSQL(pa(i))
'            Next i
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            If Not GetData Then Exit Sub 'Added by Morgan 2012/7/17
            
            'edit by nickc 2006/06/07 搬出 dll
            'If objPublicData.SaveNewPatentDatabase(pa, intWhere, False) Then
            If SaveNewPatentDatabase(pa, intWhere, False) Then
               'Modify by Morgan 2007/4/24 改共用函式
               'If Not SavePriority(pa, strPriority(1), strPriority(2), strPriority(3)) Then
               'Modify by Amy 2014/03/24 +strPriority(5)
               If Not ClsPDSavePriority(pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)) Then
                  Exit Sub
               End If
            Else
               Exit Sub
            End If
            Text1(8).Enabled = True
            Text1(23).Enabled = True
         ElseIf ActionEdit = 1 Then
            'Add By Cheng 2002/01/11
            Text1_Change 27
            Text1_Change 28
            Text1_Change 29
            Text1_Change 30
            'If Not GetData Then Exit Sub 'Removed by Morgan 2012/7/17 移到下面否則欄位若未跳離會沒有更新到
'            For i = 1 To 132
'               pa(i) = ChgSQL(pa(i))
'            Next i
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'edit by nickc 2006/06/07  搬出 dll
            'If objPublicData.SavePatentDatabase(pa, intWhere, False, True) Then
            
            'Added by Lydia 2017/06/19 (存檔前)檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
            strChkCuAreaMail = PUB_ChkSameCustSales(Trim(Text1(1)), Trim(Text1(2)), Trim(Text1(3)), Trim(Text1(4)), "", Trim(Text1(26)), Trim(Text1(27)), Trim(Text1(28)), Trim(Text1(29)), Trim(Text1(30)), strChkCuAreaMailTo)
            
            If Not GetData Then Exit Sub 'Added by Morgan 2012/7/17
            
            If SavePatentDatabase(pa, intWhere, False, True) Then
               'Modify by Morgan 2007/4/24 改共用函式
               'If Not SavePriority(pa, strPriority(1), strPriority(2), strPriority(3)) Then
               'Modify by Amy 2014/03/24 +strPriority(5)
               If Not ClsPDSavePriority(pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)) Then
                  Exit Sub
               End If
               '92.05.22 nick 重新抓資料
                For i = 1 To 4
                    strExc(i) = TmpPatent(i)
                 Next
               ReadPatent strExc
               'add by nickc 2005/08/23 紀錄修改案號
               pub_ModifyCaseNum = Text1(1) & "-" & Text1(2) & "-" & Text1(3) & "-" & Text1(4)
            Else
               Exit Sub
            End If
            
            'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
            If strChkCuAreaMail <> "" Then
               PUB_SendMail strUserNum, strChkCuAreaMailTo, "", "案件收文通知--此案收文非原智權人員(區)！", strChkCuAreaMail
            End If
            'end 2017/06/19
            
            Text1(8).Enabled = True
            Text1(23).Enabled = True
         ElseIf ActionEdit = 2 Then '在查詢狀態按下Enter鍵
            If Text1(1) = "" Or Text1(2) = "" Then
               MsgBox "本所案號不可空白，請重新輸入 !", vbCritical
               Exit Sub
            End If
            If Text1(3) = "" Then Text1(3) = "0"
            If Text1(4) = "" Then Text1(4) = "00"
            m_bolFmpAuth = False 'Added by Morgan 2023/5/12
            'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
            If FMP2open = True Then
              'Modified by Morgan 2023/5/12 FMP非寰華案需可維護部分欄位
              'If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1(1), Text1(2), Text1(3), Text1(4)) = False Then Exit Sub
              If PUB_FMPtoCheck(1, 1, Pub_strUserST05, Text1(1), Text1(2), Text1(3), Text1(4)) = False Then
                  m_bolFmpAuth = True
              End If
              'end 2023/5/12
            End If
            intI = 1
            'strExc(0) = "SELECT COUNT(*) FROM PATENT WHERE PA01='" & Text1(1) & "' AND PA02='" & Text1(2) & "'  AND PA03='" & Text1(3) & "' AND PA04='" & Text1(4) & "'"
            strExc(0) = "SELECT COUNT(*),MIN(cp09) FROM PATENT,CASEPROGRESS WHERE PA01='" & Text1(1) & "' AND PA02='" & Text1(2) & "'  AND PA03='" & Text1(3) & "' AND PA04='" & Text1(4) & "' and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp12(+) LIKE 'F%' and rownum<2"
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) = 0 Then
                  MsgBox "查無此案號之記錄 !", vbCritical
                  Exit Sub
               'Added by Morgan 2012/5/17
               'FMP案權限檢查
               ElseIf m_bolFmpAuth And IsNull(RsTemp.Fields(1)) Then
                  MsgBox "無權限，此案號非FMP案 !", vbCritical
                  Exit Sub
               'end 2012/5/17
               Else
                  For i = 1 To 4
                     strExc(i) = Text1(i)
                  Next
                  ReadPatent strExc
               End If
            End If
         End If
         ActionEdit = 3 'Added by Lydia 2020/02/21
         CmdSitu True
         'ActionEdit = 3 'Mark by Lydia 2020/02/21
         ' 90.10.18 modify by louis
         Command3.Enabled = False
      Case 4 'cancel
         If ActionEdit <> 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
         End If
         CmdSitu True
         For i = 1 To 4
            strExc(i) = TmpPatent(i)
         Next
         ActionEdit = 3
         
         'Added by Morgan 2012/5/29
         '前次無查詢案件時要清除案號以免下次取消會帶出本次無權限的案件資料
         If TmpPatent(1) = "" Then
            For i = 2 To 4
               Text1(i).Text = ""
            Next
         Else
         'end 2012/5/29
         
            ReadPatent strExc
         End If 'Added by Morgan 2012/5/29
         
         Text1(8).Enabled = False
         Text1(23).Enabled = False
         ' 90.10.18 modify by louis
         Command3.Enabled = False
      Case 5 'query
         For i = 1 To 4
            TmpPatent(i) = Text1(i).Text
         Next
         ActionEdit = 2 'Added by Lydia 2020/02/21
         CmdSitu False
         TxtLock 2
         'Modify by Morgan 2006/10/18 避免畫面物件未加欄位已新增時發生錯誤
         'For i = 5 To TF_PA 'edit by nickc 2006/07/12 T_PA
         'Modify by Morgan 2007/10/26 改有物件的才放,否則畫面有欄位沒放時會錯
         'For i = 5 To Me.Text1.Count - 1
         '   Text1(i).Locked = True
         'Next
         For Each txt In Text1
            i = txt.Index
            If i >= 5 Then
               Text1(i).Locked = True
            End If
         Next
         'end 2007/10/26
         
         'Modify by Morgan 2010/3/2
         'For i = 0 To 2
         '   Command1(i).Enabled = False
         'Next
         For Each cmd In Command1
            cmd.Enabled = False
         Next
         'end 2010/3/2
      
         'ActionEdit = 2 'Mark by Lydia 2020/02/21
         'ADD BY TONI 2008/10/13
         Combo2 = ""
         'END 2008/10/13
         Combo2.Tag = "" 'Added by Lydia 2017/11/30
         
         'Add By Sindy 2010/10/27
         Combo3 = ""
         '2010/10/27 End
         
         '2010/5/19 MODIFY BY SONIA 系統類別不清空,FOCUS停在流水號
         'Text1(1).SetFocus
         Text1(1) = TmpPatent(1)
         Text1(2).SetFocus
         '2010/5/19 END
   End Select
   
   Exit Sub
CheckingErr:
    cnnConnection.RollbackTrans
    MsgBox Err.Description
    
End Sub

Private Sub RsAction(ByVal Sty As Integer)
Dim i As Integer
'On Error GoTo ErrHand
'   TxtLock 2
'   If rsMain.EOF And rsMain.BOF Then MsgBox "資料庫內無資料 !", vbInformation: Exit Sub
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case Sty
      Case 0
         strExc(0) = "SELECT PA01,PA02,PA03,PA04 FROM PATENT WHERE PA01=" & CNULL(m_SysKind) & " AND PA02='" & strRsStart & "'"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For i = 1 To 4
               strExc(i) = RsTemp.Fields(i - 1).Value
            Next
         End If
      Case 1
         If Text1(2).Text = strRsStart Then
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 6
            Exit Sub
         Else
            If Text1(2) = "" Then Text1(2) = "000000"
            If Text1(3) = "" Then Text1(3) = "0"
            If Text1(4) = "" Then Text1(4) = "00"
            intI = 1
            strExc(0) = "SELECT MAX(PA03||PA04) FROM PATENT WHERE PA01=" & CNULL(m_SysKind) & " AND PA02='" & Text1(2) & "' AND PA03||PA04<'" & Text1(3) & Text1(4) & "'"
            'edit by nickc 2007/02/05 不用 dll 了
            'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 And Not IsNull(RsTemp.Fields(0)) Then
               strExc(1) = m_SysKind
               strExc(2) = Text1(2)
               strExc(3) = Left(RsTemp.Fields(0), 1)
               strExc(4) = Right(RsTemp.Fields(0), 2)
            Else
               strExc(0) = "SELECT MAX(PA01||PA02||PA03||PA04) FROM PATENT WHERE PA01=" & CNULL(m_SysKind) & _
                  " AND PA02<'" & Text1(2) & "'"
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 Then ChgCaseNo RsTemp.Fields(0), strExc
               strExc(0) = "SELECT MAX(PA03||PA04) FROM PATENT WHERE PA01=" & CNULL(m_SysKind) & " AND PA02='" & strExc(2) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(3) = Left(RsTemp.Fields(0), 1)
                  strExc(4) = Right(RsTemp.Fields(0), 2)
               End If
            End If
         End If
      Case 2
         If Text1(2).Text = strRsEnd Then
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 7
            Exit Sub
         Else
            If Text1(2) = "" Then Text1(2) = "000000"
            If Text1(3) = "" Then Text1(3) = "0"
            If Text1(4) = "" Then Text1(4) = "00"
            strExc(0) = "SELECT PA03,PA04 FROM PATENT WHERE PA01=" & CNULL(m_SysKind) & " AND PA02='" & Text1(2) & "' AND PA03||PA04>'" & Text1(3).Text & Text1(4).Text & "'"
            intI = 1
            'edit by nickc 2007/02/05 不用 dll 了
            'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = Text1(1)
               strExc(2) = Text1(2)
               strExc(3) = RsTemp.Fields(0)
               strExc(4) = RsTemp.Fields(1)
            Else
               strExc(0) = "SELECT MIN(PA01||PA02||PA03||PA04) FROM PATENT WHERE PA01=" & CNULL(m_SysKind) & _
                  " AND PA02>'" & Text1(2) & "'"
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 Then ChgCaseNo RsTemp.Fields(0), strExc
            End If
         End If
      Case 3
         strExc(0) = "SELECT PA01,PA02,PA03,PA04 FROM PATENT WHERE PA01=" & CNULL(m_SysKind) & " AND PA02='" & strRsEnd & "'"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For i = 1 To 4
               strExc(i) = RsTemp.Fields(i - 1).Value
            Next
         End If
   End Select
   ReadPatent strExc
   Screen.MousePointer = vbDefault
   Exit Sub
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Sub CmdSitu(ByVal TF As Boolean)
 Dim i As Integer, txt As Object
   If TF = True Then
      TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         ' 90.10.18 modify by louis
         'TBar1.Buttons(i + 5).Enabled = True
         If Not IsEmptyText(strRsStart) And Not IsEmptyText(strRsEnd) Then
            TBar1.Buttons(i + 5).Enabled = True
         Else
            TBar1.Buttons(i + 5).Enabled = False
         End If
      Next
      
      'Add by Morgan 2009/9/14
      If Not m_bInsert Then
          TBar1.Buttons(1).Enabled = False
      End If
      If Not m_bUpdate Then
          TBar1.Buttons(2).Enabled = False
      End If
      If Not m_bDelete Then
          TBar1.Buttons(3).Enabled = False
      End If
      'end 2009/9/14
      
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
      TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub TxtLock(ByVal Lt As Integer)
 Dim txt As Object, i As Integer, cbo As Object
 Dim Lbl As Object
 
   Select Case Lt
      Case 0
         For Each txt In Text1
            txt.Locked = True
         Next
         For Each txt In Text3
            txt.Locked = True
         Next
'         For Each cbo In Combo1
            Combo1.Locked = True
'         Next
         For Each cmd In Command1
            cmd.Enabled = False
         Next
         Combo3.Enabled = False 'Add By Sindy 2010/10/27
         'Add By Sindy 2016/11/23
         Combo4.Locked = True
         Combo5.Locked = True
         '2016/11/23 End
         'Added by Lydia 2020/02/21 開啟「名稱有特殊字」
         If ActionEdit = 1 Then '修改
           FraPA174.Enabled = True
         Else
           FraPA174.Enabled = False
         End If
         'end 2020/02/21
         Combo6.Locked = True 'Added by Morgan 2023/3/30
         Frame1K.Enabled = False 'Add By Sindy 2025/1/7
      Case 1
         'Added by Morgan 2012/5/17
         '外專人員無P案權限者對FMP案只能改特定欄位
         If m_bolFmpAuth Then
            'Modified by Morgan 2023/5/15 只開放FC資料頁籤內欄位--Sharon
            Text1(75).Locked = False
            Text1(77).Locked = False
            Text1(89).Locked = False
            Text1(76).Locked = False
            Text1(167).Locked = False
            Text1(49).Locked = False
            Text1(50).Locked = False
            Text1(151).Locked = False
            Text1(152).Locked = False
            Text1(88).Locked = False
            Text1(71).Locked = False
            Text1(141).Locked = False
            Text1(90).Locked = False
            Text1(70).Locked = False
            Text1(78).Locked = False
            Text1(156).Locked = False
            Text1(69).Locked = False
            Text1(146).Locked = False
            Text1(143).Locked = False
            Text1(133).Locked = False
            Text1(134).Locked = False
            Text1(135).Locked = False
            Text1(159).Locked = False
            FraPA174.Enabled = False
            'end 2023/5/15
         Else
         'End 2012/5/17
            For Each txt In Text1
               'Modify by Amy 2018/07/03 只有電腦中心才可改 特殊出名公司
               If txt.Index = 161 Then
                 If Pub_StrUserSt03 = "M51" Then txt.Locked = False
               Else
                 txt.Locked = False
               End If
            Next
            For Each txt In Text3
               txt.Locked = False
            Next
'            For Each cbo In Combo1
               Combo1.Locked = False
'            Next

         'End If Removed by Morgan 2023/5/12
         
            For Each cmd In Command1
               cmd.Enabled = True
            Next
            Combo3.Enabled = True 'Add By Sindy 2010/10/27
            'Add By Sindy 2016/11/23
            Combo4.Locked = False
            Combo5.Locked = False
            '2016/11/23 End
            'Added by Lydia 2020/02/21 開啟「名稱有特殊字」
            If ActionEdit = 1 Then '修改
              FraPA174.Enabled = True
            Else
              FraPA174.Enabled = False
            End If
            'end 2020/02/21
            Combo6.Locked = False 'Added by Morgan 2023/3/30
            Frame1K.Enabled = True 'Add By Sindy 2025/1/7
         End If 'Added by Morgan 2023/5/12
      Case 2
         TxtLock 1
         For Each txt In Text1
            txt.Text = ""
            txt.Tag = "" 'Added by Lydia 2019/11/27 預設清空
         Next
         For Each txt In Text3
            txt.Text = ""
         Next
         'end 2008/10/13
         textCUID = ""
         cboContact.Clear
'         For Each cbo In Combo1
            If Combo1.ListCount > 0 Then
               Combo1.ListIndex = 0
            Else
               Combo1.ListIndex = -1
            End If
'         Next
         For Each Lbl In Label2
            Lbl.Caption = ""
         Next
         Combo3.Text = "" 'Add By Sindy 2010/10/27
         'Add By Sindy 2014/11/6
         GRD1.Clear
         'Modified by Lydia 2024/12/03
         'SetGrd
         Call SetGrd(GRD1)
         
         '2014/11/6 END
         lstPA166.Clear 'Added by Morgan 2015/6/11
         
         lblTot6.Caption = Empty 'Added by Lydia 2018/12/27
         
         ChkPA174.Value = vbUnchecked 'Added by Lydia 2020/02/21 清空「名稱有特殊字」
         
         'Add By Sindy 2016/11/23
         If Combo4.Visible = True Then
            'Modified by Lydia 2016/11/24 第二次查詢會出錯
            'Combo4.Text = ""
            'Combo5.Text = ""
             Combo4.ListIndex = 0
             Combo5.ListIndex = 0
         End If
         '2016/11/23 End
         Combo6.ListIndex = -1 'Added by Morgan 2023/3/30
         'Add By Sindy 2025/1/7
         For Each txt In Chk1K
            txt = Empty
         Next
         '2025/1/7 END
   End Select
   
   'Added by Lydia 2021/08/16 專利連結通知：僅供顯示(開放由工程師在frm090907維護)
   If Text1(177).Locked = False Then Text1(177).Locked = True
End Sub

Private Sub GetListData(ByVal strValue As String, ByRef strTxt() As String)
 Dim strTmp As String, iPos As Integer
   strTmp = LTrim(strValue)
   iPos = InStr(strTmp, " ")
   strTxt(1) = Left(strTmp, iPos - 1)
   strTxt(2) = LTrim(Mid(strTmp, iPos))
   Text3(0).Text = Trim(Mid(strTmp, 29, 8))
   Text3(1).Text = Trim(Mid(strTmp, 59, 1))
End Sub

Private Sub MSHFlexGrid1_Click()
   'Modify By Sindy 2009/06/29
   'GridClick MSHFlexGrid1, intRow, 4
   GridClick MSHFlexGrid1, intRow, 5
   
   'Add by Morgan 2010/3/2
   'Remove by Morgan 2010/3/17 不必帶舊資料
   'If intRow > 0 Then
   '   Text3(0) = MSHFlexGrid1.TextMatrix(intRow, 2)
   '   Text3(1) = MSHFlexGrid1.TextMatrix(intRow, 3)
   'End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ErrHand
   Select Case Button.Index
      Case 1 '按下新增
         RsSitu 0
      Case 2 '按下修改
         RsSitu 1
      Case 3 '按下刪除
         RsSitu 2
      Case 4 '按下查詢
         RsSitu 5
      Case 6 '第一筆
         RsAction 0
      Case 7 '前一筆
         RsAction 1
      Case 8 '後一筆
         RsAction 2
      Case 9 '最後筆
         RsAction 3
      Case 11 '按下確定
         RsSitu 3
      Case 12 '按下取消
         RsSitu 4
      Case 14 '結束
         Unload Me
   End Select

'Remove by Morgn 2009/9/14 集中到 CmdSitu 處理
'   ' Ken 90.07.16 -- Start
'   If Button.Index <> 14 And Button.Index <> 1 And Button.Index <> 2 And Button.Index <> 3 And Button.Index <> 4 Then
'      If m_bInsert Then
'          TBar1.Buttons(1).Enabled = True
'      Else
'          TBar1.Buttons(1).Enabled = False
'      End If
'      If m_bUpdate Then
'          TBar1.Buttons(2).Enabled = True
'      Else
'          TBar1.Buttons(2).Enabled = False
'      End If
'      If m_bDelete Then
'          TBar1.Buttons(3).Enabled = True
'      Else
'          TBar1.Buttons(3).Enabled = False
'      End If
'   End If
'   ' Ken 90.07.16 -- End
   Exit Sub
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

Private Function CheckRule() As Boolean
 Dim i As Integer, bolChk As Boolean, j As Integer
   CheckRule = False
   If Text1(1) = "" Or Text1(2) = "" Then
      MsgBox "本所案號不可空白 !", vbCritical
      Text1(1).SetFocus
      Exit Function
   End If
   If Text1(23) = "" Then
      MsgBox "卷宗性質不可空白 !", vbCritical
      Text1(23).SetFocus
      Exit Function
   End If
   If Text1(8) = "" Then
      MsgBox "專利種類不可空白 !", vbCritical
      Text1(8).SetFocus
      Exit Function
   End If
   If Text1(9) = "" Then
      MsgBox "申請國家不可空白 !", vbCritical
      Text1(9).SetFocus
      Exit Function
   End If
   If Text1(5) = "" And Text1(6) = "" And Text1(7) = "" Then
      MsgBox "案件名稱不可同時空白 !", vbCritical
      Text1(5).SetFocus
      Exit Function
   End If
   'If Text1(17) = "" Then
   '   MsgBox "專利權是否存在不可空白 !", vbCritical
   '   Text1(17).SetFocus
   '   Exit Function
   'End If
   '申請人
   If Text1(26) = "" Then
      '2006/3/1 MODIFY BY SONIA
'      MsgBox "申請人1不可空白 !", vbCritical
'      Text1(26).SetFocus
'      Exit Function
      If Text1(75) = "" Then
         MsgBox "申請人1和代理人不可同時空白 !", vbCritical
         Text1(26).SetFocus
         Exit Function
      End If
      '2006/3/1 END
   '92.6.28 add by sonia
   Else
      ChkKeyIn 26
   '92.6.28 end
   End If
   If Text1(24) <> "" And Text1(25) <> "" Then
      If Not ChkRange(Text1(24), Text1(25), "專用期限") Then
         Text1(24).SetFocus
         Exit Function
      End If
   End If
   
   bolChk = True
   j = 1
   For i = 1 To 5
      If Text1(i + 25).Text <> "" Then
         strExc(j) = Text1(i + 25).Text
         j = j + 1
      End If
   Next
   Sort strExc, j - 2
   For i = 1 To j - 2
      If strExc(i) = strExc(i + 1) Then
         bolChk = False
         Exit For
      End If
   Next
   If Not bolChk Then
      MsgBox "申請人不可重覆 !", vbCritical
      Text1(26).SetFocus
      Exit Function
   End If
   
   '2008/10/13 add by toni　加FCP工程師組別
   If Text1(1) = "FCP" And Text1(2) >= "035187" And Combo2 = "" Then
      MsgBox "請輸入FCP工程師組別"
      Combo2.SetFocus
      Exit Function
   End If
   'end 2008/10/13
   CheckRule = True
End Function

Private Sub Text1_Change(Index As Integer)
'Add By Cheng 2002/01/11
'若申請人為空白, 自動清除相關地址及代表人
Select Case Index
'Added by Morgan 2014/7/29
Case 1
   'Removed by Morgan 2015/7/9 改放PA60
   'If Text1(1) = "P" Then
   '   Label1(175) = "P一案兩請是否放棄新型：　　　　 ( Y：放棄新型)"
   'Else
   'end 2015/7/9
   '   Label1(175) = "是否另函通知初審核准後分割：　　( Y：是 N：否)" 'Removed by Morgan 2019/10/7
   'End If'Removed by Morgan 2015/7/9
   
     
Case 27 '申請人2
    'Add By Cheng 2002/12/03
    If Me.Text1(Index).Text <> "" Then Exit Sub
   Me.Text1(32).Text = Empty:   Me.Text1(37).Text = Empty
   Me.Text1(42).Text = Empty
   Me.Text1(109).Text = Empty:   Me.Text1(110).Text = Empty
   Me.Text1(111).Text = Empty:   Me.Text1(112).Text = Empty
   Me.Text1(113).Text = Empty:   Me.Text1(114).Text = Empty
Case 28 '申請人3
    'Add By Cheng 2002/12/03
    If Me.Text1(Index).Text <> "" Then Exit Sub
   Me.Text1(33).Text = Empty:   Me.Text1(38).Text = Empty
   Me.Text1(43).Text = Empty
   Me.Text1(115).Text = Empty:   Me.Text1(116).Text = Empty
   Me.Text1(117).Text = Empty:   Me.Text1(118).Text = Empty
   Me.Text1(119).Text = Empty:   Me.Text1(120).Text = Empty
Case 29 '申請人4
    'Add By Cheng 2002/12/03
    If Me.Text1(Index).Text <> "" Then Exit Sub
   Me.Text1(34).Text = Empty:   Me.Text1(39).Text = Empty
   Me.Text1(44).Text = Empty
   Me.Text1(121).Text = Empty:   Me.Text1(122).Text = Empty
   Me.Text1(123).Text = Empty:   Me.Text1(124).Text = Empty
   Me.Text1(125).Text = Empty:   Me.Text1(126).Text = Empty
Case 30 '申請人5
    'Add By Cheng 2002/12/03
    If Me.Text1(Index).Text <> "" Then Exit Sub
   Me.Text1(35).Text = Empty:   Me.Text1(40).Text = Empty
   Me.Text1(45).Text = Empty
   Me.Text1(127).Text = Empty:   Me.Text1(128).Text = Empty
   Me.Text1(129).Text = Empty:   Me.Text1(130).Text = Empty
   Me.Text1(131).Text = Empty:   Me.Text1(132).Text = Empty
End Select

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   If ActionEdit <> 3 Then
   Select Case Index
      Case 5, 7, 31, 32, 33, 34, 35, 41, 42, 43, 44, 45, 51, 53, 54, 56, 79, 81, 82, 84, 91, 98, 100, 109, 111, 112, 114, 115, 117, 118, 120, 121, 123, 124, 126, 127, 129, 130, 132, 139
         'edit by nickc 2007/06/06 切換輸入法改用API
         'Text1(Index).IMEMode = 1
         OpenIme
      Case Else
         'edit by nickc 2007/06/06 切換輸入法改用API
         'Text1(Index).IMEMode = 2
         CloseIme
   End Select
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As ReturnInteger)
   If ActionEdit = 3 Then Exit Sub
   '若在查詢的狀態下按下Enter鍵
   If ActionEdit = 2 And KeyAscii = 13 Then
      Select Case Index
         Case 1, 2, 3, 4
            RsSitu 3
      End Select
      Exit Sub
   End If
   Select Case Index
      'Modify by Amy 2014/03/26 +164
      'Modify By Sindy 2012/3/2 +160
      'Modify by Morgan 2014/8/29 +166
      Case 1, 26, 27, 28, 29, 30, 59, 75, 76, 86, 88, 101, 105, 133, 134, 160, 164, 166
         KeyAscii = UpperCase(KeyAscii)
      'Modified by Morgan 2022/12/1 +178
      Case 16, 178
         If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Modify by Morgan 2008/11/13 +151,152
      'Modify by Morgan 2009/9/14 +153,154
      Case 49, 50, 151, 152, 153, 154
         If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 23, 85
         If (KeyAscii < 49 Or KeyAscii > 51) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Modified by Lydia 2016/08/18 拿掉PA70
      'Case 18, 19, 46, 57, 70, 71, 78, 89, 107, 108
      'Modified by Lydia 2019/11/27 +FCP年費自動代繳PA70
      'Modified by Morgan 2025/2/7 +181
      Case 18, 19, 46, 57, 69, 70, 71, 78, 89, 107, 108, 181
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then 'Y/ null
            KeyAscii = 0
            Beep
         End If
      'Added by Lydia 2018/12/27 中文本資訊-各項頁數
      'Modified by Lydia 2019/01/10 +申請專利範圍項數(最初項數)PA172、圖式圖數PA173
      Case 64, 65, 66, 67, 68, 172, 173
        If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
           KeyAscii = 0
           Beep
        End If
      'Added by Lydia 2016/08/18 FCP年費自動代繳(Y)/寄證書後年費不續辦(N)
      'Modified by Lydia 2019/11/27 改成FCP年費特殊管制PA156: Y:年費續辦  N:寄證書/二核後年費不續辦  空白:視代理人/申請人設定
      'Case 70
      Case 156
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then 'Y/N/ null
            KeyAscii = 0
            Beep
         End If
      'Modify by Morgan 2008/6/3 +147
      'Modify by Morgan 2007/10/26 +141
      'Modify by Morgan 2008/2/27 +143
      'Modify by Morgan 2008/4/10 +146
      'Modified by Morgan 2015/5/27 +167
      'Modified by Morgan 2017/9/12 +61
      Case 141, 143, 146, 147, 167, 61
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Modify by Morgan 2010/6/17 +157
      'Modified by Morgan 2012/12/26 +162,163
      'Modified by Morgan 2015/7/9 +60
      'Modified by Morgan 2021/6/29 +176
      Case 17, 157, 162, 163, 60, 176
         KeyAscii = UpperCase(KeyAscii)
         'Added by Morgan 2014/7/29
         If Text1(1) = "P" And Index = 162 Then
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         Else
         'end 2014/7/29
            If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         End If 'Added by Morgan 2014/7/29
         
         
      'Add by Morgan 2008/1/17
      'Modify by Morgan 2009/9/14 +155
      'Modify by Morgan 2011/4/1 +90
      'Modified by Morgan 2012/8/20 +161
      Case 142, 155, 90 ', 161
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 161
         KeyAscii = UpperCase(KeyAscii)
         'Added by Lydia 2020/03/30 智慧所更名日起特殊出名公司欄修改時不可輸T及L
         If strSrvDate(1) >= 智慧所更名日 Then
            'J,空白
            If KeyAscii <> 74 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         'end 2020/03/30
         'Modify By Sindy 2013/12/16
         'Modified by Lydia 2020/03/30 +elseif
         ElseIf strSrvDate(1) >= InvoiceStartDate Then
            'J,T,空白
            If KeyAscii <> 74 And KeyAscii <> 84 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         Else
         '2013/12/16 END
            'Y,空白
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         End If
   End Select
End Sub

Private Function ChkKeyIn(ByVal iSitu As Integer) As Boolean
 Dim strTmp As String, strMain As String, bolChk As Boolean, i As Integer
   ChkKeyIn = False
   strMain = Text1(iSitu).Text
   Select Case iSitu
      Case 1
         Call AuthCheck(Text1(iSitu).Text)
         If strMain <> m_SysKind Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            ChkKeyIn = True
         End If
      Case 3
         If strMain = "" Then Text1(iSitu) = "0"
      Case 4
         If strMain = "" Then Text1(iSitu) = "00"
      Case 8
         If Text1(9).Text = 台灣國家代號 Or m_SysKind = "CFP" Then
            bolChk = False
         Else
            bolChk = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, strMain, strTmp, bolChk, Text1(9).Text) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, strMain, strTmp, bolChk, Text1(9).Text) = 1 Then
            ChkKeyIn = False
            If Text1(9).Text <> "" Then
               '2010/2/12 MODIFY BY SONIA 抓修法次數
               'If GetMoneyDate(Val(Text1(8).Text), Text1(9).Text, pa, strExc(0), strExc(1)) Then varYear = Split(strExc(1), ",")
               'Modified by Morgan 2025/2/26 ECP子案要抓EPC的繳費年度
               If GetMoneyDate(Val(Text1(8).Text), IIf(pa(4) <> "00", "221", Text1(9).Text), pa, strExc(0), strExc(1), , , m_FixNo) Then varYear = Split(strExc(1), ",")
            End If
         Else
            ChkKeyIn = True
         End If
         Label2(1) = strTmp
      Case 9
         If strMain < "009" And strMain > "000" Then
            MsgBox "國家代號不可輸入 001 至 008，請重新輸入 !", vbCritical
            Label2(0) = ""
            ChkKeyIn = True
         Else
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetNation(strMain, strTmp) Then
            If ClsPDGetNation(strMain, strTmp) Then
               ChkKeyIn = False
               If Text1(8).Text <> "" Then
                  '2010/2/12 MODIFY BY SONIA 抓修法次數
                  'If GetMoneyDate(Val(Text1(8).Text), Text1(9).Text, pa, strExc(0), strExc(1)) Then varYear = Split(strExc(1), ",")
                  'Modified by Morgan 2025/2/26 ECP子案要抓EPC的繳費年度
                  If GetMoneyDate(Val(Text1(8).Text), IIf(pa(4) <> "00", "221", Text1(9).Text), pa, strExc(0), strExc(1), , , m_FixNo) Then varYear = Split(strExc(1), ",")
               End If
            Else
               ChkKeyIn = True
            End If
            Label2(0) = strTmp
         End If
      Case 11
         If Text1(9).Text = 台灣國家代號 Then
            ChkKeyIn = Not ChkAppNo(strMain, Val(Text1(8).Text), 0, Text1(23))
         'Modify by Morgan 2004/3/26
         '大陸PCT案不檢查申請案號
         'ElseIf Text1(9).Text = 大陸國家代號 Then
         ElseIf Text1(9).Text = 大陸國家代號 And Text1(46).Text <> "Y" Then
            ChkKeyIn = Not ChkAppNo(strMain, Val(Text1(8).Text), 1, Text1(23))
         End If
      Case 26, 27, 28, 29, 30
        'Modify By Cheng 2003/08/07
        '若為更改申請人
        If m_CustNo(iSitu - 25) <> Left(Me.Text1(iSitu) & "000000000", 9) Then
            'edit by nickc 2007/02/02 不用 dll 了
            'ChkKeyIn = Not objPublicData.GetCustomerNameAndAddress(strMain, strTmp, strExc(0), strExc(1), strExc(2))
            strExc(0) = "": strExc(1) = "": strExc(2) = ""   'Added by Lydia 2024/06/13
            ChkKeyIn = Not ClsPDGetCustomerNameAndAddress(strMain, strTmp, strExc(0), strExc(1), strExc(2))
            Text1(iSitu).Text = strMain
            Label2(iSitu - 20) = strTmp
            If ActionEdit <> 3 Then
               For i = 0 To 2
                  Text1(31 + i * 5 + (iSitu - 26)).Text = strExc(i)
               Next
            End If
            If iSitu = 26 Then
               strExc(0) = "select cu72,cu79 from customer where " & ChgCustomer(strMain)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If IsNull(RsTemp.Fields(0)) Then
                     Label2(4) = ""
                  Else
                     Label2(4) = RsTemp.Fields(0)
                  End If
                  If IsNull(RsTemp.Fields(1)) Then
                     Text2(1) = ""
                  Else
                     Text2(1) = RsTemp.Fields(1)
                  End If
               End If
            End If
        Else
            ChkKeyIn = False
        End If
      Case 86, 88
         'edit by nickc 2007/02/05 不用 dll 了
         'ChkKeyIn = Not objLawDll.LawGetName(strMain, strTmp)
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         Label2(iSitu - 75) = strTmp
      Case 59
         'edit by nickc 2007/02/05 不用 dll 了
         'ChkKeyIn = Not objLawDll.GetReasonOfRelief(strMain, strTmp)
         ChkKeyIn = Not ClsLawGetReasonOfRelief(strMain, strTmp)
         Label2(5) = strTmp
      Case 75
         'Modify By Cheng 2002/07/09
'         ChkKeyIn = Not objPublicData.GetAgent(strMain, strTmp)
         ChkKeyIn = Not PUB_GetAgentName(Me.Text1(1).Text, strMain, strTmp)
         Text1(iSitu).Text = strMain
         Label2(2) = strTmp
        'Add By Cheng 2003/02/24
        '若輸入錯誤顯示訊息
        If ChkKeyIn = True Then MsgBox "代理人輸入錯誤!!!", vbExclamation + vbOKOnly
         If strMain <> "" Then
            strExc(0) = "select fa39,fa29 from fagent where " & ChgFagent(strMain)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If IsNull(RsTemp.Fields(0)) Then
                  Label2(3) = ""
               Else
                  Label2(3) = RsTemp.Fields(0)
               End If
               If IsNull(RsTemp.Fields(1)) Then
                  Text2(0) = ""
               Else
                  Text2(0) = RsTemp.Fields(1)
               End If
            End If
         End If
      Case 76
         'edit by nickc 2007/02/05 不用 dll 了
         'ChkKeyIn = Not objLawDll.LawGetName(strMain, strTmp)
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         Label2(12) = strTmp
      Case 105
         'edit by nickc 2007/02/05 不用 dll 了
         'ChkKeyIn = Not objLawDll.LawGetName(strMain, strTmp)
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         Label2(45) = strTmp
      Case 101
         'edit by nickc 2007/02/05 不用 dll 了
         'ChkKeyIn = Not objLawDll.LawGetName(strMain, strTmp)
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         'Modified by Lydia 2015/12/16
         'Label2(46) = strTmp
         Label2(46) = convForm(strTmp, 18)
      'Modified by Morgan 2012/5/9 +140
      Case 10, 12, 14, 20, 21, 58, 140
         ChkKeyIn = Not ChkDate(strMain)
      Case 24, 25
         If Len(Text1(iSitu)) <> 8 Then
            MsgBox "日期必須為西元年，請重新輸入 !", vbCritical
         Else
            ChkKeyIn = Not ChkDate(strMain)
         End If
      Case 133
         'edit by nickc 2007/02/05 不用 dll 了
         'ChkKeyIn = Not objLawDll.LawGetName(strMain, strTmp)
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         Label2(48) = strTmp
      Case 134
         'edit by nickc 2007/02/05 不用 dll 了
         'ChkKeyIn = Not objLawDll.LawGetName(strMain, strTmp)
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         Label2(49) = strTmp
    
      'Added by Morgan 2014/8/29
      'Removed by Morgan 2015/6/11
      'Case 166
      '   ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
      '   Label2(52) = strTmp
         
   End Select
End Function

Private Sub Text1_LostFocus(Index As Integer)
   If ActionEdit = 0 Then
      If Index = 4 Then
         If Not ChkCaseCode("PA", Text1(1).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text) Then
            Text1(2).SetFocus
         End If
      End If
      
   End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTmp As String, i As Integer, strTxt As String
   If ActionEdit = 3 Then Exit Sub
   
   If Text1(Index) = "" Then
         
      Select Case Index
         Case 8
            Label2(1) = ""
         Case 9
            Label2(0) = ""
         Case 26, 27, 28, 29, 30
            Label2(Index - 20) = ""
'            If Index = 30 Then GoTo A0
         Case 86, 88
            Label2(Index - 75) = ""
         Case 59
            Label2(5) = ""
         Case 75
            Label2(2) = ""
            Label2(3) = ""
            Text2(0) = ""
         Case 76
            Label2(12) = ""
         Case 105
            Label2(45) = ""
         Case 101
            Label2(46) = ""
         Case 133
            Label2(48) = ""
         Case 134
            Label2(49) = ""
         '2008/8/21 ADD BY SONIA 無專用期時專利權是否存在欄清空
         Case 24, 25
            Text1(17).Text = ""
         '有專用期時專利權是否存在欄不可空白
         Case 17
            If Text1(25).Text <> "" Then
               MsgBox "有專用期時, 專利權是否存在欄不可空白 !", vbCritical
               Cancel = True
               Text1(17).SetFocus
            End If
         '2008/8/21 END
                  
         'Added by Morgan 2014/8/29
         'Removed by Morgan2015/6/11
         'Case 166
         '   Label2(52) = ""
         
      End Select
      Exit Sub
   'Add by morgan 2004/11/4 當有輸專用期時若准駁為空白則預設為1
   Else
      Select Case Index
         'add by Toni 2008/10/17
         Case 1
            'Modified by Morgan 2012/3/15 +P
            If Text1(1) = "FCP" Or Text1(1) = "P" Then
               Combo2.Enabled = True
            Else
               Combo2.Enabled = False
            End If
         'end 2008/10/17
         'Added by Morgan 2012/8/21
         Case 11
            'Modified by Morgan 2016/3/24 +傳卷宗性質
            If PUB_ChkAppNo(Text1(11), Text1(1), Text1(2), Text1(9), Text1(23)) = False Then
               Cancel = True
               Exit Sub
            'Added by Morgan 2019/2/19
            '申請案要再檢查是否有對造號數存在(前面相同就算,因台灣案會有NXX)
            ElseIf Text1(23) = "1" And Text1(11) <> Text1(11).Tag Then
               '注意:CP36用比較語法是因為instr會很慢
               strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04),cp09" & _
                  " from caseprogress,patent where cp36>='" & Text1(11) & "' and cp36<'" & Text1(11) & "~'" & _
                  " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='" & Text1(9) & "'" & _
                  " and pa01||pa02<>'" & Text1(1) & Text1(2) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If MsgBox("本案申請號與 " & RsTemp(0) & " 案(" & RsTemp(1) & ")的對造號相同，是否確定要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                     Text1(11).Tag = Text1(11)
                  Else
                     Cancel = True
                     Exit Sub
                  End If
               End If
            'end 2019/2/19
            End If
            
         'end 2012/8/21
         Case 24
            If Text1(16).Text = "" Then Text1(16).Text = "1"
      
         'Add by Morgan 2009/10/16
         Case 142, 155
            If (Text1(142).Text = "" And Text1(155).Text = "Y") Then
               MsgBox "【EMail 同時寄紙本】為 Y 時，【以EMail 通知】欄位也必須為 Y！"
               Cancel = True
               Exit Sub
            End If
         'Add by Amy 2014/03/24 +pa164
         Case 164
            If Len(Trim(Text1(164))) <> "4" Then
                MsgBox Label1(176) & "輸入錯誤！", vbCritical
                Cancel = True
               Exit Sub
            End If
      End Select
   End If
   
   'Add By Sindy 2014/7/8
   If Index = 8 And Trim(Text1(8)) <> "" Then
      If (Trim(Text1(8)) <> "3" And Combo3.List(0) <> "1.機械") Or _
         (Trim(Text1(8)) = "3" And Combo3.List(0) <> "1.整體") Then
         Call PUB_AddCaseAttributeCombo(Combo3, Trim(Text1(8))) '專利案件屬性選單 Modify By Sindy 2020/3/10
'         Combo3.Clear
'         If Trim(Text1(8)) = "3" Then '設計案
'            Combo3.AddItem "1.整體"
'            Combo3.AddItem "2.部分"
'            Combo3.AddItem "3.圖像"
'            Combo3.AddItem "4.成組"
'         Else
'            Combo3.AddItem "1.機械"
'            Combo3.AddItem "2.電子電機"
'            Combo3.AddItem "3.化學生醫"
'         End If
      End If
   End If
   '2014/7/8 END
   
   'Added by Lydia 2018/12/27 中文本-頁數總計
   If Index >= 64 And Index <= 68 Then
        i = Val(Text1(64)) + Val(Text1(65)) + Val(Text1(66)) + Val(Text1(67)) + Val(Text1(68))
        lblTot6.Caption = i
   End If
   'end 2018/12/27
   
   '檢查中文欄位長度是否過長
   'Modify By Sindy 2014/11/6 發明人
   If CheckLengthIsOK(Text1(Index).Text, rsDefineSize.Fields(Index - 1).DefinedSize) Then
      Cancel = ChkKeyIn(Index)
      '申請人
      If Cancel = False And (Index = 26 Or Index = 27 Or Index = 28 Or Index = 29 Or Index = 30) Then
         strTxt = Combo1.Text
         strTmp = ""
         For i = 26 To 30
            'If Text1(i).Text <> "" Then strTmp = strTmp & "'" & Left(Text1(i).Text, 8) &
            '   String(8 - Len(Text1(i).Text), "0") & "',"
            ' 90.06.27 modify by louis
            If Len(Text1(i).Text) >= 8 Then
               strTmp = strTmp & "'" & Left(Text1(i).Text, 8) & "',"
            Else
               strTmp = strTmp & "'" & Left(Text1(i).Text, 8) & String(8 - Len(Text1(i).Text), "0") & "',"
            End If
         Next
         If strTmp <> "" Then strTmp = Left(strTmp, Len(strTmp) - 1)
         'Modified by Lydia 2024/12/03 改成GetCombo1Data
         'ChgCombo strTmp
         GetCombo1Data strTmp
         
'On Error Resume Next
         If strTxt <> "" Then
            Combo1.Text = strTxt
            If Err.Number = 383 Then
               Combo1.AddItem strTxt
               Combo1.ListIndex = Combo1.ListCount - 1
            End If
         Else
            Combo1.ListIndex = 0
         End If
      End If
   Else
      Cancel = True
   End If
   If Cancel = True Then TextInverse Text1(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 1 Then
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub Text3_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 And Text3(0) <> "" Then Cancel = Not ChkDate(Text3(0).Text)
End Sub

'將資料放入Grid
Public Sub InitFeeData()
Dim i As Integer, varTmp1 As Variant, varTmp2 As Variant, varTmp3 As Variant, strTmp As String
Dim strFormat As String
'Add By Sindy 2009/06/29
Dim strFeeType As String, strYF15 As String
   
   GridHead
   If pa(72) <> "" Then
      varTmp1 = Split(pa(72), ",")
      varTmp2 = Split(pa(73), ",")
      varTmp3 = Split(pa(74), ",")
      With MSHFlexGrid1
         'Modified by Morgan 2022/6/13 +pa(10)
         strFeeType = PUB_GetNa20Na22Na24(pa(9), pa(8), pa(10)) 'Add By Sindy 2009/06/29
         For i = 0 To UBound(varTmp1)
            'Modify By Sindy 2009/06/29
            'strTmp = varTmp1(i)
            '繳費次數
            strTmp = i + 1
            '年度說明
            '2010/2/12 modify by sonia
            'strYF15 = PUB_GetYF15(pa(9), pa(8), "Y0000000", strFeeType, CDbl(varTmp1(i)))
            strYF15 = PUB_GetYF15(pa(9), pa(8), "Y000000" & m_FixNo, strFeeType, CDbl(varTmp1(i)))
            strTmp = strTmp & vbTab & strYF15
            '2009/06/29 End
            '繳費日期
            If UBound(varTmp2) >= i Then
               strTmp = strTmp & vbTab & TransDate(varTmp2(i), 1)
            Else
               strTmp = strTmp & vbTab & ""
            End If
            '費用是否雙倍
            If UBound(varTmp3) >= i Then
               strTmp = strTmp & vbTab & varTmp3(i)
            Else
               strTmp = strTmp & vbTab & ""
            End If
            .AddItem strTmp
         Next
         FixGrid MSHFlexGrid1
         'Modify By Sindy 2009/06/29
         'If .Rows > 1 Then GridClick MSHFlexGrid1, 1, 4
         'Modify by Morgan 2010/3/2
         'If .Rows > 1 Then GridClick MSHFlexGrid1, 1, 5
         If .Rows > 1 Then MSHFlexGrid1_Click
      End With
   End If
End Sub

Private Sub GridHead()
   With MSHFlexGrid1
      'Modify By Sindy 2009/06/29
      'InitGrid 3, MSHFlexGrid1
      InitGrid 4, MSHFlexGrid1
      '2009/06/29 End
      .Rows = 1
      'Modify By Sindy 2009/06/29
      '.ColWidth(0) = 900:      .TextMatrix(0, 0) = "繳費年度"
      .col = 0
      .ColWidth(0) = 500:      .TextMatrix(0, 0) = "次數"
      .ColAlignment = flexAlignCenterCenter
      .col = 1
      .ColWidth(1) = 2200:      .TextMatrix(0, 1) = "年度說明"
      .ColAlignment = flexAlignLeftCenter
      '2009/06/29 End
      .col = 2
      .ColWidth(2) = 800:      .TextMatrix(0, 2) = "繳費日期"
      .ColAlignment = flexAlignRightCenter
      .col = 3
      'Modified by Morgan 2021/3/24
      '.ColWidth(3) = 800:    .TextMatrix(0, 3) = "費用雙倍"
      .ColWidth(3) = 800:    .TextMatrix(0, 3) = "逾期補繳"
      .ColAlignment = flexAlignLeftCenter
   End With
End Sub

Private Function GetFeeData() As Boolean
Dim i As Integer, varTemp As Variant
   
   For i = 72 To 74
      pa(i) = ""
   Next
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         'Modify By Sindy 2009/06/29
         'pa(72) = pa(72) & .TextMatrix(i, 0) & ","
         'pa(73) = pa(73) & TransDate(.TextMatrix(i, 1), 2) & ","
         'pa(74) = pa(74) & .TextMatrix(i, 2) & ","
         pa(72) = pa(72) & varYear(CDbl(.TextMatrix(i, 0)) - 1) & ","
         pa(73) = pa(73) & TransDate(.TextMatrix(i, 2), 2) & ","
         pa(74) = pa(74) & .TextMatrix(i, 3) & ","
         '2009/06/29 End
      Next
   End With
   For i = 72 To 74
      If Right(pa(i), 1) = "," Then pa(i) = Left(pa(i), Len(pa(i)) - 1)
   Next
   GetFeeData = True
End Function

Private Function GetData() As Boolean
Dim i As Integer
Dim txt As Object
   
   GetData = False
   If CheckRule = False Then Exit Function
   'Modify By Sindy 2014/11/6 Mark
'   For i = 0 To 9
'      'Modify By Cheng 2002/07/02
''      Text1(i + 60).Text = Combo1(i).Text
'      Text1(i + 60).Text = Replace(Right(Combo1(i).Text, 11), ")", "")
'   Next
   '2014/11/6 END
   
   'Modify by Morgan 2006/10/18 避免畫面物件未加欄位已新增時發生錯誤
   'For i = 1 To TF_PA 'edit by nickc 2006/07/12 T_PA
   'Modify by Morgan 2007/10/26 改有物件的才放,否則畫面有欄位沒放時會錯
   'For i = 1 To Me.Text1.Count - 1
   '     pa(i) = Text1(i).Text
   'Next
   For Each txt In Text1
      i = txt.Index
      If i > 0 Then
         pa(i) = Text1(i).Text
      End If
   Next
   'end 2007/10/26
   pa(24) = TransDate(pa(24), 1)
   pa(25) = TransDate(pa(25), 1)
   If pa(3) = "" Then pa(3) = "0"
   If pa(4) = "" Then pa(4) = "00"
   If Not GetFeeData Then
      MsgBox "讀取年費資料錯誤，請重新輸入 !", vbCritical
      Exit Function
   End If
   
   'Add By Sindy 2016/11/23
   pa(170) = Combo4.Text
   pa(171) = IIf(Combo5.Text <> "", Combo5.ListIndex, "")
   '2016/11/23 END
   
   pa(174) = IIf(ChkPA174.Value = 0, "", "Y") 'Added by Lydia 2020/02/21 設定「名稱有特殊字」
   'Added by Morgan 2023/3/30
   If Combo6.ListIndex > 0 Then
      pa(179) = Combo6.ListIndex
   Else
      pa(179) = ""
   End If
   'end 2023/3/30
   GetData = True
End Function

' 90.06.27 add by louis
Private Sub UpdateCustomerName()
   Dim nIndex As Integer
   For nIndex = 26 To 30
      If IsEmptyText(Text1(nIndex)) = False Then
         Label2(nIndex - 26 + 6) = GetCustomerName(Text1(nIndex), 0)
      End If
   Next nIndex
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   'Add By Sindy 2014/11/6
   cmdAddRow.Enabled = True
   cmdDelRow.Enabled = True
   '2014/11/6 END
   'Added by Lydia 2024/12/03
   cmdUp.Enabled = True
   cmdDown.Enabled = True
   Frame3.Enabled = True
   'end 2024/12/03
   
   TxtValidate = False
   
   'Add by Sindy 2021/12/07 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Function
   End If
   
   For Each objTxt In Me.Text1
      If objTxt.Enabled = True Then
         Cancel = False
         Text1_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   For Each objTxt In Me.Text3
      If objTxt.Enabled = True Then
         Cancel = False
         Text3_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   'Add by Morgan 2007/5/10
   If Not ((Text1(57).Text = "" And Text1(58).Text = "" And Text1(59).Text = "") Or (Text1(57).Text <> "" And Text1(58).Text <> "" And Text1(59).Text <> "")) Then
      MsgBox "是否閉卷、閉卷日期、閉卷原因三個欄位須同時空白或有值！", vbExclamation
      Exit Function
   End If
   'end 2007/5/10
   
      '2010/1/8 ADD BY SONIA
      If Combo2.Enabled = True Then
         Cancel = False
         Combo2_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
      '2010/1/8 END
      
      'Add By Sindy 2010/10/27
      If Combo3.Enabled = True Then
         Cancel = False
         Combo3_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
      '2010/10/27 End
   
   'Add by Morgan 2010/5/3
   strExc(1) = Text1(26)
   For intI = 1 To 4
      strExc(1) = strExc(1) & "," & Text1(26 + intI)
   Next
   'Modify By Sindy 2014/11/6 比對申請人與發明人資料是否不符
   'For intI = 0 To 9
   '   strExc(2) = Replace(Right(Combo1(intI).Text, 11), ")", "")
   '   If PUB_ChkInventor(strExc(2), strExc(1)) = False Then
   '      If MsgBox("發明人資料與目前申請人不符，是否要繼續！", vbYesNo + vbDefaultButton2) = vbNo Then
   '         If intI > 4 Then
   '            SSTab1.Tab = 5
   '         Else
   '            SSTab1.Tab = 4
   '         End If
   '         Combo1(intI).SetFocus
   '         Exit Function
   '      End If
   '   End If
   'Next
   
   If PUB_ChkCPExist(pa(), "701", 2) = False Then  'Added by Morgan 2015/1/20 有讓與發文就不必檢查
   
   If GRD1.Rows >= 2 And GRD1.TextMatrix(1, 1) <> "" Then
      For intI = 1 To GRD1.Rows - 1
         If PUB_ChkInventor(GRD1.TextMatrix(intI, 1), strExc(1)) = False Then
            'Modify By Sindy 2014/11/6
'            MsgBox "發明人資料與目前申請人不符！", vbExclamation
'            SSTab1.Tab = 4
'            Exit Function
            If MsgBox("發明人資料與目前申請人不符，是否要繼續！", vbYesNo + vbDefaultButton2) = vbNo Then
               SSTab1.Tab = 4
               Exit Function
            End If
            '2014/11/6 END
         End If
      Next
   End If
   '2014/11/6 END
   'end 2010/5/3
   
   End If 'Added by Morgan 2015/1/20
   
   'Added by Morgan 2013/3/29
   'P,CFP案沒有申請日或申請程序未發文案件, 不可在基本檔改案件屬性
   If (Text1(1) = "P" Or Text1(1) = "CFP") And Left(Combo3, 1) <> pa(158) Then
      If Text1(10) = "" Then
         strExc(0) = "select cp09 from caseprogress where cp01='" & Text1(1) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(1) & "' and cp04='" & Text1(1) & "' and instr('" & NewCasePtyList & "',cp10)>0 and cp27>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI <> 1 Then
            MsgBox "P,CFP案沒有申請日或申請程序未發文案件, 不可在基本檔改案件屬性！", vbExclamation
            Exit Function
         End If
      End If
   End If
   'end 2013/3/29
   
   'Add By Sindy 2014/7/15
   If Text1(9) <> "000" And Text1(8) = "3" And Combo3.Text <> "" Then
      MsgBox "非台灣設計案不可輸入案件屬性！", vbExclamation
      Combo3.SetFocus
      Exit Function
   End If
   '2014/7/15 END
   
   'Added by Lydia 2017/11/30 FCP案件命名電子化:未發文前，專利基本檔不可直接變更
   If strSrvDate(1) >= FCP案件命名啟用日 And Text1(1).Text = "FCP" And Combo2.Tag <> Combo2.Text Then
        strExc(0) = "SELECT CP09,CPM03 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' " & _
                    "AND CP10 IN (" & NewCasePtyList & ") AND CP158=0 AND CP159=0 AND CP01=CPM01(+) AND CP10=CPM02(+) ORDER BY CP09 DESC"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           RsTemp.MoveFirst
           If "" & RsTemp.Fields("CP09") <> "" Then
               MsgBox Trim(RsTemp.Fields("CPM03")) & "未發文前，專利基本檔不可直接變更工程師組別 !", vbExclamation
               Combo2.SetFocus
               Exit Function
           End If
        End If
   End If
   'end 2017/11/30
   
   'Added by Lydia 2020/02/21 檢查「名稱有特殊字」
   If Text1(1) = "P" Or Text1(1) = "FCP" Then
       If Pub_GetPA174toFile("2", Text1(1), Text1(2), Text1(3), Text1(4), Me, frm100101_M_1) = True Then
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
           If Text1(5) & Text1(6) & Text1(7) <> pa(5) & pa(6) & pa(7) Then
               MsgBox "名稱有特殊字，案件名稱有修改，請一併修改案件名稱Word檔。", vbInformation, "檢查資料"
               Call ProcPA174toFile("Y")
               Exit Function
           End If
       End If
   End If
   'end 2020/02/21
   
   'Added by Morgan 2016/7/19 Ex.P-092834
   If Text1(57) <> "Y" And InStr(1, Label2(44).Caption, "不續辦") > 0 Then
      If MsgBox("本案未閉卷但下次繳費日為不續辦，是否確定要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   'end 2016/7/19
   
   'Add By Sindy 2016/11/23
   If Me.Combo4.Enabled = True Then
      Cancel = False
      Combo4_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Trim(Me.Combo4.Text) <> "" Then
      '若輸入幣別就一定要選格式
      If Trim(Me.Combo5.Text) = "" Then
         ShowMsg "請款單列印幣別格式不可空白 !"
         Me.Combo5.SetFocus
         Exit Function
      End If
      '請款幣別<>NTD時不可輸入1
      If Trim(Me.Combo4.Text) <> "NTD" And Me.Combo5.ListIndex = 1 Then
         ShowMsg "請款幣別<>NTD時，請款單列印幣別格式不可選純台幣 !"
         Me.Combo5.SetFocus
         Exit Function
      End If
   End If
   '2016/11/23 ENd
   
   'Added by Lydia 2019/12/10 個案設年費自動代繳PA70，檢查年費特殊管制PA156
   If Text1(70).Text = "Y" And Text1(156).Text = "N" Then
       MsgBox "個案設年費自動代繳=Y，則年費特殊管制不可為不續辦！", vbCritical, "資料檢核"
       SSTab1.Tab = 2
       Text1(70).SetFocus
       Exit Function
   End If
   
   'Added by Lydia 2021/11/24 申請國家pa09='201'英國、卷宗性質pa23='1'申請、專利號數第1碼為'9'，若沒有建立歐盟239案的相關卷號時，則存檔時彈訊息"此為脫歐英國案，若歐盟案亦為本所案件，請建立相關卷號關聯，否則此英國案在計算結餘時會扣安全基金！"，但仍可存檔。
   If Text1(1) = "CFP" And Text1(9) = "201" And Text1(23) = "1" And Left(Text1(22), 1) = "9" Then
      strExc(0) = "select cr05 as d01,cr06 as d02,cr07 as d03,cr08 as d04 from caserelation1,patent where cr01='" & Text1(1) & "' and cr02='" & Text1(2) & "' and cr03='" & Text1(3) & "' and cr04='" & Text1(4) & "' and cr05='" & Text1(1) & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+) and pa09='239' " & _
                       "union select cr01 as d01,cr02 as d02,cr03 as d03,cr04 as d04 from caserelation1,patent where cr05='" & Text1(1) & "' and cr06='" & Text1(2) & "' and cr07='" & Text1(3) & "' and cr08='" & Text1(4) & "' and cr01='" & Text1(1) & "' and cr01=pa01(+) and cr02=pa02(+) and cr03=pa03(+) and cr04=pa04(+) and pa09='239' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
          MsgBox "此為脫歐英國案，若歐盟案亦為本所案件，請建立相關卷號關聯，否則此英國案在計算結餘時會扣安全基金！", vbExclamation, "英國脫歐案管制"
      End If
   End If
   'end 2021/11/24
   
   'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
   If Pub_ChkAppList(strExc(0), Text1(26) & "," & Text1(27) & "," & Text1(28) & "," & Text1(29) & "," & Text1(30)) = False Then
      SSTab1.Tab = 1
      Text1(Val(strExc(0)) + 25).SetFocus
      Text1_GotFocus Val(strExc(0)) + 25
      Exit Function
   End If
   'end 2024/06/14
      
   'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
   For ii = 26 To 30
      strExc(1) = ChangeCustomerL(Text1(ii))
      strExc(2) = ChangeCustomerL(pa(ii))
      If strExc(1) <> "" And strExc(1) <> strExc(2) Then
         If GetCustomerAndState(strExc(1), strExc(3), , , , pa(1), strExc(8), False, Me.Name, pa(2), pa(3), pa(4)) = False Then
            Me.SSTab1.Tab = 1
            Text1(ii).SetFocus
            Text1_GotFocus ii
            Exit Function
         End If
      End If
   Next ii
   strExc(1) = ChangeCustomerL(Text1(75))
   strExc(2) = ChangeCustomerL(pa(75))
   If strExc(1) <> "" And strExc(1) <> strExc(2) Then
      If GetAgentAndState(strExc(1), strExc(3), , , , pa(1), strExc(8), False) = False Then
         Me.SSTab1.Tab = 2
         Text1(75).SetFocus
         Text1_GotFocus 75
         Exit Function
      End If
   End If
   'end 2024/06/13
   
   'Added by Morgan 2024/10/15
   '大陸發明改專用期時檢查是否有補償期
   If pa(1) = "P" And Text1(9) = "020" And Text1(8) = "1" And Text1(25) <> "" Then
      If pa(9) = Text1(9) And DBDATE(pa(10)) = DBDATE(Text1(10)) Then
         If PUB_GetCNExtDays(pa(), Text1(25), ii, True) = True Then
            If ii > 0 Then
               Label1(22) = Label1(22).Tag & " (含補償 " & ii & " 天)"
            End If
         Else
            Exit Function
         End If
      End If
   End If
   'end 2024/10/15
   
   TxtValidate = True
End Function

'Add By Cheng 2002/07/02
'Mark by Lydia 2024/12/03 檢查過,沒有使用
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
'end ---- 'Mark by Lydia 2024/12/03 檢查過,沒有使用

'add by nickc 2006/06/07  dll 搬出，為了紀錄 log
Private Function SaveNewPatentDatabase(ByRef pa() As String, ByVal intWhere As Integer, Optional bolDate As Boolean = True) As Boolean
Dim strSql As String, i As Integer, strTmp As String, strTmp1 As String
Dim ii As Integer 'Add By Sindy 2014/11/6
   
'On Error GoTo ErrHand
   
   If intWhere <> 國外_CF Or bolDate = False Then
      pa(10) = ChangeTStringToWString(pa(10))
      pa(12) = ChangeTStringToWString(pa(12))
      pa(14) = ChangeTStringToWString(pa(14))
      pa(20) = ChangeTStringToWString(pa(20))
      pa(21) = ChangeTStringToWString(pa(21))
      pa(24) = ChangeTStringToWString(pa(24))
      pa(25) = ChangeTStringToWString(pa(25))
      pa(58) = ChangeTStringToWString(pa(58))
   End If
   pa(26) = ChangeCustomerL(pa(26))
   pa(27) = ChangeCustomerL(pa(27))
   pa(28) = ChangeCustomerL(pa(28))
   pa(29) = ChangeCustomerL(pa(29))
   pa(30) = ChangeCustomerL(pa(30))
   pa(75) = ChangeCustomerL(pa(75))
   pa(76) = ChangeCustomerL(pa(76))
   pa(86) = ChangeCustomerL(pa(86))
   pa(88) = ChangeCustomerL(pa(88))
   pa(101) = ChangeCustomerL(pa(101))
   pa(105) = ChangeCustomerL(pa(105))
   pa(133) = ChangeCustomerL(pa(133))
   pa(134) = ChangeCustomerL(pa(134))
   'add by Toni 2008/10/17
   pa(150) = Left(Combo2, 1)
   'end 2008/10/17
   'Add By Sindy 2010/10/27
   pa(158) = Left(Combo3, 1)
   '2010/10/27 End
   'pa(166) = ChangeCustomerL(pa(166)) 'Added by Morgan 2014/8/29 'Removed by Morgan 2015/6/11
   For i = 1 To 99
      Select Case i
         Case 92, 93, 94, 95, 96, 97
         Case Else
          strSql = strSql & CNULL(Replace(pa(i), "'", "''")) & ","
          strTmp = strTmp & "PA" & Format(i, "00") & ","
      End Select
   Next
   For i = 100 To TF_PA 'edit by nickc 2006/07/12 T_PA
        'add by nickc 2006/07/12
        If i <> 108 And i <> 136 And i <> 137 And i <> 138 Then
            'Add By Sindy 2025/1/7
            If i = 180 Then
               pa(i) = ""
               For Each obj In Chk1K
                  If obj.Value = 1 Then
                     pa(i) = pa(i) & "," & obj.Index + 1
                  End If
               Next
               If pa(i) <> "" Then pa(i) = Mid(pa(i), 2)
            End If
            '2025/1/7 END
            
      '      If i = 103 Then
               If InStr(pa(i), "'") > 0 Then
                  strTmp1 = Replace(pa(i), "'", "''")
                  strSql = strSql + CNULL(strTmp1) + ","
               Else
                  strSql = strSql + CNULL(pa(i)) + ","
               End If
      '      Else
      '         strSQL = strSQL + CNULL(pa(i)) + ","
      '      End If
            
            strTmp = strTmp & "PA" & Format(i) & ","
        End If
   Next
   
   strSql = Left(strSql, Len(strSql) - 1)
   strTmp = Left(strTmp, Len(strTmp) - 1)
   strSql = "insert into patent (" & strTmp & ") values (" & strSql & ")"
    'add by nickc 2006/06/07 紀錄分析語法
    'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
    'Pub_SeekTbLog strSql
    Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
   cnnConnection.Execute strSql
   
   'Add By Sindy 2014/11/6
   If cmdAddRow.Tag = "I" Or cmdDelRow.Tag = "D" Then '有異動發明人資料
      For ii = 1 To GRD1.Rows - 1
         'Modify By Sindy 2023/3/7 無發明人編號不須新增
         If GRD1.TextMatrix(ii, 1) <> "" Then
         '2023/3/7 END
            strSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                     CNULL(pa(1)) & "," & CNULL(pa(2)) & "," & CNULL(pa(3)) & "," & CNULL(pa(4)) & "," & ii & ",'" & GRD1.TextMatrix(ii, 1) & "')"
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strSql
            Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strSql
         End If
      Next ii
      
      PUB_UpdateRefCaseInventorAndCaseName pa(1), pa(2), pa(3), pa(4) 'Added by Morgan 2020/2/13
   End If
   '2014/11/6 END
   
   SaveNewPatentDatabase = True
   Exit Function
ErrHand:
   ShowMsg MsgText(9131)
   'ErrorLog
End Function

'Patent專利檔存檔(維護)
'Optional bolDate As Boolean = True 依系統別轉日期 False 不依系統別，全轉為民國日期
'Optional BolWdb As Boolean = false 寫更新時間     nick 910910
Private Function SavePatentDatabase(ByRef pa() As String, ByVal intWhere As Integer, Optional bolDate As Boolean = True, Optional BolWDB As Boolean = False) As Boolean
Dim strSql As String, i As Integer, strTmp As String, strMsg As String
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
Dim ii As Integer 'Add By Sindy 2014/11/6
Dim strLDate As String, strLtitle As String  'Added by Lydia 2015/09/09

BolTransOk = True
   
   'On Error GoTo ErrHand
   
   If intWhere <> 國外_CF Or bolDate = False Then
      pa(10) = ChangeTStringToWString(pa(10))
      pa(12) = ChangeTStringToWString(pa(12))
      pa(14) = ChangeTStringToWString(pa(14))
      pa(20) = ChangeTStringToWString(pa(20))
      pa(21) = ChangeTStringToWString(pa(21))
      pa(24) = ChangeTStringToWString(pa(24))
      pa(25) = ChangeTStringToWString(pa(25))
      pa(58) = ChangeTStringToWString(pa(58))
      pa(140) = ChangeTStringToWString(pa(140)) 'Added by Morgan 2012/5/9
   End If
   pa(26) = ChangeCustomerL(pa(26))
   pa(27) = ChangeCustomerL(pa(27))
   pa(28) = ChangeCustomerL(pa(28))
   pa(29) = ChangeCustomerL(pa(29))
   pa(30) = ChangeCustomerL(pa(30))
   pa(75) = ChangeCustomerL(pa(75))
   pa(76) = ChangeCustomerL(pa(76))
   pa(86) = ChangeCustomerL(pa(86))
   pa(88) = ChangeCustomerL(pa(88))
   pa(101) = ChangeCustomerL(pa(101))
   pa(105) = ChangeCustomerL(pa(105))
   pa(133) = ChangeCustomerL(pa(133))
   pa(134) = ChangeCustomerL(pa(134))
   'add by Toni 2008/10/17
   pa(150) = Left(Combo2, 1)
   'end 2008/10/17
   'Add By Sindy 2010/10/27
   pa(158) = Left(Combo3, 1)
   '2010/10/27 End
   'pa(166) = ChangeCustomerL(pa(166)) 'Added by Morgan 2014/8/29 'Removed by Morgan 2015/6/11
   strSql = "update patent set "
   For i = 1 To 99
      Select Case i
         'Modified by Lydia 2019/10/04 因為會造成log無法區分,所以Set 更新欄位排除本所案號 (+pa(1)~pa(4)
         Case 92, 93, 94, 95, 96, 97, 1, 2, 3, 4
         Case Else
            strSql = strSql + "pa" + Format(i, "00") + "=" + CNULL(Replace(pa(i), "'", "''")) + ","
      End Select
   Next
   For i = 100 To TF_PA 'edit by nickc 2006/07/12 T_PA
      'Add By Sindy 2025/1/7
      If i = 180 Then
         pa(i) = ""
         For Each obj In Chk1K
            If obj.Value = 1 Then
               pa(i) = pa(i) & "," & obj.Index + 1
            End If
         Next
         If pa(i) <> "" Then pa(i) = Mid(pa(i), 2)
         strSql = strSql + "pa" + Format(i) + "=" + CNULL(pa(i)) + ","
         '2025/1/7 END
      'add by nickc 2006/07/12
      ElseIf i <> 108 And i <> 136 And i <> 137 And i <> 138 Then
    '      If i = 103 Then
             If InStr(pa(i), "'") > 0 Then
                strTmp = Replace(pa(i), "'", "''")
                strSql = strSql + "pa" + Format(i) + "=" + CNULL(strTmp) + ","
             Else
                strSql = strSql + "pa" + Format(i) + "=" + CNULL(pa(i)) + ","
             End If
    '      Else
             ' 91.03.25 modify by louis (單引號)
             'strSQL = strSQL + "pa" + Format(i) + "=" + CNULL(pa(i)) + ","
    '         strSQL = strSQL + "pa" + Format(i) + "=" + CNULL(DBString(pa(i))) + ","
    '      End If
      End If
   Next
   
   strSql = Left(strSql, Len(strSql) - 1)
   strSql = strSql & " where pa01=" + CNULL(pa(1)) + " and pa02=" + CNULL(pa(2)) + " and pa03=" + CNULL(pa(3)) + " and pa04=" + CNULL(pa(4))
   '910910 nick tigger
   '***** start
   'If BolWDB = True Then
   '   StrSql = "begin user_data.user_enabled:=1;  " & StrSql & "; end; "
   'End If
   cnnConnection.BeginTrans
   '***** end
   'Add by Amy 2017/11/28 CFP-29915 EPC案 母案改出名公司別,子案也要更新
   If pa(1) = "CFP" And pa(4) = "00" And pa(9) = "221" And Text1(161).Tag <> Text1(161) Then
      strExc(1) = "Update Patent Set PA161='" & Text1(161) & "' " & _
                  "Where PA01='" & pa(1) & "' And PA02='" & pa(2) & "' And PA04<>'00' "
      cnnConnection.Execute strExc(1)
   End If
   'end 2017/11/28
   'add by nickc 2006/06/07 紀錄分析語法
   'Modified by Lydia 2019/10/04 名稱可能有IN
   'Pub_SeekTbLog strSql
   'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
   'Pub_SeekTbLog strSql, , True
   'Modified by Lydia 2025/10/31 改用模組判斷
   'Pub_SeekTbLog strSql, , True, , Me.Caption & "(" & Me.Name & ")"
   Pub_SeekTbLog strSql, , PUB_FilterSeekSQL(strSql), , Me.Caption & "(" & Me.Name & ")"
   
   If BolWDB = True Then
      strSql = "begin user_data.user_enabled:=1;  " & strSql & "; end; "
   End If
   cnnConnection.Execute strSql
   'Add By Sindy 2014/11/6
   If cmdAddRow.Tag = "I" Or cmdDelRow.Tag = "D" Then '有異動發明人資料
      '全部刪除,重新新增
      strSql = "delete from patentInventor where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4))
      'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
      'Pub_SeekTbLog strSql
      Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
      cnnConnection.Execute strSql, intI
      For ii = 1 To GRD1.Rows - 1
         'Modify By Sindy 2023/3/7 無發明人編號不須新增
         'Modified by Lydia 2024/12/03
         'If GRD1.TextMatrix(ii, 1) <> "" Then
         '自行輸入則客戶發明人檔IN01=PA26
         strTmp = ""
         If Trim(GRD1.TextMatrix(ii, 1)) = "" And _
            (Trim(GRD1.TextMatrix(ii, 2)) <> "" Or Trim(GRD1.TextMatrix(ii, 3)) <> "" Or Trim(GRD1.TextMatrix(ii, 4)) <> "") Then
            '造字後面可能會加空白不可用Trim
            InsInventor strTmp, pa(26), LTrim(GRD1.TextMatrix(ii, 2)), Trim(GRD1.TextMatrix(ii, 3)), LTrim(GRD1.TextMatrix(ii, 4)), Trim(GRD1.TextMatrix(ii, 7))
            GRD1.TextMatrix(ii, 1) = strTmp
         End If
         'end 2024/12/03
         If GRD1.TextMatrix(ii, 1) <> "" Then 'Add By Sindy 2025/7/29 +if,被誤mark了,才導致會出現PI06空白的資料
            strSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                     CNULL(pa(1)) & "," & CNULL(pa(2)) & "," & CNULL(pa(3)) & "," & CNULL(pa(4)) & "," & ii & ",'" & GRD1.TextMatrix(ii, 1) & "')"
            'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
            'Pub_SeekTbLog strSql
            Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
            cnnConnection.Execute strSql
         End If 'Modify By Sindy 2025/7/29 被誤mark了 Mark by Lydia 2024/12/03
      Next ii
      PUB_UpdateRefCaseInventorAndCaseName pa(1), pa(2), pa(3), pa(4) 'Added by Morgan 2020/2/13
   End If
   '2014/11/6 END

   'Added by Lydia 2015/09/09
   If pa(1) = "P" And pa(9) = "020" And pa(8) = "1" And pa(23) = "1" Then
     '大陸發明案，有申請日無公開日無准駁無公告日未閉卷未銷卷也要掛下一程序公開999期限；
     If pa(10) <> "" And Trim(pa(12) & pa(16) & pa(21) & pa(57) & pa(108) & pa(136)) = "" Then
        If PUB_GetOpenLimit020(pa(1), pa(2), pa(3), pa(4), strLDate, strLtitle) Then
          '公開期限的所限(工作天)=法限
          strExc(6) = PUB_GetWorkDay1(strLDate, True)
          strExc(0) = " select * from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='999' "
          intI = 1: strExc(7) = ""
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
             If IsNull(RsTemp.Fields("NP06")) Then
                  '未處理,更新期限
                strSql = "UPDATE NEXTPROGRESS SET NP08=" & strExc(6) & " , NP09=" & strLDate & ",NP15='" & strLtitle & "' WHERE NP01='" & RsTemp.Fields("NP01") & "' AND NP07='999' "
                cnnConnection.Execute strSql, intI
             Else '已處理,另外新增
                strExc(7) = "A"
             End If
          Else
             strExc(7) = "A"
          End If
          '掛下一程序公開999
          If strExc(7) = "A" Then
             '收文號抓最近一天的A類收文
             strExc(0) = "select cp05,cp09 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and substr(cp09,1,1) = '" & strExc(7) & "' and cp57 is null order by 1 desc "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
               strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22)" & _
                  " SELECT '" & RsTemp(1) & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','999'" & _
                  "," & strExc(6) & "," & strLDate & ",'" & strUserNum & "'" & _
                  ",'" & strLtitle & "',NP22 FROM (SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) X"
               cnnConnection.Execute strSql, intI
             End If
          End If
        End If
     End If
       '大陸發明案新增公開日,下一程序999的NP06='Y'
     If pPA12 = "" And pa(12) <> "" Then
        strSql = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP02='" & pa(1) & "' and NP03='" & pa(2) & "' and NP04='" & pa(3) & "' and NP05='" & pa(4) & "' AND NP07='999' and np06 is null "
        cnnConnection.Execute strSql, intI
     End If
   End If
   'end 2015/09/09
   
   'Added by Lydia 2019/11/27 FCP年費特殊管制PA165=N => 目前案件的年費期限自動上不續辦
   strMsg = "": strTmp = ""  'Added by Lydia 2020/03/17
   If Text1(156).Text = "N" And Text1(156).Tag <> Text1(156).Text Then
       'Modified by Lydia 2020/03/17 回傳FMP案範圍，發清單通知程序
       'Call Pub_AutoUpdFCP605(pa(1) & pa(2) & pa(3) & pa(4))
       If Pub_AutoUpdFCP605(pa(1) & pa(2) & pa(3) & pa(4), strTmp, strMsg) = False Then
           GoTo ErrHand
       End If
       'end 2020/03/17
   End If
   'end 2019/11/27
   
   'Added by Lydia 2020/09/14 FCP和FMP案之中間接進來案件，檢查「專利案件」及「English_Vers」進度。
   '109/6/30發出之詢問意見後，僅David回覆，故以David的意見調整程式如下：
   '1.  中間接進來案件於收文時仍自動產生「專利案件」及「English_Vers」進度；
   '2.  程序至專利案件基本檔補輸資料存檔時，程式檢查該案沒有A類收文之新申請案且已有專用期間時，再檢查「專利案件」及「English_Vers」進度的卷宗區及原始檔都沒有任何檔案時，系統自動刪除「專利案件」及「English_Vers」進度。
   If Pub_StrUserSt03 = "F22" And (pa(1) = "P" Or pa(1) = "FCP") And Text1(24) <> "" And Text1(25) <> "" Then
      strExc(0) = ""
      If pa(1) = "P" Then
          If PUB_ChkIsFMP(pa(1), pa(2), pa(3), pa(4)) = True Then strExc(0) = "Y"
      End If
      If pa(1) = "FCP" Or strExc(0) = "Y" Then
         strExc(1) = "select cp01,cp02,cp03,cp04,cp09,cp10 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
                          "and substr(cp09,1,1)='A' and cp159=0 and cp10 in (" & NewCasePtyList & ") "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
         If intI = 0 Then
            strExc(2) = "select cp09,cp10,nvl(count(cpf02),0) cnt1,nvl(count(cpp02),0) cnt2 From caseprogress, casepaperfile,casepaperpdf " & _
                             "where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
                             "and substr(cp09,1,1)='D' and cp10 in (" & cnt專利案件 & "," & cntEnglish_Vers & " ) and cp09=cpf01(+) and cp09=cpp01(+) group by cp09,cp10 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(2))
            If intI = 1 Then
                 RsTemp.MoveFirst
                 Do While Not RsTemp.EOF
                      If Val("" & RsTemp.Fields("cnt1")) + Val("" & RsTemp.Fields("cnt2")) = 0 Then
                           strSql = "delete from caseprogress where cp09='" & RsTemp.Fields("cp09") & "' "
                           'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
                           'Pub_SeekTbLog strSql
                           Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
                           cnnConnection.Execute strSql
                      End If
                      RsTemp.MoveNext
                 Loop
            End If
         End If
      End If
   End If
   'end 2020/09/14
   
   '910910 nick tigger
   '***** start
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
   '***** end
   
    'Added by Lydia 2020/03/17 FMP案件不自動上年費不續辦，改發清單給程序，由各區程序逐筆產生定稿通知大陸代理人
    If strTmp <> "" Then
       If PUB_GetP605Email("1", strTmp, strMsg) = False Then
          If strMsg <> "" Then
              MsgBox strMsg, vbCritical
          End If
       End If
    End If
    'end 2020/03/17
     
   SavePatentDatabase = True
   Exit Function
ErrHand:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
   
   '910910 nick tigger
   '***** start
   cnnConnection.RollbackTrans
   '***** end
   'Modified by Lydia 2020/03/17
   'ShowMsg MsgText(9143) & vbCrLf & vbCrLf & Err.Description
   ShowMsg MsgText(9143) & vbCrLf & vbCrLf & Err.Description & vbCrLf & strMsg
   'ErrorLog
End Function

Private Function DBString(ByVal strData As String) As String
   Dim strOutput As String
   Dim nPos As Integer
   strOutput = Empty
   For nPos = 1 To Len(strData)
      If Mid(strData, nPos, 1) = "'" Then
         strOutput = strOutput & "''"
      Else
         strOutput = strOutput & Mid(strData, nPos, 1)
      End If
   Next nPos
   DBString = strOutput
End Function

'Added by Morgan 2015/6/11
Private Function AddList(pList As ListBox, pText As String) As Boolean
   Dim idx As Integer, bFound As Boolean
   
   For idx = 0 To pList.ListCount - 1
      If Left(pList.List(idx), 9) = Left(pText, 9) Then
         MsgBox "編號存在於清單中！"
         bFound = True
         Exit For
      End If
   Next
   If bFound = False Then
      pList.AddItem pText
      AddList = True
   End If
End Function

Private Sub RemoveList(pList As ListBox)
   Dim idx As Integer, ii As Integer
   If pList.ListCount > 0 Then
      ii = 0
      For idx = 0 To pList.ListCount - 1
         If pList.Selected(ii) = True Then
            pList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Next
   End If
End Sub

'Add By Sindy 2019/4/9
Private Sub txtIN11_Validate(Cancel As Boolean)
Dim strName As String
   
   If txtIN11 = "" Then Exit Sub
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

Private Sub txtNo_Change()
   If Len(txtNo) < 6 Then
      lblName = ""
   End If
End Sub

Private Sub txtNo_GotFocus()
   TextInverse txtNo
End Sub

Private Function ComposeList(pList As ListBox) As String
   Dim stRtn As String, ii As Integer
   If pList.ListCount > 0 Then
      stRtn = Left(pList.List(0), 9)
      For ii = 1 To pList.ListCount - 1
         stRtn = stRtn & "," & Left(pList.List(ii), 9)
      Next
   End If
   ComposeList = stRtn
End Function

Private Sub SetList(pList As ListBox, pText As String)
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

Private Sub txtNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNo_Validate(Cancel As Boolean)
   If Len(txtNo) > 5 Then
      txtNo = Left(txtNo & "000", 9)
      lblName = GetName(txtNo)
   End If
   If txtNo <> "" And lblName = "" Then
      MsgBox "代碼輸入錯誤!!", vbCritical
      Cancel = True
   End If
End Sub

Private Function GetName(pNo As String) As String
   
   If Left(pNo, 1) = "Y" Then
      strExc(0) = "select nvl(fa06, nvl(rtrim(fa05||' '||fa63||' '||fa64||' '||fa65),fa04)) from fagent where fa01='" & Left(pNo, 8) & "' and fa02='" & Mid(pNo, 9) & "'"
   Else
      strExc(0) = "select nvl(cu06, nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu04)) from customer where cu01='" & Left(pNo, 8) & "' and cu02='" & Mid(pNo, 9) & "'"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetName = "" & RsTemp(0)
   End If
End Function

'Added by Lydia 2020/02/21
Private Sub CmdPA174_Click()
    Call ProcPA174toFile("N")
End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟/維護FCP0xxxxx.新案性質.案件名稱.doc
Private Sub ProcPA174toFile(ByVal pKind As String)
Dim strKind As String

   If ActionEdit = 0 Or ActionEdit = 2 Then
        '因為無完整的本所案號，所以不可執行
        MsgBox IIf(ActionEdit = 0, "新增", "查詢") & "時，不可執行!", vbInformation + vbOKOnly, Me.Caption
   Else
        If ChkPA174.Value = vbUnchecked Then
            MsgBox "請先按修改，再勾選「有特殊字」!", vbInformation + vbOKOnly, Me.Caption
        Else
            If ActionEdit = 1 And Me.Text1(1) = "P" Then '修改才問
                If MsgBox("請問是否為FMP案或寰華案？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            End If
            If pKind = "Y" Then 'bolAskPA174
                strKind = "3"
            Else
                strKind = IIf(ActionEdit = 1, "1", "0")
            End If
            If Pub_GetPA174toFile(strKind, Me.Text1(1), Me.Text1(2), Me.Text1(3), Me.Text1(4), Me, frm100101_M_1) = True Then
            End If
        End If
   End If
End Sub

'Added by Lydia 2020/02/21
Public Sub PubShowNextData()
   '原始檔Word檔維護，上傳後直接進入存檔
   If bolAskPA174 = True Then
        RsSitu 3 '確定->存檔
   End If
End Sub
'Added by Morgan 2022/11/30
'檢查FCP案第1道收文程序是否為年費
Private Function IsFCP605Case() As Boolean
   'Modified by Lydia 2024/09/05 +P案；ex.P-134271
   If pa(1) = "FCP" Or pa(1) = "P" Then
      strExc(0) = "select * from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp31='Y' and cp10='605'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         IsFCP605Case = True
      End If
   End If
End Function

'Added by Lydia 2024/12/03
Private Sub txtInvField_GotFocus(Index As Integer)
    
    If Combo1 <> "" Then
        Combo1.SetFocus
    Else
        If txtInvField(0) = "" And txtInvField(1) = "" And txtInvField(2) = "" Then
            txtIN11.Text = ""
            Lb_IN11N.Caption = ""
        End If
    End If
End Sub

'Added by Lydia 2024/12/03 發明人輸入比對兼自動代入(模糊比對)
Private Sub txtInvField_Validate(Index As Integer, Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim tRec As Integer, tSearch As Boolean
   Dim tInx As Integer, tSno As Integer 'tInx =Combo1(index), tSno=List編號

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
         Combo1.ListIndex = tSno + 1 '讀發明人List=>call Combo1_click
         Combo1.SetFocus  '移到比對出的發明人combo List
      End If
   End If
End Sub

'Added by Lydia 2024/12/03 向上移
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

'Added by Lydia 2024/12/03 向下移
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

'Added by Lydia 2024/12/03 發明人輸入比對兼自動代入(模糊比對)
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

'Added by Lydia 2024/12/03
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
         Call SetComboData(.TextMatrix(nRow, 1))
         If .TextMatrix(nRow, 1) = "" Then
            txtInvField(0) = .TextMatrix(nRow, 2)
            txtInvField(1) = .TextMatrix(nRow, 3)
            txtInvField(2) = .TextMatrix(nRow, 4)
            Lb_IN11N = .TextMatrix(nRow, 5)
            txtIN11 = .TextMatrix(nRow, 7)
            cmdUpdRow.Enabled = True
            cmdAddRow.Enabled = False
         End If
      End If
      .Visible = True
   End With
End Sub

'Added by Lydia 2024/12/03 更新發明人
Private Sub InsInventor(ByRef m_PI06, ByVal InvNo As String, ByVal InvCh As String, ByVal InvEng As String, ByVal InvJP As String, ByVal IN11 As String)
   Dim strIns As String, m_IN01 As String, m_IN02 As String
   
   m_IN01 = Left(ChangeCustomerL(InvNo), 8)
   m_IN02 = PUB_GetNewIN02(m_IN01)
   m_PI06 = m_IN01 & m_IN02
   strIns = "Insert Into Inventor (IN01,IN02,IN04,IN05,IN06,IN11) Values(" & CNULL(ChgSQL(m_IN01)) & "," & CNULL(ChgSQL(m_IN02)) & "," & _
               CNULL(ChgSQL(InvCh)) & "," & CNULL(ChgSQL(InvEng)) & "," & CNULL(ChgSQL(InvJP)) & "," & CNULL(ChgSQL(IN11)) & ")"
   cnnConnection.Execute strIns
End Sub
