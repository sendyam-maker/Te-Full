VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075004_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件進度檔案維護"
   ClientHeight    =   6408
   ClientLeft      =   4644
   ClientTop       =   2688
   ClientWidth     =   9144
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6408
   ScaleWidth      =   9144
   Begin VB.CommandButton CmdDot 
      Caption         =   "工作點數分配"
      Height          =   400
      Left            =   8430
      TabIndex        =   258
      Top             =   1095
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint201 
      Caption         =   "翻譯承辦單(&C)"
      Height          =   400
      Left            =   8040
      TabIndex        =   257
      Top             =   1095
      Visible         =   0   'False
      Width           =   1620
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4500
      Left            =   90
      TabIndex        =   94
      Top             =   1560
      Width           =   8985
      _ExtentX        =   15854
      _ExtentY        =   7938
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   423
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm075004_2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label50"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label47"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label31"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label26"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label16(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label28"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label17"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label14(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label20(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label23"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label30(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label32"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label13"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label21"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label35"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label9"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label4"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label18(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label19(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label16(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label15"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label25"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label6(4)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label6(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblNameAgent"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label1(12)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label1(5)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label6(6)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label52"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lbl202CP86"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label14(8)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label6(3)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label6(7)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textCP14_2"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCP13_2"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCP83_2"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textCP44_2"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textCP64"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "lstNameAgent"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textCP29_2"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textCP10_2"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textCP12_2"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textCP58_2"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textCP12"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textCP13"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "textCP14"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "textCP27"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "textCP06"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textCP43"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "textCP08"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "textCP29"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "textCP15"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "textCP21"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "textCP48"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "textCP26"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "textCP31"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "textCP57"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "textCP58"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "textCP45"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "textCP44"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "textCP28"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "textCP23"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "textCP24"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "textCP25"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "textCP07"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "textCP05"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "textCP10"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "textCP82"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "textCP83"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "textCP84"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "textCP113"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "textCP114"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "textCP118"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "textCP119"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "textCP22"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "textCP145"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "textCP152"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).ControlCount=   85
      TabCaption(1)   =   "相關資料"
      TabPicture(1)   =   "frm075004_2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "lblCP71"
      Tab(1).Control(2)=   "Label6(1)"
      Tab(1).Control(3)=   "lblCP49"
      Tab(1).Control(4)=   "Label34"
      Tab(1).Control(5)=   "Label30(3)"
      Tab(1).Control(6)=   "Label30(2)"
      Tab(1).Control(7)=   "Label12(3)"
      Tab(1).Control(8)=   "Label12(1)"
      Tab(1).Control(9)=   "Label12(2)"
      Tab(1).Control(10)=   "lblCP19"
      Tab(1).Control(11)=   "Label37"
      Tab(1).Control(12)=   "Label38"
      Tab(1).Control(13)=   "Label39"
      Tab(1).Control(14)=   "Label14(0)"
      Tab(1).Control(15)=   "Label14(2)"
      Tab(1).Control(16)=   "Label14(3)"
      Tab(1).Control(17)=   "Label1(3)"
      Tab(1).Control(18)=   "Label41"
      Tab(1).Control(19)=   "Label42"
      Tab(1).Control(20)=   "Label43"
      Tab(1).Control(21)=   "Label14(4)"
      Tab(1).Control(22)=   "Label14(5)"
      Tab(1).Control(23)=   "Label20(5)"
      Tab(1).Control(24)=   "Label20(6)"
      Tab(1).Control(25)=   "lblCP137"
      Tab(1).Control(26)=   "lblCP136"
      Tab(1).Control(27)=   "lblCP135"
      Tab(1).Control(28)=   "lblCP138"
      Tab(1).Control(29)=   "Label1(121)"
      Tab(1).Control(30)=   "Label2(2)"
      Tab(1).Control(31)=   "lblCP81"
      Tab(1).Control(32)=   "Label30(1)"
      Tab(1).Control(33)=   "Label22"
      Tab(1).Control(34)=   "Label148"
      Tab(1).Control(35)=   "textCP49"
      Tab(1).Control(36)=   "lblCP168"
      Tab(1).Control(37)=   "lblCP167"
      Tab(1).Control(38)=   "Frame1"
      Tab(1).Control(39)=   "textCP11_2"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "textCP17"
      Tab(1).Control(41)=   "textCP19"
      Tab(1).Control(42)=   "textCP33"
      Tab(1).Control(43)=   "textCP34"
      Tab(1).Control(44)=   "textCP46"
      Tab(1).Control(45)=   "textCP47"
      Tab(1).Control(46)=   "textCP32"
      Tab(1).Control(47)=   "textCP20"
      Tab(1).Control(48)=   "textCP11"
      Tab(1).Control(49)=   "textCP59"
      Tab(1).Control(50)=   "textCP30"
      Tab(1).Control(51)=   "textCP60"
      Tab(1).Control(52)=   "textCP61"
      Tab(1).Control(53)=   "textCP62"
      Tab(1).Control(54)=   "textCP63"
      Tab(1).Control(55)=   "textCP18"
      Tab(1).Control(56)=   "textCP16"
      Tab(1).Control(57)=   "textCP81"
      Tab(1).Control(58)=   "textCP88"
      Tab(1).Control(59)=   "textCP87"
      Tab(1).Control(60)=   "textCP120"
      Tab(1).Control(61)=   "textCP121"
      Tab(1).Control(62)=   "textCP135"
      Tab(1).Control(63)=   "textCP136"
      Tab(1).Control(64)=   "textCP137"
      Tab(1).Control(65)=   "textCP138"
      Tab(1).Control(66)=   "textCP140"
      Tab(1).Control(67)=   "textCP148"
      Tab(1).Control(68)=   "textCP71_2"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "textCP71"
      Tab(1).Control(70)=   "cmdMail"
      Tab(1).Control(71)=   "textCP168"
      Tab(1).Control(72)=   "textCP167"
      Tab(1).Control(73)=   "cmdPage"
      Tab(1).ControlCount=   74
      TabCaption(2)   =   "移轉/授權"
      TabPicture(2)   =   "frm075004_2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20(2)"
      Tab(2).Control(1)=   "Label20(1)"
      Tab(2).Control(2)=   "Label33(0)"
      Tab(2).Control(3)=   "Label20(3)"
      Tab(2).Control(4)=   "Label20(4)"
      Tab(2).Control(5)=   "Label7(7)"
      Tab(2).Control(6)=   "Label7(8)"
      Tab(2).Control(7)=   "Label7(1)"
      Tab(2).Control(8)=   "Label7(2)"
      Tab(2).Control(9)=   "Label7(3)"
      Tab(2).Control(10)=   "Label7(4)"
      Tab(2).Control(11)=   "Label7(6)"
      Tab(2).Control(12)=   "Label7(9)"
      Tab(2).Control(13)=   "Label7(11)"
      Tab(2).Control(14)=   "Label7(10)"
      Tab(2).Control(15)=   "Label20(7)"
      Tab(2).Control(16)=   "textCP55_2"
      Tab(2).Control(17)=   "textCP93_2"
      Tab(2).Control(18)=   "textCP94_2"
      Tab(2).Control(19)=   "textCP95_2"
      Tab(2).Control(20)=   "Line1"
      Tab(2).Control(21)=   "textCP96_2"
      Tab(2).Control(22)=   "textCP56_2"
      Tab(2).Control(23)=   "textCP89_2"
      Tab(2).Control(24)=   "textCP90_2"
      Tab(2).Control(25)=   "textCP91_2"
      Tab(2).Control(26)=   "textCP92_2"
      Tab(2).Control(27)=   "textCP50"
      Tab(2).Control(28)=   "textCP51"
      Tab(2).Control(29)=   "textCP52"
      Tab(2).Control(30)=   "textCP54"
      Tab(2).Control(31)=   "textCP53"
      Tab(2).Control(32)=   "textCP72"
      Tab(2).Control(33)=   "textCP53_2"
      Tab(2).Control(34)=   "textCP54_2"
      Tab(2).Control(35)=   "textCP92"
      Tab(2).Control(36)=   "textCP91"
      Tab(2).Control(37)=   "textCP90"
      Tab(2).Control(38)=   "textCP89"
      Tab(2).Control(39)=   "textCP96"
      Tab(2).Control(40)=   "textCP95"
      Tab(2).Control(41)=   "textCP94"
      Tab(2).Control(42)=   "textCP93"
      Tab(2).Control(43)=   "textCP55"
      Tab(2).Control(44)=   "textCP56"
      Tab(2).ControlCount=   45
      TabCaption(3)   =   "對造/其他"
      TabPicture(3)   =   "frm075004_2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label18(5)"
      Tab(3).Control(1)=   "Label33(1)"
      Tab(3).Control(2)=   "Label18(4)"
      Tab(3).Control(3)=   "Label18(1)"
      Tab(3).Control(4)=   "Label19(3)"
      Tab(3).Control(5)=   "Label18(2)"
      Tab(3).Control(6)=   "Label18(3)"
      Tab(3).Control(7)=   "Label19(1)"
      Tab(3).Control(8)=   "Label19(2)"
      Tab(3).Control(9)=   "Label33(2)"
      Tab(3).Control(10)=   "textCP144"
      Tab(3).Control(11)=   "textCP37"
      Tab(3).Control(12)=   "textCP38"
      Tab(3).Control(13)=   "textCP39"
      Tab(3).Control(14)=   "textCP40"
      Tab(3).Control(15)=   "textCP41"
      Tab(3).Control(16)=   "textCP42"
      Tab(3).Control(17)=   "textCP37_1"
      Tab(3).Control(18)=   "Label51"
      Tab(3).Control(19)=   "Label24"
      Tab(3).Control(20)=   "lblCP86_1"
      Tab(3).Control(21)=   "lblCP86"
      Tab(3).Control(22)=   "textCP80"
      Tab(3).Control(23)=   "textCP36"
      Tab(3).Control(24)=   "textCP117"
      Tab(3).Control(25)=   "textCP35"
      Tab(3).Control(26)=   "textCP86"
      Tab(3).ControlCount=   27
      TabCaption(4)   =   "收據帳目"
      TabPicture(4)   =   "frm075004_2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label3"
      Tab(4).Control(1)=   "Label27"
      Tab(4).Control(2)=   "Label44"
      Tab(4).Control(3)=   "Label45"
      Tab(4).Control(4)=   "Label46"
      Tab(4).Control(5)=   "Label48"
      Tab(4).Control(6)=   "Label49"
      Tab(4).Control(7)=   "lblCP73"
      Tab(4).Control(8)=   "lblCP74"
      Tab(4).Control(9)=   "lblCP75"
      Tab(4).Control(10)=   "lblCP76"
      Tab(4).Control(11)=   "lblCP77"
      Tab(4).Control(12)=   "lblCP78"
      Tab(4).Control(13)=   "lblCP79"
      Tab(4).ControlCount=   14
      TabCaption(5)   =   "發文室"
      TabPicture(5)   =   "frm075004_2.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label14(6)"
      Tab(5).Control(1)=   "Label16(2)"
      Tab(5).Control(2)=   "Label16(3)"
      Tab(5).Control(3)=   "Label16(4)"
      Tab(5).Control(4)=   "Label16(5)"
      Tab(5).Control(5)=   "Label14(7)"
      Tab(5).Control(6)=   "Label16(6)"
      Tab(5).Control(7)=   "Label16(7)"
      Tab(5).Control(8)=   "Label16(8)"
      Tab(5).Control(9)=   "Label16(9)"
      Tab(5).Control(10)=   "textCP131"
      Tab(5).Control(11)=   "textCP123"
      Tab(5).Control(12)=   "textCP124"
      Tab(5).Control(13)=   "textCP127"
      Tab(5).Control(14)=   "textCP125"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "textCP128"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "textCP126"
      Tab(5).Control(17)=   "textCP129"
      Tab(5).Control(18)=   "textCP132"
      Tab(5).Control(19)=   "lstNameOrg"
      Tab(5).ControlCount=   20
      Begin VB.CommandButton cmdPage 
         BackColor       =   &H80000010&
         Caption         =   "增刪頁數"
         Height          =   285
         Left            =   -69930
         Style           =   1  '圖片外觀
         TabIndex        =   290
         Top             =   3570
         Width           =   1065
      End
      Begin VB.TextBox textCP167 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67875
         TabIndex        =   287
         Top             =   3270
         Width           =   420
      End
      Begin VB.TextBox textCP168 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67875
         TabIndex        =   286
         Top             =   3570
         Width           =   420
      End
      Begin VB.TextBox textCP86 
         Height          =   285
         Left            =   -73035
         MaxLength       =   1
         TabIndex        =   110
         Top             =   4080
         Width           =   255
      End
      Begin VB.CommandButton cmdMail 
         Caption         =   "修改費用通知信"
         Height          =   375
         Left            =   -72570
         TabIndex        =   283
         Top             =   390
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox textCP35 
         Height          =   285
         Left            =   -73230
         MaxLength       =   32
         TabIndex        =   108
         Top             =   3750
         Width           =   3585
      End
      Begin VB.TextBox textCP117 
         Height          =   285
         Left            =   -68250
         MaxLength       =   15
         TabIndex        =   109
         Top             =   3750
         Width           =   2025
      End
      Begin VB.TextBox textCP56 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   79
         Top             =   1825
         Width           =   1212
      End
      Begin VB.TextBox textCP55 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   74
         Top             =   360
         Width           =   1212
      End
      Begin VB.TextBox textCP93 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   75
         Top             =   653
         Width           =   1212
      End
      Begin VB.TextBox textCP94 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   76
         Top             =   946
         Width           =   1212
      End
      Begin VB.TextBox textCP95 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   77
         Top             =   1239
         Width           =   1212
      End
      Begin VB.TextBox textCP96 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   78
         Top             =   1532
         Width           =   1212
      End
      Begin VB.TextBox textCP89 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   80
         Top             =   2118
         Width           =   1212
      End
      Begin VB.TextBox textCP90 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   81
         Top             =   2411
         Width           =   1212
      End
      Begin VB.TextBox textCP91 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   82
         Top             =   2704
         Width           =   1212
      End
      Begin VB.TextBox textCP92 
         Height          =   285
         Left            =   -73110
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   83
         Top             =   2997
         Width           =   1212
      End
      Begin VB.TextBox textCP71 
         Height          =   285
         Left            =   -70965
         MaxLength       =   7
         TabIndex        =   51
         Top             =   2073
         Width           =   1092
      End
      Begin VB.TextBox textCP71_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   -69840
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   2073
         Width           =   1815
      End
      Begin VB.TextBox textCP152 
         Height          =   270
         Left            =   7920
         MaxLength       =   7
         TabIndex        =   14
         Top             =   1489
         Width           =   870
      End
      Begin VB.TextBox textCP145 
         Height          =   270
         Left            =   8115
         MaxLength       =   1
         TabIndex        =   23
         Top             =   2392
         Width           =   255
      End
      Begin VB.TextBox textCP148 
         Height          =   285
         Left            =   -66870
         MaxLength       =   1
         TabIndex        =   42
         Top             =   881
         Width           =   255
      End
      Begin VB.TextBox textCP140 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67200
         MaxLength       =   10
         TabIndex        =   48
         Top             =   1477
         Width           =   1100
      End
      Begin VB.TextBox textCP22 
         Height          =   270
         Left            =   6930
         MaxLength       =   1
         TabIndex        =   35
         Top             =   3540
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textCP138 
         Height          =   285
         Left            =   -66525
         TabIndex        =   64
         Top             =   3563
         Width           =   420
      End
      Begin VB.TextBox textCP137 
         Height          =   285
         Left            =   -66525
         TabIndex        =   62
         Top             =   3265
         Width           =   420
      End
      Begin VB.TextBox textCP136 
         Height          =   285
         Left            =   -66525
         TabIndex        =   59
         Top             =   2967
         Width           =   420
      End
      Begin VB.TextBox textCP135 
         Height          =   285
         Left            =   -67875
         TabIndex        =   58
         Top             =   2967
         Width           =   420
      End
      Begin VB.TextBox textCP54_2 
         Height          =   285
         Left            =   -67680
         MaxLength       =   7
         TabIndex        =   88
         Top             =   2997
         Width           =   1095
      End
      Begin VB.TextBox textCP53_2 
         Height          =   285
         Left            =   -69240
         MaxLength       =   7
         TabIndex        =   87
         Top             =   2997
         Width           =   1095
      End
      Begin VB.ListBox lstNameOrg 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         IntegralHeight  =   0   'False
         ItemData        =   "frm075004_2.frx":00A8
         Left            =   -72075
         List            =   "frm075004_2.frx":00B2
         Sorted          =   -1  'True
         Style           =   1  '項目包含核取方塊
         TabIndex        =   125
         Top             =   2145
         Width           =   3150
      End
      Begin VB.TextBox textCP132 
         Height          =   285
         Left            =   -72075
         MaxLength       =   7
         TabIndex        =   127
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox textCP129 
         Height          =   285
         Left            =   -72075
         MaxLength       =   7
         TabIndex        =   124
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox textCP126 
         Height          =   285
         Left            =   -72075
         MaxLength       =   1
         TabIndex        =   122
         Top             =   1140
         Width           =   255
      End
      Begin VB.TextBox textCP128 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -68385
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox textCP125 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -68385
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   810
         Width           =   1095
      End
      Begin VB.TextBox textCP127 
         Height          =   285
         Left            =   -72075
         MaxLength       =   7
         TabIndex        =   123
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox textCP124 
         Height          =   285
         Left            =   -72075
         MaxLength       =   7
         TabIndex        =   121
         Top             =   810
         Width           =   1095
      End
      Begin VB.TextBox textCP123 
         Height          =   285
         Left            =   -72075
         MaxLength       =   1
         TabIndex        =   120
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox textCP121 
         Height          =   285
         Left            =   -66930
         MaxLength       =   1
         TabIndex        =   55
         Top             =   2669
         Width           =   255
      End
      Begin VB.TextBox textCP120 
         Height          =   285
         Left            =   -69810
         MaxLength       =   1
         TabIndex        =   54
         Top             =   2669
         Width           =   255
      End
      Begin VB.TextBox textCP119 
         Height          =   270
         Left            =   3100
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   4
         Top             =   586
         Width           =   735
      End
      Begin VB.TextBox textCP118 
         Height          =   270
         Left            =   7020
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1188
         Width           =   315
      End
      Begin VB.TextBox textCP114 
         Height          =   270
         Left            =   8190
         MaxLength       =   4
         TabIndex        =   20
         Top             =   2091
         Width           =   600
      End
      Begin VB.TextBox textCP113 
         Height          =   270
         Left            =   8190
         MaxLength       =   5
         TabIndex        =   18
         Top             =   1790
         Width           =   600
      End
      Begin VB.TextBox textCP72 
         Height          =   285
         Left            =   -73950
         MaxLength       =   9
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   84
         Top             =   3290
         Width           =   1212
      End
      Begin VB.TextBox textCP87 
         Height          =   285
         Left            =   -70650
         MaxLength       =   15
         TabIndex        =   57
         Top             =   2967
         Width           =   1770
      End
      Begin VB.TextBox textCP88 
         Height          =   285
         Left            =   -70650
         MaxLength       =   15
         TabIndex        =   61
         Top             =   3265
         Width           =   1770
      End
      Begin VB.TextBox textCP84 
         Height          =   270
         Left            =   5985
         TabIndex        =   17
         Top             =   1790
         Width           =   1305
      End
      Begin VB.TextBox textCP83 
         Appearance      =   0  '平面
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   6660
         TabIndex        =   193
         Top             =   902
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.TextBox textCP82 
         Appearance      =   0  '平面
         BorderStyle     =   0  '沒有框線
         Enabled         =   0   'False
         Height          =   255
         Left            =   8145
         MaxLength       =   8
         TabIndex        =   6
         Top             =   638
         Width           =   705
      End
      Begin VB.TextBox textCP36 
         Height          =   285
         Left            =   -73080
         MaxLength       =   200
         TabIndex        =   98
         Top             =   390
         Width           =   2775
      End
      Begin VB.TextBox textCP80 
         Height          =   285
         Left            =   -73080
         MaxLength       =   39
         TabIndex        =   106
         Top             =   2775
         Width           =   6855
      End
      Begin VB.TextBox textCP81 
         Height          =   285
         Left            =   -67590
         MaxLength       =   1
         TabIndex        =   45
         Top             =   1179
         Width           =   255
      End
      Begin VB.TextBox textCP53 
         Height          =   285
         Left            =   -69240
         MaxLength       =   7
         TabIndex        =   85
         Top             =   3290
         Width           =   1095
      End
      Begin VB.TextBox textCP54 
         Height          =   264
         Left            =   -67680
         MaxLength       =   7
         TabIndex        =   86
         Top             =   3300
         Width           =   1095
      End
      Begin VB.TextBox textCP16 
         Height          =   285
         Left            =   -73920
         MaxLength       =   8
         TabIndex        =   36
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox textCP18 
         Height          =   285
         Left            =   -73920
         MaxLength       =   8
         TabIndex        =   38
         Top             =   583
         Width           =   1095
      End
      Begin VB.TextBox textCP63 
         Height          =   285
         Left            =   -73650
         MaxLength       =   15
         TabIndex        =   63
         Top             =   3563
         Width           =   1770
      End
      Begin VB.TextBox textCP62 
         Height          =   285
         Left            =   -73650
         MaxLength       =   15
         TabIndex        =   60
         Top             =   3265
         Width           =   1770
      End
      Begin VB.TextBox textCP61 
         Height          =   285
         Left            =   -73650
         MaxLength       =   15
         TabIndex        =   56
         Top             =   2967
         Width           =   1770
      End
      Begin VB.TextBox textCP60 
         Height          =   285
         Left            =   -73230
         MaxLength       =   15
         TabIndex        =   53
         Top             =   2669
         Width           =   1770
      End
      Begin VB.TextBox textCP30 
         Height          =   285
         Left            =   -70890
         MaxLength       =   20
         TabIndex        =   52
         Top             =   2371
         Width           =   4725
      End
      Begin VB.TextBox textCP59 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73890
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2073
         Width           =   1935
      End
      Begin VB.TextBox textCP11 
         Height          =   285
         Left            =   -73575
         MaxLength       =   2
         TabIndex        =   49
         Top             =   1775
         Width           =   375
      End
      Begin VB.TextBox textCP20 
         Height          =   285
         Left            =   -69885
         MaxLength       =   1
         TabIndex        =   47
         Top             =   1477
         Width           =   375
      End
      Begin VB.TextBox textCP32 
         Height          =   285
         Left            =   -73440
         MaxLength       =   1
         TabIndex        =   46
         Top             =   1477
         Width           =   375
      End
      Begin VB.TextBox textCP47 
         Height          =   285
         Left            =   -69885
         MaxLength       =   7
         TabIndex        =   44
         Top             =   1179
         Width           =   1095
      End
      Begin VB.TextBox textCP46 
         Height          =   285
         Left            =   -72720
         MaxLength       =   7
         TabIndex        =   43
         Top             =   1179
         Width           =   900
      End
      Begin VB.TextBox textCP34 
         Height          =   285
         Left            =   -69885
         MaxLength       =   8
         TabIndex        =   41
         Top             =   881
         Width           =   1095
      End
      Begin VB.TextBox textCP33 
         Height          =   285
         Left            =   -73920
         MaxLength       =   8
         TabIndex        =   40
         Top             =   881
         Width           =   1095
      End
      Begin VB.TextBox textCP19 
         Height          =   285
         Left            =   -69885
         MaxLength       =   8
         TabIndex        =   39
         Top             =   583
         Width           =   1095
      End
      Begin VB.TextBox textCP17 
         Height          =   285
         Left            =   -69885
         MaxLength       =   8
         TabIndex        =   37
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox textCP10 
         Height          =   270
         Left            =   4950
         MaxLength       =   4
         TabIndex        =   2
         Top             =   285
         Width           =   1092
      End
      Begin VB.TextBox textCP05 
         Height          =   270
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   1
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox textCP07 
         Height          =   270
         Left            =   4950
         MaxLength       =   7
         TabIndex        =   5
         Top             =   586
         Width           =   1095
      End
      Begin VB.TextBox textCP25 
         Height          =   270
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   27
         Top             =   2970
         Width           =   1005
      End
      Begin VB.TextBox textCP24 
         Height          =   270
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   24
         Top             =   2670
         Width           =   255
      End
      Begin VB.TextBox textCP23 
         Height          =   285
         Left            =   4350
         MaxLength       =   1
         TabIndex        =   25
         Top             =   2663
         Width           =   260
      End
      Begin VB.TextBox textCP28 
         Height          =   270
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   16
         Top             =   1790
         Width           =   1692
      End
      Begin VB.TextBox textCP44 
         Height          =   270
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   19
         Top             =   2091
         Width           =   1095
      End
      Begin VB.TextBox textCP45 
         Height          =   270
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2392
         Width           =   2895
      End
      Begin VB.TextBox textCP58 
         Height          =   270
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   33
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox textCP57 
         Height          =   270
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   32
         Top             =   3596
         Width           =   1095
      End
      Begin VB.TextBox textCP31 
         Height          =   270
         Left            =   3630
         MaxLength       =   1
         TabIndex        =   28
         Top             =   2970
         Width           =   255
      End
      Begin VB.TextBox textCP26 
         Height          =   270
         Left            =   7875
         MaxLength       =   1
         TabIndex        =   26
         Top             =   2670
         Width           =   255
      End
      Begin VB.TextBox textCP48 
         Height          =   270
         Left            =   4950
         MaxLength       =   7
         TabIndex        =   10
         Top             =   1188
         Width           =   1095
      End
      Begin VB.TextBox textCP21 
         Height          =   270
         Left            =   6240
         MaxLength       =   1
         TabIndex        =   29
         Top             =   2970
         Width           =   255
      End
      Begin VB.TextBox textCP15 
         Height          =   270
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   12
         Top             =   1489
         Width           =   1092
      End
      Begin VB.TextBox textCP29 
         Height          =   270
         Left            =   4950
         MaxLength       =   6
         TabIndex        =   13
         Top             =   1489
         Width           =   1092
      End
      Begin VB.TextBox textCP08 
         Height          =   270
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   30
         Top             =   3300
         Width           =   4770
      End
      Begin VB.TextBox textCP43 
         Height          =   270
         Left            =   5400
         MaxLength       =   9
         TabIndex        =   22
         Top             =   2392
         Width           =   1335
      End
      Begin VB.TextBox textCP06 
         Height          =   270
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   3
         Top             =   586
         Width           =   1092
      End
      Begin VB.TextBox textCP27 
         Height          =   270
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   15
         Top             =   1790
         Width           =   1095
      End
      Begin VB.TextBox textCP14 
         Height          =   270
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1188
         Width           =   1092
      End
      Begin VB.TextBox textCP13 
         Height          =   270
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   7
         Top             =   887
         Width           =   1092
      End
      Begin VB.TextBox textCP12 
         Height          =   270
         Left            =   4950
         MaxLength       =   3
         TabIndex        =   8
         Top             =   887
         Width           =   495
      End
      Begin VB.TextBox textCP11_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   -73170
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   1775
         Width           =   1440
      End
      Begin VB.TextBox textCP58_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   180
         Left            =   4290
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   3645
         Width           =   2610
      End
      Begin VB.TextBox textCP12_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   5460
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   902
         Width           =   1155
      End
      Begin VB.TextBox textCP10_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   6120
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   300
         Width           =   1845
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   615
         Left            =   -70440
         TabIndex        =   243
         Top             =   1710
         Width           =   4380
         Begin VB.CheckBox chkCP176 
            Caption         =   "暫不送"
            Height          =   250
            Left            =   3510
            TabIndex        =   291
            Top             =   120
            Width           =   850
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "指定日期"
            Height          =   180
            Index           =   3
            Left            =   1710
            TabIndex        =   245
            Top             =   135
            Width           =   1035
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "收款後"
            Height          =   180
            Index           =   2
            Left            =   888
            TabIndex        =   246
            Top             =   135
            Width           =   850
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '沒有框線
            Caption         =   "Frame2"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2400
            TabIndex        =   279
            Top             =   390
            Width           =   1965
            Begin VB.OptionButton Option1 
               Caption         =   "當天"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   280
               Top             =   0
               Width           =   675
            End
            Begin VB.OptionButton Option1 
               Caption         =   "之前"
               Height          =   195
               Index           =   1
               Left            =   690
               TabIndex        =   281
               Top             =   0
               Width           =   675
            End
            Begin VB.OptionButton Option1 
               Caption         =   "之後"
               Height          =   195
               Index           =   2
               Left            =   1350
               TabIndex        =   282
               Top             =   0
               Width           =   705
            End
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "不限制"
            Height          =   180
            Index           =   1
            Left            =   45
            TabIndex        =   247
            Top             =   135
            Width           =   850
         End
         Begin VB.TextBox textCP142 
            Height          =   270
            Left            =   2745
            MaxLength       =   7
            TabIndex        =   244
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Label lblCP167 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪未審頁："
         Height          =   180
         Left            =   -68790
         TabIndex        =   289
         Top             =   3315
         Width           =   900
      End
      Begin VB.Label lblCP168 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪已審頁："
         Height          =   180
         Left            =   -68790
         TabIndex        =   288
         Top             =   3615
         Width           =   900
      End
      Begin VB.Label lblCP86 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "收到分所接洽單紀錄："
         Height          =   180
         Left            =   -74880
         TabIndex        =   285
         Top             =   4125
         Width           =   1800
      End
      Begin VB.Label lblCP86_1 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是N:自動收文)"
         Height          =   180
         Left            =   -72675
         TabIndex        =   284
         Top             =   4125
         Width           =   1350
      End
      Begin MSForms.TextBox textCP131 
         Height          =   300
         Left            =   -72075
         TabIndex        =   126
         Top             =   2880
         Width           =   5895
         VariousPropertyBits=   -1467989989
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "10398;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP29_2 
         Height          =   285
         Left            =   6060
         TabIndex        =   277
         TabStop         =   0   'False
         Top             =   1500
         Width           =   855
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "1508;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "審查委員/法院案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   276
         Top             =   3795
         Width           =   1665
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "審查委員編號："
         Height          =   180
         Left            =   -69495
         TabIndex        =   275
         Top             =   3795
         Width           =   1260
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   825
         Left            =   7320
         TabIndex        =   31
         Top             =   3015
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;1455"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   960
         Left            =   -73080
         TabIndex        =   99
         Top             =   717
         Width           =   6855
         VariousPropertyBits=   -1467989989
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "12091;1693"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73080
         TabIndex        =   105
         Top             =   2430
         Width           =   6855
         VariousPropertyBits=   -1467989989
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP41 
         Height          =   300
         Left            =   -73080
         TabIndex        =   104
         Top             =   2085
         Width           =   6855
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73080
         TabIndex        =   103
         Top             =   1743
         Width           =   6855
         VariousPropertyBits=   -1467989989
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP39 
         Height          =   300
         Left            =   -73080
         TabIndex        =   102
         Top             =   1401
         Width           =   6855
         VariousPropertyBits=   -1467989989
         MaxLength       =   160
         ScrollBars      =   2
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP38 
         Height          =   300
         Left            =   -73080
         TabIndex        =   101
         Top             =   1059
         Width           =   6855
         VariousPropertyBits=   -1467989989
         MaxLength       =   250
         ScrollBars      =   2
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37 
         Height          =   300
         Left            =   -73080
         TabIndex        =   100
         Top             =   717
         Width           =   6855
         VariousPropertyBits=   -1467989989
         MaxLength       =   160
         ScrollBars      =   2
         Size            =   "12091;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP144 
         Height          =   570
         Left            =   -73890
         TabIndex        =   107
         Top             =   3120
         Width           =   7665
         VariousPropertyBits=   -1467989989
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "13520;1005"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP52 
         Height          =   285
         Left            =   -73095
         TabIndex        =   91
         Top             =   4140
         Width           =   6975
         VariousPropertyBits=   -1467989989
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12303;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP51 
         Height          =   285
         Left            =   -73095
         TabIndex        =   90
         Top             =   3861
         Width           =   6975
         VariousPropertyBits=   -1467989989
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12303;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP50 
         Height          =   285
         Left            =   -73095
         TabIndex        =   89
         Top             =   3583
         Width           =   6975
         VariousPropertyBits=   -1467989989
         MaxLength       =   60
         ScrollBars      =   2
         Size            =   "12303;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP92_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   274
         TabStop         =   0   'False
         Top             =   3012
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP91_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   273
         TabStop         =   0   'False
         Top             =   2719
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP90_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   272
         TabStop         =   0   'False
         Top             =   2426
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP89_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   271
         TabStop         =   0   'False
         Top             =   2133
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP56_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   270
         TabStop         =   0   'False
         Top             =   1840
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP96_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   269
         TabStop         =   0   'False
         Top             =   1547
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line1 
         X1              =   -68040
         X2              =   -67830
         Y1              =   3420
         Y2              =   3420
      End
      Begin MSForms.TextBox textCP95_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   268
         TabStop         =   0   'False
         Top             =   1254
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP94_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   267
         TabStop         =   0   'False
         Top             =   961
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP93_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   266
         TabStop         =   0   'False
         Top             =   668
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP55_2 
         Height          =   285
         Left            =   -71820
         TabIndex        =   265
         TabStop         =   0   'False
         Top             =   375
         Width           =   5700
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "10054;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP49 
         Height          =   525
         Left            =   -73410
         TabIndex        =   65
         Top             =   3870
         Width           =   7260
         VariousPropertyBits=   -1467989989
         MaxLength       =   249
         ScrollBars      =   2
         Size            =   "12806;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   525
         Left            =   1020
         TabIndex        =   34
         Top             =   3900
         Width           =   7725
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13626;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   285
         Left            =   2250
         TabIndex        =   264
         TabStop         =   0   'False
         Top             =   2085
         Width           =   5070
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "8943;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP83_2 
         Height          =   285
         Left            =   7350
         TabIndex        =   263
         TabStop         =   0   'False
         Top             =   638
         Width           =   855
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "1508;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP13_2 
         Height          =   285
         Left            =   2250
         TabIndex        =   262
         TabStop         =   0   'False
         Top             =   887
         Width           =   1635
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "2884;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   285
         Left            =   2250
         TabIndex        =   261
         TabStop         =   0   'False
         Top             =   1188
         Width           =   1635
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "2884;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "第　　　　期登記期"
         Height          =   180
         Index           =   7
         Left            =   -69450
         TabIndex        =   256
         Top             =   3342
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label6 
         Caption         =   " ( Y:是, W:待確認,    A:自動扣款 )"
         Height          =   390
         Index           =   7
         Left            =   7380
         TabIndex        =   255
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "扣款日期："
         Height          =   180
         Index           =   3
         Left            =   7020
         TabIndex        =   254
         Top             =   1541
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否收到副本：       (Y:是)"
         Height          =   180
         Index           =   8
         Left            =   6840
         TabIndex        =   253
         Top             =   2444
         Width           =   2040
      End
      Begin VB.Label Label148 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "是否為一申請書多件："
         Height          =   180
         Left            =   -68700
         TabIndex        =   252
         Top             =   933
         Width           =   1800
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Left            =   -66570
         TabIndex        =   251
         Top             =   947
         Width           =   465
      End
      Begin VB.Label Label30 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "電子表單單號："
         Height          =   180
         Index           =   1
         Left            =   -68475
         TabIndex        =   249
         Top             =   1529
         Width           =   1260
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "報價備註："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   250
         Top             =   3180
         Width           =   900
      End
      Begin VB.Label lblCP81 
         AutoSize        =   -1  'True
         Caption         =   "是否可減免：       (Y/N)"
         Height          =   180
         Left            =   -68685
         TabIndex        =   181
         Top             =   1231
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(N:不收)"
         Height          =   180
         Index           =   2
         Left            =   -69330
         TabIndex        =   164
         Top             =   1529
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "送件方式："
         Height          =   180
         Index           =   121
         Left            =   -71310
         TabIndex        =   248
         Top             =   1827
         Width           =   900
      End
      Begin VB.Label lbl202CP86 
         BackColor       =   &H00FFFFFF&
         Caption         =   "（複委任）"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7980
         TabIndex        =   242
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lblCP138 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪已審項："
         Height          =   180
         Left            =   -67425
         TabIndex        =   241
         Top             =   3615
         Width           =   900
      End
      Begin VB.Label lblCP135 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "增加頁數："
         Height          =   180
         Left            =   -68790
         TabIndex        =   240
         Top             =   3015
         Width           =   900
      End
      Begin VB.Label lblCP136 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "增加項數："
         Height          =   180
         Left            =   -67425
         TabIndex        =   239
         Top             =   3015
         Width           =   900
      End
      Begin VB.Label lblCP137 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "刪未審項："
         Height          =   180
         Left            =   -67425
         TabIndex        =   238
         Top             =   3315
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "取消發文日："
         Height          =   180
         Index           =   9
         Left            =   -74685
         TabIndex        =   237
         Top             =   3292
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發文室取消發文備註："
         Height          =   180
         Index           =   8
         Left            =   -74685
         TabIndex        =   236
         Top             =   2962
         Width           =   1800
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "主管機關名稱："
         Height          =   180
         Index           =   7
         Left            =   -74685
         TabIndex        =   235
         Top             =   2145
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "分所發文日："
         Height          =   180
         Index           =   6
         Left            =   -74685
         TabIndex        =   234
         Top             =   1845
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否算發文室件數-非主管機關：        (  Y: 是  N:否 空白:未經發文室 )"
         Height          =   180
         Index           =   7
         Left            =   -74685
         TabIndex        =   233
         Top             =   1200
         Width           =   5370
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發文室發文時間-非主管機關："
         Height          =   180
         Index           =   5
         Left            =   -70770
         TabIndex        =   232
         Top             =   1515
         Width           =   2400
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發文室發文時間-主管機關："
         Height          =   180
         Index           =   4
         Left            =   -70770
         TabIndex        =   231
         Top             =   855
         Width           =   2220
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發文室發文日-非主管機關："
         Height          =   180
         Index           =   3
         Left            =   -74685
         TabIndex        =   230
         Top             =   1515
         Width           =   2220
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發文室發文日-主管機關："
         Height          =   180
         Index           =   2
         Left            =   -74685
         TabIndex        =   229
         Top             =   855
         Width           =   2040
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否算發文室件數-主管機關：            ( Y: 是  N:否 空白:未經發文室 )"
         Height          =   180
         Index           =   6
         Left            =   -74685
         TabIndex        =   228
         Top             =   525
         Width           =   5325
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "說明書電子檔已上傳：       (Y:是)"
         Height          =   180
         Index           =   6
         Left            =   -68760
         TabIndex        =   227
         Top             =   2721
         Width           =   2580
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "說明書要電子檔：      (Y:是)"
         Height          =   180
         Index           =   5
         Left            =   -71280
         TabIndex        =   226
         Top             =   2721
         Width           =   2175
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "來函櫃台收文日："
         Height          =   180
         Left            =   2400
         TabIndex        =   225
         Top             =   337
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "電子送件："
         Height          =   180
         Index           =   6
         Left            =   6090
         TabIndex        =   224
         Top             =   1240
         Width           =   900
      End
      Begin VB.Label lblCP79 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   -73725
         TabIndex        =   221
         Top             =   2100
         Width           =   1755
      End
      Begin VB.Label lblCP78 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   -73725
         TabIndex        =   220
         Top             =   1830
         Width           =   1755
      End
      Begin VB.Label lblCP77 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   -73725
         TabIndex        =   219
         Top             =   1545
         Width           =   1755
      End
      Begin VB.Label lblCP76 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   -73725
         TabIndex        =   218
         Top             =   1275
         Width           =   1755
      End
      Begin VB.Label lblCP75 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   -73725
         TabIndex        =   217
         Top             =   1005
         Width           =   1755
      End
      Begin VB.Label lblCP74 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   -73725
         TabIndex        =   216
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label lblCP73 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   180
         Left            =   -73725
         TabIndex        =   215
         Top             =   450
         Width           =   1755
      End
      Begin VB.Label Label49 
         Alignment       =   1  '靠右對齊
         Caption         =   "未收金額："
         Height          =   180
         Left            =   -74850
         TabIndex        =   214
         Top             =   2100
         Width           =   1080
      End
      Begin VB.Label Label48 
         Alignment       =   1  '靠右對齊
         Caption         =   "已退費金額："
         Height          =   180
         Left            =   -74850
         TabIndex        =   213
         Top             =   1830
         Width           =   1080
      End
      Begin VB.Label Label46 
         Alignment       =   1  '靠右對齊
         Caption         =   "已銷帳金額："
         Height          =   180
         Left            =   -74850
         TabIndex        =   212
         Top             =   1545
         Width           =   1080
      End
      Begin VB.Label Label45 
         Alignment       =   1  '靠右對齊
         Caption         =   "已扣繳金額："
         Height          =   180
         Left            =   -74850
         TabIndex        =   211
         Top             =   1275
         Width           =   1080
      End
      Begin VB.Label Label44 
         Alignment       =   1  '靠右對齊
         Caption         =   "已收金額："
         Height          =   180
         Left            =   -74850
         TabIndex        =   210
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label Label27 
         Alignment       =   1  '靠右對齊
         Caption         =   "已收規費："
         Height          =   180
         Left            =   -74850
         TabIndex        =   209
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   1  '靠右對齊
         Caption         =   "已收服務費："
         Height          =   180
         Left            =   -74850
         TabIndex        =   208
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "核稿時數:"
         Height          =   180
         Index           =   5
         Left            =   7380
         TabIndex        =   207
         Top             =   2143
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作時數:"
         Height          =   180
         Index           =   12
         Left            =   7380
         TabIndex        =   206
         Top             =   1842
         Width           =   765
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人："
         Height          =   180
         Left            =   6270
         TabIndex        =   205
         Top             =   3360
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人4："
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   203
         Top             =   2756
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人5："
         Height          =   180
         Index           =   11
         Left            =   -74880
         TabIndex        =   202
         Top             =   3049
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人3："
         Height          =   180
         Index           =   9
         Left            =   -74880
         TabIndex        =   201
         Top             =   2463
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人2："
         Height          =   180
         Index           =   6
         Left            =   -74880
         TabIndex        =   200
         Top             =   2170
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人5："
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   199
         Top             =   1584
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人4："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   198
         Top             =   1291
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人3："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   197
         Top             =   998
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人2："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   196
         Top             =   705
         Width           =   1395
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號4："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   5
         Left            =   -71835
         TabIndex        =   195
         Top             =   3015
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號5："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   4
         Left            =   -71835
         TabIndex        =   194
         Top             =   3315
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Index           =   5
         Left            =   5085
         TabIndex        =   192
         Top             =   1842
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "發文人員時間："
         Height          =   180
         Index           =   4
         Left            =   6120
         TabIndex        =   191
         Top             =   638
         Width           =   1260
      End
      Begin VB.Label Label19 
         Caption         =   "對造名稱(日)："
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   190
         Top             =   2453
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "對造名稱(英)："
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   189
         Top             =   2108
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件名稱(日)："
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   188
         Top             =   1424
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件名稱(英)："
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   187
         Top             =   1082
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "對造名稱(中)："
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   186
         Top             =   1766
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件名稱(中)："
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   185
         Top             =   740
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "對造號數："
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   184
         Top             =   405
         Width           =   1335
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "對造案件商品類別："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   183
         Top             =   2820
         Width           =   1620
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件名稱："
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   182
         Top             =   740
         Width           =   1575
      End
      Begin VB.Label Label43 
         Caption         =   "若有收據或請款編號或帳單編號 ， 只能由電腦中心人員修改!!!"
         ForeColor       =   &H000000FF&
         Height          =   630
         Left            =   -68700
         TabIndex        =   180
         Top             =   285
         Width           =   1980
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "受讓人/移轉申請人1："
         Height          =   180
         Index           =   8
         Left            =   -74880
         TabIndex        =   167
         Top             =   1877
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "讓與人/移轉人1："
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   168
         Top             =   412
         Width           =   1395
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "被授權人(日)："
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   174
         Top             =   4185
         Width           =   1200
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "被授權人(英)："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   173
         Top             =   3906
         Width           =   1200
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "被授權人："
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   171
         Top             =   3342
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "被授權人(中)/收件人："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   170
         Top             =   3628
         Width           =   1785
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "授權期間/質權設定期間(起/迄)/聘任期間："
         Height          =   180
         Index           =   2
         Left            =   -72510
         TabIndex        =   169
         Top             =   3342
         Width           =   3315
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "(N:不開)"
         Height          =   180
         Left            =   -72990
         TabIndex        =   166
         Top             =   1529
         Width           =   645
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "是否開電腦收據："
         Height          =   180
         Left            =   -74880
         TabIndex        =   165
         Top             =   1529
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "是否向客戶收款："
         Height          =   180
         Index           =   3
         Left            =   -71340
         TabIndex        =   163
         Top             =   1529
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號3："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   162
         Top             =   3615
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號2："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   161
         Top             =   3317
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "CF帳單編號1："
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   160
         Top             =   3019
         Width           =   1200
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "收據編號/請款單號："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -74910
         TabIndex        =   159
         Top             =   2721
         Width           =   1665
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "費        用："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -74880
         TabIndex        =   158
         Top             =   337
         Width           =   900
      End
      Begin VB.Label Label37 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "規        費："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -70800
         TabIndex        =   157
         Top             =   337
         Width           =   900
      End
      Begin VB.Label lblCP19 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "後        金："
         Height          =   180
         Left            =   -70800
         TabIndex        =   156
         Top             =   635
         Width           =   900
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "底        價："
         Height          =   180
         Index           =   2
         Left            =   -70800
         TabIndex        =   155
         Top             =   933
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "標  準  價："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   154
         Top             =   933
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "結餘註記："
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   153
         Top             =   2125
         Width           =   900
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "代理人收達日/回執收受日："
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   152
         Top             =   1231
         Width           =   2205
      End
      Begin VB.Label Label30 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "代理人提申日："
         Height          =   180
         Index           =   3
         Left            =   -71160
         TabIndex        =   151
         Top             =   1231
         Width           =   1260
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "大陸申請案號/延展註冊號數/股別/下一程序序號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   150
         Top             =   2423
         Width           =   3915
      End
      Begin VB.Label lblCP49 
         AutoSize        =   -1  'True
         Caption         =   "條款/當事人稱謂："
         Height          =   180
         Left            =   -74880
         TabIndex        =   149
         Top             =   4020
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "案件來源代號："
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   148
         Top             =   1827
         Width           =   1260
      End
      Begin VB.Label lblCP71 
         AutoSize        =   -1  'True
         Caption         =   "機關代號："
         Height          =   180
         Left            =   -71880
         TabIndex        =   147
         Top             =   2125
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "點        數："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   -74880
         TabIndex        =   146
         Top             =   635
         Width           =   900
      End
      Begin VB.Label Label25 
         Caption         =   "(1:准/勝,2:駁/敗,3:部分勝敗)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   7.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1350
         TabIndex        =   145
         Top             =   2715
         Width           =   1900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "(1:准/勝,2:駁/敗,3:部分勝敗)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   7.8
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   160
         Left            =   4620
         TabIndex        =   144
         Top             =   2720
         Width           =   1900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "准駁/勝敗/判決日："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   143
         Top             =   3045
         Width           =   1530
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "實際結果："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   142
         Top             =   2715
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "預估結果："
         Height          =   180
         Index           =   0
         Left            =   3420
         TabIndex        =   141
         Top             =   2720
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "發文字號："
         Height          =   180
         Left            =   2325
         TabIndex        =   140
         Top             =   1842
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "代  理  人："
         Height          =   180
         Left            =   120
         TabIndex        =   139
         Top             =   2143
         Width           =   900
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   180
         Left            =   120
         TabIndex        =   138
         Top             =   2444
         Width           =   900
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "承辦期限："
         Height          =   180
         Left            =   3990
         TabIndex        =   137
         Top             =   1240
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "新案："
         Height          =   180
         Left            =   3030
         TabIndex        =   136
         Top             =   3030
         Width           =   600
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "取消收文原因："
         Height          =   180
         Left            =   2580
         TabIndex        =   135
         Top             =   3645
         Width           =   1260
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Index           =   0
         Left            =   3930
         TabIndex        =   134
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Left            =   6540
         TabIndex        =   133
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "(N:不算)"
         Height          =   180
         Index           =   0
         Left            =   8205
         TabIndex        =   132
         Top             =   2715
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "支援時數 ："
         Height          =   180
         Left            =   120
         TabIndex        =   131
         Top             =   1541
         Width           =   945
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數："
         Height          =   180
         Index           =   1
         Left            =   6600
         TabIndex        =   130
         Top             =   2715
         Width           =   1260
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "是否多國/取締案："
         Height          =   180
         Left            =   4710
         TabIndex        =   129
         Top             =   3000
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "收  文  日："
         Height          =   180
         Left            =   120
         TabIndex        =   128
         Top             =   337
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   180
         Index           =   0
         Left            =   3990
         TabIndex        =   119
         Top             =   337
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "機關文號："
         Height          =   180
         Left            =   480
         TabIndex        =   118
         Top             =   3360
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   117
         Top             =   638
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         Height          =   180
         Index           =   0
         Left            =   3990
         TabIndex        =   116
         Top             =   638
         Width           =   900
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "取消收文日期："
         Height          =   180
         Left            =   120
         TabIndex        =   115
         Top             =   3648
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發  文  日："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   114
         Top             =   1842
         Width           =   900
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號："
         Height          =   180
         Left            =   4080
         TabIndex        =   113
         Top             =   2444
         Width           =   1260
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "進度備註:"
         Height          =   180
         Left            =   120
         TabIndex        =   112
         Top             =   3930
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "承  辦  人："
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   111
         Top             =   1240
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   97
         Top             =   939
         Width           =   900
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "業  務  區："
         Height          =   180
         Left            =   3990
         TabIndex        =   96
         Top             =   939
         Width           =   900
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "繪圖人員/協辦人員："
         Height          =   180
         Left            =   3225
         TabIndex        =   95
         Top             =   1541
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "變更事項(&E)"
      Height          =   405
      Left            =   7710
      TabIndex        =   204
      Top             =   210
      Width           =   1245
   End
   Begin VB.TextBox textCP01 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4710
      MaxLength       =   3
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   603
      Width           =   492
   End
   Begin VB.TextBox textCP02 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5190
      MaxLength       =   6
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   603
      Width           =   732
   End
   Begin VB.TextBox textCP03 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5910
      MaxLength       =   1
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   603
      Width           =   252
   End
   Begin VB.TextBox textCP04 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6150
      MaxLength       =   2
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   603
      Width           =   372
   End
   Begin VB.TextBox textCP09 
      Height          =   300
      Left            =   1110
      MaxLength       =   9
      TabIndex        =   0
      Top             =   603
      Width           =   1572
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Width           =   9144
      _ExtentX        =   16129
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
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6756
      Top             =   408
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
            Picture         =   "frm075004_2.frx":00D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":03EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":0708
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":08E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":0C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":0F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":1238
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":1554
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":1870
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":1B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075004_2.frx":1EA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCP43 
      Caption         =   "相關收文資料(&F)"
      Height          =   400
      Left            =   7470
      TabIndex        =   70
      Top             =   645
      Width           =   1620
   End
   Begin VB.CommandButton cmdPrintCForm 
      Caption         =   "C類接洽記錄單(&C)"
      Height          =   400
      Left            =   7470
      TabIndex        =   179
      Top             =   1095
      Width           =   1620
   End
   Begin MSForms.TextBox textCUID 
      Height          =   285
      Left            =   150
      TabIndex        =   278
      TabStop         =   0   'False
      Top             =   6120
      Width           =   8865
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "15637;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1140
      TabIndex        =   260
      Top             =   1200
      Width           =   6225
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "10980;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   270
      Left            =   1500
      TabIndex        =   259
      TabStop         =   0   'False
      Top             =   930
      Width           =   5865
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "10345;476"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收  文  號： "
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   93
      Top             =   645
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   3630
      TabIndex        =   92
      Top             =   645
      Width           =   900
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      Caption         =   "申請人/當事人："
      Height          =   180
      Left            =   150
      TabIndex        =   73
      Top             =   945
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   72
      Top             =   1252
      Width           =   900
   End
End
Attribute VB_Name = "frm075004_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/08 改成Form2.0 ; cmbTM05、textTM23、textCUID、textCP13_2、textCP14_2、textCP83_2、textCP44_2、textCP64、textCP49、textCP50~CP52
                                                                'textCP55_2、textCP56_2、textCP89_2~CP96_2、textCP144、lstNameAgent、textCP29_2、textCP131
                                                                '「審查委員/法院案號」、「審查委員編號」從基本資料頁籤移到對造頁籤
'end 2021/10/08
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

' 本所案號
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
' 申請國家
Dim m_Nation As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
'edit by nick 2004/08/18 已不用常數
'Dim m_FieldList(T_CP) As FIELDITEM
Dim m_FieldList() As FIELDITEM

' 變數宣告區
Dim m_EditMode As Integer  'Memo by Lydia 2020/04/10  1-新增, 2-修改
Dim m_SubMode As Integer

' 儲存單筆記錄的結構
Private Type DATAITEM
   diCP09 As String
End Type
' 記錄所有資料的串列
Dim m_DataList() As DATAITEM
' 記錄所有可瀏覽資料的總筆數
Dim m_DataListCount As Integer

' 原專用期間
Dim m_CP53 As String
Dim m_CP54 As String

' 目前正在作用的資料項目索引
Dim m_CurrDL As Integer

Dim m_AddData As Boolean

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

'Add by Morgan 2004/2/12
'控制是否恢復下一程序管制期限
Dim m_CP43 As String
Dim m_CP10 As String
'add by nickc 2006/01/27
Dim m_CP110 As String
Dim m_CP130 As String 'Add By Sindy 2009/05/11
Dim m_cp109 As String  'Added by Lydia 2021/11/29 可結餘日期
'Public m_PrevFormNm As String 'Add By Sindy 2013/10/25 前一畫面
Dim m_PrevForm As Form '前一畫面 'Add By Sindy 2018/10/9
Dim bolUpdCP82 As Boolean 'Add By Sindy 2014/3/25
Dim m_SK02 As Integer 'Added by Lydia 2016/05/30
Dim m_CP65 As String 'Added by Morgan 2016/6/2
Dim m_DelLD01 As String, m_DelLD04 As String, m_DelLD10 As String   'Addded by Lydia 2018/07/03 要刪除收款寄證的定稿建檔人,定稿收文號,定稿別
Dim m_CP85 As String 'Added by Lydia 2019/06/28 承辦人發文日/FCP定稿日期
'Added by Lydia 2021/01/14 法律所案源收文
Dim m_LOS01 As String '案源總收文號
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '案源總收文號之本所案號
Dim m_LOS02 As String '案源案件類型
Dim m_MeTrackMode  As String 'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序
Dim m_CP162 As String 'Added by Lydia 2023/08/14 (案件進度)案源單號


'Add By Sindy 2018/10/9
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCP43_Click()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nRecordCount As Integer

   Select Case m_EditMode
      Case 0:
         If IsEmptyText(textCP43) = False Then
            strSql = "SELECT * FROM CASEPROGRESS " & _
                     "WHERE CP01 = '" & m_CP01 & "' AND " & _
                           "CP02 = '" & m_CP02 & "' AND " & _
                           "CP03 = '" & m_CP03 & "' AND " & _
                           "CP04 = '" & m_CP04 & "' AND " & _
                           "CP09 = '" & textCP43 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            nRecordCount = rsTmp.RecordCount
            rsTmp.Close
            Set rsTmp = Nothing
            
            If nRecordCount <= 0 Then
               strTit = "檢核資料"
               strMsg = "相關收文資料與本案不符或是不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Else
               SetDataListItem textCP43
               ShowCurrRecord textCP43
            End If
         End If
      Case Else:
   End Select
End Sub

'Add By Sindy 2022/12/5 修改費用通知信
Private Sub cmdMail_Click()
Dim objOutLook As Object
Dim objMail As Object
'Dim myForward As Object
'Dim jj As Integer
'Dim ArrStr As Variant
Dim strContent As String, strSubject As String

   '產生outlook草稿
   Set objOutLook = CreateObject("Outlook.Application")
   Set objMail = objOutLook.CreateItem(0)
   
   '主旨
   'Modify By Sindy 2025/3/24 +◎
   strSubject = "◎請修改 " & GetMailSubject & " 費用"
   '內容
   'Modify By Sindy 2025/3/24 +內文加入原點數：   修改為：
   strContent = GetMailContent & _
                convForm("原費用： " & Val(textCP16) & " 元", 17) & "  修改為：  元" & vbCrLf & _
                convForm("原規費： " & Val(textCP17) & " 元", 17) & "  修改為：  元" & vbCrLf & _
                convForm("原點數： " & textCP18, 17) & "  修改為：  " & vbCrLf & vbCrLf & _
                "修改原因： " & vbCrLf
   
   '轉HTML格式
   strContent = Replace(strContent, "新細明體", "Times New Roman")
   '&nbsp; 不換行空格
   '&thinsp; 窄空格
   '單純只是想要輸入空白？ &nbsp; 就對了
   '&emsp; 全形空格
   '&ensp; 半形空格
   'strContent = Replace(strContent, "　", "&emsp;") '&emsp; 全形空格
   strContent = Replace(strContent, " ", "&thinsp;") '&ensp; 半形空格
   strContent = Replace(strContent, vbCrLf, "<BR>")
'      If TypeName(objOutLook.Assistant) <> "Nothing" Then
'         objOutLook.ActiveWindow.WindowState = 1 '0.最大化 1.視窗小點
'      End If
   With objMail
      '.BodyFormat = 2 '2=olFormatHTML 1=olFormatPlain 3=olFormatRichText
      'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個
      If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
          .To = Pub_GetSpecMan("財務處應收處理人員")
      Else
         .To = Pub_GetSpecMan("財務處總帳人員")
      End If
      .cc = textCP13.Text '智權人員
'      '加入附件
'      If m_strContactSheetA4 <> "" Then
'         ArrStr = Split(m_strContactSheetA4, ";")
'         For jj = 0 To UBound(ArrStr)
'            If Dir(ArrStr(jj)) <> "" Then
'               .Attachments.add ArrStr(jj) '加附件
'            End If
'         Next jj
'      End If
      .Subject = strSubject
      .HTMLBody = strContent
      .Display
   End With

   Set objMail = Nothing
   Set objOutLook = Nothing
   '*** END
End Sub
Private Function GetMailSubject() As String
   GetMailSubject = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & " 案(" & textCP10_2 & ")「" & cmbTM05.Text & "」"
End Function
Private Function GetMailContent() As String
   GetMailContent = "本所案號： " & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & vbCrLf & _
                "案件名稱： " & cmbTM05.Text & vbCrLf & _
                "總收文號： " & textCP09.Text & vbCrLf & _
                "收 文 日： " & ChangeTStringToTDateString(TransDate(textCP05, 1)) & vbCrLf & _
                "案件性質： " & textCP10 & textCP10_2 & vbCrLf & vbCrLf
End Function

'add by nick 2004/10/14 變更事項
Private Sub cmdok_Click()
        Screen.MousePointer = vbHourglass
        frm050706.Show
        frm050706.Hide
        frm050706.IsCall = True
        frm050706.IsMod = True
        frm050706.textCE01.Text = textCP09.Text
        If frm050706.QueryRecord = True Then
            Me.Hide
            frm050706.Show
            Set frm050706.frmParent = Me 'Add by Morgan 2007/5/22
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            Exit Sub
        'add by nickc 2005/06/09
        Else
            MsgBox "無變更事項！", vbCritical, "警告！"
        End If
        Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2023/4/13 增刪頁數
Private Sub cmdPage_Click()
   If m_EditMode = 1 Or m_EditMode = 2 Then '新增,修改
      frm075004_3.bolModify = True
   Else
      frm075004_3.bolModify = False
   End If
   frm075004_3.strReceiveNo = textCP09.Text
   Set frm075004_3.objAddPage = textCP135
   Set frm075004_3.objCP167 = textCP167
   Set frm075004_3.objCP168 = textCP168
   Call frm075004_3.SetParent(Me)
   frm075004_3.Show vbModal
End Sub

Private Sub cmdPrintCForm_Click()
'Added by Lydia 2015/07/17
Dim pPA(1 To 4) As String
Dim pPA75 As String, pPA26 As String, strKind As String, bolReCall As Boolean
Dim m_strMemo As String
'Added by Lydia 2019/03/06 逐筆判斷Y代理人+X申請人1~5
Dim bolTmp As Boolean
Dim pCust(1 To 4) As String 'PA27~PA30
Dim bolGrant As Boolean 'Added by Lydia 2019/07/30

'Add By Cheng 2002/01/28
'列印C類接洽記錄單
'92.1.28 MODIFY BY SONIA
'2006/9/11 MODIFY BY SONIA 加入 FG
'If Me.textCP01.Text = "FCP" Then
'Modify by Morgan 2010/6/18 +FMP案
If (Me.textCP01.Text = "FCP" Or Me.textCP01.Text = "FG" Or (textCP01 = "P" And Left(textCP12, 1) = "F")) Then
   'Added by Lydia 2015/07/17 讓核對已准專利926的備註代入列印
   'g_PrtForm001.PrintCForm Me.textCP09.Text
   'If Me.textCP01.Text = "FCP" And Me.textCP10.Text = "926" Then  'Mark by Lydia 2022/08/02 不限制
      pPA(1) = Me.textCP01.Text: pPA(2) = Me.textCP02.Text: pPA(3) = Me.textCP03.Text: pPA(4) = Me.textCP04.Text
      'If PUB_CheckAuto926(pPA) = True Then  'Mark by Lydia 2022/08/02 不限制
            m_strMemo = "": bolReCall = False
            'Modified by Lydia 2019/03/06 +pa27,pa28,pa29,pa30
            'Modified by Lydia 2019/07/30 +pa08
            'Modified by Lydia 2022/08/02
            'strSql = "SELECT pa16,pa57,pa163,pa75,pa26,c2.cp10,pa27,pa28,pa29,pa30,pa08 FROM patent,caseprogress c1,caseprogress c2 " & _
                     "WHERE pa01='" & pPA(1) & "' and pa02='" & pPA(2) & "' and pa03='" & pPA(3) & "' and pa04='" & pPA(4) & "' " & _
                     "and pa01=c1.cp01(+) and pa02=c1.cp02(+) and pa03=c1.cp03(+) and pa04=c1.cp04(+) and c1.cp09='" & Me.textCP43.Text & "' " & _
                     "and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c1.cp43=c2.cp09(+)"
            strSql = "SELECT pa16,pa57,pa163,pa75,pa26,pa27,pa28,pa29,pa30,pa08,c1.cp09 as c1cp09, c1.cp10 as c1cp10,c2.cp09 as c2cp09,c2.cp10 as c2cp10 " & _
                     "FROM patent,caseprogress c1,caseprogress c2 WHERE pa01='" & pPA(1) & "' and pa02='" & pPA(2) & "' and pa03='" & pPA(3) & "' and pa04='" & pPA(4) & "' " & _
                     "and pa01=c1.cp01(+) and pa02=c1.cp02(+) and pa03=c1.cp03(+) and pa04=c1.cp04(+) and c1.cp09='" & Me.textCP43.Text & "' " & _
                     "and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c1.cp43=c2.cp09(+)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               pPA26 = RsTemp.Fields("pa26"): pPA75 = RsTemp.Fields("pa75")
               'Added by Lydia 2019/03/06
               pCust(1) = "" & RsTemp.Fields("pa27")
               pCust(2) = "" & RsTemp.Fields("pa28")
               pCust(3) = "" & RsTemp.Fields("pa29")
               pCust(4) = "" & RsTemp.Fields("pa30")
               'end 2019/03/06
               'Modified by Lydia 2022/08/02
               'strKind = "" & RsTemp.Fields("cp10")
               If Left("" & RsTemp.Fields("c1cp09"), 1) < "C" Then
                  strKind = "" & RsTemp.Fields("c1cp10")
               Else
                  strKind = "" & RsTemp.Fields("c2cp10")
               End If
               'end 2022/08/02
               If ((strKind >= "101" And strKind <= "105") Or strKind = "107" Or strKind = "125" Or (strKind >= "301" And strKind <= "308")) Then
                  bolReCall = True
               End If
            End If
            'If bolReCall = True Then 'Mark by Lydia 2022/08/02 不限制
                '初審核准=>核准管制分割期限
                'Modified by Lydia 2019/07/30 因108.11.1修法分割管制期限設定
                 '1. 於108.8.1收到之核准函：
                 '　1.1. 發明初審核准：維持原有設定之分割期限
                 '　1.2. 發明再審核准、新型核准：原有設定分割期限之客戶編號，增加控管行事曆期限，原則照初審核准，期限為收到核准函後３個月期限，並帶備註至通知告准之進度備註。
                 '2. 於108.10.1收到之核准函：發明初審核准、發明再審核准、新型核准：皆設定收到核准函後３個月期限。
                'If "" & RsTemp.Fields("pa16") = "1" And (strKind = "101" Or (strKind = "307" And "" & RsTemp.Fields("pa163") = "Y")) And Len("" & RsTemp.Fields("pa57")) = 0 Then
                bolGrant = False
                If "" & RsTemp.Fields("pa16") = "1" And (strKind = "101" Or (strKind = "307" And "" & RsTemp.Fields("pa163") = "Y")) And Len("" & RsTemp.Fields("pa57")) = 0 Then
                    bolGrant = True
                ElseIf DBDATE(textCP05) >= "20190801" And (strKind = "102" Or (strKind = "107" And "" & RsTemp.Fields("pa08") = "1")) Then
                    bolGrant = True
                End If
                If bolGrant = True Then
                'end 2019/07/30
                   'Modified by Lydia 2019/03/06 逐筆判斷Y代理人+X申請人1~5
                   'm_strMemo = PUB_GetApprMemo(pPA(1) & pPA(2) & pPA(3) & pPA(4), "1001", pPA75, pPA26, "4")
                   'Modified by Lydia 2022/08/02 整合模組：修改一般備註、核對已准備註為複數新規則
                   'm_strMemo = PUB_GetApprMemo(pPA(1) & pPA(2) & pPA(3) & pPA(4), "1001", pPA75, pPA26, "4", bolTmp)
                   'If bolTmp = False Then '不存在個案備註
                   '   strExc(1) = "": strExc(2) = ""
                   '   For intI = 1 To 4
                   '      If pCust(intI) <> "" Then
                   '          strExc(1) = PUB_GetApprMemo(pPA(1) & pPA(2) & pPA(3) & pPA(4), "1001", pPA75, pCust(intI), "4", bolTmp)
                   '          If strExc(1) <> "" And (m_strMemo = "" Or (m_strMemo <> "" And InStr(m_strMemo, strExc(1)) = 0)) Then
                   '              m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
                   '          End If
                   '      End If
                   '   Next intI
                   'End If
                   'end 2019/03/06
                   strExc(1) = PUB_GetApprMemo2("4", pPA(1) & pPA(2) & pPA(3) & pPA(4), "1001", pPA75, pPA26 & "," & pCust(1) & "," & pCust(2) & "," & pCust(3) & "," & pCust(4))
                   If strExc(1) <> "" And InStr(m_strMemo & ",", strExc(1)) = 0 Then
                        m_strMemo = m_strMemo & strExc(1) & vbCrLf
                   End If
                   'end 2022/08/02
                End If
                '一般核准
                'Modified by Lydia 2019/03/06 逐筆判斷Y代理人+X申請人1~5
                'm_strMemo = IIf(Len(m_strMemo) > 0, m_strMemo & vbCrLf, "") & PUB_GetApprMemo(pPA(1) & pPA(2) & pPA(3) & pPA(4), "1001", pPA75, pPA26, "1")
                'Modified by Lydia 2022/08/02 整合模組：修改一般備註、核對已准備註為複數新規則
                'm_strMemo = IIf(Len(m_strMemo) > 0, m_strMemo & vbCrLf, "") & PUB_GetApprMemo(pPA(1) & pPA(2) & pPA(3) & pPA(4), "1001", pPA75, pPA26, "1", bolTmp)
                'If bolTmp = False Then '不存在個案備註
                '      strExc(1) = "": strExc(2) = ""
                '      For intI = 1 To 4
                '         If pCust(intI) <> "" Then
                '             strExc(1) = PUB_GetApprMemo(pPA(1) & pPA(2) & pPA(3) & pPA(4), "1001", pPA75, pCust(intI), "1", bolTmp)
                '             If strExc(1) <> "" And (m_strMemo = "" Or (m_strMemo <> "" And InStr(m_strMemo, strExc(1)) = 0)) Then
                '                 m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
                '             End If
                '         End If
                '      Next intI
                'End If
                ''end 2019/03/06
                strExc(1) = PUB_GetApprMemo2("1", pPA(1) & pPA(2) & pPA(3) & pPA(4), "1001", pPA75, pPA26 & "," & pCust(1) & "," & pCust(2) & "," & pCust(3) & "," & pCust(4))
                If strExc(1) <> "" And InStr(m_strMemo & ",", strExc(1)) = 0 Then
                  m_strMemo = m_strMemo & strExc(1) & vbCrLf
                End If
                'end 2022/08/02
                '核對已准
                'Modified by Lydia 2019/03/06 逐筆判斷Y代理人+X申請人1~5
                'm_strMemo = IIf(Len(m_strMemo) > 0, m_strMemo & vbCrLf, "") & PUB_GetApprMemo(pPA(1) & pPA(2) & pPA(3) & pPA(4), "926,1001", pPA75, pPA26, "2")
                'Modified by Lydia 2022/08/02 整合模組：修改一般備註、核對已准備註為複數新規則
                'm_strMemo = IIf(Len(m_strMemo) > 0, m_strMemo & vbCrLf, "") & PUB_GetApprMemo(pPA(1) & pPA(2) & pPA(3) & pPA(4), "926,1001", pPA75, pPA26, "2", bolTmp)
                'If bolTmp = False Then '不存在個案備註
                '      strExc(1) = "": strExc(2) = ""
                '      For intI = 1 To 4
                '         If pCust(intI) <> "" Then
                '             strExc(1) = PUB_GetApprMemo(pPA(1) & pPA(2) & pPA(3) & pPA(4), "926,1001", pPA75, pCust(intI), "2", bolTmp)
                '             If strExc(1) <> "" And (m_strMemo = "" Or (m_strMemo <> "" And InStr(m_strMemo, strExc(1)) = 0)) Then
                '                 m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
                '             End If
                '         End If
                '      Next intI
                'End If
                ''end 2019/03/06
                If textCP10 = "926" Then
                   '因為前面已抓一般核准, 所以限定傳入案件性質
                   strExc(1) = PUB_GetApprMemo2("2", pPA(1) & pPA(2) & pPA(3) & pPA(4), "926", pPA75, pPA26 & "," & pCust(1) & "," & pCust(2) & "," & pCust(3) & "," & pCust(4))
                   If strExc(1) <> "" And InStr(m_strMemo & ",", strExc(1)) = 0 Then
                     m_strMemo = m_strMemo & strExc(1) & vbCrLf
                   End If
                End If
                'end 2022/08/02
            'End If 'Mark by Lydia 2022/08/02 不限制
      'End If  'Mark by Lydia 2022/08/02 不限制
   'End If 'Mark by Lydia 2022/08/02 不限制
   'end 2015/07/17
   
   'Modified by Lydia 2019/03/18 改成開啟Word
   'g_PrtForm001.PrintCForm Me.textCP09.Text, m_strMemo
   g_PrtForm001.PrintCFormNew Me.textCP09.Text, m_strMemo
Else
   g_PrtForm001.PrintCFForm Me.textCP09.Text
End If
End Sub

Private Sub Form_Initialize()
   'add by nick 2004/08/18 重新定義
   ReDim m_FieldList(TF_CP)
End Sub

'add by nickc 2006/11/10 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
   
   Select Case KeyAscii
      Case 13:
         'Remove by Lydia 2021/11/22 取消以ENTER控制為換行的功能 (Form2.0修改之維護資料功能Toolbar之修改統一)
         'If m_EditMode <> 0 Then
         '   Call PUB_SaveMeTrackMode(m_MeTrackMode, 1, vbKeyF9)  'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序
         '   KeyAscii = 0
         '   OnAction vbKeyF9
         'End If
         'end 2021/11/22
   End Select
End Sub

'Added by Lydia 2021/10/20
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Call PUB_SaveMeTrackMode(m_MeTrackMode, 1, KeyCode)  'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序
    
'Memo by Lydia 2021/10/20 從Form_KeyDown搬來
Select Case KeyCode
      ' 90.07.13 modify by louis
      ' 新增
      'Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
      '   If m_EditMode = 0 Then
      '      OnAction KeyCode
      '      KeyCode = 0
      '   End If
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
'edit by nickc 2006/11/10
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

' Load Form
Private Sub Form_Load()

   SSTab1.Tab = 0
   
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm075004_2", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm075004_2", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm075004_2", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm075004_2", strFind, False)
   
   'Added by Lydia 2018/04/11 外專翻譯承辦單列印
   cmdPrint201.Left = cmdPrintCForm.Left
   'Added by Lydia 2020/04/10 法務工作點數分配
   CmdDot.Left = cmdPrintCForm.Left
   CmdDot.Visible = False
   'Mark by Lydia 2020/04/20 先隱藏
   'Remove Mark by Lydia 2021/08/18 法務系統的工作點數分配功能先上線
   If InStr(m_CP01, "L") > 0 Then
       CmdDot.Visible = True
   End If
   'end 2020/04/20
   'end 2020/04/10
   
   textTM23.BackColor = &H8000000F
   textCP01.BackColor = &H8000000F
   textCP02.BackColor = &H8000000F
   textCP03.BackColor = &H8000000F
   textCP04.BackColor = &H8000000F
   textCP10_2.BackColor = &H8000000F
   textCP11_2.BackColor = &H8000000F
   textCP12_2.BackColor = &H8000000F
   textCP13_2.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP29_2.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   textCP55_2.BackColor = &H8000000F
   textCP56_2.BackColor = &H8000000F
   textCP58_2.BackColor = &H8000000F
   textCP71_2.BackColor = &H8000000F
   textCUID.BackColor = &H8000000F
   'add by nick 2004/08/18
   textCP82.BackColor = &H8000000F
   textCP83_2.BackColor = &H8000000F
   textCP93_2.BackColor = &H8000000F
   textCP94_2.BackColor = &H8000000F
   textCP95_2.BackColor = &H8000000F
   textCP96_2.BackColor = &H8000000F
   textCP89_2.BackColor = &H8000000F
   textCP90_2.BackColor = &H8000000F
   textCP91_2.BackColor = &H8000000F
   textCP92_2.BackColor = &H8000000F
   
   'Add by Morgan 2009/3/18
   lblCP73.BackColor = &H8000000F
   lblCP74.BackColor = &H8000000F
   lblCP75.BackColor = &H8000000F
   lblCP76.BackColor = &H8000000F
   lblCP77.BackColor = &H8000000F
   lblCP78.BackColor = &H8000000F
   lblCP79.BackColor = &H8000000F
   lbl202CP86.BackColor = &H8000000F 'Add by Morgan 2010/5/13
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   'Added by Lydia 2021/10/08 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 825
   lstNameAgent.Width = 1300
   
   'Added by Lydia 2021/11/09 區分送件方式CP141和指定送件日期方式CP164
   Frame1.BackColor = &H8000000F
   Frame2.BackColor = &H8000000F
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Add By Cheng 2002/03/06
'避免程式無法完全釋放
tlbar_ButtonClick Me.tlbar.Buttons(14) 'OnAction會跑2次
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2017/11/29
   ClearDataList
   ClearFieldList
   Set m_PrevForm = Nothing 'Add By Sindy 2018/10/9
   Set frm075004_2 = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' 設定資料
Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_CP01 = Empty
      m_CP02 = Empty
      m_CP03 = Empty
      m_CP04 = Empty
      ClearDataList
      m_AddData = False
   End If
   
   Select Case nType
      ' 本所案號
      Case 0: m_CP01 = strData
      Case 1: m_CP02 = strData
      Case 2: m_CP03 = strData
      Case 3: m_CP04 = strData
      Case 4: SetDataListItem strData
      Case 99: textCP43 = strData
   End Select
End Sub

' 刪除資料串列
Private Sub ClearDataList()
   If m_DataListCount > 0 Then
      Erase m_DataList
   End If
   m_DataListCount = 0
End Sub

' 設定資料串列
Private Sub SetDataListItem(ByVal strData As String)
   Dim nIndex As Integer
   Dim bFind As Boolean

   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex).diCP09 = strData Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_DataList(m_DataListCount + 1)
      m_DataList(m_DataListCount).diCP09 = strData
      m_DataListCount = m_DataListCount + 1
   End If
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   'edit by nick 2004/08/18
   'For nIndex = 1 To T_CP
   For nIndex = 1 To TF_CP
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CP" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      'edit by nick 2004/08/18 定義數字
      Select Case nIndex
         'edit by nick 2004/08/18
         'Case 5, 6, 7, 15, 16, 17, 18, 19, 25, 27, 33, 34, 46, 47, 48, 53, 54, 57, 66, 67, 69, 70:
         'Modify by Morgan 2007/7/19
         'Case 5, 6, 7, 15, 16, 17, 18, 19, 25, 27, 33, 34, 46, 47, 48, 53, 54, 57, 66, 67, 69, 70, 73, 74, 75, 76, 77, 78, 79, 82, 84, 85:
         'Modify by Morgan 2009/3/18 +124,125,127,128,129
         'Modify by Morgan 2010/1/5 +135~138
         Case 5, 6, 7, 15, 16, 17, 18, 19, 25, 27, 33, 34, 46, 47, 48, 53, 54, 57, 66, 67, 69, 70, 73, 74, 75, 76, 77, 78, 79, 82, 84, 85, 113, 114, 124, 125, 127, 128, 129, 135, 136, 137, 138
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

Private Sub ClearFieldList()
   Erase m_FieldList
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   'edit by nick 2004/08/18
   'For nIndex = 0 To T_CP - 1
   For nIndex = 0 To TF_CP - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   Dim strCP09 As String
   
   'add by nickc 2006/01/27
   Dim ii As Integer
   Dim bolCheck As Boolean
   bolCheck = False
   
   '出名代理人
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ITEMDATA(ii)
         'Modified by Lydia 2021/10/08 改模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   
   'Add By Sindy 2009/05/11
   '發文主管機關名稱
   m_CP130 = ""
   For ii = 0 To lstNameOrg.ListCount - 1
      If lstNameOrg.Selected(ii) = True Then
         m_CP130 = m_CP130 & "," & lstNameOrg.List(ii)
      End If
   Next
   If Left(m_CP130, 1) = "," Then m_CP130 = Mid(m_CP130, 2)
   '2009/05/11 End
   
   '2008/8/5 ADD BY SONIA 未發文或申請國家非台灣案者不存CP22
   If IsEmptyText(textCP27) = True Then bolCheck = True
   If m_Nation <> "000" Then bolCheck = True
   '2008/8/5 END

   'Modified by Morgan 2024/1/8
   'If bolCheck = True Then
   If bolCheck = True Or textCP09 > "C" Then
      Me.textCP22 = ""
   Else
      textCP22 = "N"
   End If
   
   ' 本所案號
   SetFieldNewData "CP01", m_CP01
   SetFieldNewData "CP02", m_CP02
   SetFieldNewData "CP03", m_CP03
   SetFieldNewData "CP04", m_CP04
   ' 收文日
   If IsEmptyText(textCP05) = False Then
      SetFieldNewData "CP05", DBDATE(textCP05)
   Else
      SetFieldNewData "CP05", textCP05
   End If
   ' 本所期限
   If IsEmptyText(textCP06) = False Then
      SetFieldNewData "CP06", DBDATE(textCP06)
   Else
      SetFieldNewData "CP06", textCP06
   End If
   ' 法定期限
   If IsEmptyText(textCP07) = False Then
      SetFieldNewData "CP07", DBDATE(textCP07)
   Else
      SetFieldNewData "CP07", textCP07
   End If
   ' 機關文號
   SetFieldNewData "CP08", textCP08
   ' 總收文號
   SetFieldNewData "CP09", textCP09
   ' 案件性質
   SetFieldNewData "CP10", textCP10
   SetFieldNewData "CP11", textCP11
   SetFieldNewData "CP12", textCP12
   SetFieldNewData "CP13", textCP13
   SetFieldNewData "CP14", textCP14
   SetFieldNewData "CP15", textCP15
    'Modify By Cheng 2002/11/19
    '若CP60, CP61, Cp62, CP63,cp87,cp88有值時, CP16,CP17,CP18此三欄位只有電腦中心人員才可修改
    If m_EditMode = 2 Then
        'Modify By Cheng 2002/11/29
        '若無收據及帳單編號
        '2007/3/2 modify by sonia 加入cp87,cp88
        If "" & m_FieldList(59).fiOldData = "" And "" & m_FieldList(60).fiOldData = "" And "" & m_FieldList(61).fiOldData = "" And "" & m_FieldList(62).fiOldData = "" And "" & m_FieldList(86).fiOldData = "" And "" & m_FieldList(87).fiOldData = "" Then
            SetFieldNewData "CP16", textCP16
            SetFieldNewData "CP17", textCP17
            SetFieldNewData "CP18", textCP18
            SetFieldNewData "CP60", textCP60
            SetFieldNewData "CP61", textCP61
            SetFieldNewData "CP62", textCP62
            SetFieldNewData "CP63", textCP63
            SetFieldNewData "CP87", textCP87
            SetFieldNewData "CP88", textCP88
        '若有收據或帳單編號
        ElseIf Pub_StrUserSt03 = "M51" Then
            SetFieldNewData "CP16", textCP16
            SetFieldNewData "CP17", textCP17
            SetFieldNewData "CP18", textCP18
            SetFieldNewData "CP60", textCP60
            SetFieldNewData "CP61", textCP61
            SetFieldNewData "CP62", textCP62
            SetFieldNewData "CP63", textCP63
            SetFieldNewData "CP87", textCP87
            SetFieldNewData "CP88", textCP88
        End If
    Else
       SetFieldNewData "CP16", textCP16
       SetFieldNewData "CP17", textCP17
       SetFieldNewData "CP18", textCP18
       SetFieldNewData "CP60", textCP60
       SetFieldNewData "CP61", textCP61
       SetFieldNewData "CP62", textCP62
       SetFieldNewData "CP63", textCP63
       SetFieldNewData "CP87", textCP87
       SetFieldNewData "CP88", textCP88
    End If
   SetFieldNewData "CP19", textCP19
   SetFieldNewData "CP20", textCP20
   SetFieldNewData "CP21", textCP21
   SetFieldNewData "CP22", textCP22
   SetFieldNewData "CP23", textCP23
   SetFieldNewData "CP24", textCP24
   ' 准駁日
   If IsEmptyText(textCP25) = False Then
      SetFieldNewData "CP25", DBDATE(textCP25)
   Else
      SetFieldNewData "CP25", textCP25
   End If
   SetFieldNewData "CP26", textCP26
   ' 發文日
   If IsEmptyText(textCP27) = False Then
      SetFieldNewData "CP27", DBDATE(textCP27)
   Else
      SetFieldNewData "CP27", textCP27
   End If
   SetFieldNewData "CP28", textCP28
   SetFieldNewData "CP29", textCP29
   SetFieldNewData "CP30", textCP30
   SetFieldNewData "CP31", textCP31
   SetFieldNewData "CP32", textCP32
   SetFieldNewData "CP33", textCP33
   SetFieldNewData "CP34", textCP34
   SetFieldNewData "CP35", textCP35
   SetFieldNewData "CP36", textCP36
    Select Case Me.textCP01.Text
    Case "T", "FCT", "CFT", "TF"
        SetFieldNewData "CP37", textCP37_1
    Case Else
        SetFieldNewData "CP37", textCP37
        SetFieldNewData "CP38", textCP38
        SetFieldNewData "CP39", textCP39
    End Select
   SetFieldNewData "CP40", textCP40
   SetFieldNewData "CP41", textCP41
   SetFieldNewData "CP42", textCP42
   SetFieldNewData "CP43", textCP43
   ' 代理人
   If IsEmptyText(textCP44) = False Then
      'Modify by Morgan 2008/5/14 +聯絡人CP116
      intI = InStr(textCP44, "-")
      If intI > 0 Then
         SetFieldNewData "CP44", ChangeCustomerL(Left(textCP44, intI - 1))
         SetFieldNewData "CP116", Format(Mid(textCP44, intI + 1), "00")
      Else
         SetFieldNewData "CP44", ChangeCustomerL(textCP44)
         SetFieldNewData "CP116", Empty
      End If
   Else
      SetFieldNewData "CP44", textCP44
      SetFieldNewData "CP116", Empty
   End If
   SetFieldNewData "CP45", textCP45
   ' 代理人收達日
   If IsEmptyText(textCP46) = False Then
      SetFieldNewData "CP46", DBDATE(textCP46)
   Else
      SetFieldNewData "CP46", textCP46
   End If
   ' 代理人提申日
   If IsEmptyText(textCP47) = False Then
      SetFieldNewData "CP47", DBDATE(textCP47)
   Else
      SetFieldNewData "CP47", textCP47
   End If
   ' 承辦期限
   If IsEmptyText(textCP48) = False Then
      SetFieldNewData "CP48", DBDATE(textCP48)
   Else
        'Modify By Cheng 2002/11/12
'      SetFieldNewData "CP48", textCP48
      SetFieldNewData "CP48", DBDATE(textCP48)
   End If
   SetFieldNewData "CP49", textCP49
   SetFieldNewData "CP50", textCP50
   SetFieldNewData "CP51", textCP51
   SetFieldNewData "CP52", textCP52
   
   'Modify By Sindy 2009/07/06
   If textCP53_2.Visible = True And textCP54_2.Visible = True Then
      ' 繳費年度/次數(起)
      If IsEmptyText(textCP53_2) = False Then
         SetFieldNewData "CP53", DBDATE(textCP53_2)
      Else
         SetFieldNewData "CP53", textCP53_2
      End If
      ' 繳費年度/次數(迄)
      If IsEmptyText(textCP54_2) = False Then
         SetFieldNewData "CP54", DBDATE(textCP54_2)
      Else
         SetFieldNewData "CP54", textCP54_2
      End If
   '2009/07/06 End
   'Added by Lydia 2017/08/24 TB條碼案繳年費708,服務業務結果1801-第?期登記期
   ElseIf m_CP01 = "TB" And (m_CP10 = "708" Or m_CP10 = "1801") Then
      SetFieldNewData "CP53", textCP53
      SetFieldNewData "CP54", textCP54
   'end 2017/08/24
   Else
      ' 授權期間
      If IsEmptyText(textCP53) = False Then
         SetFieldNewData "CP53", DBDATE(textCP53)
      Else
         SetFieldNewData "CP53", textCP53
      End If
      ' 授權期間
      If IsEmptyText(textCP54) = False Then
         SetFieldNewData "CP54", DBDATE(textCP54)
      Else
         SetFieldNewData "CP54", textCP54
      End If
   End If
   
   ' 移轉人
   If IsEmptyText(textCP55) = False Then
      SetFieldNewData "CP55", textCP55 & String(9 - Len(textCP55), "0")
   Else
      SetFieldNewData "CP55", textCP55
   End If
   ' 移轉申請人
   If IsEmptyText(textCP56) = False Then
      SetFieldNewData "CP56", textCP56 & String(9 - Len(textCP56), "0")
   Else
      SetFieldNewData "CP56", textCP56
   End If
   ' 取消收文日期
   If IsEmptyText(textCP57) = False Then
      SetFieldNewData "CP57", DBDATE(textCP57)
   Else
      SetFieldNewData "CP57", textCP57
   End If
   SetFieldNewData "CP58", textCP58
   SetFieldNewData "CP59", textCP59
    'Modify By Cheng 2003/07/31
'   SetFieldNewData "CP60", textCP60
'   SetFieldNewData "CP61", textCP61
'   SetFieldNewData "CP62", textCP62
'   SetFieldNewData "CP63", textCP63
   SetFieldNewData "CP64", textCP64
   'add by sonia 2013/6/26 T-184682異議案對方減縮商品但我們未撤回異議案,後收到異議敗訴函,因不計敗訴故人工改來函性質為其他來函,加註進度備註
   If m_EditMode = 2 And InStr(m_CP01, "T") And m_CP10 = "1004" And textCP10 = "1706" Then
      If MsgBox("是否為對方減縮商品異議之不成立案，是否要在進度備註加註 因對方減縮商品而異議不成立？", vbYesNo) = vbYes Then
         SetFieldNewData "CP64", "對方減縮商品而異議不成立,因不計敗訴故改為其他來函；" & textCP64
      End If
   End If
   '2013/6/26 end
   
   SetFieldNewData "CP71", textCP71
   ' 被授權人代號
   If IsEmptyText(textCP72) = False Then
      SetFieldNewData "CP72", textCP72 & String(9 - Len(textCP72), "0")
   Else
      SetFieldNewData "CP72", textCP72
   End If
   
   SetFieldNewData "CP80", textCP80
   
   SetFieldNewData "CP81", textCP81 'Add by Morgan 2004/6/11
   'add by nick 2004/08/18
   'SetFieldNewData "CP82", textCP82
   SetFieldNewData "CP84", textCP84
   'SetFieldNewData "CP85", textCP85 'Remove by Morgan 2010/12/30 目前沒用
   SetFieldNewData "CP86", textCP86
   SetFieldNewData "CP87", textCP87
   SetFieldNewData "CP88", textCP88
   If IsEmptyText(textCP89) = False Then
      SetFieldNewData "CP89", textCP89 & String(9 - Len(textCP89), "0")
   Else
      SetFieldNewData "CP89", textCP89
   End If
   If IsEmptyText(textCP90) = False Then
      SetFieldNewData "CP90", textCP90 & String(9 - Len(textCP90), "0")
   Else
      SetFieldNewData "CP90", textCP90
   End If
   If IsEmptyText(textCP91) = False Then
      SetFieldNewData "CP91", textCP91 & String(9 - Len(textCP91), "0")
   Else
      SetFieldNewData "CP91", textCP91
   End If
   If IsEmptyText(textCP92) = False Then
      SetFieldNewData "CP92", textCP92 & String(9 - Len(textCP92), "0")
   Else
      SetFieldNewData "CP92", textCP92
   End If
   If IsEmptyText(textCP93) = False Then
      SetFieldNewData "CP93", textCP93 & String(9 - Len(textCP93), "0")
   Else
      SetFieldNewData "CP93", textCP93
   End If
   If IsEmptyText(textCP94) = False Then
      SetFieldNewData "CP94", textCP94 & String(9 - Len(textCP94), "0")
   Else
      SetFieldNewData "CP94", textCP94
   End If
   If IsEmptyText(textCP95) = False Then
      SetFieldNewData "CP95", textCP95 & String(9 - Len(textCP95), "0")
   Else
      SetFieldNewData "CP95", textCP95
   End If
   If IsEmptyText(textCP96) = False Then
      SetFieldNewData "CP96", textCP96 & String(9 - Len(textCP96), "0")
   Else
      SetFieldNewData "CP96", textCP96
   End If
   'add by nickc 2006/01/27
   SetFieldNewData "CP110", m_CP110
   'Add By Sindy 2009/05/11
   SetFieldNewData "CP130", m_CP130
   
   'Add by Morgan 2007/7/19
   SetFieldNewData "CP113", textCP113
   SetFieldNewData "CP114", textCP114
   'Add by Morgan 2008/5/14
   SetFieldNewData "CP117", textCP117
   'Add by Morgan 2008/7/11
   SetFieldNewData "CP118", textCP118
   '2008/8/27 add by sonia 櫃台收文日
   If IsEmptyText(textCP119) = False Then
      SetFieldNewData "CP119", DBDATE(textCP119)
   Else
      SetFieldNewData "CP119", textCP119
   End If
   '2008/8/27 end
   'Add by Morgan 2008/11/10
   SetFieldNewData "CP120", textCP120
   SetFieldNewData "CP121", textCP121
   SetFieldNewData "CP145", textCP145 'Add By Morgan 2016/6/2
   SetFieldNewData "CP148", textCP148 'Add By Sindy 2015/6/3
   
   'Add By Morgan 2016/6/8
   If IsEmptyText(textCP152) = False Then
      SetFieldNewData "CP152", DBDATE(textCP152)
   Else
      SetFieldNewData "CP152", textCP152
   End If
   'end 2016/6/8
   
   'Add by Morgan 2009/3/18
   '是否經發文室-主管機關
   SetFieldNewData "CP123", textCP123
   '發文室發文日-主管機關
   If textCP124 = Empty Then
      SetFieldNewData "CP124", textCP124
   Else
      SetFieldNewData "CP124", DBDATE(textCP124)
   End If
   '是否經發文室-非主管機關
   SetFieldNewData "CP126", textCP126
   '發文室發文日-非主管機關
   If textCP127 = Empty Then
      SetFieldNewData "CP127", textCP127
   Else
      SetFieldNewData "CP127", DBDATE(textCP127)
   End If
   '分所發文日
   If textCP129 = Empty Then
      SetFieldNewData "CP129", textCP129
   Else
      SetFieldNewData "CP129", DBDATE(textCP129)
   End If
   'end 2009/3/18
   
   'Add by Sindy 2009/05/04
   '發文室取消發文備註
   SetFieldNewData "CP131", textCP131
   '取消發文日
   SetFieldNewData "CP132", DBDATE(textCP132)
   '2009/05/04 End
   
   'Add by Morgan 2010/1/5
   SetFieldNewData "CP135", textCP135
   SetFieldNewData "CP136", textCP136
   SetFieldNewData "CP137", textCP137
   SetFieldNewData "CP138", textCP138
   'end 2010/1/5
   'Add By Sindy 2023/4/13
   SetFieldNewData "CP167", textCP167
   SetFieldNewData "CP168", textCP168
   '2023/4/13 END
   
   SetFieldNewData "CP144", textCP144  '2011/5/26 add by sonia
   
   ' 若為新增資料, 則新產生一筆B類的總收文號
   If m_EditMode = 1 Then
      strCP09 = AutoNo("B", 6)
      SetFieldNewData "CP09", strCP09
   End If
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   'edit by nick 2004/08/18
   'For nIndex = 0 To T_CP - 1
   For nIndex = 0 To TF_CP - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            'add by nickc 2006/03/10
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Public Sub QueryDB()
   InitialField
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   textCP01 = m_CP01
   textCP02 = m_CP02
   textCP03 = m_CP03
   textCP04 = m_CP04
   textCP05 = Empty
   textCP06 = Empty
   textCP07 = Empty
   textCP08 = Empty
   textCP09 = Empty
   textCP10 = Empty
   textCP10_2 = Empty
   textCP11 = Empty
   textCP11_2 = Empty
   textCP12 = Empty
   textCP12_2 = Empty
   textCP13 = Empty
   textCP13_2 = Empty
   textCP14 = Empty
   'Add By Cheng 2003/10/29
   '記錄原承辦人欄位
   Me.textCP14.Tag = Empty
   'ENd
   textCP14_2 = Empty
   textCP15 = Empty
   textCP16 = Empty
   textCP16.Tag = Empty 'Add By Sindy 2022/12/5
   textCP17 = Empty
   textCP17.Tag = Empty 'Add By Sindy 2022/12/5
   textCP18 = Empty
   textCP18.Tag = Empty 'Added By Lydia 2022/12/21
   textCP19 = Empty
   textCP20 = Empty
   textCP21 = Empty
   textCP22 = Empty
   textCP23 = Empty
   textCP24 = Empty
   textCP25 = Empty
   textCP26 = Empty
   textCP27 = Empty
   textCP28 = Empty
   textCP29 = Empty
   textCP29_2 = Empty
   textCP30 = Empty
   textCP31 = Empty
   textCP32 = Empty
   textCP33 = Empty
   textCP34 = Empty
   textCP35 = Empty
   textCP36 = Empty
   textCP37 = Empty
   textCP37_1 = Empty
   textCP38 = Empty
   textCP39 = Empty
   textCP40 = Empty
   textCP41 = Empty
   textCP42 = Empty
   textCP43 = Empty
   textCP44 = Empty
   textCP44_2 = Empty
   textCP45 = Empty
   textCP46 = Empty
   textCP47 = Empty
   textCP48 = Empty
   textCP49 = Empty
   textCP50 = Empty
   textCP51 = Empty
   textCP52 = Empty
   textCP53 = Empty
   textCP54 = Empty
   'Add By Sindy 2009/07/06
   textCP53_2 = Empty
   textCP54_2 = Empty
   '2009/07/06 End
   textCP55 = Empty
   textCP55_2 = Empty
   textCP56 = Empty
   textCP56_2 = Empty
   textCP57 = Empty
   textCP58 = Empty
   textCP58_2 = Empty
   textCP59 = Empty
   textCP60 = Empty
   textCP61 = Empty
   textCP62 = Empty
   textCP63 = Empty
   textCP64 = Empty
   textCP71 = Empty
   textCP72 = Empty
   textCP71_2 = Empty
   textCUID = Empty
   textCP80 = Empty
   textCP81 = Empty
   'add by nick 2004/08/18 加欄位
   textCP82 = Empty
   textCP83 = Empty
   textCP83_2 = Empty
   textCP84 = Empty
   'textCP85 = Empty 'Remove by Morgan 2010/12/30 目前沒用
   textCP86 = Empty
   textCP87 = Empty
   textCP88 = Empty
   textCP89 = Empty
   textCP89_2 = Empty
   textCP90 = Empty
   textCP90_2 = Empty
   textCP91 = Empty
   textCP91_2 = Empty
   textCP92 = Empty
   textCP92_2 = Empty
   textCP93 = Empty
   textCP93_2 = Empty
   textCP94 = Empty
   textCP94_2 = Empty
   textCP95 = Empty
   textCP95_2 = Empty
   textCP96 = Empty
   textCP96_2 = Empty
   'Add by Morgan 2007/7/19
   textCP113 = Empty
   textCP114 = Empty
   'Add by Morgan 2008/5/14
   textCP117 = Empty
   'Add by Morgan 2008/7/11
   textCP118 = Empty
   '2008/8/27 add by sonia
   textCP119 = Empty
   'Add by Morgan 2008/11/10
   textCP120 = Empty
   textCP121 = Empty
   textCP145 = Empty 'Add By Morgan 2016/6/2
   textCP148 = Empty 'Add By Sindy 2015/6/3
   textCP152 = Empty 'Add By Morgan 2016/6/8
   
   'Add by Morgan 2009/3/18
   textCP123 = Empty
   textCP124 = Empty
   textCP125 = Empty
   textCP126 = Empty
   textCP127 = Empty
   textCP128 = Empty
   textCP129 = Empty
   
   'Add By Sindy 2009/04/27
   textCP131 = Empty
   textCP132 = Empty
   
   'Add by Morgan 2010/1/5
   textCP135 = Empty
   textCP136 = Empty
   textCP137 = Empty
   textCP138 = Empty
   'Add By Sindy 2023/4/13
   textCP167 = Empty
   textCP168 = Empty
   '2023/4/13 END
   
   'Add by Morgan 2010/12/30
   textCP140 = Empty
   OptSendType(1).Value = False
   OptSendType(2).Value = False
   OptSendType(3).Value = False
   textCP142 = Empty: Option1(0).Value = False: Option1(1).Value = False: Option1(2).Value = False
   textCP144 = Empty    '2011/5/26 add by sonia
   
   'add by nickc 2006/01/27
   Me.lstNameAgent.Clear
   'Add By Sindy 2009/05/11
   Me.lstNameOrg.Clear
   
   'add by nickc 2008/01/31
   lblCP73.Caption = Empty
   lblCP74.Caption = Empty
   lblCP75.Caption = Empty
   lblCP76.Caption = Empty
   lblCP77.Caption = Empty
   lblCP78.Caption = Empty
   lblCP79.Caption = Empty
   
   'edit by nick 2004/08/18
   'For nIndex = 0 To T_CP - 1
   For nIndex = 0 To TF_CP - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2014/3/25
Dim intEEP02 As Integer 'Add By Sindy 2014/3/25
   
   textCP05.Locked = bEnable
   textCP06.Locked = bEnable
   textCP07.Locked = bEnable
   textCP08.Locked = bEnable
   textCP09.Locked = bEnable
   textCP10.Locked = bEnable
   textCP11.Locked = bEnable
   textCP12.Locked = bEnable
   textCP13.Locked = bEnable
   textCP14.Locked = bEnable
   textCP15.Locked = bEnable
   textCP16.Locked = bEnable
   textCP17.Locked = bEnable
   textCP18.Locked = bEnable
   textCP60.Locked = bEnable
   textCP61.Locked = bEnable
   textCP62.Locked = bEnable
   textCP63.Locked = bEnable
   'add by nick 2004/08/18
   textCP87.Locked = bEnable
   textCP88.Locked = bEnable
   textCP82.Locked = True
   textCP83.Locked = True
   textCP83_2.Locked = True
   textCP82.Locked = bEnable
   textCP84.Locked = bEnable
   'textCP85.Locked = bEnable 'Remove by Morgan 2010/12/30 目前沒用
   textCP86.Locked = bEnable
   textCP89.Locked = bEnable
   textCP90.Locked = bEnable
   textCP91.Locked = bEnable
   textCP92.Locked = bEnable
   textCP93.Locked = bEnable
   textCP94.Locked = bEnable
   textCP95.Locked = bEnable
   textCP96.Locked = bEnable
   cmdMail.Visible = False
    'Add By Cheng 2002/11/26
    If m_EditMode = 2 Then
        '若有收據請款或帳單編號且使用者非電腦中心人員, 則鎖住費用, 規費, 點數三欄位
        '2007/3/2 modify by sonia 加入cp87,cp88
        If (m_FieldList(59).fiOldData <> "" Or m_FieldList(60).fiOldData <> "" Or m_FieldList(61).fiOldData <> "" Or _
            m_FieldList(62).fiOldData <> "" Or m_FieldList(86).fiOldData <> "" Or m_FieldList(87).fiOldData <> "") Then
            If Pub_StrUserSt03 <> "M51" Then
               textCP16.Locked = True
               textCP17.Locked = True
               textCP18.Locked = True
               textCP60.Locked = True
               textCP61.Locked = True
               textCP62.Locked = True
               textCP63.Locked = True
               'add by nick 2004/08/18
               textCP87.Locked = True
               textCP88.Locked = True
               '2010/10/19 ADD BY SONIA 非請款單加入CP12,CP13
               If "" & Left(m_FieldList(59).fiOldData, 1) = "E" Then
                  textCP12.Locked = True
                  textCP13.Locked = True
               End If
               '2010/10/19 END
            End If
            'Add By Sindy 2022/12/5
            If strSrvDate(1) >= 接洽單電子收文啟用日 And Len(textCP140) = 10 And Left(Trim(m_FieldList(59).fiOldData), 1) = "E" Then
               cmdMail.Visible = True
            End If
            '2022/12/5 END
        End If
    End If
   textCP19.Locked = bEnable
   textCP20.Locked = bEnable
   textCP21.Locked = bEnable
   textCP22.Locked = bEnable
   textCP23.Locked = bEnable
   textCP24.Locked = bEnable
   textCP25.Locked = bEnable
   textCP26.Locked = bEnable
   textCP27.Locked = bEnable
   textCP28.Locked = bEnable
   textCP29.Locked = bEnable
   textCP30.Locked = bEnable
   textCP31.Locked = bEnable
   textCP32.Locked = bEnable
   textCP33.Locked = bEnable
   textCP34.Locked = bEnable
   textCP35.Locked = bEnable
   textCP36.Locked = bEnable
   textCP37.Locked = bEnable
   textCP37_1.Locked = bEnable
   textCP38.Locked = bEnable
   textCP39.Locked = bEnable
   textCP40.Locked = bEnable
   textCP41.Locked = bEnable
   textCP42.Locked = bEnable
   textCP43.Locked = bEnable
   textCP44.Locked = bEnable
   textCP45.Locked = bEnable
   textCP46.Locked = bEnable
   textCP47.Locked = bEnable
   textCP48.Locked = bEnable
   'Modify by Amy 2014/09/16 承辦期限設灰
   textCP48.BackColor = &H80000005
   'Add by Morgan 2010/9/28 承辦期限不可修改
   If bolNewPromoterRule Then
      If (textCP01 = "P" And Left(textCP12, 1) <> "F") Or textCP01 = "CFP" Then
         textCP48.Locked = True
         textCP48.BackColor = &H8000000F
      End If
   End If
   'end 2010/9/28
   'end 2014/09/16
   textCP49.Locked = bEnable
   textCP50.Locked = bEnable
   textCP51.Locked = bEnable
   textCP52.Locked = bEnable
   textCP53.Locked = bEnable
   textCP54.Locked = bEnable
   'Add By Sindy 2009/07/06
   textCP53_2.Locked = bEnable
   textCP54_2.Locked = bEnable
   '2009/07/06 End
   textCP55.Locked = bEnable
   textCP56.Locked = bEnable
   textCP57.Locked = bEnable
   textCP58.Locked = bEnable
'edit by nickc 2006/10/30
'   textCP59.Locked = bEnable
   textCP64.Locked = bEnable
   textCP71.Locked = bEnable
   textCP72.Locked = bEnable
   textCP80.Locked = bEnable
   textCP81.Locked = bEnable
   'add by nickc 2006/01/27
   'edit by nickc 2006/02/06 因為會有多個，所以不能鎖
   'Me.lstNameAgent.Enabled = Not bEnable
   'Add by Morgan 2008/5/14
   textCP117.Locked = bEnable
   'Add by Morgan 2008/7/11
   textCP118.Locked = bEnable
   '2008/8/27 add by sonia
   textCP119.Locked = bEnable
   'Add by Morgan 2008/11/10
   textCP120.Locked = bEnable
   textCP121.Locked = bEnable
   textCP145.Locked = bEnable 'Add By Morgan 2016/6/2
   textCP148.Locked = bEnable 'Add By Sindy 2015/6/3
   textCP152.Locked = bEnable 'Add By Morgan 2016/6/8
   
   'Add by Morgan 2009/3/18
   textCP123.Locked = bEnable
   textCP124.Locked = bEnable
   textCP125.Locked = True
   textCP126.Locked = bEnable
   textCP127.Locked = bEnable
   textCP128.Locked = True
   textCP129.Locked = bEnable
   
   'Add By Sindy 2009/04/27
   textCP131.Locked = bEnable
   textCP132.Locked = bEnable
      
   'Add by Morgan 2010/1/5 一般只能看
   'Modify by Morgan 2010/4/13 開放頁數項數能改
   textCP135.Locked = bEnable
   textCP136.Locked = bEnable
   '取消限制，申請書或發文時可能輸錯，實務上會需要修改--淑華
   'If Pub_StrUserSt03 = "M51" Then
   '   textCP137.Locked = bEnable
   '   textCP138.Locked = bEnable
   'Else
   '   textCP137.Locked = True
   '   textCP138.Locked = True
   'End If
      textCP137.Locked = bEnable
      textCP138.Locked = bEnable
      'Add By Sindy 2023/4/13
      textCP167.Locked = bEnable
      textCP168.Locked = bEnable
      '2023/4/13 END
   'end 2021/8/11
   
   'Add By Sindy 2009/05/11
   If m_EditMode = 2 Then
      'Add By Sindy 2014/3/25 檢查若有承辦歷程判發或退件重送時,且已無歷程附件則鎖住發文日
      bolUpdCP82 = False
      'Modify By Sindy 2014/4/14
      strSql = "select eep01,eep02 from empelectronprocess" & _
               " where eep01='" & textCP09 & "'" & _
               " order by eep02 desc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         bolUpdCP82 = True
         rsTmp.Close
      '2014/4/14 END
         strSql = "select eep01,eep02 from empelectronprocess" & _
                  " where eep01='" & textCP09 & "'" & _
                    " and eep04 in('" & EMP_判發 & "','" & EMP_退件重送 & "','" & EMP_送件 & "','" & EMP_發文歸檔 & "')" & _
                  " order by eep02 desc"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            intEEP02 = rsTmp.Fields("eep02")
            rsTmp.Close
            'Modify By Sindy 2018/9/4 + and eef12 is not null
            strSql = "select count(*) from empelectronfile" & _
                     " where eef01='" & textCP09 & "'" & _
                       " and eef02=" & intEEP02 & " and eef12 is not null"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               If rsTmp.Fields(0) = 0 Then '已無歷程附件則鎖住發文日
                  If Pub_StrUserSt03 <> "M51" Then 'Add By Sindy 2014/4/14 +if
                     'Modify By Sindy 2014/5/7 讓使用者自己可以拿掉發文日
'                     textCP27.Locked = True '發文日
                  End If
               End If
            End If
         End If
      End If
      rsTmp.Close
      '2014/3/25 END
      
      '若有發文室發文日-主管機關或發文室發文日-非主管機關(發文人非QPGMR)且使用者非電腦中心人員
      '則鎖住發文日欄位
      If (m_FieldList(123).fiOldData <> "" Or (m_FieldList(126).fiOldData <> "" And m_FieldList(153).fiOldData <> "QPGMR")) And Pub_StrUserSt03 <> "M51" Then
         textCP27.Locked = True '發文日
      End If
   End If
   '使用者非電腦中心人員, 則鎖住發文室相關欄位
   If Pub_StrUserSt03 <> "M51" Then
      If Not m_bDelete Then '內專催提申可刪除也開放可修改發文號
         textCP28.Locked = True '發文字號
      End If
      textCP124.Locked = True '發文室發文日-主管機關
      textCP127.Locked = True '發文室發文日-非主管機關
      textCP129.Locked = True '分所發文日
      textCP131.Locked = True '發文室取消發文備註
      textCP132.Locked = True '取消發文日
   End If
   '2009/05/11 End

   textCP144.Locked = bEnable  '2011/5/26 add by sona
   
   'Added by Lydia 2020/04/10 法務工作點數分配
   If CmdDot.Visible = True Then
       CmdDot.Enabled = bEnable
   End If
   'end 2020/04/10
   
   Set rsTmp = Nothing 'Add By Sindy 2014/3/25
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCP09.Locked = bEnable
End Sub

' 讀取商標基本檔
Private Function QueryTradeMark(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryTradeMark = False
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryTradeMark = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_Nation = rsTmp.Fields("TM10")
      End If
      ' 專用期間
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_CP53 = rsTmp.Fields("TM21")
      End If
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_CP54 = rsTmp.Fields("TM22")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取服務業務基本檔
Private Function QueryServicePractice(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryServicePractice = False
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & strSP01 & "' AND " & _
                  "SP02 = '" & strSP02 & "' AND " & _
                  "SP03 = '" & strSP03 & "' AND " & _
                  "SP04 = '" & strSP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryServicePractice = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_Nation = rsTmp.Fields("SP09")
      End If
      ' 專用期間
      If IsNull(rsTmp.Fields("SP20")) = False Then
         m_CP53 = rsTmp.Fields("SP20")
      End If
      If IsNull(rsTmp.Fields("SP21")) = False Then
         m_CP54 = rsTmp.Fields("SP21")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取專利基本檔
Private Function QueryPatent(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryPatent = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & strPA01 & "' AND " & _
                  "PA02 = '" & strPA02 & "' AND " & _
                  "PA03 = '" & strPA03 & "' AND " & _
                  "PA04 = '" & strPA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryPatent = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("PA05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("PA05")
      End If
      If IsNull(rsTmp.Fields("PA06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("PA06")
      End If
      If IsNull(rsTmp.Fields("PA07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("PA07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("PA26")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("PA26"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("PA09")) = False Then
         m_Nation = rsTmp.Fields("PA09")
      End If
      ' 專用期間
      If IsNull(rsTmp.Fields("PA24")) = False Then
         m_CP53 = rsTmp.Fields("PA24")
      End If
      If IsNull(rsTmp.Fields("PA25")) = False Then
         m_CP54 = rsTmp.Fields("PA25")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取法務基本檔
Private Function QueryLawCase(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryLawCase = False
   strSql = "SELECT * FROM LAWCASE " & _
            "WHERE LC01 = '" & strLC01 & "' AND " & _
                  "LC02 = '" & strLC02 & "' AND " & _
                  "LC03 = '" & strLC03 & "' AND " & _
                  "LC04 = '" & strLC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryLawCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("LC05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("LC05")
      End If
      If IsNull(rsTmp.Fields("LC06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("LC06")
      End If
      If IsNull(rsTmp.Fields("LC07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("LC07")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("LC11")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("LC11"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("LC15")) = False Then
         m_Nation = rsTmp.Fields("LC15")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取顧問案件基本檔
Private Function QueryHireCase(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryHireCase = False
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & strHC01 & "' AND " & _
                  "HC02 = '" & strHC02 & "' AND " & _
                  "HC03 = '" & strHC03 & "' AND " & _
                  "HC04 = '" & strHC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryHireCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("HC06")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("HC05")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("HC05"), 0)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim tmpArr As Variant 'Added by Lydia 2016/06/06
   'Added by Lydia 2021/10/08
   Dim tmpOA As String, intQ As Integer
   Dim pStart As String '第一個列出的出名代理人
   'end 2021/10/108
   
   If m_DataListCount <= 0 Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & m_DataList(m_CurrDL).diCP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then: textCP05 = ChangeWStringToTString(rsTmp.Fields("CP05"))
      ' 本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then: textCP06 = ChangeWStringToTString(rsTmp.Fields("CP06"))
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then: textCP07 = ChangeWStringToTString(rsTmp.Fields("CP07"))
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then: textCP08 = rsTmp.Fields("CP08")
      ' 總收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then: textCP09 = rsTmp.Fields("CP09")
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then: textCP10 = rsTmp.Fields("CP10")
      'Add by Morgan 2004/2/11
      m_CP10 = textCP10
      'Add By Sindy 2023/3/15
      If m_CP01 = "T" And m_CP10 = "210" Then '210.陳述意見書
         Label11.Caption = "是否為快軌案件："
         Label23.Caption = "(Y/N)"
      End If
      '2023/3/15 END
      
      If IsNull(rsTmp.Fields("CP11")) = False Then: textCP11 = rsTmp.Fields("CP11")
      If IsNull(rsTmp.Fields("CP12")) = False Then: textCP12 = rsTmp.Fields("CP12")
      If IsNull(rsTmp.Fields("CP13")) = False Then: textCP13 = rsTmp.Fields("CP13")
      textCP13.Tag = textCP13 'Add by Morgan 2010/1/5 紀錄原智權人員
      
      If IsNull(rsTmp.Fields("CP14")) = False Then: textCP14 = rsTmp.Fields("CP14")
      'Add By Cheng 2003/10/29
      '記錄原承辦人
      Me.textCP14.Tag = "" & rsTmp.Fields("CP14").Value
      'End
      If IsNull(rsTmp.Fields("CP15")) = False Then: textCP15 = rsTmp.Fields("CP15")
      If IsNull(rsTmp.Fields("CP16")) = False Then: textCP16 = rsTmp.Fields("CP16")
      textCP16.Tag = textCP16 'Add by Sindy 2022/12/5 紀錄費用
      If IsNull(rsTmp.Fields("CP17")) = False Then: textCP17 = rsTmp.Fields("CP17")
      textCP17.Tag = textCP17 'Add by Sindy 2022/12/5 紀錄規費
      If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
      textCP18.Tag = textCP18 'Add by Sindy 2022/12/5 紀錄點數
      If IsNull(rsTmp.Fields("CP19")) = False Then: textCP19 = rsTmp.Fields("CP19")
      If IsNull(rsTmp.Fields("CP20")) = False Then: textCP20 = rsTmp.Fields("CP20")
      If IsNull(rsTmp.Fields("CP21")) = False Then: textCP21 = rsTmp.Fields("CP21")
      If IsNull(rsTmp.Fields("CP22")) = False Then: textCP22 = rsTmp.Fields("CP22")
      If IsNull(rsTmp.Fields("CP23")) = False Then: textCP23 = rsTmp.Fields("CP23")
      If IsNull(rsTmp.Fields("CP24")) = False Then: textCP24 = rsTmp.Fields("CP24")
      ' 准駁日
      If IsNull(rsTmp.Fields("CP25")) = False Then: textCP25 = ChangeWStringToTString(rsTmp.Fields("CP25"))
      If IsNull(rsTmp.Fields("CP26")) = False Then: textCP26 = rsTmp.Fields("CP26")
      ' 發文日
      If IsNull(rsTmp.Fields("CP27")) = False Then: textCP27 = ChangeWStringToTString(rsTmp.Fields("CP27"))
      textCP27.Tag = textCP27 'Added by Morgan 2024/11/18
      If IsNull(rsTmp.Fields("CP28")) = False Then: textCP28 = rsTmp.Fields("CP28")
      If IsNull(rsTmp.Fields("CP29")) = False Then: textCP29 = rsTmp.Fields("CP29")
      textCP29.Tag = textCP29 'Added by Morgan 2025/2/25
      If IsNull(rsTmp.Fields("CP30")) = False Then: textCP30 = rsTmp.Fields("CP30")
      If IsNull(rsTmp.Fields("CP31")) = False Then: textCP31 = rsTmp.Fields("CP31")
      If IsNull(rsTmp.Fields("CP32")) = False Then: textCP32 = rsTmp.Fields("CP32")
      If IsNull(rsTmp.Fields("CP33")) = False Then: textCP33 = rsTmp.Fields("CP33")
      If IsNull(rsTmp.Fields("CP34")) = False Then: textCP34 = rsTmp.Fields("CP34")
      If IsNull(rsTmp.Fields("CP35")) = False Then: textCP35 = rsTmp.Fields("CP35")
      If IsNull(rsTmp.Fields("CP36")) = False Then: textCP36 = rsTmp.Fields("CP36")
        Select Case Me.textCP01.Text
        Case "T", "FCT", "CFT", "TF"
            If IsNull(rsTmp.Fields("CP37")) = False Then: textCP37_1 = rsTmp.Fields("CP37")
        Case Else
            If IsNull(rsTmp.Fields("CP37")) = False Then: textCP37 = rsTmp.Fields("CP37")
            If IsNull(rsTmp.Fields("CP38")) = False Then: textCP38 = rsTmp.Fields("CP38")
            If IsNull(rsTmp.Fields("CP39")) = False Then: textCP39 = rsTmp.Fields("CP39")
        End Select
      If IsNull(rsTmp.Fields("CP40")) = False Then: textCP40 = rsTmp.Fields("CP40")
      If IsNull(rsTmp.Fields("CP41")) = False Then: textCP41 = rsTmp.Fields("CP41")
      If IsNull(rsTmp.Fields("CP42")) = False Then: textCP42 = rsTmp.Fields("CP42")
      If IsNull(rsTmp.Fields("CP43")) = False Then: textCP43 = rsTmp.Fields("CP43")
      'Add by Morgan 2004/2/12
      m_CP43 = textCP43
      
      'Modify By Cheng 2002/04/22
      '若代理人代號最後面三碼為"000", 則顯示六碼即可
'      If IsNull(rsTmp.Fields("CP44")) = False Then: textCP44 = rsTmp.Fields("CP44")
      If IsNull(rsTmp.Fields("CP44")) = False Then: textCP44 = IIf(Len(rsTmp.Fields("CP44")) = 9 And Right(rsTmp.Fields("CP44"), 3) = "000", Left(rsTmp.Fields("CP44"), 6), rsTmp.Fields("CP44"))
      If IsNull(rsTmp.Fields("CP45")) = False Then: textCP45 = rsTmp.Fields("CP45")
      ' 代理人收達日
      If IsNull(rsTmp.Fields("CP46")) = False Then: textCP46 = ChangeWStringToTString(rsTmp.Fields("CP46"))
      ' 代理人提申日
      If IsNull(rsTmp.Fields("CP47")) = False Then: textCP47 = ChangeWStringToTString(rsTmp.Fields("CP47"))
      'Added by Lydia 2016/05/30 法務或顧問案件時，回執收受日為111111和110101，修改'代理人提申日'欄的Label名稱
      Label30(3) = "代理人提申日："
      If ClsPDGetSystemKind(m_CP01, m_SK02) Then
         If m_SK02 = 3 Or m_SK02 = 4 Then
            If Val(textCP46) = 111111 Then
               Label30(3) = "回執退件日："
            ElseIf Val(textCP46) = 110101 Then
               Label30(3) = "回執未回郵局送達日："
            End If
         End If
      End If
      'end 2016/05/30
      
      ' 承辦期限
      If IsNull(rsTmp.Fields("CP48")) = False Then: textCP48 = ChangeWStringToTString(rsTmp.Fields("CP48"))
      If IsNull(rsTmp.Fields("CP49")) = False Then: textCP49 = rsTmp.Fields("CP49")
      If IsNull(rsTmp.Fields("CP50")) = False Then: textCP50 = rsTmp.Fields("CP50")
      If IsNull(rsTmp.Fields("CP51")) = False Then: textCP51 = rsTmp.Fields("CP51")
      If IsNull(rsTmp.Fields("CP52")) = False Then: textCP52 = rsTmp.Fields("CP52")
       
      'Added by Lydia 2017/08/24 預設顯示
      Label20(7).Visible = False
      Label20(2).Visible = True
      textCP53.Width = 1095
      'end 2017/08/24
      
      'Modify By Sindy 2009/07/06
      'Modify by Morgan 2010/1/21 +908
      'Modify by Morgan 2010/6/22 +P 1001
      'Modified by Morgan 2012/10/4 +FCP
      'Modify by Amy 2018/04/10 +612 年費移作次年
      If (m_CP01 = "P" Or m_CP01 = "CFP" Or m_CP01 = "FCP") And (m_CP10 = "601" Or m_CP10 = "1001" Or m_CP10 = "605" Or m_CP10 = "606" Or m_CP10 = "607" Or m_CP10 = "908" Or m_CP10 = "612") Then
         Label20(2).Caption = "繳費年度/次數(起/迄)："
         textCP53_2.Visible = True
         textCP54_2.Visible = True
         textCP53_2.Top = textCP53.Top 'Added by Morgan 2017/9/25
         textCP54_2.Top = textCP54.Top 'Added by Lydia 2021/10/08
         textCP53.Visible = False
         textCP54.Visible = False
         ' 繳費年度/次數(起)
         If IsNull(rsTmp.Fields("CP53")) = False Then: textCP53_2 = rsTmp.Fields("CP53")
         ' 繳費年度/次數(迄)
         If IsNull(rsTmp.Fields("CP54")) = False Then: textCP54_2 = rsTmp.Fields("CP54")
      'Added by Lydia 2017/08/24 TB條碼案繳年費708,服務業務結果1801-第?期登記期
      ElseIf m_CP01 = "TB" And (m_CP10 = "708" Or m_CP10 = "1801") Then
         Label20(2).Visible = False
         Label20(7).Visible = True
         textCP53_2.Visible = False
         textCP54_2.Visible = False
         textCP53.Visible = True
         textCP54.Visible = False
         textCP53.Width = 500
         If IsNull(rsTmp.Fields("CP53")) = False Then textCP53 = rsTmp.Fields("CP53")
         If IsNull(rsTmp.Fields("CP54")) = False Then textCP54 = rsTmp.Fields("CP54")
      'end 2017/08/24
      Else
         Label20(2).Caption = "授權期間/質權設定期間(起/迄)/聘任期間："
         textCP53_2.Visible = False
         textCP54_2.Visible = False
         textCP53.Visible = True
         textCP54.Visible = True
      '2009/07/06 End
         ' 質權設定期間(起)
         If IsNull(rsTmp.Fields("CP53")) = False Then: textCP53 = ChangeWStringToTString(rsTmp.Fields("CP53"))
         ' 質權設定期間(迄)
         If IsNull(rsTmp.Fields("CP54")) = False Then: textCP54 = ChangeWStringToTString(rsTmp.Fields("CP54"))
      End If
      
      '2011/5/26 add by sonia
      If Pub_StrUserSt03 = "M51" Or Mid(Pub_StrUserSt03, 1, 2) = "P1" Then
         Label33(2).Visible = True
         textCP144.Visible = True
         textCP144.Enabled = True
      Else
         Label33(2).Visible = False
         textCP144.Visible = False
         textCP144.Enabled = False
      End If
      '2011/5/26 end
      
      'Add by Morgan 2009/10/8
      If m_CP01 = "FCP" And m_CP10 = "908" Then
         lblCP19.Caption = "退費金額："
         lblCP49.Caption = "特定退款人名稱："
         lblCP86.Caption = "是否同意扣除服務費："
         lblCP86_1.Caption = "(N:不同意)"
      Else
         lblCP19.Caption = "後　　金："
         lblCP49.Caption = "條款/當事人稱謂："
         'Add by Morgan 2010/7/1
         If m_CP01 = "FCP" And m_CP10 = "202" Then
            lblCP86.Caption = "是否為複委任："
         Else
         'end 2010/7/1
            lblCP86.Caption = "收到分所接洽單紀錄："
         End If
         lblCP86_1.Caption = "(Y:是)"
      End If
      'end 2009/10/8
      
      If IsNull(rsTmp.Fields("CP55")) = False Then: textCP55 = rsTmp.Fields("CP55")
      If IsNull(rsTmp.Fields("CP56")) = False Then: textCP56 = rsTmp.Fields("CP56")
      ' 取消收文日期
      If IsNull(rsTmp.Fields("CP57")) = False Then: textCP57 = ChangeWStringToTString(rsTmp.Fields("CP57"))
      If IsNull(rsTmp.Fields("CP58")) = False Then: textCP58 = rsTmp.Fields("CP58")
      If IsNull(rsTmp.Fields("CP59")) = False Then: textCP59 = rsTmp.Fields("CP59")
      If IsNull(rsTmp.Fields("CP60")) = False Then: textCP60 = rsTmp.Fields("CP60")
      If IsNull(rsTmp.Fields("CP61")) = False Then: textCP61 = rsTmp.Fields("CP61")
      If IsNull(rsTmp.Fields("CP62")) = False Then: textCP62 = rsTmp.Fields("CP62")
      If IsNull(rsTmp.Fields("CP63")) = False Then: textCP63 = rsTmp.Fields("CP63")
      If IsNull(rsTmp.Fields("CP64")) = False Then: textCP64 = rsTmp.Fields("CP64")
      'Modified by Lydia 2025/02/12
      'If IsNull(rsTmp.Fields("CP71")) = False Then: textCP71 = rsTmp.Fields("CP71")
      If "" & rsTmp.Fields("CP71") <> "" Then
         'P臺灣與大陸案若申請延緩審查，請於發文時讓user輸入延緩審查日期; FCP案在核准時輸入
         If ("" & rsTmp.Fields("CP01") = "P" And m_Nation = "000" And "" & rsTmp.Fields("CP10") = "245") Or ("" & rsTmp.Fields("CP01") = "FCP" And "" & rsTmp.Fields("CP10") = "1924") Then
            textCP71 = TransDate(rsTmp.Fields("cp71"), 1)
         Else
            textCP71 = rsTmp.Fields("CP71")
         End If
      End If
      'end 2025/02/12
      If IsNull(rsTmp.Fields("CP72")) = False Then: textCP72 = rsTmp.Fields("CP72")
      If IsNull(rsTmp.Fields("CP80")) = False Then: textCP80 = rsTmp.Fields("CP80")
      'add & edit by nick 2004/08/18
      'textCP81 = "" & rsTmp.Fields("CP81") 'Add by Morgan 2004/6/11
      If IsNull(rsTmp.Fields("CP81")) = False Then: textCP81 = rsTmp.Fields("CP81")
      If IsNull(rsTmp.Fields("CP82")) = False Then: textCP82 = Format(rsTmp.Fields("CP82"), "00:00:00")
      If IsNull(rsTmp.Fields("CP83")) = False Then: textCP83 = rsTmp.Fields("CP83")
      If IsNull(rsTmp.Fields("CP84")) = False Then: textCP84 = rsTmp.Fields("CP84")
      'If IsNull(rsTmp.Fields("CP85")) = False Then: textCP85 = rsTmp.Fields("CP85")'Remove by Morgan 2010/12/30 目前沒用
      m_CP85 = "" & rsTmp.Fields("CP85") 'Added by Lydia 2019/06/28 承辦人發文日/FCP定稿日期
      If IsNull(rsTmp.Fields("CP86")) = False Then: textCP86 = rsTmp.Fields("CP86")
      If IsNull(rsTmp.Fields("CP87")) = False Then: textCP87 = rsTmp.Fields("CP87")
      If IsNull(rsTmp.Fields("CP88")) = False Then: textCP88 = rsTmp.Fields("CP88")
      If IsNull(rsTmp.Fields("CP89")) = False Then: textCP89 = rsTmp.Fields("CP89")
      If IsNull(rsTmp.Fields("CP90")) = False Then: textCP90 = rsTmp.Fields("CP90")
      If IsNull(rsTmp.Fields("CP91")) = False Then: textCP91 = rsTmp.Fields("CP91")
      If IsNull(rsTmp.Fields("CP92")) = False Then: textCP92 = rsTmp.Fields("CP92")
      If IsNull(rsTmp.Fields("CP93")) = False Then: textCP93 = rsTmp.Fields("CP93")
      If IsNull(rsTmp.Fields("CP94")) = False Then: textCP94 = rsTmp.Fields("CP94")
      If IsNull(rsTmp.Fields("CP95")) = False Then: textCP95 = rsTmp.Fields("CP95")
      If IsNull(rsTmp.Fields("CP96")) = False Then: textCP96 = rsTmp.Fields("CP96")
      'Add by Morgan 2007/7/19
      If IsNull(rsTmp.Fields("CP113")) = False Then: textCP113 = rsTmp.Fields("CP113")
      If IsNull(rsTmp.Fields("CP114")) = False Then: textCP114 = rsTmp.Fields("CP114")
      'Add by Morgan 2008/5/14
      If IsNull(rsTmp.Fields("CP116")) = False Then: textCP44 = textCP44 & "-" & rsTmp.Fields("CP116")
      If IsNull(rsTmp.Fields("CP117")) = False Then: textCP117 = rsTmp.Fields("CP117")
      'Add by Morgan 2008/7/11
      If IsNull(rsTmp.Fields("CP118")) = False Then: textCP118 = rsTmp.Fields("CP118")
      '2008/8/27 add by sonia 櫃台收文日
      If IsNull(rsTmp.Fields("CP119")) = False Then: textCP119 = ChangeWStringToTString(rsTmp.Fields("CP119"))
      'Add by Morgan 2008/11/10
      If IsNull(rsTmp.Fields("CP120")) = False Then: textCP120 = rsTmp.Fields("CP120")
      If IsNull(rsTmp.Fields("CP121")) = False Then: textCP121 = rsTmp.Fields("CP121")
      If IsNull(rsTmp.Fields("CP145")) = False Then: textCP145 = rsTmp.Fields("CP145") 'Added by Morgan 2016/6/2
      'Add By Sindy 2015/6/3
      If IsNull(rsTmp.Fields("CP148")) = False Then: textCP148 = rsTmp.Fields("CP148")
      If InStr(m_CP01, "P") > 0 Then
         'Modify By Sindy 2015/9/23
         'Label148 = "是否有檢索："
         If Left(textCP09, 1) = "C" Then
            Label148 = "是否有檢索："
         Else
            Label148 = "是否為特殊請款："
         End If
         '2015/9/23 END
      Else
         Label148 = "是否為一申請書多件:"
      End If
      '2015/6/3 END
      
      If IsNull(rsTmp.Fields("CP152")) = False Then: textCP152 = ChangeWStringToTString(rsTmp.Fields("CP152")) 'Added by Morgan 2016/6/8
      'Add by Morgan 2009/3/18
      If IsNull(rsTmp.Fields("CP123")) = False Then: textCP123 = rsTmp.Fields("CP123")
      If IsNull(rsTmp.Fields("CP124")) = False Then: textCP124 = ChangeWStringToTString(rsTmp.Fields("CP124"))
      If IsNull(rsTmp.Fields("CP125")) = False Then: textCP125 = Format(rsTmp.Fields("CP125"), "00:00:00")
      If IsNull(rsTmp.Fields("CP126")) = False Then: textCP126 = rsTmp.Fields("CP126")
      If IsNull(rsTmp.Fields("CP127")) = False Then: textCP127 = ChangeWStringToTString(rsTmp.Fields("CP127"))
      If IsNull(rsTmp.Fields("CP128")) = False Then: textCP128 = Format(rsTmp.Fields("CP128"), "00:00:00")
      If IsNull(rsTmp.Fields("CP129")) = False Then: textCP129 = ChangeWStringToTString(rsTmp.Fields("CP129"))
      
      'Add By Sindy 2009/04/27
      If IsNull(rsTmp.Fields("CP131")) = False Then: textCP131 = rsTmp.Fields("CP131")
      If IsNull(rsTmp.Fields("CP132")) = False Then: textCP132 = ChangeWStringToTString(rsTmp.Fields("CP132"))
      
      'Add by Morgna 2010/12/30
      If IsNull(rsTmp.Fields("cp140")) = False Then textCP140 = rsTmp.Fields("cp140")
      If IsNull(rsTmp.Fields("cp141")) = False Then
         'Add By Sindy 2024/5/27 取消4
         If Val("" & rsTmp.Fields("cp141")) <> 4 Then
         '2024/5/27 END
            OptSendType(Val(rsTmp.Fields("cp141"))).Value = True
         End If
      End If
      If IsNull(rsTmp.Fields("cp142")) = False Then textCP142 = TransDate(rsTmp.Fields("cp142"), 1)
      OptSendType(1).Caption = PUB_GetCP114Opt1Desc(textCP01, textCP10)  'Added by Morgan 2024/1/22
      'Add By Sindy 2021/4/20
      'Memo by Lydia 2021/11/09 區分送件方式CP141和指定送件日期方式CP164：因為放在同一Frame的Option選項只能有一項點選
      If "" & rsTmp.Fields("CP164") = "1" Then
         Option1(0).Value = True
      ElseIf "" & rsTmp.Fields("CP164") = "2" Then
         Option1(1).Value = True
      'Add By Sindy 2021/10/20
      ElseIf "" & rsTmp.Fields("CP164") = "3" Then
         Option1(2).Value = True
      End If
      '2021/4/20 END
      'Add By Sindy 2024/5/27 暫不送
      If IsNull(rsTmp.Fields("cp176")) = False Then
         chkCP176.Value = 1
      Else
         chkCP176.Value = 0
      End If
      '2024/5/27 END
      
      If IsNull(rsTmp.Fields("cp144")) = False Then textCP144 = rsTmp.Fields("cp144")   '2011/5/26 add by sonia
      
      'Add by Morgan 2010/1/5
      textCP135 = "" & rsTmp.Fields("CP135")
      'Modified by Morgan 2015/8/5 +209,235 --靜芳
      'Modified by Morgan 2022/7/5 +107
      'Modified by Sindy 2023/10/12 +307
      If textCP10 = "416" Or textCP10 = "201" Or textCP10 = "209" Or _
         textCP10 = "235" Or textCP10 = "107" Or textCP10 = "307" Then
         lblCP136 = "總項數："
         lblCP135 = "總頁數：" 'Add By Sindy 2023/4/13
      Else
         lblCP136 = "增加項數："
         lblCP135 = "增加頁數：" 'Add By Sindy 2023/4/13
         textCP135.Enabled = False 'Add By Sindy 2023/4/13
      End If
      textCP136 = "" & rsTmp.Fields("CP136")
      textCP137 = "" & rsTmp.Fields("CP137")
      textCP138 = "" & rsTmp.Fields("CP138")
      'end 2010/1/5
      'Add By Sindy 2023/4/13
      textCP167 = "" & rsTmp.Fields("CP167")
      textCP168 = "" & rsTmp.Fields("CP168")
      strSql = "select * from pagedetail where pd01='" & textCP09.Text & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         cmdPage.BackColor = &HC0C0FF '粉紅色
      Else
         cmdPage.BackColor = &H80000010
      End If
      '2023/4/13 END
      
      'Add by Morgan 2010/5/13
      If textCP01 = "FCP" And textCP10 = "202" And textCP86 = "Y" Then
         lbl202CP86 = "( 複委任 )"
      Else
         lbl202CP86 = ""
      End If
      
      'add by nickc 2006/01/27
      m_CP110 = CheckStr(rsTmp.Fields("cp110"))
      'Add By Sindy 2009/05/11
      m_CP130 = CheckStr(rsTmp.Fields("cp130"))
      'Added by  Lydida 2021/11/29 可結餘日期
      m_cp109 = "" & rsTmp.Fields("cp109")
      
      m_CP65 = "" & rsTmp.Fields("cp65") 'Added by Morgan 2016/6/2
      m_CP162 = "" & rsTmp.Fields("CP162") 'Added by Lydia 2023/08/14 (案件進度)案源單號
      
      'add by nickc 2008/01/31
      If IsNull(rsTmp.Fields("CP73")) = False Then: lblCP73.Caption = rsTmp.Fields("CP73")
      If IsNull(rsTmp.Fields("CP74")) = False Then: lblCP74.Caption = rsTmp.Fields("CP74")
      If IsNull(rsTmp.Fields("CP75")) = False Then: lblCP75.Caption = rsTmp.Fields("CP75")
      If IsNull(rsTmp.Fields("CP76")) = False Then: lblCP76.Caption = rsTmp.Fields("CP76")
      If IsNull(rsTmp.Fields("CP77")) = False Then: lblCP77.Caption = rsTmp.Fields("CP77")
      If IsNull(rsTmp.Fields("CP78")) = False Then: lblCP78.Caption = rsTmp.Fields("CP78")
      If IsNull(rsTmp.Fields("CP79")) = False Then: lblCP79.Caption = rsTmp.Fields("CP79")
      DoEvents
      
      '出名代理人
      DoEvents
      'Modify by Morgan 2007/6/12 加考慮代理人無法再出名的情形 Ex.65002
      'strSQL = "select st01,st02,OA03 from ouragent,staff where oa01='" & Me.textCP01.Text & "' and st01=oa02 order by 3 DESC, 1 desc"
      If m_CP110 = "" Then
         strSql = "select st01,st02,OA03 from ouragent,staff where oa01='" & Me.textCP01.Text & "' and st01=oa02 order by 3 DESC, 1 desc"
      Else
         '2010/5/6 MODIFY BY SONIA 因會改變出名順序故取消ST01的排序FCT-030306
         'strSql = "select st01,st02,0 from staff where instr('" & m_CP110 & "',st01)>0" & _
            " union select st01,st02,OA03 from ouragent,staff where oa01='" & Me.textCP01.Text & "' and instr('" & m_CP110 & "',st01)=0 and st01=oa02" & _
            " order by 3 DESC, 1 desc"
         'Modified by Lydia 2016/06/06 依出名順序排序
         'strSql = "select st01,st02,0,OA03 from staff,OURAGENT where instr('" & m_CP110 & "',st01)>0 AND oa01='" & Me.textCP01.Text & "' AND ST01=OA02 " & _
            " union select st01,st02,OA03,OA03 from ouragent,staff where oa01='" & Me.textCP01.Text & "' and instr('" & m_CP110 & "',st01)=0 and st01=oa02" & _
            " order by 3 DESC, 1 desc"
         'Modified by Morgan 2020/3/17 排序改用資料順序不要再抓設定
         'Modified by Morgan 2020/4/14 再改回抓設定
         'strSql = "SELECT ST01,ST02,INSTR('" & m_CP110 & "',ST01) OA03,'3' SRT2 FROM OURAGENT,STAFF WHERE OA01='" & Me.textCP01.Text & "' AND INSTR('" & m_CP110 & "',ST01)=0 AND ST01=OA02 "
         strSql = "SELECT ST01,ST02,OA03,'3' SRT2 FROM OURAGENT,STAFF WHERE OA01='" & Me.textCP01.Text & "' AND INSTR('" & m_CP110 & "',ST01)=0 AND ST01=OA02 "
         'end 2020/3/17
         tmpArr = Split(m_CP110, ",")
         For intI = 0 To UBound(tmpArr)
            If Trim(tmpArr(intI)) <> "" Then
               'Modified by Morgan 2019/5/17 要考慮代理人無法再出名的情形 Ex:FCP-023657(65002)
               'strSql = strSql & "UNION SELECT ST01,ST02,0 OA03 ,'" & intI + 1 & "' SRT2 FROM OURAGENT,STAFF WHERE OA01='" & Me.textCP01.Text & "' AND ST01='" & tmpArr(intI) & "' AND ST01=OA02 "
               strSql = strSql & "UNION SELECT ST01,ST02,0 OA03 ,'" & intI + 1 & "' SRT2 FROM STAFF,OURAGENT WHERE ST01='" & tmpArr(intI) & "' AND OA02(+)=st01 AND  OA01(+)='" & Me.textCP01.Text & "'"
               'end 2019/5/17
            End If
         Next
         strSql = strSql & "ORDER BY OA03 DESC,SRT2 DESC"
      End If
      'end 2007/6/12
      CheckOC
      Me.lstNameAgent.Clear
      'Added by Lydia 2021/10/08
      Me.lstNameAgent.Tag = ""  '原本放在ItemData,改放在Tag
      tmpOA = ""
      pStart = ""
      intQ = 0
      'end 2021/10/08
      lstNameAgent.Enabled = False
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      Dim iNum As Integer
      iNum = 0
      If adoRecordset.RecordCount > 0 Then
         Do While Not adoRecordset.EOF
            If InStr(m_CP110, "" & adoRecordset.Fields(0)) > 0 Then
               lstNameAgent.AddItem "" & adoRecordset.Fields(1), 0
               'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
               'lstNameAgent.ITEMDATA(0) = adoRecordset.Fields(0) '員工編號
               'Modified by Lydia 2021/10/08 參考PUB_SetOurAgent；原本放在ItemData,改放在Tag
               'lstNameAgent.ItemData(0) = PUB_Id2Num(adoRecordset.Fields(0)) '員工編號
                tmpOA = adoRecordset.Fields(0) & IIf(tmpOA <> "", ",", "") & tmpOA
               lstNameAgent.Selected(0) = True
               iNum = iNum + 1
            Else
               lstNameAgent.AddItem "" & adoRecordset.Fields(1), iNum
               'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
               'lstNameAgent.ITEMDATA(iNum) = adoRecordset.Fields(0) '員工編號
               'Modified by Lydia 2021/10/08 參考PUB_SetOurAgent；
               'lstNameAgent.ItemData(iNum) = PUB_Id2Num(adoRecordset.Fields(0)) '員工編號
               If pStart <> "" Then
                   intQ = InStr(tmpOA, pStart & ",")
                   If intQ = 0 Then
                       tmpOA = tmpOA & "," & adoRecordset.Fields(0)
                   Else
                       tmpOA = Mid(tmpOA, 1, intQ + Len(pStart)) & adoRecordset.Fields(0) & "," & Mid(tmpOA, intQ + Len(pStart & ","))
                   End If
               Else
                    tmpOA = adoRecordset.Fields(0) & IIf(tmpOA <> "", ",", "") & tmpOA
               End If
               'end 2021/10/08
            End If
            adoRecordset.MoveNext
         Loop
      End If
      lstNameAgent.Enabled = True
      'Added by Lydia 2021/10/08
      If tmpOA <> "" Then
          lstNameAgent.Tag = tmpOA
          lstNameAgent.ListIndex = 0
      End If
      'end 2021/10/08
      DoEvents
      CheckOC
      
      'Add By Sindy 2009/05/11
      '發文主管機關名稱
      DoEvents
      'modify by sonia 2015/11/5
      'strSql = "SELECT Distinct(CF10) FROM CaseFee WHERE CF01='" & Me.textCP01.Text & "' AND CF02='000' AND length(CF03)=3 "
      strSql = "SELECT Distinct(CF10) FROM CaseFee WHERE CF01='" & Me.textCP01.Text & "' AND CF02='" & m_Nation & "' AND length(CF03)=3 "
      'add by sonia 2018/6/26
      If Me.textCP01.Text = "CFT" And Me.textCP10.Text = "304" Then
         strSql = strSql & " union select '經濟部智慧財產局' from dual"
      'Added
      End If
      'end  2018/6/26
      CheckOC
      Me.lstNameOrg.Clear
      lstNameOrg.Enabled = False
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      iNum = 0
      strExc(3) = "" 'Added by Lydia 2025/01/15
      strExc(5) = "" 'Added by Lydia 2025/01/16
      Dim temp_CP130() As String, intCnt As Integer, ii As Integer
      ii = 1: intCnt = 0
      Do While InStr(ii, m_CP130, ",") > 0
         intCnt = intCnt + 1
         ii = InStr(ii, m_CP130, ",") + 1
      Loop
      If m_CP130 <> "" Then intCnt = intCnt + 1
      If adoRecordset.RecordCount > 0 Then
         'Added by Lydia 2025/01/16
         strExc(5) = adoRecordset.GetString(adClipString, , , ",")
         adoRecordset.MoveFirst
         'end 2025/01/16
         Do While Not adoRecordset.EOF
            If Not IsNull(adoRecordset.Fields(0)) Then
               iNum = lstNameOrg.ListCount
               lstNameOrg.AddItem adoRecordset.Fields(0), iNum
               temp_CP130 = Split(m_CP130, ",")
               For ii = 0 To intCnt - 1
                  If Trim(temp_CP130(ii)) = Trim(adoRecordset.Fields(0)) Then
                     lstNameOrg.Selected(iNum) = True
                     Exit For
                  'Added by Lydia 2025/01/15
                  Else '將舊名稱也列出;
                     If Trim(temp_CP130(ii)) <> "" And InStr(strExc(5), Trim(temp_CP130(ii))) = 0 Then 'Added by Lydia 2025/01/16 FCT案有多筆主管機關
                        strExc(3) = strExc(3) & IIf(strExc(3) <> "", ",", "") & Trim(temp_CP130(ii))
                     End If
                  'end 2025/01/15
                  End If
               Next ii
            End If
            adoRecordset.MoveNext
         Loop
      End If
      'Added by Lydia 2025/01/15 將舊名稱也列出;
      If strExc(3) <> "" Then
         iNum = lstNameOrg.ListCount
         temp_CP130 = Split(m_CP130, ",")
         For intI = 0 To UBound(temp_CP130)
            lstNameOrg.AddItem temp_CP130(intI), iNum
            lstNameOrg.Selected(iNum) = True
            iNum = iNum + 1
         Next intI
      End If
      'end 2025/01/15
      lstNameOrg.Enabled = True
      DoEvents
      CheckOC
      '2009/05/11 End
      
      'Added by Lydia 2021/01/14 法律所案源收文：讀取案源
      If InStr(textCP01, "L") > 0 Then
          Call ReadLOS
      End If
      'end 2021/01/14
      
      'Added by Lydia 2021/05/05 ACS智財顧問專業分配比例管制
      'Modified by Lydia 2024/04/15 +LA之顧問聘任0
      'If m_CP01 = "ACS" And m_CP10 = "112" Then
      If (m_CP01 = "ACS" And m_CP10 = "112") Or (m_CP01 = "LA" And m_CP10 = "0") Then
          Label8.Caption = "簽約時數："
      End If
      'end 2021/05/05

      ' 更新欄位的內容
      UpdateCUID rsTmp
      ' 更新欄位的內容
      UpdateFieldOldData rsTmp

      textCP10_Validate False
      textCP11_Validate False
      textCP12_Validate False
      textCP13_Validate False
      textCP14_Validate False
      textCP29_Validate False
      textCP44_Validate False
      textCP55_Validate False
      textCP56_Validate False
      textCP58_Validate False
      textCP71_Validate False
      'add & edit by nick 2004/08/18
      textCP83_Validate False
      textCP89_Validate False
      textCP90_Validate False
      textCP91_Validate False
      textCP92_Validate False
      textCP93_Validate False
      textCP94_Validate False
      textCP95_Validate False
      textCP96_Validate False
   End If
   'add by nickc 2006/12/28 秀玲說發文日跟發文規費有值時，不可以修改規費
   If textCP84 <> "" And textCP27 <> "" Then
      textCP17.Locked = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
EXITSUB:
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim stRPTxt As String 'Add byAmy 2022/09/02 '顯示對造/相關人 label名稱
   
   m_CP53 = Empty
   m_CP54 = Empty
   ' 清除欄位內容
   ClearField
   ' 依本所案號讀取基本檔案
   Select Case m_CP01
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         QueryTradeMark m_CP01, m_CP02, m_CP03, m_CP04
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         QueryPatent m_CP01, m_CP02, m_CP03, m_CP04
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/24 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         QueryLawCase m_CP01, m_CP02, m_CP03, m_CP04
      ' 讀取顧問案件基本檔
      Case "LA":
         QueryHireCase m_CP01, m_CP02, m_CP03, m_CP04
      ' 讀取服務業務基本檔
      Case Else:
         QueryServicePractice m_CP01, m_CP02, m_CP03, m_CP04
   End Select
   ' 讀取案件進度檔
   QueryCaseProgress
   
   'Add By Cheng 2002/01/28
   '若為CFP案且案件性質為1001,1002,1201-1203,1210,1211,1301-1307,1401,1502,1504,1507,1801,1802,1805-1808,1903
   '2006/9/11 MODIFY BY SONIA 加入 FG
   'If Me.textCP01.Text = "FCP" Then
   If (Me.textCP01.Text = "FCP" Or Me.textCP01.Text = "FG") Then
      If Me.textCP10.Text = "1001" Or Me.textCP10.Text = "1002" Or _
         (Me.textCP10.Text >= "1201" Or Me.textCP10.Text <= "1203") Or _
         Me.textCP10.Text = "1210" Or Me.textCP10.Text = "1211" Or _
         (Me.textCP10.Text >= "1301" Or Me.textCP10.Text <= "1307") Or _
         Me.textCP10.Text = "1401" Or Me.textCP10.Text = "1502" Or Me.textCP10.Text = "1504" Or Me.textCP10.Text = "1507" Or Me.textCP10.Text = "1801" Or Me.textCP10.Text = "1802" Or _
         (Me.textCP10.Text >= "1805" Or Me.textCP10.Text <= "1808") Or Me.textCP10.Text = "1903" Then
         Me.cmdPrintCForm.Visible = True
      Else
         Me.cmdPrintCForm.Visible = False
      End If
   '92.1.28 ADD BY SONIA
   ElseIf Me.textCP01.Text = "CFP" And Me.textCP10.Text <> "1001" And Mid(Me.textCP10.Text, 4, 1) <> "" Then
         Me.cmdPrintCForm.Visible = True
   '92.1.28 END
   'Add by Morgan 2010/6/18
   ElseIf Left(textCP12, 1) = "F" And textCP01 = "P" And textCP09 > "C" And textCP07 <> "" Then
      cmdPrintCForm.Visible = True
   Else
      Me.cmdPrintCForm.Visible = False
   End If
   
   'Add By Sindy 2023/10/3
   If strSrvDate(1) >= 外專承辦歷程啟用日 Then
      cmdPrint201.Visible = False '取消此功能
'      If Me.textCP01.Text = "FCP" And InStr("209,235", Me.textCP10.Text) > 0 Then
'         cmdPrint201.Caption = "送排版"
'         cmdPrint201.Visible = True
'      Else
'         cmdPrint201.Visible = False
'      End If
   Else
   '2023/10/3 END
      'Added by Lydia 2018/04/11 外專翻譯承辦單列印
      'Memo by Lydia 2020/07/27 拿掉”列印”
      If Me.textCP01.Text = "FCP" And InStr("201,209,235,210", Me.textCP10.Text) > 0 Then
           cmdPrint201.Visible = True
      'Added by Lydia 2019/05/02 會稿Claims/說明書承辦單列印
      ElseIf Me.textCP01.Text = "FCP" And Me.textCP10.Text = "924" Then
           cmdPrint201.Visible = True
           'Modified by Lydia 2020/07/27 拿掉”列印”
           cmdPrint201.Caption = "會稿承辦單(&C)"
      'end 2019/05/02
      Else
           cmdPrint201.Visible = False
      End If
      'end 2018/04/11
   End If
   
   'Added by Morgan 2022/3/18
   If Me.textCP01.Text = "CFP" And Me.textCP10.Text = "605" Then
      Label14(8) = "是否收到收據：       (Y:是)"
   ElseIf Me.textCP01.Text = "T" And Me.textCP10.Text = "102" Then
      Label14(8) = "可辦前已通知：       (Y:是)"
   Else
      Label14(8) = "是否收到副本：       (Y:是)"
   End If
   'end 2022/3/18
   
   ' 顯示商標名稱
   If cmbTM05.ListCount > 0 Then
      cmbTM05.ListIndex = 0
   End If
   
   SetCP71 'Add by Morgan 2004/10/13
   
   'Added by Morgan 2016/6/2
   '開放催提申、收達B類收文可由建立人員自行刪除
   m_bDelete = IsUserHasRightOfFunction("frm075004_2", strDel, False)
   If m_bDelete = False And (textCP01 = "P" Or textCP01 = "CFP" Or textCP01 = "PS" Or textCP01 = "CPS") Then
      'Modified by Morgan 2018/10/12 +催公開954 --郭
      If m_CP65 = strUserNum And Left(textCP09, 1) = "B" And (textCP10 = "952" Or textCP10 = "953" Or textCP10 = "954") Then
         m_bDelete = True
      End If
      UpdateToolbarState
   End If
   'end 2016/6/2
   'Add by Amy 2022/09/02 若為「其他相關人」對造/其他 頁籤 顯示 關係案/其他,「對造」文字->對方
   stRPTxt = "對造"
   SSTab1.TabCaption(3) = "對造/其他"
   If Pub_ChkRelevantPeople(1, textCP09.Text) = True Then
        SSTab1.TabCaption(3) = "關係案/其他"
        stRPTxt = "對方"
   End If
   Call SetLabTxt(stRPTxt)
   'end 2022/09/02
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   If IsNull(rsSrcTmp.Fields("CP65")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP65")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("CP65"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP66")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP66")) = False Then
         strTemp = ChangeWStringToTString(rsSrcTmp.Fields("CP66"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP67")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP67")) = False Then
         strTemp = rsSrcTmp.Fields("CP67")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP68")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP68")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("CP68"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP69")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP69")) = False Then
         strTemp = ChangeWStringToTString(rsSrcTmp.Fields("CP69"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CP70")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CP70")) = False Then
         strTemp = rsSrcTmp.Fields("CP70")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strCP09 As String)
   Dim strTemp As String
   Dim nIndex As Integer
   
   If IsRecordExist(strCP09) = True Then
      For nIndex = 0 To m_DataListCount - 1
         If m_DataList(nIndex).diCP09 = strCP09 Then
            m_CurrDL = nIndex
            Exit For
         End If
      Next nIndex
   Else
      m_CurrDL = 0
      strTemp = Empty
      For nIndex = 0 To m_DataListCount - 1
         If strCP09 > strTemp Then
            m_CurrDL = nIndex
            Exit For
         End If
         strTemp = m_DataList(nIndex).diCP09
      Next nIndex
   End If
   UpdateCtrlData
   
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   If m_DataListCount > 0 Then
      m_CurrDL = 0
   End If
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   If m_CurrDL = 0 Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   If m_CurrDL > 0 Then
      m_CurrDL = m_CurrDL - 1
   End If
   
   UpdateCtrlData
   
EXITSUB:
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   If m_CurrDL >= m_DataListCount - 1 Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   If m_CurrDL < (m_DataListCount - 1) Then
      m_CurrDL = m_CurrDL + 1
   End If
   
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   If m_DataListCount > 0 Then
      m_CurrDL = m_DataListCount - 1
   End If
   
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         ' 90.07.13 modify by louis (依照權限設定其工具列的按紐狀態)
         'tlbar.Buttons(1).Enabled = True
         'tlbar.Buttons(2).Enabled = True
         'tlbar.Buttons(3).Enabled = True
         'tlbar.Buttons(4).Enabled = True
         'tlbar.Buttons(6).Enabled = True
         'tlbar.Buttons(7).Enabled = True
         'tlbar.Buttons(8).Enabled = True
         'tlbar.Buttons(9).Enabled = True
         'tlbar.Buttons(11).Enabled = False
         'tlbar.Buttons(12).Enabled = False
         'tlbar.Buttons(14).Enabled = True
         
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(6).Enabled = True
            tlbar.Buttons(7).Enabled = True
            tlbar.Buttons(8).Enabled = True
            tlbar.Buttons(9).Enabled = True
         Else
            tlbar.Buttons(6).Enabled = False
            tlbar.Buttons(7).Enabled = False
            tlbar.Buttons(8).Enabled = False
            tlbar.Buttons(9).Enabled = False
         End If
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         tlbar.Buttons(1).Enabled = False
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(3).Enabled = False
         tlbar.Buttons(4).Enabled = False
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
   End Select
   
End Sub

Private Sub lstNameAgent_ItemCheck(Item As Integer)
'add by nickc 2006/02/06 不是新增和修改要回復原狀
If lstNameAgent.Enabled Then
    If m_EditMode <> 1 And m_EditMode <> 2 Then
        lstNameAgent.Selected(Item) = Not lstNameAgent.Selected(Item)
    End If
End If
End Sub

'add by nickc 2006/01/26
'檢查並設定cp110資料
'Removed by Morgan 2024/1/8 存檔前也會設定，此處可取消
'Private Sub lstNameAgent_Validate(Cancel As Boolean)
'   Dim ii As Integer, bolCheck As Boolean
'   If m_EditMode <> 1 And m_EditMode <> 2 Then
'      bolCheck = False
'      m_CP110 = ""
'      For ii = 0 To lstNameAgent.ListCount - 1
'         If lstNameAgent.Selected(ii) = True Then
'            'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
'            'm_CP110 = m_CP110 & "," & lstNameAgent.ITEMDATA(ii)
'            'Modified by Lydia 2021/10/08 改模組
'            'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
'            m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
'            bolCheck = True
'         End If
'      Next
'      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
'      '2008/8/5 ADD BY SONIA 未發文或申請國家非台灣案者不存CP22
'      If IsEmptyText(textCP27) = True Then bolCheck = True
'      If m_Nation <> "000" Then bolCheck = True
'      '2008/8/5 END
'      If bolCheck = True Then
'         Me.textCP22 = ""
'      Else
'         textCP22 = "N"
'      End If
'   End If
'End Sub

'Add By Sindy 2009/05/11
Private Sub lstNameOrg_ItemCheck(Item As Integer)
'不是新增和修改要回復原狀
If lstNameOrg.Enabled Then
    If m_EditMode <> 1 And m_EditMode <> 2 Then
        lstNameOrg.Selected(Item) = Not lstNameOrg.Selected(Item)
    End If
End If
End Sub
'檢查並設定cp130資料
Private Sub lstNameOrg_Validate(Cancel As Boolean)
   Dim ii As Integer
   If m_EditMode <> 1 And m_EditMode <> 2 Then
      m_CP130 = ""
      For ii = 0 To lstNameOrg.ListCount - 1
         If lstNameOrg.Selected(ii) = True Then
            m_CP130 = m_CP130 & "," & lstNameOrg.ITEMDATA(ii)
         End If
      Next
      If Left(m_CP130, 1) = "," Then m_CP130 = Mid(m_CP130, 2)
   End If
End Sub
'2009/05/11 End

Private Sub SSTab1_Click(PreviousTab As Integer)
'add by nick 2004/08/18
If frm075004_2.Visible = True Then
    Select Case Me.SSTab1.Tab
    Case 0 '基本資料
        If Me.textCP05.Enabled = True Then
            Me.textCP05.SetFocus
        End If
    Case 1 '相關資料
        If Me.textCP16.Enabled = True Then
            Me.textCP16.SetFocus
        End If
    Case 2 '移轉/授權
        If Me.textCP55.Enabled = True Then
            Me.textCP55.SetFocus
        End If
    'add by nick 2004/08/18
    Case 3 '對造
        If Me.textCP36.Enabled = True Then
            Me.textCP36.SetFocus
        End If
    End Select
End If
End Sub

Private Sub textCP01_Change()
    Select Case Me.textCP01
    Case "T", "FCT", "CFT", "TF"
        Me.Label18(1).Visible = False
        Me.textCP37.Visible = False
        Me.textCP37.Enabled = False
        Me.Label18(2).Visible = False
        Me.textCP38.Visible = False
        Me.textCP38.Enabled = False
        Me.Label18(3).Visible = False
        Me.textCP39.Visible = False
        Me.textCP39.Enabled = False
        Me.Label18(5).Visible = True
        Me.textCP37_1.Visible = True
        Me.textCP37_1.Enabled = True
    Case Else
        Me.Label18(1).Visible = True
        Me.textCP37.Visible = True
        Me.textCP37.Enabled = True
        Me.Label18(2).Visible = True
        Me.textCP38.Visible = True
        Me.textCP38.Enabled = True
        Me.Label18(3).Visible = True
        Me.textCP39.Visible = True
        Me.textCP39.Enabled = True
        Me.Label18(5).Visible = False
        Me.textCP37_1.Visible = False
        Me.textCP37_1.Enabled = False
    End Select
End Sub

' 收文日
Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsTaiwanDate(textCP05, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
      End If
      'Add By Sindy 2019/8/29
      If Val(DBDATE(textCP05)) > Val(strSrvDate(1)) Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文日不可大於系統日！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
      End If
      '2019/8/29 END
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If m_EditMode = 1 Or m_EditMode = 2 Then  'Added by Lydia 2020/07/13 新增或修改
        If IsEmptyText(textCP06) = False Then
           If CheckIsTaiwanDate(textCP06, False) = False Then
              Cancel = True
              strTit = "檢核資料"
              strMsg = "本所期限日期格式不正確"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              textCP06_GotFocus
           'Added by Lydia 2020/07/13 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天;因為FCP案有輸入非工作日的需求,彈訊息
           Else
               If DBDATE(textCP06.Text) <> PUB_GetWorkDay1(textCP06.Text, 1) Then
                  'Added by Lydia 2025/03/14 建檔日為114/1/1以後,直接鎖定
                  If "" & m_FieldList(65).fiOldData >= "20250101" Then
                     MsgBox "本所期限只能輸入工作天！", vbExclamation + vbOKOnly
                     Cancel = True
                     textCP06.SetFocus
                     textCP06_GotFocus
                     Exit Sub
                  End If
                  'end 2025/03/14
                  If MsgBox("本所期限非工作天，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                     Cancel = True
                     textCP06.SetFocus
                     textCP06_GotFocus
                     Exit Sub
                  'Add By Sindy 2023/8/25
                  Else
                     Me.textCP09.SetFocus 'SetFocus離開此欄位,訊息才不會一直重覆詢問,離不開
                     '2023/8/25 END
                  End If
               End If
           'end 2020/07/13
           End If
        End If
   End If 'Added by Lydia 2020/07/13
End Sub

' 法定期限
Private Sub textCP07_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsTaiwanDate(textCP07, False) = False Then
    '     Cancel = True
         strTit = "檢核資料"
         strMsg = "法定期限日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07.SetFocus
         Exit Sub
      End If
   End If
   
   'Modified by Morgan 2012/12/12 有變更才要檢查
   If textCP06.Tag <> Me.textCP06 Or textCP07.Tag <> Me.textCP07 Then
      ' 本所期限不可超過法定期限
      If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
         If Val(DBDATE(textCP06)) > Val(DBDATE(textCP07)) Then
        '    Cancel = True
            strTit = "檢核資料"
            strMsg = "本所期限不可超過法定期限"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP06.SetFocus
         End If
      End If
   End If
End Sub

' 機關文號
Private Sub textCP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modify by Morgan 2011/1/3 機關文號欄位改長度(百年問題)改抓MaxLength屬性控制
   If CheckLengthIsOK(textCP08, textCP08.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "機關文號內容太長"
      textCP08_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP08.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCP09_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 案件性質
Private Sub textCP10_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
    Cancel = False
    textCP10_2 = Empty
    If IsEmptyText(textCP10) = False Then
        If m_Nation < "010" Then
            textCP10_2 = GetCaseTypeName(m_CP01, textCP10, 0)
        Else
            textCP10_2 = GetCaseTypeName(m_CP01, textCP10, 1)
        End If
        Select Case m_EditMode
        Case 1, 2:
            If IsEmptyText(textCP10_2) = True Then
                Cancel = True
                strTit = "檢核資料"
                strMsg = "案件性質代號不存在"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textCP10_GotFocus
            End If
        End Select
    End If
    'Add By Cheng 2003/08/20
    Select Case Me.textCP01.Text
    Case "P", "CFP", "FCP"
        'edit by nickc 2006/09/01 不鎖定只有核准，因為不續辦、取消收文、閉卷也要
        'Select Case Me.textCP10.Text
        'Case "1001", "1002"
            If Me.textCP10_2.Text <> "" Then Me.textCP10_2.Text = Me.textCP10_2.Text & PUB_GetRelateCasePropertyName(Me.textCP09.Text, "1")
        'End Select
    Case "T", "TF", "CFT", "FCT"
        'edit by nickc 2006/09/01 不鎖定只有核准，因為不續辦、取消收文、閉卷也要
        'Select Case Me.textCP10.Text
        'Case "1001", "1002", "1003", "1004"
            If Me.textCP10_2.Text <> "" Then Me.textCP10_2.Text = Me.textCP10_2.Text & PUB_GetRelateCasePropertyName(Me.textCP09.Text, "1")
        'End Select
    End Select
    'Modify by Morgan 2004/10/13 改sub
    SetCP71
    OptSendType(1).Caption = PUB_GetCP114Opt1Desc(textCP01, textCP10)  'Added by Morgan 2024/1/22
End Sub
'Add by Morgan 2004/10/13
Private Sub SetCP71()
   'Add by Morgan 2004/9/29
   If textCP01.Text = "P" And textCP10.Text = "412" Then
      lblCP71.Caption = "延緩公告月數/日期："
      textCP71.MaxLength = 7
   'Add by Morgan 2010/6/3
   ElseIf textCP01.Text = "CFP" And textCP10.Text = "106" Then
      lblCP71.Caption = "是否需直譯本："
      textCP71.MaxLength = 1
   'Added by Morgan 2012/4/25
   'modify by sonia 2013/3/21 加436
   'modify by sonia 2016/7/28 加437
   ElseIf textCP01.Text = "P" And (textCP10.Text = "405" Or textCP10.Text = "436" Or textCP10.Text = "437") Then
      lblCP71.Caption = "優先權份數："
      textCP71.MaxLength = 2
   'end 2012/4/25
   
   'Added by Morgan 2020/2/4
   ElseIf textCP01.Text = "P" And textCP10.Text = "404" Then
      lblCP71.Caption = "延期月數："
      textCP71.MaxLength = 2
   'end 2020/2/4
   'Added by Lydia 2025/02/12 P臺灣與大陸案若申請延緩審查，請於發文時讓user輸入延緩審查日期; FCP案在核准時輸入
   ElseIf (textCP01 = "P" And textCP10.Text = "245") Or (textCP01 = "FCP" And textCP10.Text = "1924") Then
      If m_Nation = "000" Then
         lblCP71.Caption = "延緩審查日期："
         textCP71.MaxLength = 7
      Else
         lblCP71.Caption = "延緩審查日期(年度)："
         textCP71.MaxLength = 1
      End If

   'end 2025/02/12
   Else
      lblCP71.Caption = "機關代號："
      textCP71.MaxLength = 5
   End If
   textCP71.Left = lblCP71.Left + lblCP71.Width + 50 'Added by Morgan 2012/4/25
End Sub
' 案件來源代號
Private Sub textCP11_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCP11_2 = Empty
   If IsEmptyText(textCP11) = False Then
      strSql = "SELECT * FROM CASESOURCEMAP " & _
               "WHERE CSM01 = '" & textCP11 & "' "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CSM02")) = False Then
            textCP11_2 = rsTmp.Fields("CSM02")
         End If
      Else
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "案件來源代號不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP11_GotFocus
         End Select
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Sub

Private Sub textCP118_GotFocus()
   TextInverse textCP118
End Sub

Private Sub textCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add by Amy 2016/07/07 台灣不能設定電子送件--郭雅娟
   'Modified by Morgan 2018/10/8 行政訴訟503改可電子送--陳玲玲 Ex:P-96988
   If textCP01 = "P" And m_Nation = "000" And (textCP10 = "803" Or textCP10 = "804" Or textCP10 = "501") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   'Modified by Morgan 2013/5/16 +W
   ElseIf KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("A") And KeyAscii <> Asc("W") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCP120_GotFocus()
   TextInverse textCP120
End Sub

Private Sub textCP120_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCP121_GotFocus()
   TextInverse textCP121
End Sub

Private Sub textCP121_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCP117_GotFocus()
   TextInverse textCP117
End Sub

Private Sub textCP119_GotFocus()
   InverseTextBox textCP119
End Sub

'2008/8/27 add by sonia 櫃台收文日
Private Sub textCP119_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP119) = False Then
      If CheckIsTaiwanDate(textCP119, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "來函櫃台收文日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP119_GotFocus
      End If
   End If
End Sub

Private Sub textCP12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 業務區別
Private Sub textCP12_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP12_2 = Empty
   If IsEmptyText(textCP12) = False Then
      textCP12_2 = GetDepartmentName(textCP12)
      If IsEmptyText(textCP12_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "業務區別代號不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP12_GotFocus
         End Select
      End If
   End If
End Sub

Private Sub textCP123_GotFocus()
   TextInverse textCP123
End Sub

Private Sub textCP123_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub


Private Sub textCP124_GotFocus()
   TextInverse textCP124
End Sub

Private Sub textCP124_Validate(Cancel As Boolean)
   If textCP124 <> Empty Then
      If Not ChkDate(textCP124) Then
         Cancel = True
      End If
   End If
End Sub

Private Sub textCP126_GotFocus()
   TextInverse textCP126
End Sub

Private Sub textCP126_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCP127_GotFocus()
   TextInverse textCP127
End Sub

Private Sub textCP127_Validate(Cancel As Boolean)
   If textCP127 <> Empty Then
      If Not ChkDate(textCP127) Then
         Cancel = True
      End If
   End If
End Sub

Private Sub textCP129_GotFocus()
   TextInverse textCP129
End Sub

Private Sub textCP129_Validate(Cancel As Boolean)
   If textCP129 <> Empty Then
      If Not ChkDate(textCP129) Then
         Cancel = True
      End If
   End If
End Sub

'Add By Sindy 2010/11/25
Private Sub textCP13_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2009/05/04
Private Sub textCP131_GotFocus()
   TextInverse textCP131
   OpenIme
End Sub
Private Sub textCP131_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   Cancel = False
   If IsEmptyText(textCP131) = False Then
      If CheckLengthIsOK(textCP131, 100) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "發文室取消發文備註內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP131_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub
Private Sub textCP132_GotFocus()
   TextInverse textCP132
End Sub
Private Sub textCP132_Validate(Cancel As Boolean)
   If textCP132 <> Empty Then
      If Not ChkDate(textCP132) Then
         Cancel = True
      End If
   End If
End Sub
'2009/05/04 End

' 智權人員代號
Private Sub textCP13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP13_2 = Empty
   If IsEmptyText(textCP13) = False Then
      textCP13_2 = GetStaffName(textCP13, True)
      If IsEmptyText(textCP13_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "智權人員代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP13_GotFocus
         End Select
      Else
         Select Case m_EditMode
            Case 1, 2:
               'Modify by Morgan 2010/1/5 改智權人員也要重抓
               'If textCP12 = "" Then
               If textCP12 = "" Or textCP13.Tag <> textCP13 Then
               'end 2010/1/5
                  textCP12 = GetST15(textCP13)
               End If
               textCP12_Validate False
            Case Else:
         End Select
      End If
   End If
   If Cancel = False Then
      textCP13.Tag = textCP13
   End If
End Sub

Private Sub textCP135_GotFocus()
   TextInverse textCP135
   CloseIme
End Sub

Private Sub textCP135_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub textCP136_GotFocus()
   TextInverse textCP136
   CloseIme
End Sub

Private Sub textCP136_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub textCP137_GotFocus()
   TextInverse textCP137
   CloseIme
End Sub

Private Sub textCP137_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub textCP138_GotFocus()
   TextInverse textCP138
   CloseIme
End Sub

Private Sub textCP138_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

'Add By Sindy 2023/4/13
Private Sub textCP167_GotFocus()
   TextInverse textCP167
   CloseIme
End Sub
Private Sub textCP167_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub textCP168_GotFocus()
   TextInverse textCP168
   CloseIme
End Sub
Private Sub textCP168_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
'2023/4/13 END

'Add By Sindy 2010/11/25
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14, True)
      If IsEmptyText(textCP14_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "承辦人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP14_GotFocus
         End Select
      End If
   End If
End Sub

Private Sub textCP144_GotFocus()
   InverseTextBox textCP144
   OpenIme
End Sub

'報價備註 2011/5/26 add by sonia
Private Sub textCP144_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP144, textCP144.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "報價備註內容太長"
      textCP144_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub
'2011/5/26 end

'Added by Morgan 2016/6/2
Private Sub textCP145_GotFocus()
   TextInverse textCP145
End Sub

Private Sub textCP145_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2016/6/2

'Add By Sindy 2015/6/3
Private Sub TextCP148_GotFocus()
   TextInverse textCP148
End Sub
Private Sub TextCP148_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'2015/6/3 END

Private Sub textCP152_GotFocus()
   InverseTextBox textCP152
End Sub

Private Sub textCP152_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP152) = False Then
      If CheckIsTaiwanDate(textCP152, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "扣款日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP152_GotFocus
      End If
   End If
End Sub

' 費用
Private Sub textCP16_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP16) = False Then
      If IsNumeric(textCP16) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP16_GotFocus
      End If
   End If
End Sub

' 規費
Private Sub textCP17_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP17) = False Then
      If IsNumeric(textCP17) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "規費只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP17_GotFocus
      End If
   End If
End Sub

' 點數
Private Sub textCP18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP18) = False Then
      If IsNumeric(textCP18) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "點數只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP18_GotFocus
      End If
   End If
End Sub

Private Sub textCP20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

' 是否向客戶收款
Private Sub textCP20_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP20) = False Then
      Select Case textCP20
         Case "N":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否向客戶收款只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP20_GotFocus
      End Select
   End If
End Sub

Private Sub textCP21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      'Add By Sindy 2023/3/15
      If InStr(Label11.Caption, "快軌") > 0 Then
         If KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      Else
      '2023/3/15 END
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

' 是否多國/是否取締案
Private Sub textCP21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP21) = False Then
      'Add By Sindy 2023/3/15
      If InStr(Label11.Caption, "快軌") > 0 Then
         Select Case textCP21
            Case "Y", "N"
            Case Else:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "是否為快軌案件只可輸入Y或N"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP21_GotFocus
         End Select
      Else
      '2023/3/15 END
         Select Case textCP21
            Case "Y":
            Case Else:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "是否多國/是否取締案只可輸入空白或Y"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP21_GotFocus
         End Select
      End If
   Else
      'Add By Sindy 2023/3/15
      If InStr(Label11.Caption, "快軌") > 0 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "是否為快軌案件不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP21_GotFocus
      End If
      '2023/3/15 END
   End If
End Sub

Private Sub textCP22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCP23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modify By Sindy 98/04/13
   'If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'edit by nickc 2006/01/27
'Private Sub textCP22_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

' 是否出名
'Private Sub textCP22_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP22) = False Then
'      Select Case textCP22
'         Case "N":
'         Case Else:
'            Cancel = True
'            strTit = "檢核資料"
'            strMsg = "是否出名只可輸入空白或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCP22_GotFocus
'      End Select
'   End If
'End Sub

' 預估結果
Private Sub textCP23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP23) = False Then
      Select Case textCP23
         Case "1", "2", "3": 'Modify By Sindy 98/04/13 增加3
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "預估結果只可輸入1或2或3" 'Modify By Sindy 98/04/13 增加3
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP23_GotFocus
      End Select
   End If
End Sub

Private Sub textCP24_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'modify by sonia 2024/8/2 加可輸入3
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
      Beep
   End If
End Sub

' 實際結果
Private Sub textCP24_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP24) = False Then
      Select Case textCP24
         Case "1", "2", "3":    'modify by sonia 2024/8/2 加可輸入3
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "實際結果只可輸入1或2或3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP24_GotFocus
      End Select
   End If
End Sub

' 准駁日
Private Sub textCP25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP25) = False Then
      If CheckIsTaiwanDate(textCP25, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "准駁日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25_GotFocus
      End If
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case "N":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否算案件數只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP27) = False Then
      If CheckIsTaiwanDate(textCP27, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "發文日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2010/11/25
Private Sub textCP29_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 繪圖人員/協辦人員
Private Sub textCP29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP29_2 = Empty
   If IsEmptyText(textCP29) = False Then
      textCP29_2 = GetStaffName(textCP29, True)
      If IsEmptyText(textCP29_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               'Modified by Lydia 2015/10/05
               'strMsg = "繪圖人員/法務人員代號不存在"
               strMsg = "繪圖人員/協辦人員代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP29_GotFocus
         End Select
      End If
   End If
End Sub

Private Sub textCP31_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

' 是否為新案件
Private Sub textCP31_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP31) = False Then
      Select Case textCP31
         Case "Y":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否為新案件只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP31_GotFocus
      End Select
   End If
End Sub

Private Sub textCP32_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否開立電腦收據
Private Sub textCP32_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP32) = False Then
      Select Case textCP32
         Case "N":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否開立電腦收據只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP32_GotFocus
      End Select
   End If
End Sub

' 標準價
Private Sub textCP33_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP33) = False Then
      If IsNumeric(textCP33) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "標準價只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP33_GotFocus
      End If
   End If
End Sub

' 底價
Private Sub textCP34_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP34) = False Then
      If IsNumeric(textCP34) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "底價只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP34_GotFocus
      End If
   End If
End Sub

Private Sub textCP37_1_GotFocus()
    TextInverse Me.textCP37_1
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textCP37_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP37_1) = False Then
      If CheckLengthIsOK(textCP37_1, 140) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "對造案件名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP37_1_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

' 對造案件名稱(中)
Private Sub textCP37_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP37) = False Then
      'Modify by Amy 2015/03/11 長度原:100
      'Removed by Morgan 2023/8/7  欄位已改char長度等於字數不必再檢查(欄位長度自動會限制內容)
      'If CheckLengthIsOK(textCP37, 140) = False Then
      '   Cancel = True
      '   strTit = "檢核資料"
      '   strMsg = "對造案件名稱(中)內容太長"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP37_GotFocus
      'End If
      'end 2023
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP37.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 對造案件名稱(英)
Private Sub textCP38_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP38) = False Then
      'Removed by Morgan 2023/8/7  欄位已改char長度等於字數不必再檢查(欄位長度自動會限制內容)
      'If CheckLengthIsOK(textCP38, 100) = False Then
      '   Cancel = True
      '   strTit = "檢核資料"
      '   strMsg = "對造案件名稱(英)內容太長"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP38_GotFocus
      'End If
      'end 2023/8/7
   End If
End Sub

' 對造案件名稱(日)
Private Sub textCP39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP39) = False Then
      'Removed by Morgan 2023/8/7  欄位已改char長度等於字數不必再檢查(欄位長度自動會限制內容)
      'If CheckLengthIsOK(textCP39, 100) = False Then
      '   Cancel = True
      '   strTit = "檢核資料"
      '   strMsg = "對造案件名稱(日)內容太長"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP39_GotFocus
      'End If
      'end 2023/8/7
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP39.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 對造名稱(中)
Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP40) = False Then
      If CheckLengthIsOK(textCP40, 600) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "對造名稱(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP40_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP40.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 對造名稱(英)
Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP41) = False Then
      If CheckLengthIsOK(textCP41, 600) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "對造名稱(英)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP41_GotFocus
      End If
   End If
End Sub

' 對造名稱(日)
Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP42) = False Then
      If CheckLengthIsOK(textCP42, 600) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "對造名稱(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP42_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP42.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCP43_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 相關總收文號
Private Sub textCP43_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP43) = False Then
      Select Case m_EditMode
         Case 1, 2:
            'strSQL = "SELECT * FROM CASEPROGRESS " & _
            '         "WHERE CP01 = '" & m_CP01 & "' AND " & _
            '               "CP02 = '" & m_CP02 & "' AND " & _
            '               "CP03 = '" & m_CP03 & "' AND " & _
            '               "CP04 = '" & m_CP04 & "' AND " & _
            '               "CP09 = '" & textCP43 & "' "
            'Modified by Morgan 2012/5/25 +考慮香港大陸案
            strSql = "SELECT CP01,CP02,CP03,CP04,CM01,CM02,CM03,CM04 FROM CASEPROGRESS,CASEMAP " & _
                     "WHERE CP09 = '" & textCP43 & "' AND CM05(+)=CP01 AND CM06(+)=CP02 AND CM07(+)=CP03 AND CM08(+)=CP04 AND CM10(+)='4'"
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               'Modify by Morgan 2010/2/24 不必判斷多國號(子案的相關總收文號可能是母案收文號)
               'If rsTmp.Fields("CP01") <> m_CP01 Or rsTmp.Fields("CP02") <> m_CP02 Or rsTmp.Fields("CP03") <> m_CP03 Or rsTmp.Fields("CP04") <> m_CP04 Then
               If rsTmp("CP01") & rsTmp("CP02") & rsTmp("CP03") <> m_CP01 & m_CP02 & m_CP03 And rsTmp("CM01") & rsTmp("CM02") & rsTmp("CM03") <> m_CP01 & m_CP02 & m_CP03 Then
                  
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "相關總收文號與本案不符"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textCP43_GotFocus
               End If
            Else
               Cancel = True
               strTit = "檢核資料"
               strMsg = "相關總收文號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP43_GotFocus
            End If
            rsTmp.Close
            Set rsTmp = Nothing
      End Select
   End If
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 代理人
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   Cancel = False
   textCP44_2 = Empty
   If IsEmptyText(textCP44) = False Then
      'Add by Morgan 2008/5/14 +聯絡人
      If InStr(textCP44, "-") > 0 Then
         If ClsPDGetContact(textCP44, strTempName) Then
            'modify by sonia 2017/11/15
            'textCP44_2 = strTempName
            If PUB_GetAgentName(Me.textCP01.Text, Left(textCP44, InStr(textCP44, "-") - 1), strTempName) = True Then
               textCP44_2 = strTempName
            End If
            If ClsPDGetContact(textCP44, strTempName) Then
               textCP44_2 = textCP44_2 & "(" & strTempName & ")"
            End If
            'end 2017/11/15
         Else
            If textCP44.Locked = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "聯絡人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP44_GotFocus
            End If
         End If
      Else
      'end 2008/5/14
         If PUB_GetAgentName(Me.textCP01.Text, Me.textCP44.Text, strTempName) = True Then
            textCP44_2 = strTempName
         Else
            textCP44_2 = ""
         End If
         If IsEmptyText(textCP44_2) = True Then
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "代理人代號不存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textCP44_GotFocus
            End Select
         End If
      End If
   End If
End Sub

' 代理人收達日
Private Sub textCP46_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP46) = False Then
      If CheckIsTaiwanDate(textCP46, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         'Modified by Lydia 2016/05/30
         'strMsg = "代理人收達日日期格式不正確"
         strMsg = Mid(Label30(2), 1, Len(Label30(2)) - 1) & "日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP46_GotFocus
      'Added by Lydia 2016/05/30
      Else
         If m_SK02 = 3 Or m_SK02 = 4 Then
            If Val(textCP46) = 111111 Then
               Label30(3) = "回執退件日："
            ElseIf Val(textCP46) = 110101 Then
               Label30(3) = "回執未回郵局送達日："
            End If
         End If
      End If
   End If
End Sub

' 代理人提申日
Private Sub textCP47_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP47) = False Then
      If CheckIsTaiwanDate(textCP47, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         'Modified by Lydia 2016/05/30
         'strMsg = "代理人提申日日期格式不正確"
         strMsg = Mid(Label30(3), 1, Len(Label30(3)) - 1) & "日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP47_GotFocus
      End If
   End If
End Sub

' 承辦期限
Private Sub textCP48_Validate(Cancel As Boolean)
   If m_EditMode = 0 Then Exit Sub 'Add by Morgan 2009/6/25
   
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "承辦期限日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
         Exit Sub
      End If
   End If
   
   
   If textCP48.Locked = False Then 'Add by Morgan 2011/1/19
      If GetOldData("CP48") <> DBDATE(textCP48) Or GetOldData("CP06") <> DBDATE(textCP06) Then 'Added by Morgan 2013/5/9 欄位值有變更才檢查
         'Add By Cheng 2002/05/07
         '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
         If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
            If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
               Cancel = True
               MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
               textCP48_GotFocus
               Exit Sub
            End If
         End If
         'Added by Lydia 2025/03/14 若承辦期限非工作天,彈訊息;建檔日為114/1/1以後,直接鎖定
         If DBDATE(textCP48.Text) <> PUB_GetWorkDay1(textCP48.Text, 1) Then
            If "" & m_FieldList(65).fiOldData >= "20250101" Then
               MsgBox "承辦期限只能輸入工作天！", vbExclamation + vbOKOnly
               Cancel = True
               textCP48.SetFocus
               textCP48_GotFocus
               Exit Sub
            Else
               If MsgBox("承辦期限非工作天，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Cancel = True
                  textCP48.SetFocus
                  textCP48_GotFocus
                  Exit Sub
               Else
                  Me.textCP09.SetFocus 'SetFocus離開此欄位,訊息才不會一直重覆詢問,離不開
               End If
            End If
         End If
         'end 2025/03/14
      End If 'Added by Morgan 2013/5/9 欄位值有變更才檢查
   End If 'Add by Morgan 2011/1/19
   
End Sub

' 被授權人(中)
Private Sub textCP50_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP50) = False Then
      If CheckLengthIsOK(textCP50, 60) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "被授權人(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP50_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP50.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 被授權人(英)
Private Sub textCP51_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP51) = False Then
      If CheckLengthIsOK(textCP51, 60) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "被授權人(英)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP51_GotFocus
      End If
   End If
End Sub

' 被授權人(日)
Private Sub textCP52_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP52) = False Then
      If CheckLengthIsOK(textCP52, 60) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "被授權人(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP52_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP52.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Added by Lydia 2017/08/24
Private Sub textCP53_KeyPress(KeyAscii As Integer)
   'TB條碼案繳年費708,服務業務結果1801-第?期登記期
   If m_CP01 = "TB" And (m_CP10 = "708" Or m_CP10 = "1801") Then
      KeyAscii = Pub_NumAscii(KeyAscii)
   End If
End Sub

' 授權期間(起)/質權期間(起)
Private Sub textCP53_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2017/08/24 TB條碼案繳年費708,服務業務結果1801-第?期登記期
   'If IsEmptyText(textCP53) = False Then
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If m_CP01 = "TB" And (m_CP10 = "708" Or m_CP10 = "1801") Then
      If IsEmptyText(textCP53) = True Then
           Cancel = True
           strTit = "檢核資料"
           strMsg = "登記期不可空白"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textCP53_GotFocus
           GoTo EXITSUB
      Else
        If Val(textCP53) > 100 Then
           Cancel = True
           strTit = "檢核資料"
           strMsg = "登記期不正確"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textCP53_GotFocus
           GoTo EXITSUB
        Else
           textCP54.Text = textCP53.Text
        End If
      End If
   ElseIf IsEmptyText(textCP53) = False Then
   'end 2017/08/24
        If CheckIsTaiwanDate(textCP53, False) = False Then
           Cancel = True
           strTit = "檢核資料"
           strMsg = "授權期間(起)/質權期間(起)日期格式不正確"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textCP53_GotFocus
           GoTo EXITSUB
        End If
      '91.5.8 MODIFY SONIA不知為何要檢查此項,
      'If IsEmptyText(m_CP53) = True Then
      '   Cancel = True
      '   strTit = "檢核資料"
      '   strMsg = "授權期間(起)/質權期間(起)日期不正確"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP53_GotFocus
      '   GoTo EXITSUB
      'End If
      'If Val(DBDATE(textCP53)) < Val(DBDATE(m_CP53)) Then
      '   Cancel = True
      '   strTit = "檢核資料"
      '   strMsg = "授權期間(起)/質權期間(起)日期不正確"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP53_GotFocus
      '   GoTo EXITSUB
      'End If
      'If IsEmptyText(m_CP54) = False Then
      '   If Val(DBDATE(textCP53)) > Val(DBDATE(m_CP54)) Then
      '      Cancel = True
      '      strTit = "檢核資料"
      '      strMsg = "授權期間(起)/質權期間(起)日期不正確"
      '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '      textCP53_GotFocus
      '      GoTo EXITSUB
      '   End If
      'End If
      '91.5.8 END
   End If
EXITSUB:
End Sub

' 質權期間(迄)
Private Sub textCP54_lostfocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Cancel = False
   If IsEmptyText(textCP54) = False Then
      If CheckIsTaiwanDate(textCP54, False) = False Then
    '     Cancel = True
         strTit = "檢核資料"
         strMsg = "質權期間(迄)日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP54.SetFocus
         GoTo EXITSUB
      End If
      '91.5.8 MODIFY BY SONIA不知為何要做此檢查
      'If IsEmptyText(m_CP54) = True Then
     ''    Cancel = True
      '   strTit = "檢核資料"
      '   strMsg = "授權期間(迄)/質權期間(迄)日期不正確"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP54.SetFocus
      '   GoTo EXITSUB
      'End If
      'If Val(DBDATE(textCP54)) > Val(DBDATE(m_CP54)) Then
     ''   Cancel = True
      '   strTit = "檢核資料"
      '   strMsg = "授權期間(迄)/質權期間(迄)日期不正確"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP54.SetFocus
      '   GoTo EXITSUB
      'End If
      'If IsEmptyText(m_CP53) = False Then
      '   If Val(DBDATE(textCP54)) < Val(DBDATE(m_CP53)) Then
      '     ' Cancel = True
      '      strTit = "檢核資料"
      '      strMsg = "授權期間(迄)/質權期間(迄)日期不正確"
      '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '      textCP54.SetFocus
      '      GoTo EXITSUB
      '   End If
      'End If
      '91.5.8 END
   End If
   ' 授權期間不正確
   If IsEmptyText(textCP53) = False And IsEmptyText(textCP54) = False Then
      If Val(DBDATE(textCP53)) > Val(DBDATE(textCP54)) Then
        ' Cancel = True
         strTit = "檢核資料"
         strMsg = "授權期間起日不可超過止日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP53.SetFocus
      End If
   End If
EXITSUB:
End Sub

Private Sub textCP55_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 移轉人
Private Sub textCP55_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP55_2 = Empty
   If IsEmptyText(textCP55) = False Then
      textCP55_2 = GetCustomerName(textCP55, 0)
      If IsEmptyText(textCP55_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP55_GotFocus
         End Select
      End If
   End If
End Sub

Private Sub textCP56_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 移轉申請人
Private Sub textCP56_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP56_2 = Empty
   If IsEmptyText(textCP56) = False Then
      textCP56_2 = GetCustomerName(textCP56, 0)
      If IsEmptyText(textCP56_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉申請人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP56_GotFocus
         End Select
      End If
   End If
End Sub

' 取消收文日期
Private Sub textCP57_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP57) = False Then
      If CheckIsTaiwanDate(textCP57, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "取消收文日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP57_GotFocus
      End If
   End If
End Sub

' 取消收文原因
Private Sub textCP58_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCP58_2 = Empty
   If IsEmptyText(textCP58) = False Then
      strSql = "SELECT * FROM REASONOFRELIEF " & _
               "WHERE ROR01 = '" & textCP58 & "' "
      Set rsTmp = New ADODB.Recordset
      ' 讀取資料庫
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      ' 檢查讀取的資料筆數
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("ROR02")) = False Then
            textCP58_2 = rsTmp.Fields("ROR02")
         End If
      Else
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "取消收文原因代號不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP58_GotFocus
         End Select
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Sub

Private Sub textCP59_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 結餘註記
Private Sub textCP59_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP59) = False Then
'      Select Case textCP59
'         Case "1", "2", "", " ":
'         Case Else:
'            Cancel = True
'            strTit = "檢核資料"
'            strMsg = "結餘註記只可輸入空白或1或2!!"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCP59_GotFocus
'      End Select
'   End If
End Sub

'Added by Lydia 2021/10/20 Form 2.0的TextBox增加右鍵選單功能
Private Sub textCP64_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 textCP64
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      textCP64_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP64.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCP71_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add by Morgan 2010/6/3
   If textCP71.MaxLength = 1 Then
      'Added by Lydia 2025/02/12 排除P臺灣與大陸案延緩審查; FCP案在核准時輸入
      If (textCP01 = "P" And textCP10.Text = "245") Or (textCP01 = "FCP" And textCP10.Text = "1924") Then
      Else
      'end 2025/02/12
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
      End If
   End If
End Sub

' 機關代號
Private Sub textCP71_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP71_2 = Empty
   
   'Modify by Morgan 2004/9/30
   '加判斷若為延緩公告時
   If textCP01.Text = "P" And textCP10.Text = "412" Then
   
      If textCP14.Text <> "" And textCP71.Text = "" Then
         MsgBox "該程序已分案，延緩月數/日期不可空白！"
         Cancel = True
         textCP71_GotFocus
      ElseIf Len(textCP71) = 1 Then
         'Modified by Morgan 2016/3/11 105/3/9日起延緩公告最長改6個月(原3個月)
         If Val(textCP71) < 1 Or Val(textCP71) > 6 Then
            MsgBox "延緩公告月份只可輸入1~6！"
            Cancel = True
            textCP71_GotFocus
         End If
         'end 2016/3/11
      ElseIf ChkDate(textCP71) = False Then
         Cancel = True
         textCP71_GotFocus
      End If
   'Add by Morgan 2010/6/3
   ElseIf textCP01.Text = "CFP" And textCP10.Text = "106" Then
   '不用檢查
   'Added by Morgan 2012/9/10
   'modify by sonia 2013/3/21 加436
   'modify by sonia 2016/7/28 加437
   'Modified by Morgan 2020/2/4 加404
   ElseIf (textCP01.Text = "CFP" Or textCP01.Text = "P" Or textCP01.Text = "FCP") And (textCP10.Text = "405" Or textCP10.Text = "436" Or textCP10.Text = "437" Or textCP10.Text = "404") Then
   '不用檢查
   'Added by Lydia 2025/02/12 P臺灣與大陸案若申請延緩審查，請於發文時讓user輸入延緩審查日期; FCP案在核准時輸入
   ElseIf (textCP01 = "P" And textCP10.Text = "245" And textCP27 <> "") Or (textCP01 = "FCP" And textCP10.Text = "1924") Then
      If m_Nation = "000" Then
         '臺灣：可指定日期
         If CheckIsTaiwanDate(textCP71) = False Then
            Cancel = True
            textCP71_GotFocus
         End If
      Else
         '大陸：以年度計，只可延緩1或2或3年
         If textCP71 <> "1" And textCP71 <> "2" And textCP71 <> "3" Then
            MsgBox "大陸案申請延緩審查以年度計，只可延緩1或2或3年", vbCritical
            Cancel = True
            textCP71_GotFocus
         End If
      End If
   'end 2025/02/12
   Else
   
      If IsEmptyText(textCP71) = False Then
         strSql = "SELECT * FROM ORGANIZATION " & _
                  "WHERE OR01 = '" & textCP71 & "' "
         Set rsTmp = New ADODB.Recordset
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields("OR02")) = False Then
               textCP71_2 = rsTmp.Fields("OR02")
            End If
         Else
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "機關代號不存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textCP71_GotFocus
            End Select
         End If
         rsTmp.Close
         Set rsTmp = Nothing
      End If
   End If
End Sub

Private Sub textCP72_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 被授權人
Private Sub textCP72_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strData As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP72) = False Then
      Select Case m_EditMode
         Case 1, 2:
            strData = textCP72 & String(9 - Len(textCP72), "0")
            strSql = "SELECT * FROM CUSTOMER " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '" & Mid(strData, 9, 1) & "' "
            Set rsTmp = New ADODB.Recordset
            ' 讀取資料庫
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
            ' 檢查讀取的資料 當被授權人中英日為空白時才取代
            If rsTmp.RecordCount > 0 Then
               If IsEmptyText(textCP50) = True Then
                  If IsNull(rsTmp.Fields("CU04")) = False Then
                     textCP50 = rsTmp.Fields("CU04")
                  End If
               End If
               If IsEmptyText(textCP51) = True Then
                  If IsNull(rsTmp.Fields("CU05")) = False Then
                     textCP51 = rsTmp.Fields("CU05")
                  End If
                  '2008/2/21 ADD BY SONIA
                  If IsNull(rsTmp.Fields("CU88")) = False Then
                     textCP51 = textCP51 & " " & rsTmp.Fields("CU88")
                  End If
                  If IsNull(rsTmp.Fields("CU89")) = False Then
                     textCP51 = textCP51 & " " & rsTmp.Fields("CU89")
                  End If
                  If IsNull(rsTmp.Fields("CU90")) = False Then
                     textCP51 = textCP51 & " " & rsTmp.Fields("CU90")
                  End If
                  '2008/2/21 END
               End If
               If IsEmptyText(textCP52) = True Then
                  If IsNull(rsTmp.Fields("CU06")) = False Then
                     textCP52 = rsTmp.Fields("CU06")
                  End If
               End If
            Else
               Cancel = True
               strTit = "檢核資料"
               strMsg = "被授權人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP72_GotFocus
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         Case Else:
      End Select
   End If
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Memo by Lydia 2021/10/20 原程式搬到Form_KeyUp

   Call PUB_SaveMeTrackMode(m_MeTrackMode, 0, KeyCode)  'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   m_SubMode = 0
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         'Modified by Morgan 2013/4/24 外專程序FMP案只能修改是否向客戶收款欄位
         'SetCtrlReadOnly False
'         If Pub_StrUserSt03 = "F22" And textCP01 <> "FCP" And textCP01 <> "FG" Then
'            textCP20.Locked = False
'            '2015/1/30 ADD BY SONIA 再開放外專程序可改FMP會稿924的本所期限及法定期限
'            If textCP10 = "924" Then
'               textCP06.Locked = False
'               textCP07.Locked = False
'            End If
'            '2015/1/30 END
'         Else
'            SetCtrlReadOnly False
'         End If
         'Modified by Lydia 2015/02/04 外專程序FMP案開放可全部修改
         SetCtrlReadOnly False
         
         'end 2013/4/24
         SetKeyReadOnly True
         UpdateToolbarState
         'SetInputEntry
         'Add By Cheng 2002/01/14
         '記錄本所期限及法定期限的初始值
         Me.textCP06.Tag = Me.textCP06.Text
         Me.textCP07.Tag = Me.textCP07.Text
      ' 刪除
      Case vbKeyF5:
        'Modify By Cheng 2003/03/14
        '已請款資料不可刪除
        If CP60IsNull(Me.textCP09.Text) = True Then
            strTit = "詢問"
            strMsg = "是否要刪除此筆資料?"
            'Added by Morgan 2012/11/2
            If textCP28 <> "" Then
               strMsg = "本案已有發文字號，" & strMsg
            End If
            'end 2012/11/2
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbYes Then
               m_EditMode = 3
               OnWork
               If m_DataListCount <= 0 Then
                  GoTo EXITSUB
               Else
                  UpdateToolbarState
               End If
            End If
        End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         'Added by Lydia 2021/10/20 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkMeTrackMode(m_MeTrackMode) = False Then
             Exit Sub
         End If
         'end 2021/10/20
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         'UpdateFieldNewData
         OnWork
         UpdateToolbarState
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         'Unload Me
         'Modify By Sindy 2013/10/25
         'If UCase(m_PrevFormNm) = UCase("frm040111") Then
         If UCase(TypeName(m_PrevForm)) = UCase("frm040111") Then 'Modify by Sindy 2018/10/9
            m_PrevForm.Show
            Call m_PrevForm.cmdQuery_Click
         Else
         '2013/10/25 END
            frm075004_1.ClearRemark
            frm075004_1.Show
            If m_AddData = True Then: frm075004_1.RefreshList
            'Add By Cheng 2002/12/11
            frm075004_1.textCP02.SetFocus
            TextInverse frm075004_1.textCP02
         End If
         Unload Me
   End Select
EXITSUB:
End Sub

Private Sub textCP80_GotFocus()
   InverseTextBox textCP80
End Sub

'對造案件商品類別
Private Sub textCP80_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP80) = False Then
      If CheckLengthIsOK(textCP80, 39) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "對造案件商品類別內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   If Cancel Then TextInverse textCP80
End Sub

'add by nick 2004/08/18
Private Sub textCP82_GotFocus()
   InverseTextBox textCP82
End Sub

Private Sub textCP82_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'add by nick 2004/08/23
   If textCP82.Enabled = False Then
        Exit Sub
   End If
   If IsEmptyText(textCP82) = False Then
      If IsNumeric(textCP82) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "時間只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP82_GotFocus
      Else
        If Len(textCP82) <> 6 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "時間必須 6 碼"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP82_GotFocus
        Else
            If Val(Mid(textCP82, 1, 2)) > 24 Or Val(Mid(textCP82, 3, 2)) > 60 Or Val(Mid(textCP82, 5, 2)) > 60 Then
                Cancel = True
                strTit = "檢核資料"
                strMsg = "時間格式不對"
                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                textCP82_GotFocus
            End If
        End If
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP83_GotFocus()
   InverseTextBox textCP83
End Sub

Private Sub textCP83_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP83_2 = Empty
   If IsEmptyText(textCP83) = False Then
      textCP83_2 = GetStaffName(textCP83, True)
      If IsEmptyText(textCP83_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "發文操作人員代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP13_GotFocus
         End Select
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP84_GotFocus()
   InverseTextBox textCP84
End Sub

Private Sub textCP84_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP84) = False Then
      If IsNumeric(textCP84) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "發文規費只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP84_GotFocus
      End If
   End If
End Sub

'Remove by Morgan 2010/12/30 目前沒用
''add by nick 2004/08/18
'Private Sub textCP85_GotFocus()
'   InverseTextBox textCP85
'End Sub
'
'Private Sub textCP85_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP85) = False Then
'      If CheckIsTaiwanDate(textCP85, False) = False Then
'         Cancel = True
'         strTit = "檢核資料"
'         strMsg = "文卷室發文日日期格式不正確"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP85_GotFocus
'      End If
'   End If
'
'End Sub

'add by nick 2004/08/18
Private Sub textCP86_GotFocus()
   InverseTextBox textCP86
End Sub

Private Sub textCP86_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If textCP01 = "FCP" And textCP10 = "908" Then
      If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
         KeyAscii = 0
         Beep
      End If
   Else
      If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub textCP86_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP86) = False Then
      If Not (m_CP01 = "FCP" And m_CP10 = "908") Then
      Select Case textCP86
         Case "Y", "N":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "收到分所接洽單紀錄只可輸入空白或Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP86_GotFocus
      End Select
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP87_GotFocus()
   InverseTextBox textCP87
End Sub
'add by nick 2004/08/18
Private Sub textCP88_GotFocus()
   InverseTextBox textCP88
End Sub
'add by nick 2004/08/18
Private Sub textCP89_GotFocus()
   InverseTextBox textCP89
End Sub

Private Sub textCP89_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP89_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP89_2 = Empty
   If IsEmptyText(textCP89) = False Then
      textCP89_2 = GetCustomerName(textCP89, 0)
      If IsEmptyText(textCP89_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉申請人2代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP89_GotFocus
         End Select
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP90_GotFocus()
   InverseTextBox textCP90
End Sub

Private Sub textCP90_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP90_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP90_2 = Empty
   If IsEmptyText(textCP90) = False Then
      textCP90_2 = GetCustomerName(textCP90, 0)
      If IsEmptyText(textCP90_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉申請人3代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP90_GotFocus
         End Select
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP91_GotFocus()
   InverseTextBox textCP91
End Sub

Private Sub textCP91_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP91_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP91_2 = Empty
   If IsEmptyText(textCP91) = False Then
      textCP91_2 = GetCustomerName(textCP91, 0)
      If IsEmptyText(textCP91_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉申請人4代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP91_GotFocus
         End Select
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP92_GotFocus()
   InverseTextBox textCP92
End Sub

Private Sub textCP92_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP92_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP92_2 = Empty
   If IsEmptyText(textCP92) = False Then
      textCP92_2 = GetCustomerName(textCP92, 0)
      If IsEmptyText(textCP92_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉申請人5代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP92_GotFocus
         End Select
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP93_GotFocus()
   InverseTextBox textCP93
End Sub

Private Sub textCP93_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP93_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP93_2 = Empty
   If IsEmptyText(textCP93) = False Then
      textCP93_2 = GetCustomerName(textCP93, 0)
      If IsEmptyText(textCP93_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉人2代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP93_GotFocus
         End Select
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP94_GotFocus()
   InverseTextBox textCP94
End Sub

Private Sub textCP94_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP94_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP94_2 = Empty
   If IsEmptyText(textCP94) = False Then
      textCP94_2 = GetCustomerName(textCP94, 0)
      If IsEmptyText(textCP94_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉人3代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP94_GotFocus
         End Select
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP95_GotFocus()
   InverseTextBox textCP95
End Sub

Private Sub textCP95_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP95_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP95_2 = Empty
   If IsEmptyText(textCP95) = False Then
      textCP95_2 = GetCustomerName(textCP95, 0)
      If IsEmptyText(textCP95_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉人4代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP95_GotFocus
         End Select
      End If
   End If
End Sub

'add by nick 2004/08/18
Private Sub textCP96_GotFocus()
   InverseTextBox textCP96
End Sub

Private Sub textCP96_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP96_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP96_2 = Empty
   If IsEmptyText(textCP96) = False Then
      textCP96_2 = GetCustomerName(textCP96, 0)
      If IsEmptyText(textCP96_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "移轉人5代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP96_GotFocus
         End Select
      End If
   End If
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Call Pub_SaveMeToolBar(m_MeTrackMode, Me.tlbar, Button.Index) 'Added by Lydia 2021/10/20 若有交錯使用Function鍵和Toolbar鍵會失去記錄造成無法判斷，所以ToolBar鍵另外記錄
   
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

' 檢查資料庫中該筆記錄是否存在
Private Function IsDataBaseExist(ByVal strCP09) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   IsDataBaseExist = False
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP01 = '" & m_CP01 & "' AND " & _
                  "CP02 = '" & m_CP02 & "' AND " & _
                  "CP03 = '" & m_CP03 & "' AND " & _
                  "CP04 = '" & m_CP04 & "' AND " & _
                  "CP09 = '" & strCP09 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsDataBaseExist = True
   Else
      IsDataBaseExist = False
   End If
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strCP09) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   IsRecordExist = False
   bFind = False
   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex).diCP09 = strCP09 Then
         bFind = True
      End If
   Next nIndex
   If bFind = False Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP01 = '" & m_CP01 & "' AND " & _
                  "CP02 = '" & m_CP02 & "' AND " & _
                  "CP03 = '" & m_CP03 & "' AND " & _
                  "CP04 = '" & m_CP04 & "' AND " & _
                  "CP09 = '" & strCP09 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Function

' 新增記錄
Private Sub AddRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strCP09 As String
   
   'edit by nick 2004/08/18
   'For nIndex = 0 To T_CP - 1
   For nIndex = 0 To TF_CP - 1
      If m_FieldList(nIndex).fiName = "CP09" Then
         strCP09 = m_FieldList(nIndex).fiNewData
         Exit For
      End If
   Next nIndex

   bFirst = True
   bDifference = False
   strSql = "INSERT INTO CASEPROGRESS ("
   'edit by nick 2004/08/18
   'For nIndex = 0 To T_CP - 1
   For nIndex = 0 To TF_CP - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   'edit by nick 2004/08/18
   'For nIndex = 0 To T_CP - 1
   For nIndex = 0 To TF_CP - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = "'" & m_FieldList(nIndex).fiNewData & "'"
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
   If bDifference = True Then
      'add by nickc 2006/03/16 紀錄分析語法
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
        'Add By Cheng 2003/10/29
        '若有更改承辦人, 則更新ENG的核稿人
        If Me.textCP14.Text <> Me.textCP14.Tag Then
             'edit by nickc 2007/08/16 修正更新欄位
             'strSQL = "UPDATE ENGINEERPROGRESS SET EP03=(" & _
                            "SELECT PP04 FROM CASEPROGRESS,PROMOTERPROOFREADER WHERE CP09='" & Me.textCP09.Text & _
                            "' AND CP01=PP01(+) AND '" & Me.textCP14.Text & "'=PP02(+) AND CP10=PP03(+)) WHERE EP02='" & Me.textCP09.Text & "'"
             strSql = "UPDATE ENGINEERPROGRESS SET EP04=(" & _
                            "SELECT PP04 FROM CASEPROGRESS,PROMOTERPROOFREADER WHERE CP09='" & Me.textCP09.Text & _
                            "' AND CP01=PP01(+) AND '" & Me.textCP14.Text & "'=PP02(+) AND CP10=PP03(+)) WHERE EP02='" & Me.textCP09.Text & "'"
            'add by nickc 2006/03/16 紀錄分析語法
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
        End If
      ' 將新增的資料加入到串列中
      SetDataListItem strCP09
      ' 顯示該筆記錄
      ShowCurrRecord strCP09
      ' 通知前畫面有新增的記錄
      frm075004_1.ModRecord strCP09
      ' 設定有新增資料
      m_AddData = True
   End If
EXITSUB:
End Sub

' 修改記錄
Private Sub ModRecord()
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strCP09 As String
Dim strReason As String
Dim strUpdDate As String, strUpdTime As String 'Add By Sindy 2023/5/18
Dim strModifyNote As String 'Add By Sindy 2023/5/18
Dim bolModifyCRL As Boolean
Dim strModCP10 As String, strModCP16 As String, strModCP17 As String
Dim tmpBol As Boolean 'Added by Lydia 2025/03/19

   strCP09 = m_DataList(m_CurrDL).diCP09
   '910910  nick tigger
   '***** start
   'strSQL = "UPDATE CASEPROGRESS SET "
   'edit by nickc 2006/03/16
   'strSQL = "begin user_data.user_enabled:=1;  UPDATE CASEPROGRESS SET "
   strSql = "UPDATE CASEPROGRESS SET "
   '***** end
   bFirst = True
   bDifference = False
   'edit by nick 2004/08/18
   'For nIndex = 0 To T_CP - 1
   For nIndex = 0 To TF_CP - 1
      strTmp = Empty
      '92.05.22 nick 跳過create & update 相關項目
      If nIndex < 64 Or nIndex > 69 Then
        If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
           If m_FieldList(nIndex).fiType = 0 Then
              If Trim(m_FieldList(nIndex).fiNewData) = Empty Then
                 strTmp = m_FieldList(nIndex).fiName & " = NULL "
              Else
                 ' 91.03.25 modify by louis (單引號)
                 'strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
                 strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
              End If
           Else
              If Trim(m_FieldList(nIndex).fiNewData) = Empty Then
                 strTmp = m_FieldList(nIndex).fiName & " = NULL "
              Else
                 strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
              End If
           End If
        End If
        If strTmp <> Empty Then
           bDifference = True
           If bFirst = True Then
              strSql = strSql & strTmp
              bFirst = False
           Else
              strSql = strSql & "," & strTmp
           End If
        End If
      End If
   Next nIndex
   '910910 nick tigger
   
   'Add By Sindy 2014/3/25 檢查是否有拿掉發文日,若有,檢查是否需要一併拿掉發文時間
   'Remove by Lydia 2019/12/10 因為FCP和P案在電子送件發文時會自動上傳檔案到卷宗區，使用發文時間CP82判斷是否重新發文；
                         '問過Sindy目前歷程不用拿掉發文時間
   'If InStr(UCase(strSql), "CP27 = NULL") > 0 Then
   '   If bolUpdCP82 = True Then
   '      strSql = strSql & ",CP82 = NULL"
   '   End If
   'End If
   '2014/3/25 END
   
   '***** start
   'edit by nickc 2006/03/16 回覆改在下面做
   strSql = strSql & " " & _
            "WHERE CP01 = '" & m_CP01 & "' AND " & _
                  "CP02 = '" & m_CP02 & "' AND " & _
                  "CP03 = '" & m_CP03 & "' AND " & _
                  "CP04 = '" & m_CP04 & "' AND " & _
                  "CP09 = '" & strCP09 & "' "
   'strSQL = strSQL & " " & _
            "WHERE CP01 = '" & m_CP01 & "' AND " & _
                  "CP02 = '" & m_CP02 & "' AND " & _
                  "CP03 = '" & m_CP03 & "' AND " & _
                  "CP04 = '" & m_CP04 & "' AND " & _
                  "CP09 = '" & strCP09 & "'; end; "
    '***** end
'910910 nick tigger
'***** start
On Error GoTo ErrHand
'***** end
   
   If bDifference = True Then
      '910910 nick tigger
      '**** start
      cnnConnection.BeginTrans
      '***** end
      'add by nickc 2006/03/16 紀錄分析語法
      Pub_SeekTbLog strSql
      'edit by nickc 2006/03/16
      'cnnConnection.Execute strSQL
      cnnConnection.Execute "begin user_data.user_enabled:=1; " & strSql & "; end;"
      
      'Add By Cheng 2003/10/29
      '若有更改承辦人, 則更新ENG的核稿人
      If Me.textCP14.Text <> Me.textCP14.Tag Then
           'edit by nickc 2007/08/16 修正更新欄位
           'strSQL = "UPDATE ENGINEERPROGRESS SET EP03=(" & _
                          "SELECT PP04 FROM CASEPROGRESS,PROMOTERPROOFREADER WHERE CP09='" & Me.textCP09.Text & _
                          "' AND CP01=PP01(+) AND '" & Me.textCP14.Text & "'=PP02(+) AND CP10=PP03(+)) WHERE EP02='" & Me.textCP09.Text & "'"
           strSql = "UPDATE ENGINEERPROGRESS SET EP04=(" & _
                          "SELECT PP04 FROM CASEPROGRESS,PROMOTERPROOFREADER WHERE CP09='" & Me.textCP09.Text & _
                          "' AND CP01=PP01(+) AND '" & Me.textCP14.Text & "'=PP02(+) AND CP10=PP03(+)) WHERE EP02='" & Me.textCP09.Text & "'"
          'add by nickc 2006/03/16 紀錄分析語法
          Pub_SeekTbLog strSql
          cnnConnection.Execute strSql
      End If
      
      'Modify By Sindy 2023/9/21 寫成共用函數,有修改資料同時需記錄在接洽單上
      'Add By Sindy 2023/1/11 有修改案件性質
      bolModifyCRL = False
      If m_CP10 <> textCP10 Then
         bolModifyCRL = True
         strModCP10 = textCP10.Text
      End If
      'Add By Sindy 2022/12/5 修改費用,規費時; 發Mail通知正本財務處,副本智權人員
      If (Val(Me.textCP16.Text) <> Val(Me.textCP16.Tag) Or Val(Me.textCP17.Text) <> Val(Me.textCP17.Tag)) _
           And _
         (Trim(textCP60.Text) = "" Or Pub_StrUserSt03 = "M51") Then
         bolModifyCRL = True
         strModCP16 = Me.textCP16.Text
         strModCP17 = Me.textCP17.Text
      End If
      If bolModifyCRL = True Then
         If PUB_ModCrLCRCData(textCP09, textCP140, strModCP10, m_CP10 _
            , m_Nation, textCP64, strModCP16, strModCP17, Me.textCP16.Tag, Me.textCP17.Tag) = False Then
            GoTo ErrHand
         End If
      End If
      '2023/1/11 END
      
      'Add By Sindy 2025/2/4
      If m_CP10 <> textCP10 Then
         Dim douStPrice As Double, douLowPrice As Double
         Call ClsPDGetCaseLowPrice(m_CP01, m_Nation, textCP10, douStPrice, douLowPrice, "", "", textCP140, "", m_CP01, m_CP02, m_CP03, m_CP04)
         ' 更新案件進度檔的標準價及底價欄位
         strSql = "UPDATE CaseProgress SET CP33 = " & douStPrice & ", " & _
                                          "CP34 = " & douLowPrice & " " & _
                  "WHERE CP09 = '" & textCP09 & "' "
         cnnConnection.Execute strSql
      End If
      '2025/2/4 END
      
      'Added by Lydia 2022/12/21 已發文之法律案點數改成0，一併刪除工作點數；ex.FCL-010968(AB1021131) 5/31發文自動分配工作點數,9/1進度檔拿掉所有費用
      If InStr(textCP01, "L") > 0 And Trim(textCP82) <> "" And Val(textCP18.Tag) > 0 And Val(textCP18) = 0 Then
          strSql = "delete acc1n0 where a1n01='" & textCP09 & "' and a1n02='3' "
          cnnConnection.Execute strSql
      End If
      'end 2022/12/21
      
      'Added by Lydia 2025/03/19 出庭費領取：Email通知承辦律師確認是否領取出庭費
      If InStr(textCP01, "L") > 0 And textCP01 <> "LA" And textCP27.Text <> textCP27.Tag Then
         strExc(0) = "select cl02 from caselawer,staff where cl01='" & Trim(textCP09) & "' and nvl(cl03,0) > 0 and cl02=st01(+) and st04='1' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(2) = RsTemp.GetString(adClipString, , , ";")
            '取消發文日
            If Trim(textCP27) = "" And textCP27.Tag <> "" And textCP27.Tag <> "111111" Then
               strSql = "Update CaseLawer set CL07=null where CL01='" & Trim(textCP09) & "' and instr ('" & strExc(2) & "',cl02) > 0"
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql
            End If
            '上發文日
            If textCP27.Text <> "111111" And textCP27.Tag = "" Then
               If Right(strExc(2), 1) = ";" Then strExc(2) = Mid(strExc(2), 1, Len(strExc(2)) - 1)
               strExc(0) = textCP01 & "-" & textCP02 & IIf(textCP03 & textCP03 <> "000", "-" & textCP03 & "-" & textCP04, "") & IIf(Left(m_LOS02, 1) = "B", "(案源案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & "-" & m_LOS01cp03 & "-" & m_LOS01cp04 & ")", "") & "通知領取出庭費事(" & Me.Name & ")"
               strExc(1) = "本所案號：" & textCP01 & "-" & textCP02 & IIf(textCP03 & textCP03 <> "000", "-" & textCP03 & "-" & textCP04, "") & vbCrLf & _
                           IIf(m_LOS01cp01 <> "" And m_LOS01cp01 <> "TT", "案源案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 & m_LOS01cp03 <> "000", "-" & m_LOS01cp03 & "-" & m_LOS01cp04, "") & vbCrLf, "") & _
                           "案件名稱：" & Mid(cmbTM05.List(0), 3) & vbCrLf & _
                           "案件性質：" & Trim(textCP10_2) & vbCrLf
               tmpBol = PUB_ChkIsPaid(textCP09, strExc(3))
               If strExc(3) <> "" Then strExc(1) = strExc(1) & IIf(Left(strExc(3), 1) = "E", "收據號碼：", "請款單號：") & strExc(3) & vbCrLf
               strExc(1) = strExc(1) & "收款狀態：" & IIf(strExc(3) = "", "未請款", IIf(tmpBol = True, "已收款", "未收款")) & vbCrLf
               strExc(1) = strExc(1) & vbCrLf & "已完成委任狀發文程序，請至【法務系統->內法->資料處理->出庭費確認維護】確認開庭費領取事宜。"
   
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13) " & _
                          "values('" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                          ",'" & ChgSQL(strExc(0)) & "','" & ChgSQL(strExc(1)) & "',null,'" & Trim(textCP09) & "') "
               cnnConnection.Execute strSql
               strSql = "Update CaseLawer Set CL07=sqldatet(to_char(sysdate,'yyyymmdd'))||decode(cl07,null,'',',')||cl07 Where CL01='" & Trim(textCP09) & "' and instr ('" & strExc(2) & "',cl02) > 0"
               cnnConnection.Execute strSql
            End If
         End If
      End If
      'end 2025/03/19
      
      'Added by Morgan 2025/2/25
      '程序人員在點選繪圖人員後，發信通知繪圖主管陳翔龍--游協理
      If textCP29 <> "" And textCP29 <> textCP29.Tag Then
         If PUB_GetST03(textCP29) = "P13" Then
            strExc(1) = Pub_GetSpecMan("設定繪圖人員通知對象")
            strExc(2) = textCP01 & "-" & textCP02 & IIf(textCP03 & textCP04 = "", "", textCP03 & textCP04) & textCP10_2 & "(" & ChangeWStringToTDateString(DBDATE(textCP05)) & ")已" & IIf(textCP29.Tag = "", "設定", "變更") & "繪圖人員為" & textCP29_2
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & strExc(2) & "','如旨')"
            cnnConnection.Execute strSql
         End If
      End If
      'end 2025/2/25
      
      '910910 nick tigger
      '***** start
      cnnConnection.CommitTrans
      '***** end
      
      'Added by Lydia 2022/12/21  已發文之法律案點數變更
      If InStr(textCP01, "L") > 0 And Trim(textCP82) <> "" And CmdDot.Visible = True And _
          ((Val(textCP18.Tag) = 0 And Val(textCP18) > 0) Or (Val(textCP18.Tag) > 0 And Val(textCP18) > 0 And Val(textCP18.Tag) <> Val(textCP18))) Then
          MsgBox "請進入工作點數分配執行點數分配！", vbExclamation, "點數異動"
      End If
      'end 2022/12/21
      
      'Added by Lydia 2021/01/14 (10/5) 若案件性質或案件屬性有改時Email通知秀玲提醒確認案源及金額是否需調整。案件屬性第1次設定時要與接洽單檔比較是否不同。
                                                    'A3類案源為非訴訟案，點數都回智慧所，與屬性是否為智財權無關; ex.L-006316
      'Modified by Lydia 2023/04/24 拿掉限制A3類;同分案作業,改成有案源就通知
      'If m_LOS01 <> "" And m_LOS02 <> "A3" Then
      If m_LOS01 <> "" Then
         strExc(0) = "": strExc(1) = ""
         If m_CP01 <> "LA" Then  '排除顧問案
             If m_CP10 <> textCP10 Then
                 Call ClsPDGetCaseProperty(textCP01, m_CP10, strExc(3))
                 strExc(1) = strExc(1) & "、案件性質"
                 strExc(0) = strExc(0) & vbCrLf & "原案件性質：" & m_CP10 & strExc(3) & vbCrLf & "現案件性質：" & textCP10 & textCP10_2
             End If
         End If
         If strExc(0) <> "" Then
            '主旨
            strExc(1) = "法務分案" & m_CP01 & "-" & m_CP02 & IIf(m_CP03 <> "0", "-" & m_CP03, "") & IIf(m_CP04 <> "00", "-" & m_CP04, "") & "，改變" & Mid(strExc(1), 2)
            '內文
            strExc(2) = "法律所案號：" & m_CP01 & "-" & m_CP02 & IIf(m_CP03 <> "0", "-" & m_CP03, "") & IIf(m_CP04 <> "00", "-" & m_CP04, "") & "(" & textCP09 & ")" & vbCrLf & _
                             "專業部案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 <> "0", "-" & m_LOS01cp03, "") & IIf(m_LOS01cp04 <> "00", "-" & m_LOS01cp04, "") & "(" & m_LOS01 & ")" & vbCrLf & _
                              strExc(0)
            PUB_SendMail strUserNum, "83002", "", strExc(1), strExc(2)
         End If
      End If
      'end 2021/01/14
      
      'Added by Lydia 2024/11/08 內專內商人員(ST03=P22,P12)在進度檔、下一程序修改本所期限和法定期限時，Email通知主管
      If InStr("P22,P12", Pub_StrUserSt03) > 0 And (Me.textCP06.Tag <> Me.textCP06.Text Or Me.textCP07.Tag <> Me.textCP07.Text) Then
         'Modified by Lydia 2025/05/19 在系統特殊設定區分要不要發通知; ex. 員工編號(N)=>不寄, 員工編號(Y)=>要寄
         'strExc(0) = "select s1.st01,s1.st02,s1.st93 from setspecman,staff s1 where ocode='期限修改郵件收受者' and instr(oman,s1.st01) > 0 and s1.st93='" & Pub_StrUserSt93 & "' "
         strExc(0) = "select s1.st01,s1.st02,s1.st93,substr(oman,instr(oman,s1.st01)+length(s1.st01),3) o1 from setspecman,staff s1 where ocode='期限修改郵件收受者' and instr(oman,s1.st01) > 0 and s1.st93='" & Pub_StrUserSt93 & "' "
         intI = 1
         strExc(1) = ""
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modified by Lydia 2025/07/04 debug: 本人操作才不用發Email; or "" & RsTemp.Fields("st01") <> strUserNum
            If "" & RsTemp.Fields("O1") <> "(N)" Or "" & RsTemp.Fields("st01") <> strUserNum Then     'Added by Lydia 2025/05/19 在系統特殊設定區分要不要發通知;
               strExc(1) = "" & RsTemp.Fields("st01")
            End If
         Else
            strExc(1) = Pub_GetSpecMan("程式管理人員")
         End If
         If strExc(1) <> "" Then
            '主旨：P/CFP-XXXXXX「XXXXXXX」(案件性質)-->進度檔/下一程序之期限有更動！ [（更動之程序人員）]
            strExc(2) = m_CP01 & "-" & m_CP02 & IIf(m_CP03 <> "0", "-" & m_CP03, "") & IIf(m_CP04 <> "00", "-" & m_CP04, "") & "「" & Trim(textCP10_2) & "」-->進度檔之期限有更動！ [" & strUserName & "]"
            strExc(3) = ""
            If Me.textCP06.Tag <> Me.textCP06.Text Then
               strExc(3) = strExc(3) & "修改前本所期限：" & ChangeWStringToTDateString(DBDATE(textCP06.Tag)) & vbCrLf & _
                           "修改後本所期限：" & ChangeWStringToTDateString(DBDATE(textCP06.Text)) & vbCrLf & vbCrLf
            End If
            If Me.textCP07.Tag <> Me.textCP07.Text Then
               strExc(3) = strExc(3) & "修改前法定期限：" & ChangeWStringToTDateString(DBDATE(textCP07.Tag)) & vbCrLf & _
                           "修改後法定期限：" & ChangeWStringToTDateString(DBDATE(textCP07.Text)) & vbCrLf & vbCrLf
            End If
            PUB_SendMail strUserNum, strExc(1), textCP09, strExc(2), strExc(3)
         End If
      End If
      'end 2024/11/08
      
      ShowCurrRecord strCP09
      'Modify By Sindy 2013/10/25
      'If UCase(m_PrevFormNm) <> UCase("frm040111") Then
      If UCase(TypeName(m_PrevForm)) <> UCase("frm040111") Then 'Modify by Sindy 2018/10/9
      '2013/10/25 END
         '通知前畫面有更新的記錄
         frm075004_1.ModRecord strCP09
      End If
   End If
'910910 nick tigger
'***** start

   PUB_SendMailCache '發信 Add By Sindy 2022/12/6
   Exit Sub
   
ErrHand:
    cnnConnection.RollbackTrans
    If Err.Number <> 0 Then MsgBox (Err.Description)
'******* end
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strCP09 As String
   Dim strDataList() As String
   Dim nDataListCount As Integer
   Dim nIndex As Integer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nPos As Integer
   Dim strDD23 As String
   Dim strDD24 As String
   'Add By Cheng 2003/05/30
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   'Add by Morgan 2006/4/24
   Dim bolReControl As Boolean '是否恢復管制
   Dim bolTrans As Boolean '是否有Transaction
   Dim strTM17 As String 'Add By Sindy 2012/9/5
   Dim m_AttachPath As String, iStr As String 'Added by Lydia 2015/12/03
   Dim iRtn As Integer 'Add By Sindy 2016/4/22
   Dim bolShowForm As Boolean 'Added by Morgan 2016/6/2
   Dim bolDelRefCP As Boolean 'Added  by Morgan 2016/6/2
   Dim fso As New FileSystemObject 'Added by Lydia 2017/12/27
   Dim strAdd01 As String, strAdd02 As String 'Added by Lydia 2020/02/13
   Dim strConSql As String 'Add By Sindy 2020/5/26
   Dim bolResetMail As Boolean 'Added by Morgan 2020/9/15 是否恢復信件為未處理狀態
   
   nPos = 0
   strCP09 = m_DataList(m_CurrDL).diCP09
   
   'Add By Sindy 2017/6/27 檢查是否有承辦歷程資料
   'Modify By Sindy 2018/7/19 + and eep04 not in('" & EMP_附加流程 & "')
   strSql = "select eep01,eep02 from empelectronprocess" & _
            " where eep01='" & strCP09 & "' and eep04 not in('" & EMP_附加流程 & "')"
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If MsgBox("此筆文號已有承辦歷程，請和（通知刪除的人員）確認是否（確定）要刪除？" & vbCrLf & vbCrLf & _
                "（因為工程師若文已辦出去，就不能刪除！）", vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
         Exit Sub
      End If
   End If
   rsA.Close
   '2017/6/27 END
   
   'Added by Morgan 2025/3/20 台灣商標分割核准(電子公文)，要先刪除子案的分割核准
   If (m_CP01 = "T" Or m_CP01 = "FCT") And m_Nation = "000" And textCP10 = "1001" Then
      strExc(0) = "select dc01||'-'||dc02||decode(dc03||dc04,'000','','-'||dc03||'-'||dc04) DivCNo from caseprogress a,divisioncase,caseprogress b,caseprogress c" & _
         " where a.cp09='" & textCP43 & "' and a.cp10='308'" & _
         " and dc05(+)=a.cp01 and dc06(+)=a.cp02 and dc07(+)=a.cp03 and dc08(+)=a.cp04" & _
         " and b.cp01(+)=dc01 and b.cp02(+)=dc02 and b.cp03(+)=dc03 and b.cp04(+)=dc04 and b.cp10='1001'" & _
         " and c.cp09(+)=b.cp43 and c.cp10='308'" & _
         " and exists(select * from edocument where ed11='" & strCP09 & "')"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "本來函為電子公文且子案也有[核准-分割]，母案的[核准-分割]不可刪除！" & vbCrLf & vbCrLf & "有[核准-分割]的子案:" & vbCrLf & rsA.GetString, vbCritical
         Exit Sub
      End If
      rsA.Close 'Add By Sindy 2025/5/8
   End If
   'end 2025/3/20
   
   'Added by Lydia 2023/10/06 檢查該收文號CP09不可存在於其他筆的CP43，也不可以存在於支援記錄檔supporthour的SH12，若存在則不可刪除該收文號。
   strSql = "select cp01,cp02,cp03,cp04,cp09,sqldatet(cp05) cp05t from caseprogress where cp43='" & strCP09 & "' order by cp05,cp09 "
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      strMsg = ""
      rsA.MoveFirst
      Do While Not rsA.EOF
         strMsg = strMsg & vbCrLf & rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & IIf(rsA.Fields("cp03") & rsA.Fields("cp04") <> "000", "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04"), "") & "　收文日：" & rsA.Fields("cp05t") & "　收文號：" & rsA.Fields("cp09")
         rsA.MoveNext
      Loop
      If strMsg <> "" Then
         MsgBox "收文號不可存在於其他收文之相關收文號：" & strMsg, vbExclamation
         Exit Sub
      End If
   End If
   rsA.Close
   strSql = "select sqldatet(sh01) sh01t,sh02||' '||s1.st02 as sh02t,sh03||' '||s2.st02 as sh03t,sh04 from supporthour,staff s1, staff s2 where sh02=s1.st01(+) and sh03=s2.st01(+) and sh12='" & strCP09 & "' order by 1,4 "
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      strMsg = ""
      rsA.MoveFirst
      Do While Not rsA.EOF
         strMsg = strMsg & vbCrLf & rsA.Fields("sh01t") & "　序號" & rsA.Fields("sh04") & "　支援人員：" & rsA.Fields("sh02t") & "　智權人員：" & rsA.Fields("sh03t")
         rsA.MoveNext
      Loop
      If strMsg <> "" Then
         MsgBox "收文號不可存在於支援時數記錄檔：" & strMsg, vbExclamation
         Exit Sub
      End If
   End If
   rsA.Close
   'end 2023/10/06
   
   'Added by Lydia 2017/12/27
   'Modified by Lydia 2018/05/09 +FMP案
   'If strSrvDate(1) >= FCP案件命名啟用日 And textCP01 = "FCP" And InStr(NewCasePtyList, textCP10) > 0 Then
   '    strExc(1) = Pub_GetFCPcaseFilePath(m_CP02)
   'Modified by Lydia 2019/06/05 排除B類假收文(P-122172誤收文101)
   'If (textCP01 = "FCP" Or textCP01 = "P") And InStr(NewCasePtyList, textCP10) > 0 Then
   'Modified by Lydia 2023/10/03 限制CP31=Y ; ex.FCP-63230刪除307分割,一併刪除English_vers992、專利案件991
   If (textCP01 = "FCP" Or textCP01 = "P") And InStr(NewCasePtyList, textCP10) > 0 And Left(textCP09, 1) = "A" And textCP31 = "Y" Then
       'Remove by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
       'strExc(1) = Pub_GetFCPcaseFilePath(m_CP02, , m_CP01)
   'end 2018/05/09
       'If fso.FolderExists(strExc(1)) Then
       '     MsgBox "外專案件資料夾" & strExc(1) & " 存在，請相關人員移除資料夾後才可刪除! "
       '     Exit Sub
       'End If
       'end 2021/12/06
       'Added by Lydia 2020/02/13 外專：專利案件和English_Vers檔案
       strExc(1) = m_CP01: strExc(2) = m_CP02: strExc(3) = m_CP03: strExc(4) = m_CP04
       Call PUB_ChkCPExist(strExc, cnt專利案件, , strAdd01, , "D")  '專利案件991
       Call PUB_ChkCPExist(strExc, cntEnglish_Vers, , strAdd02, , "D") 'English_Vers992
        If strAdd01 & strAdd02 <> "" Then
             strSql = "select cpf01 from casepaperfile" & _
                      " where cpf01 in (" & GetAddStr(strAdd01 & "," & strAdd02) & ") and substr(upper(cpf02),-4)<>upper('.del')"
             rsA.CursorLocation = adUseClient
             rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If rsA.RecordCount > 0 Then
                'Modified by Lydia 2023/10/03 +收文號
                If MsgBox("原始檔區已有資料，請和外專承辦確認是否（確定）要刪除？" & vbCrLf & vbCrLf & _
                          "（電子檔他們是否要先搬移或做什麼處理...）" & vbCrLf & vbCrLf & _
                          "專利案件和English_Vers收文號:" & strAdd01 & "、" & strAdd02, vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
                   strAdd01 = "": strAdd02 = ""  'Added by Lydia 2023/10/03
                   Exit Sub
                End If
             End If
             rsA.Close
        End If
        'end 2020/02/13
   End If
   'end 2017/12/27
   
   'Added by Lydia 2019/06/28 FCP判斷有定稿日期詢問是否刪除(FCP-059622因為刪最後一筆專利證書但是保留的沒有定稿日期,所以整批發文無法抓到資料)
   If textCP01 = "FCP" And m_CP85 <> "" Then
      If MsgBox("此收文號有FCP定稿日期，可能會造成整批發文無法抓到該案，請確定是否刪除？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
           Exit Sub
      End If
   End If
   'end 2019/06/28

   'Add By Sindy 2017/11/16 檢查是否有卷宗區和原始檔區資料
   'Modified by Lydia 2018/01/15 排除查名結果, 後面會還原資料
   'Modified by Lydia 2018/10/05 先排除.Menu
   'strSql = "select cpp01 from casepaperpdf" & _
            " where cpp01='" & strCP09 & "' and instr(upper(cpp02),'" & UCase("." & TMQ_查名作業 & ".menu") & "')=0 and substr(upper(cpp02),-4)<>upper('.del')"
   'Modified by Morgan 2019/8/16 先排除系統(QPGMR)產生的客戶函(.CUS.PDF),系統產生的所有pdf或許都該排除
   strSql = "select cpp01 from casepaperpdf" & _
            " where cpp01='" & strCP09 & "' and upper(cpp02) not like '%.MENU' and substr(upper(cpp02),-4)<>upper('.del') and not (cpp05='QPGMR' and upper(substr(cpp02,-8))='.CUS.PDF')"
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If MsgBox("此筆文號卷宗區(原始檔區也要查看)已有資料，請和（通知刪除的人員）確認是否（確定）要刪除？" & vbCrLf & vbCrLf & _
                "（電子檔他們是否要先搬移或做什麼處理...）", vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
         Exit Sub
      End If
   End If
   rsA.Close
   'Added by Lydia 2018/10/05 另外處理.Menu
   strSql = "select cpp01,cpp02 from casepaperpdf " & _
               "where cpp01='" & strCP09 & "' and upper(cpp02) like '%.MENU'  and instr(upper(cpp02),'" & "." & UCase(TMQ_查名作業 & ".menu") & "')=0 and instr(upper(cpp02),'" & "." & UCase("CASE.Menu") & "')=0 " & _
               "and substr(upper(cpp02),-4)<>upper('.del')"
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      strExc(1) = ""
      rsA.MoveFirst
      Do While Not rsA.EOF
           strExc(1) = strExc(1) & vbCrLf & rsA.Fields("cpp02")
           rsA.MoveNext
      Loop
      If MsgBox("此筆文號卷宗區有.Menu資料，確認是否（確定）要刪除？" & strExc(1), vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
            Exit Sub
      End If
   End If
   rsA.Close
   'end 2018/10/05
   strSql = "select cpf01 from casepaperfile" & _
            " where cpf01='" & strCP09 & "' and substr(upper(cpf02),-4)<>upper('.del')"
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      'Added by Lydia 2020/02/13 外專：專利案件和English_Vers檔案不可直接刪除
      If (textCP01 = "FCP" Or textCP01 = "P") And Left(textCP09, 1) = "D" And (textCP10 = cntEnglish_Vers Or textCP10 = cnt專利案件) Then
            MsgBox "此筆文號原始檔區已有資料，請外專承辦或程序人員先處理後才可刪除! ", vbCritical + vbOKOnly
            Exit Sub
      Else
      'end 2020/02/13
            If MsgBox("此筆文號原始檔區已有資料，請和（通知刪除的人員）確認是否（確定）要刪除？" & vbCrLf & vbCrLf & _
                      "（電子檔他們是否要先搬移或做什麼處理...）", vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
               Exit Sub
            End If
      End If 'end 2020/02/13
   End If
   rsA.Close
   '2017/11/16 END
   
   'Added by Lydia 2018/07/03 若刪除D類收款寄證,一併刪除定稿檔(ex.T-142048,T-199535 桂英在4/25輸入核准,又請電腦中心刪除C類核准和D類收款寄證,但是沒有刪除定稿,後面雅雯輸入核准,在發文收款寄證時造成錯誤)
   'Mark by Lydia 2024/06/05 因為basLetter有「2019/11/21 新增前先把舊資料刪除」這段不用了;
   'm_DelLD01 = "": m_DelLD04 = "": m_DelLD10 = ""
   'If Left(strCP09, 1) = "D" And m_CP10 = "1728" Then
   '   strSql = "select c1.cp65 ld01, nvl(d2.ld04,d1.ld04) ld04,nvl(d2.ld10,d1.ld10) ld10 " & _
                  "from caseprogress a1,caseprogress c1,caseprogress c2,letterdemand d1,letterdemand d2 " & _
                  "where a1.cp09='" & strCP09 & "' and a1.cp43=c1.cp09(+) and c1.cp43=c2.cp09(+) " & _
                  "and c1.cp09=d1.ld04(+) and c1.cp10=d1.ld09(+) " & _
                  "and c2.cp09=d2.ld04(+) and c2.cp10=d2.ld09(+) "
   '   intI = 1
   '   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '   If intI = 1 Then
   '        If "" & RsTemp.Fields(1) <> "" And "" & RsTemp.Fields(2) <> "" Then
   '             m_DelLD01 = "" & RsTemp.Fields(0)
   '             m_DelLD04 = "" & RsTemp.Fields(1)
   '             m_DelLD10 = "" & RsTemp.Fields(2)
   '        End If
    '  End If
   'End If
   'end 2024/06/05
   'end 2018/07/03
   
   'Add by Morgan 2004/2/12
   'Modify by Sindy 2023/3/9 +) or NP24='" & strCP09 & "')
   'Modified by Lydia 2023/03/22 法律案(L,FCL)不同審級收文期限沖銷：案件不同審級會收文-1、-2…案號，但實際是同一案件
   If m_CP01 = "L" Or m_CP01 = "FCL" Then
        strSql = "SELECT NP22 FROM NEXTPROGRESS WHERE ((NP01='" & m_CP43 & "' AND NP02='" & m_CP01 & "' AND NP03='" & m_CP02 & "' AND NP07='" & m_CP10 & "') or NP24='" & strCP09 & "') AND NP06='Y'"
   Else
   'end 2023/03/22
        strSql = "SELECT NP22 FROM NEXTPROGRESS WHERE ((NP01='" & m_CP43 & "' AND NP02='" & m_CP01 & "' AND NP03='" & m_CP02 & "' AND NP04='" & m_CP03 & "' AND NP05='" & m_CP04 & "' AND NP07='" & m_CP10 & "') or NP24='" & strCP09 & "') AND NP06='Y'"
   End If  'Added by Lydia 2023/03/22
   If rsA.State = adStateOpen Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       If MsgBox("此筆下一程序期限已上續辦，是否要恢復期限管制？", vbYesNo) = vbYes Then
         bolReControl = True
       End If
   'add by sonia 2016/5/13 進度檔有期限但無下一程序期限時提醒 P-112758刪除進度檔實審期限
   Else
      If Val(textCP07) <> 0 And Left(strCP09, 1) < "C" Then
         MsgBox "此收文程序有期限, 但無相關之下一程序期限, 是否需補期限, 請通知專業部自行決定！"
      End If
   'end 2016/5/13
   End If
   rsA.Close
   'Added by Lydia 2023/04/06 (公告)FCT,S案請控制內部收文739更換智權人員存檔時，同時更換下一程序資料維護中「是否續辦」欄為空白之「智權人員」
   If (m_CP01 = "FCT" Or m_CP01 = "S") And m_CP10 = "739" And Left(strCP09, 1) = "B" Then
      strSql = "SELECT NP22 FROM NEXTPROGRESS WHERE NP02='" & m_CP01 & "' AND NP03='" & m_CP02 & "' AND NP04='" & m_CP03 & "' AND NP05='" & m_CP04 & "' AND NP06 is null AND INSTR(NP15,'更換智權人員') >0 "
      If rsA.State = adStateOpen Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
          MsgBox "此收文有相關之下一程序更換智權人員, 是否需變更人員, 請通知專業部自行更正！"
      End If
      rsA.Close
   End If
   'end 2023/04/06
   
   'Add By Sindy 2020/5/6
   'TT-999999案刪進度時，檢查法律所案源檔，必須為未收法務案且未放棄始可刪除。
   If m_CP01 = "TT" And m_CP02 = "999999" Then
      strSql = "SELECT * FROM lawofficesource WHERE los10='" & strCP09 & "'"
      If rsA.State = adStateOpen Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         If "" & rsA.Fields("los06") <> "" Then
            MsgBox "法務案已收文，不可刪除！", vbCritical + vbOKOnly
            Exit Sub
         ElseIf "" & rsA.Fields("los08") <> "" Then
            MsgBox "法務案已填放棄收文，不可刪除！", vbCritical + vbOKOnly
            Exit Sub
         End If
      End If
      rsA.Close
   End If
   '2020/5/6 END
   
   'Added by Lydia 2021/11/29
   If m_cp109 <> "" Then
      MsgBox "此案已上可結餘，請通知專業部確認是否要取消可結餘日期！", vbExclamation
   End If
   'end 2021/11/29

   'Modify By Sindy 2016/8/30 Mark
'   'Add By Sindy 2011/8/11
'   strSql = "SELECT * FROM COURTYARDPERIOD WHERE CDP01='" & strCP09 & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      If "" & RsTemp("CDP16") <> "" Then
'         MsgBox "有庭期資料附件，請先至庭期資料維護做刪除附件！"
'         GoTo EXITSUB
'      End If
'   End If
'   '2011/8/11 End

   'Added by Lydia 2015/12/03 國內T案申請時，若有收文查名作業，先做還原
   'Added by lydia 2016/04/25 +TS案
   'If m_CP01 = "T" And m_Nation = "000" And m_CP10 = 申請 Then
   'Modified by Lydia 2021/11/19 增加737智財協作之T案
   'If (m_CP01 = "T" And m_Nation = "000" And m_CP10 = TMQ_T案) Or (m_CP01 = "TS" And m_Nation = "000" And m_CP10 = TMQ_TS案) Then
   If (m_CP01 = "T" And m_Nation = "000" And InStr(TMQ_T案, m_CP10) > 0) Or (m_CP01 = "TS" And m_Nation = "000" And InStr(TMQ_TS案, m_CP10) > 0) Then
      intI = 1: iStr = ""
      strSql = "select cpp02 from casepaperpdf where cpp01='" & strCP09 & "' and instr(upper(cpp02),'" & UCase("." & TMQ_查名作業 & ".menu") & "') > 0 "
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            strExc(1) = Mid(RsTemp(0), 1, InStr(UCase(RsTemp(0)), UCase("." & TMQ_查名作業 & ".menu")) - 1)
            '減號=從卷宗區還原到查名作業
            iStr = iStr & "-" & Mid(strExc(1), InStr(strExc(1), "H"), 9) & ","
            RsTemp.MoveNext
         Loop
         'Remove by Lydia 2016/12/07 查名作業固定放在卷宗區,共用模組處理對應的記錄
         'Move by Lydia 2016/03/25 有搬移才做
         'm_AttachPath = App.path & "\" & strUserNum
         'If Dir(m_AttachPath, vbDirectory) = "" Then
         '   MkDir m_AttachPath
         'End If
         'iStr = IIf(Right(iStr, 1) = ",", Mid(iStr, 1, Len(iStr) - 1), iStr)
         'If PUB_TMQtoCP(m_AttachPath, strCP09, iStr, "D") = True Then
         '   MsgBox "已收文的查名結果還原到查覆區！"
         'End If
      End If
   End If
   'end 2015/12/03
   
   'Added by Morgan 2016/6/2
   bolShowForm = True
   bolDelRefCP = False
   '催提申、收達若刪除的是第一案時提醒會一併刪除同發文字號的其他案
   'Modified by Morgan 2018/10/12 +954
   If (textCP01 = "P" Or textCP01 = "CFP" Or textCP01 = "PS" Or textCP01 = "CPS") And (textCP10 = "952" Or textCP10 = "953" Or textCP10 = "954") Then
      If textCP28 = textCP09 Then
         strSql = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'('||cp09||') '||cpm04 from caseprogress,casepropertymap where cp28='" & textCP28 & "' and cp09<>'" & textCP28 & "' and cp10='" & textCP10 & "' and cpm01(+)=cp01 and cpm02(+)=cp10"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If MsgBox("下列相同發文字號的程序將一併刪除！是否要繼續？" & vbCrLf & vbCrLf & RsTemp.GetString, vbYesNo + vbQuestion) = vbNo Then
               GoTo EXITSUB
            End If
            bolDelRefCP = True
         End If
      End If
      
      If m_CP65 = strUserNum Then
         strDD23 = strUserNum
         strDD24 = "失誤"
         bolShowForm = False
      End If
   End If
   'end 2016/6/2
   
   'Modified by Morgan 2016/6/2 +bolShowForm
   If OnDataDeleteRecord(1, strCP09, m_CP01 & m_CP02 & m_CP03 & m_CP04, strDD23, strDD24, bolShowForm) <> 0 Then
      GoTo EXITSUB
   End If
   
   'Added by Lydia 2016/12/07
   If iStr <> "" Then
      'Modified by Lydia 2024/03/14 +False
      'If PUB_TMQtoCP("", strCP09, iStr, "D", , True) = True Then
      If PUB_TMQtoCP(False, "", strCP09, iStr, "D", , True) = True Then
      End If
   End If
   'end 2016/12/07
   
'Add by Morgan 2004/2/12
On Error GoTo ErrHnd
    cnnConnection.BeginTrans
    bolTrans = True
'Add end 2004/2/12

   'Added by Morgan 2020/9/15 專利已處理信件還原為未處理狀態
   'Modify By Sindy 2023/6/12 +CFT,CFC,S
   If Left(strCP09, 1) = "C" And (textCP01 = "P" Or textCP01 = "CFP" Or textCP01 = "PS" Or textCP01 = "CPS" Or _
                                  textCP01 = "CFT" Or textCP01 = "CFC" Or textCP01 = "S") Then
      'Modified by Morgan 2023/4/18 + and ir16='8'
      'Modified by Morgan 2023/7/21
      '抓(專利處信箱)或(國外部信箱且處理結果為9輸入)+ 處理狀態(ir16)= 1輸入, 7已確認, 8退回2
      'strExc(0) = "select * from InputRecord where ir21='" & strCP09 & "' and ir16='8'"
      strExc(0) = "select * from InputRecord,PatentInput,IPdeptInput" & _
         " where ir21='" & strCP09 & "' and ir16 in ('1','7','8')" & _
         " and pi01(+)=ir01 and pi02(+)=ir02 and pi03(+)=ir03" & _
         " and ii01(+)=ir01 and ii02(+)=ir02 and ii03(+)=ir03" & _
         " and (pi01 is not null or ii27='9' or ii29='9')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '2次確認退回
         If RsTemp("ir16") = "8" Then
            'Modify by Amy 2023/11/02 bug原:WHERE  WHERE
            strSql = "UPDATE InputRecord SET ir21=null  WHERE ir01=" & RsTemp("ir01") & " and ir02=" & RsTemp("ir02") & " and ir03='" & RsTemp("ir03") & "' and ir04='" & RsTemp("ir04") & "' and ir21='" & strCP09 & "'"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
         
         'Modify By Sindy 2023/7/13 備註:II27是外專使用 II29是外商使用; 若有調整要同時考慮
         'Removed by Morgan 2023/4/18 取消退回收件區，因狀況不一改人工自行處理
         'Modified by Morgan 2023/7/21 再改為2次確認的來函都要要退回收件區
         ElseIf RsTemp("ir16") = "1" Or RsTemp("ir16") = "7" Then
         
            '專利處2次確認(有C類收文號的都是)
            If Left(RsTemp("ir03"), 1) = "P" Then
               If RsTemp("ir08") > 0 And RsTemp("pi16") > 0 Then
                  strSql = "UPDATE PatentInput SET pi16=0" & _
                     " WHERE pi01=" & RsTemp("ir01") & " and pi02=" & RsTemp("ir02") & " and pi03='" & RsTemp("ir03") & "'"
                  Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql, intI
               End If
               
               strSql = "UPDATE InputRecord SET" & _
                  " ir08=0,ir09=null,ir10=null,ir16=null,ir17=0,ir18=null,ir19=null,ir20=null,ir21=null" & _
                  " WHERE ir01=" & RsTemp("ir01") & " and ir02=" & RsTemp("ir02") & " and ir03='" & RsTemp("ir03") & "' and ir04='" & RsTemp("ir04") & "' and ir21='" & strCP09 & "'"
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql, intI
               
            '國外部2次確認(有法限者才是)
            ElseIf Left(RsTemp("ir03"), 1) = "F" And textCP07 <> "" Then
               '外專
               If InStr(textCP01, "P") > 0 Then
                  '外專處理結果(ii27)=9輸入
                  If RsTemp("ii27") = "9" Then
                     If RsTemp("ir08") > 0 And RsTemp("ii16") > 0 Then
                        strSql = "UPDATE IPdeptInput SET ii16=0,ii27=null" & _
                           " WHERE ii01=" & RsTemp("ir01") & " and ii02=" & RsTemp("ir02") & " and ii03='" & RsTemp("ir03") & "'"
                        Pub_SeekTbLog strSql
                        cnnConnection.Execute strSql, intI
                     End If
                     
                     strSql = "UPDATE InputRecord SET" & _
                        " ir08=0,ir09=null,ir10=null,ir16=null,ir17=0,ir18=null,ir19=null,ir20=null,ir21=null" & _
                        " WHERE ir01=" & RsTemp("ir01") & " and ir02=" & RsTemp("ir02") & " and ir03='" & RsTemp("ir03") & "' and ir04='" & RsTemp("ir04") & "' and ir21='" & strCP09 & "'"
                     Pub_SeekTbLog strSql
                     cnnConnection.Execute strSql, intI
                  End If
               '外商
               Else
                  '外商處理結果(ii29)=9輸入
                  If RsTemp("ii29") = "9" Then
                     If RsTemp("ir08") > 0 And RsTemp("ii16") > 0 Then
                        strSql = "UPDATE IPdeptInput SET ii16=0,ii29=null" & _
                           " WHERE ii01=" & RsTemp("ir01") & " and ii02=" & RsTemp("ir02") & " and ii03='" & RsTemp("ir03") & "'"
                        Pub_SeekTbLog strSql
                        cnnConnection.Execute strSql, intI
                     End If
                     
                     strSql = "UPDATE InputRecord SET" & _
                        " ir08=0,ir09=null,ir10=null,ir16=null,ir17=0,ir18=null,ir19=null,ir20=null,ir21=null" & _
                        " WHERE ir01=" & RsTemp("ir01") & " and ir02=" & RsTemp("ir02") & " and ir03='" & RsTemp("ir03") & "' and ir04='" & RsTemp("ir04") & "' and ir21='" & strCP09 & "'"
                     Pub_SeekTbLog strSql
                     cnnConnection.Execute strSql, intI
                  End If
               End If
            End If
         End If
         'end 2023/4/18
      End If
      
   End If
   'end 2020/9/15
   
   'Add By Sindy 2015/1/30 增加刪除案件表單相關資料
   'Call PUB_CloseFlowDataDel(textCP140, m_CP01, m_CP02, m_CP03, m_CP04, strCP09)
   'Modify By Sindy 2016/4/22 摩根發現電子結案單在實作上,應該只有程序人員產生的進度,需要是恢復上一處理動作的狀況;而不是刪資料
   If textCP140 <> "" And m_CP01 <> "" And m_CP02 <> "" And m_CP03 <> "" And m_CP04 <> "" Then
      If Len(textCP140) = 8 Then '電子結案單
         'Modify By Sindy 2018/10/16 重覆按解除期限,產生2筆不續辦,ex:P119898
         strSql = "SELECT cp09 FROM caseprogress WHERE cp140='" & textCP140 & "'"
         If rsA.State = adStateOpen Then rsA.Close
         rsA.CursorLocation = adUseClient
         rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 1 Then '一筆以上,代表重覆了
            '有回覆單的不可直接刪除
            strSql = "SELECT cpp01,cpp02 FROM casepaperpdf WHERE cpp01='" & strCP09 & "'" & _
                     " and instr(Upper(cpp02),upper('." & EMP_回覆單 & ".pdf'))>0"
            If rsA.State = adStateOpen Then rsA.Close
            rsA.CursorLocation = adUseClient
            rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               MsgBox "電子結案單「重覆解除期限」,但此筆「有回覆單」不可直接刪除!"
               Exit Sub
            End If
         Else
         '2018/10/16 END
            '檢查是否程序人員要恢復處理動作
            strSql = "SELECT F0309 FROM Flow003 WHERE F0301='" & textCP140 & "'"
            If rsA.State = adStateOpen Then rsA.Close
            rsA.CursorLocation = adUseClient
            rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               'Modify By Sindy 2025/8/28 + Or rsA.Fields("F0309") = Flow_歸檔
               '                          TF-00075-0-5-06已歸檔,但程序人員想改操作,原為不續辦要改閉卷,回待處理區
               If rsA.Fields("F0309") = Flow_已完成 Or rsA.Fields("F0309") = Flow_指示信判發中 _
                  Or rsA.Fields("F0309") = Flow_歸檔 Then
                  '恢復程序人員上一個處理動作
                  '流程主檔:
                  strSql = "UPDATE Flow003 SET F0308=F0307,F0309='" & Flow_處理中 & "' WHERE F0301='" & textCP140 & "'"
                  Pub_SeekTbLog strSql '記錄Log
                  cnnConnection.Execute strSql
                  '簽核檔:
                  strSql = "UPDATE Flow002 SET F0205=null,F0206=null,F0207=null WHERE F0201='" & textCP140 & "' and F0202='2'"
                  Pub_SeekTbLog strSql '記錄Log
                  cnnConnection.Execute strSql
                  'Add By Sindy 2025/8/28
                  If rsA.Fields("F0309") = Flow_歸檔 Then
                     '簽核檔:
                     strSql = "UPDATE Flow002 SET F0205=null,F0206=null,F0207=null WHERE F0201='" & textCP140 & "' and F0202='3'"
                     Pub_SeekTbLog strSql '記錄Log
                     cnnConnection.Execute strSql
                  End If
                  '2025/8/28 END
                  Call PUB_CloseRestoreLimit(strCP09) 'Modify By Sindy 2020/12/29 改成函數,結案單恢復解除期限
'                  '更新下一程序:
'                  'Modify By Sindy 2018/10/16 + NP06=null,
'                  strSql = "Update NextProgress Set NP06=null,NP24='" & textCP140 & "' WHERE NP24='" & strCP09 & "'"
'                  Pub_SeekTbLog strSql '記錄Log
'                  cnnConnection.Execute strSql
'                  '檢查卷宗區是否有回覆單要恢復
'                  strSql = "SELECT cpp02 FROM casepaperpdf WHERE cpp01='" & strCP09 & "' and instr(Upper(cpp02),upper('." & EMP_回覆單 & ".pdf'))>0"
'                  If rsA.State = adStateOpen Then rsA.Close
'                  rsA.CursorLocation = adUseClient
'                  rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                  If rsA.RecordCount > 0 Then
'                     '更新卷宗區回覆單:
'                     strSql = "Update casepaperpdf Set CPP01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "',CPP02=replace(CPP02,'." & textCP10 & ".','.'),CPP10='U' WHERE cpp01='" & strCP09 & "' and instr(Upper(cpp02),upper('." & EMP_回覆單 & ".pdf'))>0"
'                     Pub_SeekTbLog strSql '記錄Log
'                     cnnConnection.Execute strSql
'                  End If
               'Add By Sindy 2016/10/27 程序人員要退回給智權人員,取消結案動作
               ElseIf rsA.Fields("F0309") = Flow_判發退回 Then
                  '恢復程序人員上一個處理動作
                  '流程主檔:
                  strSql = "UPDATE Flow003 SET F0309='" & Flow_處理中 & "' WHERE F0301='" & textCP140 & "'"
                  Pub_SeekTbLog strSql '記錄Log
                  cnnConnection.Execute strSql
                  
                  Call PUB_CloseRestoreLimit(strCP09) 'Modify By Sindy 2020/12/29 改成函數,結案單恢復解除期限
'                  '更新下一程序:
'                  'Modify By Sindy 2018/10/16 + NP06=null,
'                  strSql = "Update NextProgress Set NP06=null,NP24='" & textCP140 & "' WHERE NP24='" & strCP09 & "'"
'                  Pub_SeekTbLog strSql '記錄Log
'                  cnnConnection.Execute strSql
'                  '檢查卷宗區是否有回覆單要恢復
'                  strSql = "SELECT cpp02 FROM casepaperpdf WHERE cpp01='" & strCP09 & "' and instr(Upper(cpp02),upper('." & EMP_回覆單 & ".pdf'))>0"
'                  If rsA.State = adStateOpen Then rsA.Close
'                  rsA.CursorLocation = adUseClient
'                  rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                  If rsA.RecordCount > 0 Then
'                     '更新卷宗區回覆單:
'                     strSql = "Update casepaperpdf Set CPP01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "',CPP02=replace(CPP02,'." & textCP10 & ".','.'),CPP10='U' WHERE cpp01='" & strCP09 & "' and instr(Upper(cpp02),upper('." & EMP_回覆單 & ".pdf'))>0"
'                     Pub_SeekTbLog strSql '記錄Log
'                     cnnConnection.Execute strSql
'                  End If
               '2016/10/27 END
               Else
                  'Modify By Sindy 2025/8/28
'                  iRtn = MsgBox("此筆進度為「電子結案單」，確定是否要刪除電子結案單資料嗎？" & vbCrLf & vbCrLf & _
'                                "（是:一併刪除結案單資料  否:不異動結案單資料  取消:放棄）" & vbCrLf & vbCrLf & _
'                                "（有此訊息時請反應出來，因為要再檢視電子結案單流程）" & vbCrLf, vbYesNoCancel + vbDefaultButton3 + vbCritical)
'                  If iRtn = vbCancel Then
'                     cnnConnection.RollbackTrans
'                     Exit Sub
'                  ElseIf iRtn = vbYes Then
'                     Call PUB_CloseFlowDataDel(textCP140, m_CP01, m_CP02, m_CP03, m_CP04, strCP09)
'                  End If
                  MsgBox "此筆進度為「電子結案單」" & vbCrLf & vbCrLf & _
                         "有此訊息時請反應出來，因為要再檢視電子結案單流程，" & _
                         "不可直接刪除!"
                  Exit Sub
                  '2025/8/28 END
               End If
            End If
         End If
      'Modify By Sindy 2022/10/21 改寫到下面程式段更新
'      'Add By Sindy 2016/7/15 接洽記錄單自動收文刪除時要清NP24=Null
'      ElseIf Len(textCP140) = 10 Then '自動收文
'         '更新下一程序:
'         'Modify By Sindy 2018/10/16 + NP06=null,
'         strSql = "Update NextProgress Set NP06=null,NP24=null WHERE NP24='" & strCP09 & "'"
'         Pub_SeekTbLog strSql '記錄Log
'         cnnConnection.Execute strSql, intI
'      '2016/7/15 END
      End If
   End If
   '2016/4/22 END
   
   strSql = "DELETE FROM CASEPROGRESS " & _
            "WHERE CP01 = '" & m_CP01 & "' AND " & _
                  "CP02 = '" & m_CP02 & "' AND " & _
                  "CP03 = '" & m_CP03 & "' AND " & _
                  "CP04 = '" & m_CP04 & "' AND " & _
                  "CP09 = '" & strCP09 & "' "
    'add by nickc 2006/03/16 紀錄分析語法
    Pub_SeekTbLog strSql
    cnnConnection.Execute strSql
    
    'Added by Lydia 2017/01/17 若刪除的收文號為 EPC母案時要一併刪除相關收文號為該收文號且案件性質相同的子案。
    If m_CP01 = "CFP" And m_CP04 = "00" And m_Nation = "221" Then
       'Add By Sindy 2020/5/26
       strConSql = "cpf01 in(SELECT cp09 FROM CASEPROGRESS WHERE CP43='" & strCP09 & "' AND CP10='" & m_CP10 & "' AND CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03='" & m_CP03 & "' AND CP04<>'00')"
       PUB_DelFtpFile2 "", strConSql, "CASEPAPERFILE"  '檔案改放 FTP,必須在DB資料刪除前執行
       strSql = "delete from casepaperfile where " & strConSql
       Pub_SeekTbLog strSql '記錄Log
       cnnConnection.Execute strSql
       
       strConSql = "cpp01 in(SELECT cp09 FROM CASEPROGRESS WHERE CP43='" & strCP09 & "' AND CP10='" & m_CP10 & "' AND CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03='" & m_CP03 & "' AND CP04<>'00')"
       PUB_DelFtpFile2 "", strConSql '檔案改放 FTP,必須在DB資料刪除前執行
       strSql = "delete from casepaperpdf where " & strConSql
       Pub_SeekTbLog strSql '記錄Log
       cnnConnection.Execute strSql, intI
       '2020/5/26 END
       
       strSql = "DELETE FROM CASEPROGRESS WHERE CP43='" & strCP09 & "' AND CP10='" & m_CP10 & "' AND CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03='" & m_CP03 & "' AND CP04<>'00'"
       cnnConnection.Execute strSql
    End If
    'end 2017/01/17
    
    'Added by Lydia 2017/02/16 若刪除的收文號為 TF母案時要一併刪除相關收文號和案件性質皆相同的子案進度。
    If m_CP01 = "TF" And m_CP03 & m_CP04 = "000" And m_CP43 <> "" Then
       If Mid(m_CP02, 6, 1) = "0" Then '非領土延伸的母案 Ex.TF-000490
           strSql = "DELETE FROM CASEPROGRESS WHERE CP43='" & m_CP43 & "' AND CP10='" & m_CP10 & "' AND CP01='" & m_CP01 & "' AND SUBSTR(CP02,1,5)='" & Mid(m_CP02, 1, 5) & "' AND SUBSTR(CP02,6,1)||CP03||CP04<>'0000' "
       Else
           strSql = "DELETE FROM CASEPROGRESS WHERE CP43='" & m_CP43 & "' AND CP10='" & m_CP10 & "' AND CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03||CP04<>'000' "
       End If
       cnnConnection.Execute strSql
    End If
    'end 2017/02/16
    
    'Add By Cheng 2003/05/30
    '刪除相關的下一程序資料
    'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
    StrSQLa = "Select Decode(PA09,'000',CPM03,CPM04), NP08, NP09 From NextProgress, CasePropertyMap, Patent Where NP02=CPM01(+) And NP07=CPM02(+) And NP02=PA01 And NP03=PA02 And NP04=PA03 And NP05=PA04 And NP01='" & Me.textCP09.Text & "' And " & ChgNextProgress(Me.textCP01.Text & Me.textCP02.Text & Me.textCP03.Text & Me.textCP04.Text)
    StrSQLa = StrSQLa & " Union Select Decode(TM10,'000',CPM03,CPM04), NP08, NP09 From NextProgress, CasePropertyMap, Trademark Where NP02=CPM01(+) And NP07=CPM02(+) And NP02=TM01 And NP03=TM02 And NP04=TM03 And NP05=TM04 And NP01='" & Me.textCP09.Text & "' And " & ChgNextProgress(Me.textCP01.Text & Me.textCP02.Text & Me.textCP03.Text & Me.textCP04.Text)
    StrSQLa = StrSQLa & " Union Select Decode(LC15,'000',CPM03,CPM04), NP08, NP09 From NextProgress, CasePropertyMap, Lawcase Where NP02=CPM01(+) And NP07=CPM02(+) And NP02=LC01 And NP03=LC02 And NP04=LC03 And NP05=LC04 And NP01='" & Me.textCP09.Text & "' And " & ChgNextProgress(Me.textCP01.Text & Me.textCP02.Text & Me.textCP03.Text & Me.textCP04.Text)
    StrSQLa = StrSQLa & " Union Select Decode('000','000',CPM03,CPM04), NP08, NP09 From NextProgress, CasePropertyMap, Hirecase Where NP02=CPM01(+) And NP07=CPM02(+) And NP02=HC01 And NP03=HC02 And NP04=HC03 And NP05=HC04 And NP01='" & Me.textCP09.Text & "' And " & ChgNextProgress(Me.textCP01.Text & Me.textCP02.Text & Me.textCP03.Text & Me.textCP04.Text)
    StrSQLa = StrSQLa & " Union Select Decode(SP09,'000',CPM03,CPM04), NP08, NP09 From NextProgress, CasePropertyMap, Servicepractice Where NP02=CPM01(+) And NP07=CPM02(+) And NP02=SP01 And NP03=SP02 And NP04=SP03 And NP05=SP04 And NP01='" & Me.textCP09.Text & "' And " & ChgNextProgress(Me.textCP01.Text & Me.textCP02.Text & Me.textCP03.Text & Me.textCP04.Text)
    'end 2018/06/05
    If rsA.State = adStateOpen Then rsA.Close
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        strMsg = ""
        While Not rsA.EOF
            strMsg = strMsg & rsA.Fields(0).Value & ", (本所)" & ChangeTStringToTDateString(ChangeWStringToTString("" & rsA.Fields(1).Value)) & ", (法定)" & ChangeTStringToTDateString(ChangeWStringToTString("" & rsA.Fields(2).Value)) & vbCrLf
            rsA.MoveNext
        Wend
        If strMsg <> "" Then
            strTit = "詢問"
            strMsg = "是否要刪除相關下一程序資料?" & vbCrLf & strMsg
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbYes Then
                'Add By Cheng 2003/06/10
                '將下一程序資料新增資料至資料刪除記錄檔
                strSql = "Insert Into DataDeleteRecord (Select NP02, NP03, NP04, NP05, '', '', '', '', '', '', '', '', '', NP01, NP07, NP08, NP09, '', NP10, '', '', '', '" & strDD23 & "', '" & strDD24 & "', NP17, NP16," & CDbl(strSrvDate(1)) & ", " & CDbl(GetMaxNumber) & " From NextProgress " & _
                            " Where NP02='" & m_CP01 & "' And NP03='" & m_CP02 & "' And NP04='" & m_CP03 & "' And NP05='" & m_CP04 & "' And NP01='" & strCP09 & "') "
                'add by nickc 2006/03/16 紀錄分析語法
                Pub_SeekTbLog strSql
                cnnConnection.Execute strSql
                '刪除下一程序資料
                'Modify By Sindy 2013/10/2 發明案來函要同時刪新型案的放棄權用權,所以取消本所案號
                strSql = "DELETE FROM NextProgress " & _
                         "WHERE NP01 = '" & strCP09 & "' "
                'add by nickc 2006/03/16 紀錄分析語法
                Pub_SeekTbLog strSql
                cnnConnection.Execute strSql
            End If
        End If
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
   
   'Add By Sindy 2012/9/5
   'T,TF,CFT,FCT若刪除案件性質1704專用權消滅,商標基本檔的TM17='Y'
   'P,CFP,FCP   若刪除案件性質1604專利權消滅,專利基本檔的PA17='Y'
   '增加判斷有沒有專用起止日,有存Y,無存null
   If m_CP53 <> "" And m_CP54 <> "" Then
      strTM17 = "Y"
   Else
      strTM17 = ""
   End If
   If (m_CP01 = "T" Or m_CP01 = "TF" Or m_CP01 = "CFT" Or m_CP01 = "FCT") And textCP10.Text = "1704" Then
      strSql = "update trademark set tm17=" & CNULL(strTM17) & _
               " WHERE tm01 = '" & m_CP01 & "' AND " & _
                      "tm02 = '" & m_CP02 & "' AND " & _
                      "tm03 = '" & m_CP03 & "' AND " & _
                      "tm04 = '" & m_CP04 & "' "
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   ElseIf (m_CP01 = "P" Or m_CP01 = "CFP" Or m_CP01 = "FCP") And textCP10.Text = "1604" Then
      strSql = "update patent set pa17=" & CNULL(strTM17) & _
               " WHERE pa01 = '" & m_CP01 & "' AND " & _
                      "pa02 = '" & m_CP02 & "' AND " & _
                      "pa03 = '" & m_CP03 & "' AND " & _
                      "pa04 = '" & m_CP04 & "' "
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   '2012/9/5 End
   
   'Added by Morgan 2025/10/2
   '刪除P台灣專利通知擇一申復1232要清除基本檔的PA60(一案兩請是否放棄新型)，否則會導致後續發明核准會沒提醒要確認是否維持一案兩請並於發明發證後新型被自動閉卷。 Ex:P-133809 --韻丞
   If m_CP01 = "P" And m_Nation = "000" And textCP10.Text = "1232" Then
      strSql = "update patent set pa60=''" & _
               " WHERE pa01 = '" & m_CP01 & "' AND " & _
                      "pa02 = '" & m_CP02 & "' AND " & _
                      "pa03 = '" & m_CP03 & "' AND " & _
                      "pa04 = '" & m_CP04 & "' and pa60 is not null"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   'end 2025/10/2
   
   ' 通知前畫面有刪除的記錄
   frm075004_1.DelRecord strCP09

   ' 刪除記錄時, 除了要刪除資料庫中的記錄還要刪除在本程式模組中所記錄的資料
   ' 以供瀏覽資料時會更新資料串列的順序及正確性
   nDataListCount = 0
   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex).diCP09 <> strCP09 Then
         ReDim Preserve strDataList(nDataListCount + 1)
         strDataList(nDataListCount) = m_DataList(nIndex).diCP09
         nDataListCount = nDataListCount + 1
      Else
         nPos = nIndex
      End If
   Next nIndex
   ' 清除資料串列
   ClearDataList
   ' 通知前畫面刪除該筆資料
   frm075004_1.DelRecord strCP09
   ' 將資料更新回到資料串列記錄中
   For nIndex = 0 To nDataListCount - 1
      SetDataListItem strDataList(nIndex)
   Next nIndex
   
   'Added by Lydi 2017/11/29 刪除FCP案件命名:因為要抓FCP管制人,所以放在刪除基本檔前面
   If strSrvDate(1) >= FCP案件命名啟用日 And textCP01 = "FCP" And InStr(NewCasePtyList, textCP10) > 0 Then
      If PUB_GetTCTmail(True, 3, textCP01, textCP02, textCP03, textCP04, strCP09) Then
         strSql = "delete from TransCaseTitle where tct01='" & strCP09 & "' "
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2017/11/29
   
   'Added by Lydia 2018/01/09 刪除翻譯費用檔
   If InStr(FCPHaveEP04, m_CP10) > 0 Then '分析有tranfee的收文號案件性質：201,209,210,927
       strSql = "delete from transfee where tf01='" & strCP09 & "' "
       cnnConnection.Execute strSql, intI
   End If
   'end 2018/01/09
   
   'Added by Lydia 2018/02/02 若人員要求刪除D類收文進度，一併刪除客戶提供文件記錄；而客戶提供文件處理只刪D類收文進度。
   If Left(strCP09, 1) = "D" And m_CP10 = "1920" Then
       strSql = "delete from CustSupportDoc where CSD05='" & strCP09 & "' "
       cnnConnection.Execute strSql, intI
   End If
   'end 2018/02/02
   
   'Added by Lydia 2024/11/21 一併刪除；內商大陸之部份核駁商品異動，記錄各類的比對結果
   If m_CP01 = "T" And m_Nation = "020" Then
       strSql = "delete from tmgoods where tg01='" & Mid(strCP09, 1, 3) & "' and tg02='" & Mid(strCP09, 4, 6) & "' and tg03='0' and tg04='00' "
       cnnConnection.Execute strSql, intI
   End If
   'end 2024/11/21
   
   'Added by Lydia 2018/07/03 若刪除D類收款寄證,一併刪除定稿檔(ex.T-142048,T-199535 桂英在4/25輸入核准,又請電腦中心刪除C類核准和D類收款寄證,但是沒有刪除定稿,後面雅雯輸入核准,在發文收款寄證時造成錯誤)
   'Mark by Lydia 2024/06/05 因為basLetter有「2019/11/21 新增前先把舊資料刪除」這段不用了;
   'If m_DelLD01 <> "" And m_DelLD04 <> "" And m_DelLD10 <> "" Then
   '     strSql = "delete from letterdemand where ld05='" & m_CP01 & "' and ld06='" & m_CP02 & "' and ld07='" & m_CP03 & "' and ld08='" & m_CP04 & "' " & _
   '                 "and ld01='" & m_DelLD01 & "' and ld04='" & m_DelLD04 & "' and ld10='" & m_DelLD10 & "'  "
   '     cnnConnection.Execute strSql, intI
   '     strSql = "delete from exceptcondition where et04='" & m_DelLD01 & "'  and et02='" & m_DelLD04 & "' and et01='" & m_DelLD10 & "' "
   '     cnnConnection.Execute strSql, intI
   'End If
   'end 2024/06/05
   'end 2018/07/03
   
   'Added by Lydia 2020/02/13 外專新案刪除，一併刪除專利案件和English_Vers檔案
   'Move by Lydia 2020/06/11 從下面搬來, 因為要在刪除基本檔的前面
   If strAdd01 <> "" Then  '專利案件
        PUB_DelFtpFile2 strAdd01, , "CASEPAPERFILE" '檔案改放 FTP,必須在DB資料刪除前執行
        strSql = "delete from casepaperfile where cpf01='" & strAdd01 & "'"
        Pub_SeekTbLog strSql '記錄Log
        cnnConnection.Execute strSql
        strSql = "delete from caseprogress where cp09='" & strAdd01 & "' "
        Pub_SeekTbLog strSql '記錄Log
        cnnConnection.Execute strSql
   End If
   If strAdd02 <> "" Then  'English_Vers
        PUB_DelFtpFile2 strAdd02, , "CASEPAPERFILE" '檔案改放 FTP,必須在DB資料刪除前執行
        strSql = "delete from casepaperfile where cpf01='" & strAdd02 & "'"
        Pub_SeekTbLog strSql '記錄Log
        cnnConnection.Execute strSql
        strSql = "delete from caseprogress where cp09='" & strAdd02 & "' "
        Pub_SeekTbLog strSql '記錄Log
        cnnConnection.Execute strSql
   End If
   'end 2020/02/13
   
   ' 若沒有案件進度資料時, 則必須刪除主檔及優先權檔及相關卷號檔...
   If ExistProgress(m_CP01, m_CP02, m_CP03, m_CP04) = False Then
      If OnDataDeleteRecord(0, m_CP01 & m_CP02 & m_CP03 & m_CP04, Empty, strDD23, strDD24, False) = 0 Then
      
         PUB_DelPatentRefData m_CP01, m_CP02, m_CP03, m_CP04, Me 'Added by Morgan 2025/6/25 以共用函數刪除相關資料
         
         Select Case m_CP01
            ' 讀取商標基本檔
            Case "T", "TF", "CFT", "FCT":
               strSql = "DELETE FROM TRADEMARK " & _
                        "WHERE TM01 = '" & m_CP01 & "' AND " & _
                              "TM02 = '" & m_CP02 & "' AND " & _
                              "TM03 = '" & m_CP03 & "' AND " & _
                              "TM04 = '" & m_CP04 & "' "
            ' 讀取專利基本檔
            Case "P", "CFP", "FCP":
               strSql = "DELETE FROM PATENT " & _
                        "WHERE PA01 = '" & m_CP01 & "' AND " & _
                              "PA02 = '" & m_CP02 & "' AND " & _
                              "PA03 = '" & m_CP03 & "' AND " & _
                              "PA04 = '" & m_CP04 & "' "
            ' 讀取法務基本檔
            'Modify By Sindy 2009/07/24 增加LIN系統類別
            'modify by sonia 2019/7/24 +ACS系統類別
            Case "L", "CFL", "FCL", "LIN", "ACS":
               strSql = "DELETE FROM LAWCASE " & _
                        "WHERE LC01 = '" & m_CP01 & "' AND " & _
                              "LC02 = '" & m_CP02 & "' AND " & _
                              "LC03 = '" & m_CP03 & "' AND " & _
                              "LC04 = '" & m_CP04 & "' "
            ' 讀取顧問案件基本檔
            Case "LA":
               strSql = "DELETE FROM HIRECASE " & _
                        "WHERE HC01 = '" & m_CP01 & "' AND " & _
                              "HC02 = '" & m_CP02 & "' AND " & _
                              "HC03 = '" & m_CP03 & "' AND " & _
                              "HC04 = '" & m_CP04 & "' "
            ' 讀取服務業務基本檔
            Case Else:
               strSql = "DELETE FROM SERVICEPRACTICE " & _
                        "WHERE SP01 = '" & m_CP01 & "' AND " & _
                              "SP02 = '" & m_CP02 & "' AND " & _
                              "SP03 = '" & m_CP03 & "' AND " & _
                              "SP04 = '" & m_CP04 & "' "
         End Select
         ' 執行刪除的指令
         'add by nickc 2006/03/16 紀錄分析語法
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
      
      
'Removed by Morgan 2025/6/25 改在前面以公用函數刪除
'      ' 刪除優先權檔
'      'add by nickc 2006/03/20
'      strSql = "SELECT * FROM PRIDATE " & _
'                   "WHERE PD01 = '" & m_CP01 & "' AND " & _
'                         "PD02 = '" & m_CP02 & "' AND " & _
'                         "PD03 = '" & m_CP03 & "' AND " & _
'                         "PD04 = '" & m_CP04 & "' "
'      If rsA.State = adStateOpen Then rsA.Close
'      rsA.CursorLocation = adUseClient
'      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'            strSql = "DELETE FROM PRIDATE " & _
'               "WHERE PD01 = '" & m_CP01 & "' AND " & _
'                     "PD02 = '" & m_CP02 & "' AND " & _
'                     "PD03 = '" & m_CP03 & "' AND " & _
'                     "PD04 = '" & m_CP04 & "' "
'            'add by nickc 2006/03/16 紀錄分析語法
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'      End If
'      ' 刪除相關卷號檔
'      'add by nickc 2006/03/20
'      strSql = "SELECT * FROM CASERELATION " & _
'               "WHERE (CR01 = '" & m_CP01 & "' AND " & _
'                      "CR02 = '" & m_CP02 & "' AND " & _
'                      "CR03 = '" & m_CP03 & "' AND " & _
'                      "CR04 = '" & m_CP04 & "') OR " & _
'                     "(CR05 = '" & m_CP01 & "' AND " & _
'                      "CR06 = '" & m_CP02 & "' AND " & _
'                      "CR07 = '" & m_CP03 & "' AND " & _
'                      "CR08 = '" & m_CP04 & "') "
'      If rsA.State = adStateOpen Then rsA.Close
'      rsA.CursorLocation = adUseClient
'      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'          strSql = "DELETE FROM CASERELATION " & _
'               "WHERE (CR01 = '" & m_CP01 & "' AND " & _
'                      "CR02 = '" & m_CP02 & "' AND " & _
'                      "CR03 = '" & m_CP03 & "' AND " & _
'                      "CR04 = '" & m_CP04 & "') OR " & _
'                     "(CR05 = '" & m_CP01 & "' AND " & _
'                      "CR06 = '" & m_CP02 & "' AND " & _
'                      "CR07 = '" & m_CP03 & "' AND " & _
'                      "CR08 = '" & m_CP04 & "') "
'         'add by nickc 2006/03/16 紀錄分析語法
'         Pub_SeekTbLog strSql
'         cnnConnection.Execute strSql
'       End If
'      ' 刪除相關卷號檔
'      'add by nickc 2006/06/22
'      strSql = "SELECT * FROM CASERELATION1 " & _
'               "WHERE (CR01 = '" & m_CP01 & "' AND " & _
'                      "CR02 = '" & m_CP02 & "' AND " & _
'                      "CR03 = '" & m_CP03 & "' AND " & _
'                      "CR04 = '" & m_CP04 & "') OR " & _
'                     "(CR05 = '" & m_CP01 & "' AND " & _
'                      "CR06 = '" & m_CP02 & "' AND " & _
'                      "CR07 = '" & m_CP03 & "' AND " & _
'                      "CR08 = '" & m_CP04 & "') "
'      If rsA.State = adStateOpen Then rsA.Close
'      rsA.CursorLocation = adUseClient
'      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'          strSql = "DELETE FROM CASERELATION1 " & _
'               "WHERE (CR01 = '" & m_CP01 & "' AND " & _
'                      "CR02 = '" & m_CP02 & "' AND " & _
'                      "CR03 = '" & m_CP03 & "' AND " & _
'                      "CR04 = '" & m_CP04 & "') OR " & _
'                     "(CR05 = '" & m_CP01 & "' AND " & _
'                      "CR06 = '" & m_CP02 & "' AND " & _
'                      "CR07 = '" & m_CP03 & "' AND " & _
'                      "CR08 = '" & m_CP04 & "') "
'         'add by nickc 2006/03/16 紀錄分析語法
'         Pub_SeekTbLog strSql
'         cnnConnection.Execute strSql
'      End If
'      'Added by Lydia 2021/11/29 刪除分割案件關係檔DivisionCase
'      strSql = "SELECT * FROM DIVISIONCASE WHERE (DC01='" & m_CP01 & "' AND DC02='" & m_CP02 & "' AND DC03='" & m_CP03 & "' AND DC04='" & m_CP04 & "') " & _
'                  "OR (DC05='" & m_CP01 & "' AND DC06='" & m_CP02 & "' AND DC07='" & m_CP03 & "' AND DC08='" & m_CP04 & "') "
'      If rsA.State = adStateOpen Then rsA.Close
'      rsA.CursorLocation = adUseClient
'      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'           strSql = "DELETE FROM DIVISIONCASE WHERE (DC01='" & m_CP01 & "' AND DC02='" & m_CP02 & "' AND DC03='" & m_CP03 & "' AND DC04='" & m_CP04 & "') " & _
'                       "OR (DC05='" & m_CP01 & "' AND DC06='" & m_CP02 & "' AND DC07='" & m_CP03 & "' AND DC08='" & m_CP04 & "') "
'           Pub_SeekTbLog strSql
'           cnnConnection.Execute strSql
'      End If
'      'end 2021/11/29
'end 2025/6/25
      
      'Add By Sindy 2011/8/11
      ' 刪除庭期資料
      strSql = "SELECT * FROM COURTYARDPERIOD " & _
               "WHERE CDP01='" & strCP09 & "'"
      If rsA.State = adStateOpen Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
          strSql = "DELETE FROM COURTYARDPERIOD " & _
               "WHERE CDP01='" & strCP09 & "'"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
      '2011/8/11 End
      
      'Add By Sindy 2022/10/21
      ' 刪除發明人資料
      strSql = "SELECT * FROM patentInventor " & _
               "WHERE pi01=" + CNULL(m_CP01) + " and pi02=" + CNULL(m_CP02) + " and pi03=" + CNULL(m_CP03) + " and pi04=" + CNULL(m_CP04)
      If rsA.State = adStateOpen Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
          strSql = "DELETE FROM patentInventor " & _
               "WHERE pi01=" + CNULL(m_CP01) + " and pi02=" + CNULL(m_CP02) + " and pi03=" + CNULL(m_CP03) + " and pi04=" + CNULL(m_CP04)
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
      '2022/10/21 End
      
      'Added by Lydia 2020/05/11 刪除特殊備註和各項指示
      If m_CP01 = "P" Or m_CP01 = "FCP" Then '特殊備註
            '下一程序固定備註(NpMemo)
            strSql = "DELETE NPMEMO WHERE NM03='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' "
            cnnConnection.Execute strSql
            '核准函輸入備註(ApprovalMemo2)
            strSql = "DELETE APPROVALMEMO2 WHERE AM03='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' "
            cnnConnection.Execute strSql
            '核駁及審查意見通知函備註(IncomMemo)
            strSql = "DELETE INCOMMEMO WHERE IM03='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' "
            cnnConnection.Execute strSql
            '請款函預設備註維護檔(DebitNotePS)
            strSql = "DELETE DEBITNOTEPS WHERE DNPS03='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' "
            cnnConnection.Execute strSql
            'end 2018/11/28
            'FCP承辦單設定維護(FcpEMPbill)
            strSql = "DELETE FCPEMPBILL WHERE FEB03='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' "
            cnnConnection.Execute strSql
            '通知告准加註(ApprovalPS)
            strSql = "DELETE APPROVALPS WHERE APS03='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' "
            cnnConnection.Execute strSql
      End If
      '各項指示
      strSql = "SELECT * FROM INSTRUCTIONS WHERE ITS02='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strSql = "DELETE FROM INSTRUCTIONS WHERE ITS02='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' "
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
      'end 2020/05/11
   End If
   
   'Modify by Morgan 2006/4/24 改在Transaction開始前詢問才不會鎖住資料
   If bolReControl = True Then
      'Modify By Sindy 2022/10/21
      'Modify By Sindy 2023/3/9 Mark:,NP24=null
      strSql = "UPDATE NEXTPROGRESS SET NP06=NULL WHERE NP24='" & strCP09 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      '2022/10/21 END
      
      'edit by nickc 2008/03/26 morgan 少寫一個條件
      'strSQL = "UPDATE NEXTPROGRESS SET NP06=NULL WHERE NP01='" & m_CP43 & "' AND NP02='" & m_CP01 & "' AND NP03='" & m_CP02 & "' AND NP04='" & m_CP03 & "' AND NP07='" & m_CP10 & "' AND NP06='Y'"
      'Modified by Lydia 2023/03/22 法律案(L,FCL)不同審級收文期限沖銷：案件不同審級會收文-1、-2…案號，但實際是同一案件。
      If m_CP01 = "L" Or m_CP01 = "FCL" Then
         strSql = "UPDATE NEXTPROGRESS SET NP06=NULL WHERE NP01='" & m_CP43 & "' AND NP02='" & m_CP01 & "' AND NP03='" & m_CP02 & "' AND NP07='" & m_CP10 & "' AND NP06='Y'"
      Else
      'end 2023/03/22
         strSql = "UPDATE NEXTPROGRESS SET NP06=NULL WHERE NP01='" & m_CP43 & "' AND NP02='" & m_CP01 & "' AND NP03='" & m_CP02 & "' AND NP04='" & m_CP03 & "' and np05='" & m_CP04 & "'" & _
                  " AND NP07='" & m_CP10 & "' AND NP06='Y'"
      End If 'Added by Lydia 2023/03/22
      'add by nickc 2006/03/16 紀錄分析語法
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   
   'Add by Amy 2013/07/02 刪除時案件命名追蹤同收文號設 案號+刪除
   strSql = "Update TrackingCaseName set TCN05='" & textCP01 & "-" & textCP02 & "-" & textCP03 & "刪除" & "' Where TCN05='" & textCP09 & "' "
   cnnConnection.Execute strSql
   'end 2013/07/02
   
   'Add By Sindy 2013/10/2 刪除承辦電子簽核相關檔案
   strSql = "delete from empelectronprocess where eep01='" & strCP09 & "'"
   cnnConnection.Execute strSql
   'Add By Sindy 2024/10/30 此文號若為附加流程而產生的進度; 則增加刪除此文號的源頭
   strSql = "delete from empelectronprocess where eep04='" & EMP_附加流程 & "' and instr(eep11,'" & strCP09 & "')>0"
   cnnConnection.Execute strSql, intI
   If intI > 1 Then
      MsgBox "刪除此文號的源頭(附加流程)有誤!!"
      GoTo ErrHnd
   End If
   '2024/10/30 END
   PUB_DelFtpFile2 strCP09, , "EMPELECTRONFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
   strSql = "delete from empelectronfile where eef01='" & strCP09 & "'"
   cnnConnection.Execute strSql
   
   strSql = "delete from empelectrondata where eed01='" & strCP09 & "'"
   cnnConnection.Execute strSql
   
   PUB_DelFtpFile2 strCP09, , "CASEPAPERFILE" 'Added by Morgan 2015/4/28 檔案改放 FTP,必須在DB資料刪除前執行
   strSql = "delete from casepaperfile where cpf01='" & strCP09 & "'"
   Pub_SeekTbLog strSql '記錄Log
   cnnConnection.Execute strSql
   
   'Add by Amy 2014/08/27 電子化-刪申請書轉檔記錄
   strSql = "Delete AppForm Where AF01='" & strCP09 & "'"
   cnnConnection.Execute strSql, intI
   'end 2014/08/27

   '卷宗區
   'Modified by Morgan 2014/1/15 +判斷若為電子機關來函時Edocument.ED11要清除且cpp01要還原為發文文號
   'strSql = "select cpp01,cpp10 from casepaperpdf where cpp01='" & strCP09 & "'"
   strSql = "select cpp01,cpp10,ed01,lp01 from casepaperpdf,edocument,letterprogress where cpp01='" & strCP09 & "' and ed11(+)=cpp01 and lp01(+)=cpp01"
   If rsA.State = adStateOpen Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      'Added by Morgan 2014/1/15
      'Modified by Morgan 2014/6/10
      '若為電子公文時要將資料還原
      If Not IsNull(rsA.Fields("ed01")) Then
         'Added by Morgan 2022/12/19 專利證書/註冊證先單獨還原
         'Modified by Morgan 2023/2/2
         'strExc(1) = PUB_GetEDocFileName(textCP01, textCP02, textCP03, textCP04, textCP10, True)
         'strSql = "update casepaperpdf a set cpp01='" & rsA.Fields("ed01") & "',cpp02='$" & rsA.Fields("ed01") & ".CERT.pdf',cpp10='U' where cpp01='" & strCP09 & "' and upper(cpp02)=upper('" & strExc(1) & "')"
         strExc(1) = PUB_CaseNo2FileName(textCP01, textCP02, textCP03, textCP04) & "." & textCP10
         strSql = "update casepaperpdf a set cpp01='" & rsA.Fields("ed01") & "',cpp02=replace(cpp02,'" & strExc(1) & "','$" & rsA.Fields("ed01") & "'),cpp10='U' where cpp01='" & strCP09 & "' and cpp02 like '" & strExc(1) & ".%.pdf'"
         'end 2023/2/2
         cnnConnection.Execute strSql, intI
         'end 2022/12/19
            
         'Modified by Morgan 2017/5/16 FCP檔名規則不同
         'strExc(1) = textCP01 & Val(textCP02) & IIf(textCP03 & textCP04 = "000", "", "-" & textCP03 & "-" & textCP04) & "." & textCP10 & ".pdf"
         'Modified by Morgan 2017/8/14 FCP改判斷CPP15='0'
         If textCP01 = "FCP" Then
            strSql = "update casepaperpdf a set cpp01='" & rsA.Fields("ed01") & "',cpp02='$" & rsA.Fields("ed01") & ".pdf',cpp10='U' where cpp01='" & strCP09 & "' and cpp10<>'D' and cpp15='0'"
            cnnConnection.Execute strSql, intI
         Else
         'end 2017/8/14
         
            strExc(1) = PUB_GetEDocFileName(textCP01, textCP02, textCP03, textCP04, textCP10)
         'end 2017/5/16
            strSql = "update casepaperpdf set cpp01='" & rsA.Fields("ed01") & "',cpp02='$" & rsA.Fields("ed01") & ".pdf',cpp10='U' where cpp01='" & strCP09 & "'  and upper(cpp02)=upper('" & strExc(1) & "')"
            cnnConnection.Execute strSql, intI
            
         End If 'Added by Morgan 2017/8/14
         
         strSql = "update edocument set ed11='C' where ed11='" & strCP09 & "'"
         cnnConnection.Execute strSql, intI
         
      End If 'Added by Morgan 2016/8/10
      
      'Added by Morgan 2014/4/8 除客戶函外,總收文號更新為本所案號
      'Modified by Morgan 2015/5/4 +判斷有信函進度者
      'Modified by Morgan 2016/8/10 不必排除電子公文,因有可能改案件性質或有其他附件
      'ElseIf strCP09 > "C" And Not IsNull(rsA.Fields("lp01")) Then
      If strCP09 > "C" And Not IsNull(rsA.Fields("lp01")) Then
         'Modified by Morgan 2014/12/2 排除有刪除註記者
         'Modified by Morgan 2015/4/2 更新後註記改放 'C'待收文來函
         'Modified by Morgan 2018/10/2 +.BLANK.PDF(回覆單),.ORDER.PDF(接洽單) 也要刪
         'Modified by Morgan 2020/9/15 +專利信件區來的也要刪(否則從收件區重輸時又轉入會導致Dupe)
         strSql = "update casepaperpdf set cpp01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "',cpp10='C' where cpp01='" & strCP09 & "' and instr(upper(cpp02),'." & m_CP10 & ".CUS.PDF')=0 and instr(upper(cpp02),'." & m_CP10 & ".BLANK.PDF')=0 and instr(upper(cpp02),'." & m_CP10 & ".ORDER.PDF')=0 and NVL(cpp12,' ')<>'P' and cpp10<>'D'"
         'Added by Morgan 2025/3/20 台灣商標分割子案核准的電子公文也要刪,否則重輸時又複製母案公文會導致Dupe
         If (m_CP01 = "T" Or m_CP01 = "FCT") And m_Nation = "000" And textCP10 = "1001" Then
            strExc(0) = "select cp09 from caseprogress where cp09='" & textCP43 & "' and cp10='308'"
            intI = 1
            Set rsA = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               '檢查相同機關文號的電子公文紀錄
               strExc(1) = UCase(PUB_GetEDocFileName(textCP01, textCP02, textCP03, textCP04, textCP10))
               strSql = strSql & " and not exists(select * from caseprogress a,edocument where cp08='" & textCP08 & "' and ed11(+)=cp09 and ed01 is not null and and upper(cpp02)='" & strExc(1) & "')"
            End If
         End If
         'end 2025/3/20
         cnnConnection.Execute strSql, intI
      'end 2014/4/8
      End If
      'end 2014/6/10
      'end 2014/1/15

'      If rsA.Fields("cpp10") = "Y" Then
'         '已合併時,必須刪除整份卷宗
'         strSql = "delete from casepaperpdf where cpp01='000000000' and upper(cpp02)='" & UCase(m_CP01 & m_CP02 & m_CP03 & m_CP04 & ".pdf") & "'"
'         cnnConnection.Execute strSql
'         '並且要將此本所案號的全部附件已合併欄位值改成X,晚上批次作業必須再全部合併一次
'         'Modified by Morgan 2014/11/24 +控制 index 不要用 cpp10 否則會很慢
'         strSql = "update casepaperpdf set cpp10='X' where cpp01 in(" & _
'                  "select cp09 from caseprogress where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "')" & _
'                  " and cpp10||''='Y'"
'         cnnConnection.Execute strSql
'      End If
      
      PUB_DelFtpFile2 strCP09 'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
      strSql = "delete from casepaperpdf where cpp01='" & strCP09 & "'"
      Pub_SeekTbLog strSql '記錄Log
      cnnConnection.Execute strSql, intI
   End If
   '2013/10/2 END
   
   'Added by Morgan 2014/4/11 電子化-刪除信函進度檔
   strSql = "delete letterprogress where lp01='" & strCP09 & "'"
   cnnConnection.Execute strSql, intI
   'end 2014/4/11
   
   'Add By Sindy 2016/1/29 增加刪除接洽記錄單相關檔案
   'Modify By Sindy 2022/10/24 + And Len(textCP140) = 10
   If textCP140 <> "" And Len(textCP140) = 10 Then '加判斷接洽單長度10碼,以免誤刪結案單
      'Modify By Sindy 2022/10/21 檢查進度檔若有其他筆文號掛同一個接洽單編號時,不能刪除
      strExc(0) = "select cp09 from caseprogress where cp140='" & textCP140 & "' and cp09<>'" & strCP09 & "'"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then '無資料才能刪除
         If PUB_DelCRLAllData(textCP140.Text) = False Then
            GoTo ErrHnd
         End If
         
      'Add By Sindy 2023/1/16 接洽單收多個案件性質只刪某一案件性質時, 接洽單做刪收文註記,和加註簽核備註
      Else
         '接洽記錄單案件性質
         strSql = "UPDATE ConsultRecCMP SET CRC04='刪收文',CRC05='刪收文',CRC06='刪收文'" & _
                  " WHERE CRC08='" & strCP09 & "'"
         cnnConnection.Execute strSql
         '流程備註檔
         strSql = GetInsertFLOW004Sql(textCP140, strUserNum, strSrvDate(1), Right("000000" & ServerTime, 6), "", _
            "刪除" & textCP10_2 & "收文")
         cnnConnection.Execute strSql
         '2023/1/16 END
      End If
      If rsA.State = adStateOpen Then rsA.Close
'      strSql = "delete from consultrecordlist where crl01='" & textCP140 & "'"
'      cnnConnection.Execute strSql, intI
'      strSql = "delete from consultrecApp where cra01='" & textCP140 & "'"
'      cnnConnection.Execute strSql, intI
'      strSql = "delete from consultrecInv where cri01='" & textCP140 & "'"
'      cnnConnection.Execute strSql, intI
'
'      PUB_DelFtpFile2 textCP140, , UCase("ConsultRecImageF") '檔案改放FTP,必須在DB資料刪除前執行
'      strSql = "delete from ConsultRecImageF where crif01='" & textCP140 & "'"
'      cnnConnection.Execute strSql, intI
      '2022/10/21 END
   End If
   '2016/1/29 END
      
   'Memo by Lydia 2020/06/11 搬到上面,因為要在刪除基本檔的前面
   
   'Added by Morgan 2016/6/2
   If bolDelRefCP Then
      'Modified by Morgan 2016/12/29 +其他相關案號也要刪除記錄
      'strSql = "delete caseprogress where cp28='" & textCP28 & "' and cp09<>'" & textCP28 & "' and cp10='" & textCP10 & "'"
      'cnnConnection.Execute strSql, intI
      strExc(0) = "select cp09 from caseprogress where cp28='" & textCP28 & "' and cp09<>'" & textCP28 & "' and cp10='" & textCP10 & "'"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Do While Not rsA.EOF
            strSql = "delete caseprogress where cp09='" & rsA(0) & "'"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
            rsA.MoveNext
         Loop
      End If
      If rsA.State = adStateOpen Then rsA.Close
      'end 2016/12/29
   End If
   'end 2016/6/2
   
   'Added by Lydia 2017/06/06 一併刪除加乘註記修改記錄
   strSql = "delete from flagstory where fs01='" & strCP09 & "'"
   cnnConnection.Execute strSql, intI
   
   'Added by Morgan 2017/10/26
   '若為期限通知,報價也要刪除 Ex:CFP-27198(DA6018489)
   If m_CP10 = "1913" Then
      'Modified by Morgan 2020/9/16 +LD18條件
      strSql = "delete from lettercache where lc01='" & textCP43 & "' and lc02='" & textCP30 & "' and lc18='" & strCP09 & "'"
      cnnConnection.Execute strSql, intI
      If intI = 1 Then
         strSql = "delete from lettercachevar where lcv01='" & textCP43 & "' and lcv02='" & textCP30 & "'"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2017/10/26
   
   'Added by Lydia 2020/04/15 一併刪除法務工作點數
   If InStr(m_CP01, "L") > 0 Then
      strSql = "delete from acc1n0 where a1n01='" & strCP09 & "' "
      cnnConnection.Execute strSql, intI
       
      'Add By Sindy 2020/5/6
      '法務案刪除時，更新法律所案源檔之放棄日期及放棄人員，放棄原因：「放棄原因：法務案刪除」。
      'AB類還要更新TT-999999進度之取消收文日，原因99，進度備註「放棄原因：法務案刪除」。
      strSql = "UPDATE lawofficesource SET los07=" & strSrvDate(1) & ",los08='" & strDD23 & "',los09='放棄原因：法務案刪除'" & _
               " WHERE los06='" & strCP09 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      'modify by sonia 2021/8/2 同時取消發文日及費用規費點數
      strSql = "UPDATE caseprogress SET CP27=NULL,CP16=0,CP17=0,CP18=0,CP57=" & strSrvDate(1) & ",CP58='99',CP64='放棄原因：法務案刪除;'||CP64" & _
               " WHERE CP09 in(SELECT LOS10 FROM lawofficesource WHERE los06='" & strCP09 & "' AND LOS10 is not null)"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      '2020/5/6 END
   End If
   'end 2020/04/15
   
   'Add By Sindy 2020/5/6
   'TT-999999案刪進度時，檢查法律所案源檔，必須為未收法務案且未放棄始可刪除，同時刪除法律所案源檔。
   '需請介紹人(智權人員)跟法律所人員確認後由法律所人員通知電腦中心刪除，刪除後Mail給通知人及介紹人。
   If m_CP01 = "TT" And m_CP02 = "999999" Then
      'add by sonia 2025/8/1 取消法務案之案源單號CP162，否則FCL-011014(收文號AB4013060)請款單X11405029在收款時會錯誤
      strSql = "update caseprogress set cp162=null WHERE cp09 in (select los06 FROM lawofficesource WHERE los10='" & strCP09 & "')"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      'end 2025/8/1
      strSql = "DELETE FROM lawofficesource WHERE los10='" & strCP09 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   '2020/5/6 END
   
   'Add By Sindy 2023/3/7 檢查接洽單的Flow是否要結束
   Call PUB_UpdateCRLFlowClose(textCP140, strCP09) 'Modify By Sindy 2023/12/13 改共用函數
   
   'Add By Sindy 2025/6/24 檢查是否有信件多案收文資料,若有一併刪除文號欄位值,才能再重新收文
   strSql = "UPDATE multiCaseRecv SET mcr11=NULL WHERE mcr11='" & strCP09 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   
   cnnConnection.CommitTrans 'Add End 2004/2/12
   
   'Modify by Morgan 2006/4/24 若點多筆時恢復期限管制會更新錯資料,改在commit後再做
   bolTrans = False 'Add End 2004/2/12
   
   ' 刪除暫存串列
   'Modified by Lydia 2020/06/11 FCP新申請案刪除,一併刪除EnglishVers和專利案件
   'If m_DataListCount <= 0 Then
   If m_DataListCount <= 0 Or strAdd01 & strAdd02 <> "" Then
      'strTit = "資料顯示"
      'strMsg = "該筆本所案號已無資料"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Unload Me
      frm075004_1.ClearRemark
      frm075004_1.Show
      'Modified by Lydia 2020/06/11
      'If m_AddData = True Then: frm075004_1.RefreshList
      If m_AddData = True Or strAdd01 & strAdd02 <> "" Then: frm075004_1.RefreshList
   Else
      If nPos <= m_DataListCount - 1 Then
         m_CurrDL = nPos
      Else
         m_CurrDL = m_DataListCount - 1
      End If
      UpdateCtrlData
      'ShowCurrRecord strCP09
   End If
   '2006/4/24 end
   
   Exit Sub
   
EXITSUB:

'Add by Morgan 2004/2/12
ErrHnd:
   If bolTrans = True Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
    
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strCP09 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
      
   QueryRecord = False
   strCP09 = textCP09
   bFind = False
   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex).diCP09 = strCP09 Then
         bFind = True
         strCP09 = m_DataList(nIndex).diCP09
         Exit For
      End If
   Next nIndex
   If bFind = True Then
      QueryRecord = True
      ShowCurrRecord strCP09
   Else
      If IsDataBaseExist(strCP09) = True Then
         SetDataListItem strCP09
         ' 通知前畫面更新此筆資料
         'frm075004_1.ModRecord strCP09
         ShowCurrRecord strCP09
         QueryRecord = True
      Else
         QueryRecord = False
      End If
   End If
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
     
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            UpdateFieldNewData
            AddRecord
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            UpdateFieldNewData
            'Add By Cheng 2002/01/14
            '若有改變本所期限或法定期限值, 則顯示是否確定修改的詢息
            If Me.textCP06.Tag <> Me.textCP06.Text Then
               If MsgBox("您確定要更改本所期限???", vbExclamation + vbOKCancel) = vbCancel Then
                  SSTab1.Tab = 0 'Added by Lydia 2019/06/13
                  Me.textCP06.SetFocus
                  Exit Sub
               End If
            End If
            If Me.textCP07.Tag <> Me.textCP07.Text Then
               If MsgBox("您確定要更改法定期限???", vbExclamation + vbOKCancel) = vbCancel Then
                  SSTab1.Tab = 0 'Added by Lydia 2019/06/13
                  Me.textCP07.SetFocus
                  Exit Sub
               End If
            End If
            ModRecord
         Else
            GoTo EXITSUB
         End If
      Case 3:
         'Added by Lydia 2024/09/23
         If ChkExceptHandle = False Then
            SSTab1.Tab = 0
            Exit Sub
         End If
         'end 2024/09/23
         UpdateFieldNewData
         DelRecord
         ' 若已無資料在串列中, 則離開
         If m_DataListCount <= 0 Then
            GoTo EXITSUB
         End If
      Case 4:
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 檢查案件進度檔是否存在
Private Function ExistProgress(ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String

   ExistProgress = False
   strSql = "SELECT CP09 FROM CASEPROGRESS " & _
            "WHERE CP01 = '" & strCP01 & "' AND " & _
                  "CP02 = '" & strCP02 & "' AND " & _
                  "CP03 = '" & strCP03 & "' AND " & _
                  "CP04 = '" & strCP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ExistProgress = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textCP05.SetFocus
      Case 2: textCP05.SetFocus
      Case 4: textCP09.SetFocus
   End Select
End Sub

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Sindy 2009/07/08
   Dim strYear As String '抓下次繳費年度
   Dim m_Nexttimes As String '抓下次繳費次數
   
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2:
         ' 收文日不可空白
         If IsEmptyText(textCP05) = True Then
            strTit = "檢核資料"
            strMsg = "收文日不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0 'Added by Lydia 2019/06/13
            textCP05.SetFocus
            GoTo EXITSUB
         End If
         ' 案件性質不可空白
         If IsEmptyText(textCP10) = True Then
            strTit = "檢核資料"
            strMsg = "案件性質不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0 'Added by Lydia 2019/06/13
            textCP10.SetFocus
            GoTo EXITSUB
         'Add by Amy 2016/07/07 台灣不能設定電子送件--郭雅娟
         'Modified by Morgan 2018/10/8 行政訴訟503改可電子送--陳玲玲 Ex:P-96988
         ElseIf textCP01 = "P" And m_Nation = "000" And (textCP10 = "803" Or textCP10 = "804" Or textCP10 = "501") And textCP118 = "Y" Then
            strTit = "檢核資料"
            strMsg = "案件性質-" & textCP10_2 & "不能電子送件"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0 'Added by Lydia 2019/06/13
            textCP118.SetFocus
            GoTo EXITSUB
         End If
         ' 業務區
         If IsEmptyText(textCP12) = True Then
            strTit = "檢核資料"
            strMsg = "業務區不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0 'Added by Lydia 2019/06/13
            textCP12.SetFocus
            GoTo EXITSUB
         End If
         ' 智權人員
         If IsEmptyText(textCP13) = True Then
            strTit = "檢核資料"
            strMsg = "智權人員不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0 'Added by Lydia 2019/06/13
            textCP13.SetFocus
            GoTo EXITSUB
         End If
         'Modified by Morgan 2012/12/12 有變更才要檢查
         If textCP06.Tag <> Me.textCP06 Or textCP07.Tag <> Me.textCP07 Then
            ' 本所期限不可超過法定期限
            If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
               If Val(DBDATE(textCP06)) > Val(DBDATE(textCP07)) Then
                  strTit = "檢核資料"
                  strMsg = "本所期限不可超過法定期限"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  SSTab1.Tab = 0 'Added by Lydia 2019/06/13
                  textCP06.SetFocus
                  GoTo EXITSUB
               End If
            End If
         End If
         
         'Modify By Sindy 2009/07/06
         If textCP53_2.Visible = True And textCP54_2.Visible = True Then
            ' 檢查繳費年度/次數是否不正確
'            If IsEmptyText(textCP53_2) = False Then
'               m_Nexttimes = PUB_Getnexttimes(m_CP01, m_CP02, m_CP03, m_CP04, strYear)
'               If m_Nexttimes <> "" Then
'                  If m_CP10 = "605" Then '繳費年度
'                     If Val(textCP53_2) <> Val(strYear) Then
'                        MsgBox "繳費(起)年度有誤，應為" & strYear & "！"
'                        textCP53_2.SetFocus
'                        GoTo EXITSUB
'                     End If
'                  Else '繳費次數
'                     If Val(textCP53_2) <> Val(m_Nexttimes) Then
'                        MsgBox "繳費(起)次數有誤，應為" & m_Nexttimes & "！"
'                        textCP53_2.SetFocus
'                        GoTo EXITSUB
'                     End If
'                  End If
'               Else
'                  If m_CP10 = "605" Then '繳費年度
'                     MsgBox "無下次繳費年度！"
'                  Else '繳費次數
'                     MsgBox "無下次繳費次數！"
'                  End If
'                  textCP53_2.SetFocus
'                  GoTo EXITSUB
'               End If
'            End If
            If IsEmptyText(textCP53_2) = False Or IsEmptyText(textCP54_2) = False Then
               If textCP53_2 = "" Or textCP53_2 = "0" Then
                  
                  If (m_CP01 = "P" And m_CP10 = "601") Or m_CP10 = "605" Then '繳費年度
                     MsgBox "無繳費(起)年度，請清空起迄年度！"
                  Else '繳費次數
                     MsgBox "無繳費(起)次數，請清空起迄次數！"
                  End If
                  SSTab1.Tab = 2 'Added by Lydia 2019/06/13
                  If textCP53_2 = "0" Then
                     textCP53_2.SetFocus
                  Else
                     textCP54_2.SetFocus
                  End If
                  GoTo EXITSUB
               End If
               If m_CP10 = "601" And textCP54_2 = "" Then
                  '不跑else段程式
               Else
                  If textCP54_2 = "" Then textCP54_2 = "0"
                  If Val(textCP53_2) > Val(textCP54_2) Then
                     If (m_CP01 = "P" And m_CP10 = "601") Or m_CP10 = "605" Then '繳費年度
                        MsgBox "繳費(迄)年度不可小於(起)年度！"
                     Else '繳費次數
                        MsgBox "繳費(迄)次數不可小於(起)次數！"
                     End If
                     SSTab1.Tab = 2 'Added by Lydia 2019/06/13
                     textCP54_2.SetFocus
                     GoTo EXITSUB
                  End If
               End If
            End If
         '2009/07/06 End
         Else
            ' 授權期間不正確
            'Modified by Lydia 2017/08/24
            'If IsEmptyText(textCP53) = False And IsEmptyText(textCP54) = False Then
            If m_CP01 = "TB" And (m_CP10 = "708" Or m_CP10 = "1801") Then
            ElseIf IsEmptyText(textCP53) = False And IsEmptyText(textCP54) = False Then
            'end 2017/08/24
               If Val(DBDATE(textCP53)) > Val(DBDATE(textCP54)) Then
                  strTit = "檢核資料"
                  strMsg = "授權期間起日不可超過止日"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  SSTab1.Tab = 2 'Added by Lydia 2019/06/13
                  textCP53.SetFocus
                  GoTo EXITSUB
               End If
            End If
         End If
         
         'Removed by Morgan 2013/5/9 txtValidate 已有檢查
         'If textCP48.Locked = False Then 'Add by Morgan 2011/1/19
         '   'Add By Cheng 2002/05/07
         '   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
         '   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
         '      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         '         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         '         textCP48.SetFocus
         '         GoTo EXITSUB
         '      End If
         '   End If
         'End If 'Add by Morgan 2011/1/19
         
         'Add By Sindy 2024/9/18 有費用,規費不可存放空白格' ',改為0
         If Val(textCP16) > 0 And IsEmptyText(textCP17) = True Then
            textCP17 = 0
         End If
         '2024/9/18 END
         '91.12.10 add by sonia
         If Format(((Val(textCP16) - Val(textCP17)) / 1000), "0.0") <> Format(Val(textCP18), "0.0") Then
            MsgBox "(費用 - 規費) / 1000 <> 點數 !!", vbExclamation + vbOKOnly
            SSTab1.Tab = 0 'Added by Lydia 2019/06/13
            textCP16.SetFocus
            GoTo EXITSUB
         End If
         '91.12.10 end
      Case Else:
   End Select
   
   'Added by Lydia 2024/09/23
   If ChkExceptHandle = False Then
      GoTo EXITSUB
   End If
   'end 2024/09/23
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCP01_GotFocus()
   InverseTextBox textCP01
End Sub

Private Sub textCP02_GotFocus()
   InverseTextBox textCP02
End Sub

Private Sub textCP03_GotFocus()
   InverseTextBox textCP03
End Sub

Private Sub textCP04_GotFocus()
   InverseTextBox textCP04
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP08.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP09_GotFocus()
   InverseTextBox textCP09
End Sub

Private Sub textCP10_GotFocus()
   InverseTextBox textCP10
End Sub

Private Sub textCP11_GotFocus()
   InverseTextBox textCP11
End Sub

Private Sub textCP12_GotFocus()
   InverseTextBox textCP12
End Sub

Private Sub textCP13_GotFocus()
   InverseTextBox textCP13
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP15_GotFocus()
   InverseTextBox textCP15
End Sub

Private Sub textCP16_GotFocus()
   InverseTextBox textCP16
End Sub

Private Sub textCP17_GotFocus()
   InverseTextBox textCP17
End Sub

Private Sub textCP18_GotFocus()
   InverseTextBox textCP18
End Sub

Private Sub textCP19_GotFocus()
   InverseTextBox textCP19
End Sub

Private Sub textCP20_GotFocus()
   InverseTextBox textCP20
End Sub

Private Sub textCP21_GotFocus()
   InverseTextBox textCP21
End Sub

Private Sub textCP22_GotFocus()
   InverseTextBox textCP22
End Sub

Private Sub textCP23_GotFocus()
   InverseTextBox textCP23
End Sub

Private Sub textCP24_GotFocus()
   InverseTextBox textCP24
End Sub

Private Sub textCP25_GotFocus()
   InverseTextBox textCP25
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP28_GotFocus()
   InverseTextBox textCP28
End Sub

Private Sub textCP29_GotFocus()
   InverseTextBox textCP29
End Sub

Private Sub textCP30_GotFocus()
   InverseTextBox textCP30
End Sub

Private Sub textCP31_GotFocus()
   InverseTextBox textCP31
End Sub

Private Sub textCP32_GotFocus()
   InverseTextBox textCP32
End Sub

Private Sub textCP33_GotFocus()
   InverseTextBox textCP33
End Sub

Private Sub textCP34_GotFocus()
   InverseTextBox textCP34
End Sub

Private Sub textCP35_GotFocus()
   InverseTextBox textCP35
End Sub

Private Sub textCP36_GotFocus()
   InverseTextBox textCP36
End Sub

Private Sub textCP37_GotFocus()
   InverseTextBox textCP37
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP37.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP38_GotFocus()
   InverseTextBox textCP38
End Sub

Private Sub textCP39_GotFocus()
   InverseTextBox textCP39
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP39.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP40.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP42.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP43_GotFocus()
   InverseTextBox textCP43
End Sub

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP45_GotFocus()
   InverseTextBox textCP45
End Sub

Private Sub textCP46_GotFocus()
   InverseTextBox textCP46
End Sub

Private Sub textCP47_GotFocus()
   InverseTextBox textCP47
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP49_GotFocus()
   InverseTextBox textCP49
End Sub

Private Sub textCP50_GotFocus()
   InverseTextBox textCP50
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP50.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP51_GotFocus()
   InverseTextBox textCP51
End Sub

Private Sub textCP52_GotFocus()
   InverseTextBox textCP52
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP52.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP53_GotFocus()
   InverseTextBox textCP53
End Sub

Private Sub textCP54_GotFocus()
   InverseTextBox textCP54
End Sub

'Add By Sindy 2009/07/06
Private Sub textCP53_2_GotFocus()
   InverseTextBox textCP53_2
End Sub
Private Sub textCP54_2_GotFocus()
   InverseTextBox textCP54_2
End Sub
Private Sub textCP53_2_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
Private Sub textCP54_2_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
'2009/07/06 End

Private Sub textCP55_GotFocus()
   InverseTextBox textCP55
End Sub

Private Sub textCP56_GotFocus()
   InverseTextBox textCP56
End Sub

Private Sub textCP57_GotFocus()
   InverseTextBox textCP57
End Sub

Private Sub textCP58_GotFocus()
   InverseTextBox textCP58
End Sub

Private Sub textCP59_GotFocus()
   InverseTextBox textCP59
End Sub

Private Sub textCP60_GotFocus()
   InverseTextBox textCP60
End Sub

Private Sub textCP61_GotFocus()
   InverseTextBox textCP61
End Sub

Private Sub textCP62_GotFocus()
   InverseTextBox textCP62
End Sub

Private Sub textCP63_GotFocus()
   InverseTextBox textCP63
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP64.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP71_GotFocus()
   InverseTextBox textCP71
End Sub

Private Sub textCP72_GotFocus()
   InverseTextBox textCP72
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strTmp As String 'Added by Lydia 2019/06/13

   TxtValidate = False
   If Me.textCP05.Enabled = True Then
      Cancel = False
      textCP05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP06.Enabled = True Then
      Cancel = False
      textCP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Added by Lydia 2025/03/14
   If Me.textCP48.Enabled = True Then
      Cancel = False
      textCP48_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2025/03/14
   
   '2008/8/27 add by sonia 櫃台收文日
   If Me.textCP119.Enabled = True Then
      Cancel = False
      textCP119_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2008/8/27 end
   
   If Me.textCP08.Enabled = True Then
      Cancel = False
      textCP08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP10.Enabled = True Then
      Cancel = False
      textCP10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      'Add by Amy 2018/10/18/16 智權人員非國外部FXX且修改案件性質時,不可改為 902(回覆代理人)
      If (textCP01 = "P" Or textCP01 = "PS" Or textCP01 = "CFP" Or textCP01 = "CPS") And m_CP10 <> textCP10 And textCP10 = "902" Then
         If Left(PUB_GetStaffST15(textCP13, 1), 1) <> "F" Then
               Cancel = True
               MsgBox "智權人員非國外部，案件性質不可改為902(回覆代理人)"
               SSTab1.Tab = 0 'Added by Lydia 2019/06/13
               textCP10.SetFocus
               Exit Function
         End If
      End If
      'Added by Lydia 2020/04/22 櫃台一開始收錯案件性質，造成TrackingNO未能正常走新案立卷流程; ex.FCP062672, AA9005089原本收416實審=>307分割
      If (textCP01 = "FCP" Or textCP01 = "P") And m_CP10 <> textCP10 And textCP31 = "Y" And InStr("101,102,103,110,112,125,307", textCP10) > 0 Then
         If Left(PUB_GetStaffST15(textCP13, 1), 1) = "F" Then
             MsgBox "修改新案的案件性質，請一併修改命名追蹤號TracknigNo；" & vbCrLf & "可參考文件：日常工作 九十三", vbInformation
         End If
      End If
      'end 2020/04/20
   End If
   
   If Me.textCP11.Enabled = True Then
      Cancel = False
      textCP11_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP12.Enabled = True Then
      Cancel = False
      textCP12_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP13.Enabled = True Then
      Cancel = False
      textCP13_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP14.Enabled = True Then
      Cancel = False
      textCP14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP16.Enabled = True Then
      Cancel = False
      textCP16_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP17.Enabled = True Then
      Cancel = False
      textCP17_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP18.Enabled = True Then
      Cancel = False
      textCP18_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP20.Enabled = True Then
      Cancel = False
      textCP20_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP21.Enabled = True Then
      Cancel = False
      textCP21_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'edit by nickc 2006/01/27
   'If Me.textCP22.Enabled = True Then
   '   Cancel = False
   '   textCP22_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   If Me.textCP23.Enabled = True Then
      Cancel = False
      textCP23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP24.Enabled = True Then
      Cancel = False
      textCP24_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP25.Enabled = True Then
      Cancel = False
      textCP25_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP26.Enabled = True Then
      Cancel = False
      textCP26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP27.Enabled = True Then
      Cancel = False
      textCP27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP29.Enabled = True Then
      Cancel = False
      textCP29_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP31.Enabled = True Then
      Cancel = False
      textCP31_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP32.Enabled = True Then
      Cancel = False
      textCP32_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP33.Enabled = True Then
      Cancel = False
      textCP33_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP34.Enabled = True Then
      Cancel = False
      textCP34_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP37.Enabled = True Then
      Cancel = False
      textCP37_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP37_1.Enabled = True Then
      Cancel = False
      textCP37_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP38.Enabled = True Then
      Cancel = False
      textCP38_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP39.Enabled = True Then
      Cancel = False
      textCP39_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP40.Enabled = True Then
      Cancel = False
      textCP40_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP41.Enabled = True Then
      Cancel = False
      textCP41_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP42.Enabled = True Then
      Cancel = False
      textCP42_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP43.Enabled = True Then
      Cancel = False
      textCP43_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP44.Enabled = True Then
      Cancel = False
      textCP44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP46.Enabled = True Then
      Cancel = False
      textCP46_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP47.Enabled = True Then
      Cancel = False
      textCP47_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Added by Lydia 2016/05/30 法務或顧問案件的回執退件日/回執未回郵局送達日輸入檢查
   If m_SK02 = "3" Or m_SK02 = "4" Then
       If Val(textCP46) = 111111 Or Val(textCP46) = 110101 Then
           If Me.textCP47 = "" Then
              MsgBox Mid(Label30(3), 1, Len(Label30(3)) - 1) & "不可空白", vbCritical
              SSTab1.Tab = 1 'Added by Lydia 2019/06/13
              textCP47.SetFocus
              textCP47_GotFocus
              Exit Function
           End If
       Else
           If Me.textCP47 <> "" Then
              MsgBox "回執收受日為111111或110101時才可輸入", vbCritical
              SSTab1.Tab = 1 'Added by Lydia 2019/06/13
              textCP46.SetFocus
              textCP46_GotFocus
              Exit Function
           End If
       End If
   End If
   'end 2016/05/30
   
   If Me.textCP48.Enabled = True Then
      Cancel = False
      textCP48_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP50.Enabled = True Then
      Cancel = False
      textCP50_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP51.Enabled = True Then
      Cancel = False
      textCP51_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP52.Enabled = True Then
      Cancel = False
      textCP52_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP53.Visible = True Then 'Add By Sindy 2009/07/06
      If Me.textCP53.Enabled = True Then
         Cancel = False
         textCP53_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   End If
   
   If Me.textCP55.Enabled = True Then
      Cancel = False
      textCP55_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP56.Enabled = True Then
      Cancel = False
      textCP56_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP57.Enabled = True Then
      Cancel = False
      textCP57_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP58.Enabled = True Then
      Cancel = False
      textCP58_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'edit by nickc 2006/10/30
   'If Me.textCP59.Enabled = True Then
   '   Cancel = False
   '   textCP59_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP71.Enabled = True Then
      Cancel = False
      textCP71_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP72.Enabled = True Then
      Cancel = False
      textCP72_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'add by nick 2004/08/18
   If Me.textCP82.Enabled = True Then
      Cancel = False
      textCP82_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP84.Enabled = True Then
      Cancel = False
      textCP84_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Remove by Morgan 2010/12/30 目前沒用
   'If Me.textCP85.Enabled = True Then
   '   Cancel = False
   '   textCP85_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   If Me.textCP86.Enabled = True Then
      Cancel = False
      textCP86_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP89.Enabled = True Then
      Cancel = False
      textCP89_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP90.Enabled = True Then
      Cancel = False
      textCP90_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP91.Enabled = True Then
      Cancel = False
      textCP91_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP92.Enabled = True Then
      Cancel = False
      textCP92_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP93.Enabled = True Then
      Cancel = False
      textCP93_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP94.Enabled = True Then
      Cancel = False
      textCP94_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP95.Enabled = True Then
      Cancel = False
      textCP95_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP96.Enabled = True Then
      Cancel = False
      textCP96_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add by Morgan 2007/7/19
   If Me.textCP113.Enabled = True Then
      Cancel = False
      textCP113_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP114.Enabled = True Then
      Cancel = False
      textCP114_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2007/7/19
   
   '2011/5/26 add by sonia
   If Me.textCP144.Enabled = True Then
      Cancel = False
      textCP144_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2011/5/26 end
   
   'Added by Morgan 2016/6/8
   If Me.textCP152.Enabled = True Then
      Cancel = False
      textCP152_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Modified by Lydia 2018/09/13 已發文並且有規費才必須輸入扣款日期(by Phoebe)
   'If textCP118 = "A" And IsEmptyText(textCP152) Then
   If textCP118 = "A" And IsEmptyText(textCP152) And Val(textCP27) > 0 And Val(textCP84) > 0 Then
      MsgBox "電子送件(自動扣款)必須輸入扣款日期！", vbExclamation + vbOKOnly, "檢核資料"
      SSTab1.Tab = 0 'Added by Lydia 2019/06/13
      textCP152.SetFocus
      textCP152_GotFocus
      Exit Function
   ElseIf textCP118 <> "A" And Not IsEmptyText(textCP152) Then
      MsgBox "電子送件(自動扣款)才可輸入扣款日期！", vbExclamation + vbOKOnly, "檢核資料"
      SSTab1.Tab = 0 'Added by Lydia 2019/06/13
      textCP118.SetFocus
      textCP118_GotFocus
      Exit Function
   End If
   'end 2016/6/8
   
   'Added by Morgan 2024/11/18
   If textCP118 = "A" And Val(textCP27) = 0 And Val(textCP27.Tag) > 0 Then
      MsgBox "取消發文時，電子送件請還原為【Y】！" & vbCrLf & vbCrLf & "※自動扣款會於發文作業時系統自動改為【A】", vbExclamation + vbOKOnly, "檢核資料"
      SSTab1.Tab = 0
      textCP118.SetFocus
      textCP118_GotFocus
      Exit Function
   End If
   'end 2024/11/18
   
   'Added by Lydia 2019/06/13 FCP勘誤公報控管:公告公報之更改或更正的核准, 進度備註有輸入勘誤日期,一定要輸入承辦期限
   If textCP01 = "FCP" And textCP10 = "1001" And Trim(textCP27) = "" And Trim(textCP64) <> "" And Trim(textCP48) = "" Then
       intI = InStr(textCP64, "勘誤日期:")
       If intI > 0 Then
          strTmp = Mid(textCP64, intI, InStr(intI + 5, textCP64, ";"))
          If (InStr(strTmp, "更正") > 0 And InStr(strTmp, "_") = 0) Or InStr(strTmp, "更正") = 0 Then
                '勾選更正時可以不輸入日期及期別，進度備註用__帶入日期及期數
                MsgBox "公告公報之更改或更正的核准，一定要輸入承辦期限！", vbExclamation + vbOKOnly, "檢核資料"
                SSTab1.Tab = 0
                textCP48.SetFocus
                textCP48_GotFocus
                Exit Function
          End If
       End If
   End If
   
   'Added by Lydia 2019/06/13 輸入指定日期，檢查承辦期限 ; FCP-50441(AA7012152)有輸入指定日期，後來人工拿掉承辦期限。
   'Modified by Morgan 2021/6/30 +判斷 FCP Ex:P-127630 年費
   If textCP01 = "FCP" And Trim(textCP27) = "" And Trim(textCP142) <> "" Then
      If Val(textCP142) < Val(textCP48) Or Val(textCP48) = 0 Then
         strTmp = ""
         If Val(textCP48) = 0 Then
             strTmp = "輸入指定日期，則承辦期限不可空白！"
         ElseIf Option1(2).Value = False Then 'Modify By Sindy 2022/4/21 + If Option1(2).Value = False Then
             strTmp = "承辦期限不可大於指定日期！"
         End If
         If strTmp <> "" Then 'Add By Sindy 2022/4/21 + If strTmp <> "" Then
            MsgBox strTmp, vbExclamation + vbOKOnly, "檢核資料"
            SSTab1.Tab = 0
            textCP48.SetFocus
            textCP48_GotFocus
            Exit Function
         End If
      End If
   End If
   
'   'Add By Sindy 2021/4/20 檢查指定送件日相關欄位
'   If Val(textCP142.Text) > 0 And Option1(0).Visible = True Then
'      If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
'         MsgBox "有輸入指定送件日，當天或之前或之後請擇一。", vbExclamation
'         Exit Function
'      End If
'   Else
'      Option1(0).Value = False
'      Option1(1).Value = False
'      Option1(2).Value = False
'   End If
'   '2021/4/20 END
         
   'add by sonia 2021/2/20 FCP-058462
   If textCP06 = "" And textCP07 <> "" Then
      If MsgBox("有法定期限，是否要輸入本所期限？", vbYesNo + vbDefaultButton1) = vbYes Then
         SSTab1.Tab = 0
         textCP06.SetFocus
         textCP06_GotFocus
         Exit Function
      End If
   End If
   'end 2021/2/20
   
   'Added by Lydia 2021/10/08 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   
   'Add By Sindy 2022/1/14 預設和檢查-所有內部收文, 若有輸入本所期限或法定期限者 : 同收文作業
   '阿蓮:有關系統類別「S」且案件性質為「查名」之案件, 其期限為本所內部管控請款用, 故本所所與法定日期相同
   '目前期限彈跳原則為本所期限當日提醒「需請款」請控管收文, 倘輸入法定期限, 一定要輸入本所期限
   'S-007325是程序去進度檔改的。
   'Modified by Lyddia 2023/11/08 傳入必需欄位
   'If PUB_CheckCP0607(0, textCP06, textCP07) = False Then Exit Function
    If PUB_CheckCP0607(0, textCP06, textCP07, textCP31, m_Nation, textCP01, textCP10) = False Then Exit Function
    
   TxtValidate = True
End Function

'Add By Cheng 2003/03/14
Private Function CP60IsNull(strCP09 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'預設未請款
CP60IsNull = True
StrSQLa = "Select * From CaseProgress Where CP09='" & strCP09 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    '若有請款單號
    If "" & rsA("CP60").Value <> "" Then
        MsgBox "此筆資料已有請款單號，不可刪除!!!", vbExclamation + vbOKOnly
        CP60IsNull = False
    'Add By Cheng 2003/05/30
    '若有帳單編號
    '2007/3/2 modify by sonia 加入cp87,cp88
    ElseIf "" & rsA("CP61").Value <> "" Or "" & rsA("CP62").Value <> "" Or "" & rsA("CP63").Value <> "" Or "" & rsA("CP87").Value <> "" Or "" & rsA("CP88").Value <> "" Then
        MsgBox "此筆資料已有帳單編號，不可刪除!!!", vbExclamation + vbOKOnly
        CP60IsNull = False
    End If
Else
        MsgBox "查無資料!!!", vbExclamation + vbOKOnly
        CP60IsNull = False
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2003/06/10
Private Function GetMaxNumber() As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   strSql = "SELECT MAX(DD28) FROM DATADELETERECORD "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields(0)) = False Then
         GetMaxNumber = rsTmp.Fields(0)
         GetMaxNumber = Val(GetMaxNumber) + 1
      End If
   End If
   If IsEmptyText(GetMaxNumber) = True Then
      GetMaxNumber = 1
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub textCP81_GotFocus()
   TextInverse textCP81
End Sub

Private Sub textCP81_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textCP113_GotFocus()
   TextInverse textCP113
End Sub

Private Sub textCP113_Validate(Cancel As Boolean)
   If textCP113 <> "" Then
      If Not IsNumeric(textCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         textCP113.SetFocus
         textCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub textCP114_GotFocus()
   TextInverse textCP114
End Sub

Private Sub textCP114_Validate(Cancel As Boolean)
   If textCP114 <> "" Then
      If Not IsNumeric(textCP114) Then
         MsgBox "請輸入數字！", vbExclamation
         textCP114.SetFocus
         textCP114_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

'Added by Morgan 2013/5/9
'讀取原始資料
Private Function GetOldData(pFiledName As String) As String
   Dim ii As Integer
   For ii = LBound(m_FieldList) To UBound(m_FieldList)
      If m_FieldList(ii).fiName = pFiledName Then
         GetOldData = m_FieldList(ii).fiOldData
         Exit For
      End If
   Next
End Function

'Added by Lydia 2018/04/11 外專翻譯承辦單列印
Private Sub cmdPrint201_Click()
Dim strYN As String 'Added by Lydia 2020/07/27
Dim strColName() As String, strColText() As String 'Add By Sindy 2023/9/14
   
'   'Add By Sindy 2023/10/3
'   If strSrvDate(1) >= 外專承辦歷程啟用日 Then
'      If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub
'      frm090202_2.Hide
'      frm090202_2.m_EEP01 = textCP09 '總收文號
'      frm090202_2.m_FlowUserNum = strUserNum '案件流程所屬人員
'      frm090202_2.intReceiveKind = 4 '4.送中說
'      frm090202_2.SetParent Me
'      If frm090202_2.QueryData = True Then
'         frm090202_2.Show
'         Me.Hide
'      End If
'   Else
'   '2023/10/3 END
      'Added by Lydia 2020/07/27 配合異地上班，檢視中說209和核對中說格式210也需要跑中說流程，需要產生電子檔
      If MsgBox("是否產生Word檔？" & vbCrLf & "選「是」產生Word檔，選「否」直接列印。", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
         strYN = "Y"
      End If
      
      'Added by Lydia 2019/05/02 會稿Claims/說明書承辦單列印
       If Me.textCP10.Text = "924" Then
           'Modified by Lydia 2020/07/27 +判斷產生Word檔 IIf(strYN = "Y", True, False)
           'Modify By Sindy 2023/9/14 +, strColName, strColText
           Call Pub_PrintFCP924Form(textCP01, textCP02, textCP03, textCP04, textCP09, strColName, strColText, IIf(strYN = "Y", True, False))
       Else
       'end 2019/05/02
           'Modified by Lydia 2020/07/27 +判斷產生Word檔 IIf(strYN = "Y", True, False)
           'Modify By Sindy 2023/9/14 +, strColName, strColText
           Call Pub_PrintFCP201Form(textCP01, textCP02, textCP03, textCP04, textCP09, strColName, strColText, IIf(strYN = "Y", True, False))
       End If
'   End If
End Sub

'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Added by Lydia 2020/04/10 法務工作點數分配
Private Sub CmdDot_Click()
   If m_EditMode <> 0 Then
       MsgBox "請先完成修改！", vbExclamation, "訊息"
       Exit Sub
   End If
   If textCP09.Text = "" Then
       MsgBox "請先完成查詢！", vbExclamation, "訊息"
       Exit Sub
   End If
   If InStr(m_CP01, "L") > 0 And Val(textCP18.Text) > 0 Then
        'Added by Lydia 2021/08/18
        If PUB_CheckFormExist("frm071021") Then
             MsgBox "請先關閉【法務工作點數分配】畫面！", vbExclamation
             Exit Sub
        End If
        'end 2021/08/18
        Set frm071021.m_PrevForm = Me
        frm071021.m_bolPrev = True
        frm071021.m_KeyList = textCP09.Text
        Me.Hide
        frm071021.Show
   Else
        MsgBox "無點數可供分配!", vbExclamation
        SSTab1.Tab = 1
   End If
End Sub

'Added by Lydia 2021/01/14 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim stSQL As String, intQ As Integer
Dim RsQ As ADODB.Recordset
   
   m_LOS01 = "": m_LOS01cp01 = "": m_LOS01cp02 = "": m_LOS01cp03 = "": m_LOS01cp04 = ""
   m_LOS02 = ""
   
   'Modified by Lydia 2023/08/14 改用(案件進度)案源單號
   'stSQL = "select nvl(X.LOS07,0) ord1,X.LOS01,X.LOS02,X.LOS04,X.LOS06,X.LOS10,X.LOS15,CP01,CP02,CP03,CP04 " & _
                "from LawOfficeSource X,CaseProgress where X.LOS06='" & textCP09 & "' and X.LOS01=CP09(+) order by ord1, X.LOS01 "
   stSQL = "select nvl(X.LOS07,0) ord1,X.LOS01,X.LOS02,X.LOS04,X.LOS06,X.LOS10,X.LOS15,CP01,CP02,CP03,CP04 " & _
                "from LawOfficeSource X,CaseProgress where X.LOS15='" & m_CP162 & "' and X.LOS01=CP09(+) order by ord1, X.LOS01 "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      '案源總收文號
      m_LOS01 = "" & RsQ.Fields("los01")
      '案源總收文號之本所案號
      m_LOS01cp01 = "" & RsQ.Fields("cp01")
      m_LOS01cp02 = "" & RsQ.Fields("cp02")
      m_LOS01cp03 = "" & RsQ.Fields("cp03")
      m_LOS01cp04 = "" & RsQ.Fields("cp04")
      '(原)案源案件類型
      m_LOS02 = "" & RsQ.Fields("LOS02")
   End If
   Set RsQ = Nothing
End Sub

'Add by Amy 2022/09/02 置換 Label文字,「其他相關人」顯示「對方」
Private Sub SetLabTxt(stRPTxt As String)
    Dim oLab
    
    If Left(Label18(4).Caption, 2) = stRPTxt Then Exit Sub
    
    For Each oLab In Label18
        If oLab.Index >= 1 And oLab.Index <= 5 Then
            oLab.Caption = stRPTxt & Mid(oLab.Caption, 3)
        End If
    Next
    
    For Each oLab In Label19
        If oLab.Index >= 1 And oLab.Index <= 3 Then
            oLab.Caption = stRPTxt & Mid(oLab.Caption, 3)
        End If
    Next
    Label33(1).Caption = stRPTxt & Mid(Label33(1).Caption, 3)
End Sub

'Added by Lydia 2024/09/23
Private Function ChkExceptHandle() As Boolean
   '檢查收文號; ex.FCP-52704的CB3055229，收文號被操作者改成.
   If Len(textCP09) <> 9 Then
      MsgBox "收文號長度為9碼！", vbCritical + vbOKOnly, "資料稽核"
      GoTo EXITSUB
   Else
      If m_EditMode = 2 Or m_EditMode = 3 Then '修改/刪除
         If m_DataList(m_CurrDL).diCP09 <> "" And m_DataList(m_CurrDL).diCP09 <> textCP09 Then
            MsgBox "收文號不正確！" & vbCrLf & "正確收文號：" & m_DataList(m_CurrDL).diCP09, vbCritical + vbOKOnly, "資料稽核"
            GoTo EXITSUB
         End If
      End If
   End If
   ChkExceptHandle = True
   Exit Function
   
EXITSUB:
   ChkExceptHandle = False
End Function

