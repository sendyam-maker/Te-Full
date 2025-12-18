VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_2 
   Caption         =   "工作進度資料維護"
   ClientHeight    =   6480
   ClientLeft      =   3400
   ClientTop       =   2950
   ClientWidth     =   11940
   ControlBox      =   0   'False
   FillColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   11940
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   610
      Left            =   9456
      ScaleHeight     =   610
      ScaleWidth      =   2420
      TabIndex        =   156
      Top             =   48
      Visible         =   0   'False
      Width           =   2424
      Begin VB.Label lblCal1 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BorderStyle     =   1  '單線固定
         Caption         =   "0.00 點"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   6
         Left            =   1488
         TabIndex        =   162
         Top             =   384
         Width           =   924
      End
      Begin VB.Label Label20 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         BorderStyle     =   1  '單線固定
         Caption         =   "累計發文收文點數"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   0
         TabIndex        =   161
         Top             =   384
         Width           =   1500
      End
      Begin VB.Label lblCal1 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BorderStyle     =   1  '單線固定
         Caption         =   "0.00 點"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   5
         Left            =   1488
         TabIndex        =   160
         Top             =   192
         Width           =   924
      End
      Begin VB.Label Label19 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         BorderStyle     =   1  '單線固定
         Caption         =   "累計會稿收文點數"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   0
         TabIndex        =   159
         Top             =   192
         Width           =   1500
      End
      Begin VB.Label lblCal1 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BorderStyle     =   1  '單線固定
         Caption         =   "0.00 點"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   4
         Left            =   1488
         TabIndex        =   158
         Top             =   0
         Width           =   924
      End
      Begin VB.Label Label9 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         BorderStyle     =   1  '單線固定
         Caption         =   "累計完稿收文點數"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   0
         TabIndex        =   157
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "達成情況(&Word)"
      Height          =   252
      Index           =   3
      Left            =   96
      TabIndex        =   153
      Top             =   240
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   8190
      TabIndex        =   29
      Top             =   50
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "本月統計(&T)"
      Height          =   400
      Index           =   0
      Left            =   6420
      TabIndex        =   27
      Top             =   50
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Height          =   400
      Index           =   1
      Left            =   7560
      TabIndex        =   28
      Top             =   50
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5955
      Left            =   30
      TabIndex        =   34
      Top             =   510
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   10495
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "瀏覽資料"
      TabPicture(0)   =   "frm090201_2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Combo3"
      Tab(0).Control(1)=   "grd1"
      Tab(0).Control(2)=   "cmdok2(1)"
      Tab(0).Control(3)=   "cmdok2(0)"
      Tab(0).Control(4)=   "Combo1"
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).Control(7)=   "Label3"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "詳細資料"
      TabPicture(1)   =   "frm090201_2.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(33)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(31)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(29)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(25)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(5)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(27)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(26)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(24)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1(23)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(4)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lbl1(23)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lbl1(19)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lbl1(17)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lbl1(15)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "lbl1(13)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "lbl1(11)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "lbl1(9)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "lbl1(7)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lbl1(5)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "lbl1(3)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label1(32)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "lbl1(30)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "lbl1(28)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "lbl1(26)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "lbl1(18)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lbl1(16)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "lbl1(14)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "lbl1(10)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "lbl1(8)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "lbl1(6)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "lbl1(4)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Label1(30)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "lbl1(29)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Label1(11)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Label1(13)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Label1(14)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Label1(15)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Label1(16)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Label1(17)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Label1(18)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Label1(19)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Label1(20)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Label1(21)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Label1(8)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "lbl1(0)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "lbl1(1)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Label1(1)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "lbl1(21)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Label1(3)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Label1(22)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Label1(12)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "lbl1(31)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "Label1(28)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "lblClose"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Label1(35)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "Label1(36)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Label6"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "Label7"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "Label8"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "lbl1(32)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "Label1(37)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "Label1(38)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "Label1(39)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "Label1(40)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "lbl1(34)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "Label1(41)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "Label1(43)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "Label1(42)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "Label1(46)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "Label1(44)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "lbl1(2)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "Label15"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "Label1(45)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "lblEApp"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "Label1(47)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "Label18"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "lblCM10"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "Label1(50)"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "Label1(49)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "lblCMboth"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "txtCP64"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "txtEP12"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "txtCP144"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "txtCP99"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "Combo2"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "Combo4"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "Combo6"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).Control(91)=   "Label1(9)"
      Tab(1).Control(91).Enabled=   0   'False
      Tab(1).Control(92)=   "lbl1(12)"
      Tab(1).Control(92).Enabled=   0   'False
      Tab(1).Control(93)=   "txt1(8)"
      Tab(1).Control(93).Enabled=   0   'False
      Tab(1).Control(94)=   "txt1(7)"
      Tab(1).Control(94).Enabled=   0   'False
      Tab(1).Control(95)=   "txt1(6)"
      Tab(1).Control(95).Enabled=   0   'False
      Tab(1).Control(96)=   "txt1(5)"
      Tab(1).Control(96).Enabled=   0   'False
      Tab(1).Control(97)=   "txt1(4)"
      Tab(1).Control(97).Enabled=   0   'False
      Tab(1).Control(98)=   "txt1(3)"
      Tab(1).Control(98).Enabled=   0   'False
      Tab(1).Control(99)=   "txt1(1)"
      Tab(1).Control(99).Enabled=   0   'False
      Tab(1).Control(100)=   "txt1(0)"
      Tab(1).Control(100).Enabled=   0   'False
      Tab(1).Control(101)=   "txt1(9)"
      Tab(1).Control(101).Enabled=   0   'False
      Tab(1).Control(102)=   "cmd1"
      Tab(1).Control(102).Enabled=   0   'False
      Tab(1).Control(103)=   "txt1(12)"
      Tab(1).Control(103).Enabled=   0   'False
      Tab(1).Control(104)=   "txt1(14)"
      Tab(1).Control(104).Enabled=   0   'False
      Tab(1).Control(105)=   "txt1(13)"
      Tab(1).Control(105).Enabled=   0   'False
      Tab(1).Control(106)=   "txt1(2)"
      Tab(1).Control(106).Enabled=   0   'False
      Tab(1).Control(107)=   "cmd(2)"
      Tab(1).Control(107).Enabled=   0   'False
      Tab(1).Control(108)=   "txt1(15)"
      Tab(1).Control(108).Enabled=   0   'False
      Tab(1).Control(109)=   "txt1(17)"
      Tab(1).Control(109).Enabled=   0   'False
      Tab(1).Control(110)=   "txt1(18)"
      Tab(1).Control(110).Enabled=   0   'False
      Tab(1).Control(111)=   "txt1(19)"
      Tab(1).Control(111).Enabled=   0   'False
      Tab(1).Control(112)=   "txt1(20)"
      Tab(1).Control(112).Enabled=   0   'False
      Tab(1).Control(113)=   "chk1"
      Tab(1).Control(113).Enabled=   0   'False
      Tab(1).Control(114)=   "cmd(3)"
      Tab(1).Control(114).Enabled=   0   'False
      Tab(1).Control(115)=   "cmdPic"
      Tab(1).Control(115).Enabled=   0   'False
      Tab(1).Control(116)=   "cmd(1)"
      Tab(1).Control(116).Enabled=   0   'False
      Tab(1).Control(117)=   "txt1(21)"
      Tab(1).Control(117).Enabled=   0   'False
      Tab(1).Control(118)=   "txt1(23)"
      Tab(1).Control(118).Enabled=   0   'False
      Tab(1).Control(119)=   "txt1(11)"
      Tab(1).Control(119).Enabled=   0   'False
      Tab(1).Control(120)=   "cmdFAmend"
      Tab(1).Control(120).Enabled=   0   'False
      Tab(1).Control(121)=   "cmd(0)"
      Tab(1).Control(121).Enabled=   0   'False
      Tab(1).Control(122)=   "cmd(4)"
      Tab(1).Control(122).Enabled=   0   'False
      Tab(1).Control(123)=   "cmd(5)"
      Tab(1).Control(123).Enabled=   0   'False
      Tab(1).ControlCount=   124
      TabCaption(2)   =   "待辦歷程"
      TabPicture(2)   =   "frm090201_2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Combo5"
      Tab(2).Control(1)=   "cmdDetail"
      Tab(2).Control(2)=   "cmdQuery"
      Tab(2).Control(3)=   "grd2"
      Tab(2).Control(4)=   "Label17"
      Tab(2).Control(5)=   "Label1(48)"
      Tab(2).Control(6)=   "Label16"
      Tab(2).ControlCount=   7
      Begin VB.CommandButton cmd 
         Caption         =   "IDS清單"
         Height          =   285
         Index           =   5
         Left            =   8910
         TabIndex        =   149
         Top             =   1620
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "未完稿暫存區"
         Height          =   285
         Index           =   4
         Left            =   6120
         Style           =   1  '圖片外觀
         TabIndex        =   148
         Top             =   330
         Width           =   1260
      End
      Begin VB.CommandButton cmd 
         Caption         =   "申請書(&A)"
         Height          =   285
         Index           =   0
         Left            =   6120
         TabIndex        =   146
         Top             =   660
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdFAmend 
         Caption         =   "免費修正事由"
         Height          =   285
         Left            =   8400
         TabIndex        =   16
         Top             =   3420
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   5850
         MaxLength       =   1
         TabIndex        =   20
         Top             =   4530
         Width           =   360
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   23
         Left            =   8520
         MaxLength       =   1
         TabIndex        =   12
         Top             =   2880
         Width           =   270
      End
      Begin VB.ComboBox Combo5 
         Height          =   276
         ItemData        =   "frm090201_2.frx":0054
         Left            =   -69120
         List            =   "frm090201_2.frx":0064
         Style           =   2  '單純下拉式
         TabIndex        =   137
         Top             =   390
         Width           =   960
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "明細資料(&D)"
         Height          =   360
         Left            =   -68100
         TabIndex        =   135
         Top             =   360
         Width           =   1125
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "畫面更新(&Q)"
         Height          =   360
         Left            =   -66930
         TabIndex        =   133
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   21
         Left            =   8100
         MaxLength       =   7
         TabIndex        =   22
         Top             =   5055
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmd 
         Caption         =   "承辦歷程(&E)"
         Height          =   285
         Index           =   1
         Left            =   7470
         TabIndex        =   123
         Top             =   660
         Width           =   1320
      End
      Begin VB.CommandButton cmdPic 
         BackColor       =   &H00C0C0C0&
         Caption         =   "代表圖(&I)"
         Height          =   285
         Left            =   7680
         Style           =   1  '圖片外觀
         TabIndex        =   122
         Top             =   4530
         Width           =   1455
      End
      Begin VB.CommandButton cmd 
         Caption         =   "撰寫信函(&L)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   7500
         TabIndex        =   121
         Top             =   1620
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CheckBox chk1 
         Caption         =   "無圖式"
         Height          =   255
         Left            =   7665
         TabIndex        =   24
         Top             =   4800
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   20
         Left            =   7305
         MaxLength       =   1
         TabIndex        =   10
         Top             =   2570
         Width           =   270
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   19
         Left            =   6810
         MaxLength       =   7
         TabIndex        =   11
         Top             =   2880
         Width           =   900
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   18
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   6
         Top             =   2280
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   17
         Left            =   8220
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1027
         Width           =   360
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   15
         Left            =   3165
         MaxLength       =   3
         TabIndex        =   13
         Top             =   3750
         Width           =   525
      End
      Begin VB.CommandButton cmd 
         Caption         =   "承辦單列印(&P)"
         Height          =   285
         Index           =   2
         Left            =   9510
         TabIndex        =   101
         Top             =   1980
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1350
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   13
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   31
         Top             =   420
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   14
         Left            =   1110
         MaxLength       =   1
         TabIndex        =   30
         Top             =   420
         Width           =   480
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         ItemData        =   "frm090201_2.frx":0083
         Left            =   -70260
         List            =   "frm090201_2.frx":0099
         TabIndex        =   98
         Top             =   390
         Width           =   2430
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   5100
         Left            =   -74940
         TabIndex        =   35
         Top             =   750
         Width           =   9825
         _ExtentX        =   17339
         _ExtentY        =   8996
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   2
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
         _Band(0).Cols   =   1
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   12
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   0
         Top             =   705
         Width           =   915
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "專利相關案件"
         Height          =   345
         Left            =   60
         TabIndex        =   95
         Top             =   4770
         Width           =   1680
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   9
         Left            =   5070
         MaxLength       =   1
         TabIndex        =   21
         Top             =   4800
         Width           =   360
      End
      Begin VB.CommandButton cmdok2 
         Caption         =   "未發文"
         Height          =   400
         Index           =   1
         Left            =   -66648
         TabIndex        =   33
         Top             =   348
         Width           =   852
      End
      Begin VB.CommandButton cmdok2 
         Caption         =   "當月資料"
         Height          =   400
         Index           =   0
         Left            =   -67656
         TabIndex        =   32
         Top             =   348
         Width           =   972
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   0
         Left            =   5010
         MaxLength       =   6
         TabIndex        =   26
         Top             =   1020
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   5010
         MaxLength       =   1
         TabIndex        =   5
         Top             =   1965
         Width           =   480
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1650
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   4
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   7
         Top             =   2595
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   5
         Left            =   7305
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1965
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   6
         Left            =   7305
         MaxLength       =   6
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   7
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   17
         Top             =   3930
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   8
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   19
         Top             =   4230
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Height          =   5055
         Left            =   -74940
         TabIndex        =   132
         Top             =   750
         Width           =   9825
         _ExtentX        =   17339
         _ExtentY        =   8908
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|目次|流程日期|本所案號|案件名稱|國家|種類|案件性質|本所期限|承辦人|承辦期限|智權人員|目前流程狀態|不顯示"
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
         _Band(0).Cols   =   14
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   12
         Left            =   3060
         TabIndex        =   155
         Top             =   3252
         Width           =   840
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1482;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "指定日期："
         Height          =   180
         Index           =   9
         Left            =   2160
         TabIndex        =   154
         Top             =   3228
         Width           =   900
      End
      Begin MSForms.ComboBox Combo6 
         Height          =   315
         Left            =   7350
         TabIndex        =   18
         Top             =   3960
         Width           =   1890
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3334;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo4 
         Height          =   315
         Left            =   7305
         TabIndex        =   9
         Top             =   2280
         Width           =   1890
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3334;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   315
         Left            =   5010
         TabIndex        =   1
         Top             =   1020
         Width           =   1890
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3334;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCP99 
         Height          =   840
         Left            =   1830
         TabIndex        =   14
         Top             =   4320
         Width           =   1890
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "3334;1482"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCP144 
         Height          =   705
         Left            =   960
         TabIndex        =   152
         Top             =   5190
         Width           =   2895
         VariousPropertyBits=   -1466941413
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "5106;1244"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEP12 
         Height          =   510
         Left            =   5010
         TabIndex        =   15
         Top             =   3420
         Width           =   3345
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "5900;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCP64 
         Height          =   510
         Left            =   5010
         TabIndex        =   23
         Top             =   5370
         Width           =   4170
         VariousPropertyBits=   -1466941409
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "7355;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   315
         Left            =   -74100
         TabIndex        =   150
         Top             =   360
         Width           =   2430
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "4286;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
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
         Left            =   2970
         TabIndex        =   145
         Top             =   750
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "報價備註："
         Height          =   180
         Index           =   49
         Left            =   60
         TabIndex        =   144
         Top             =   5190
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "核稿語文：   (1.英2.日)"
         Height          =   180
         Index           =   50
         Left            =   7800
         TabIndex        =   142
         Top             =   2940
         Width           =   1785
      End
      Begin VB.Label lblCM10 
         Caption         =   "一案兩請"
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
         Left            =   2970
         TabIndex        =   141
         Top             =   1320
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label18 
         Caption         =   "可不跑承辦歷程"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   7860
         TabIndex        =   140
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "已確認過會完日，在會完流程狀態前加註Y／N。"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   -74910
         TabIndex        =   139
         Top             =   540
         Width           =   3915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "最近聯絡："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   48
         Left            =   -70020
         TabIndex        =   138
         Top             =   450
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "判發人："
         Height          =   180
         Index           =   47
         Left            =   6570
         TabIndex        =   136
         Top             =   4035
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "註：雙擊選取時，開啟承辦歷程"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -74910
         TabIndex        =   134
         Top             =   330
         Width           =   2895
      End
      Begin VB.Label lblEApp 
         Caption         =   "電子送件"
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
         Left            =   2970
         TabIndex        =   131
         Top             =   1035
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "他所寄存：        (Y：是)"
         Height          =   180
         Index           =   45
         Left            =   7200
         TabIndex        =   130
         Top             =   5100
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "是否為複雜或特殊案件：         (Y:是)"
         Height          =   180
         Left            =   3885
         TabIndex        =   129
         Top             =   4560
         Width           =   2850
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   2
         Left            =   7410
         TabIndex        =   128
         Top             =   3180
         Width           =   585
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1032;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "修改及衍生時數："
         Height          =   180
         Index           =   44
         Left            =   5865
         TabIndex        =   127
         Top             =   3180
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "進度備註："
         Height          =   180
         Index           =   46
         Left            =   4095
         TabIndex        =   120
         Top             =   5400
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "是否暫停核稿："
         Height          =   180
         Index           =   42
         Left            =   6045
         TabIndex        =   112
         Top             =   2655
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y：暫停)"
         Height          =   180
         Index           =   43
         Left            =   7605
         TabIndex        =   113
         Top             =   2655
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "英文核完日："
         Height          =   180
         Index           =   41
         Left            =   5880
         TabIndex        =   111
         Top             =   2940
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   34
         Left            =   3375
         TabIndex        =   110
         Top             =   2895
         Width           =   330
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "582;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "會稿加乘："
         Height          =   180
         Index           =   40
         Left            =   2385
         TabIndex        =   109
         Top             =   2895
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "預定會稿日："
         Height          =   180
         Index           =   39
         Left            =   3915
         TabIndex        =   108
         Top             =   2340
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y：是)"
         Height          =   180
         Index           =   38
         Left            =   8625
         TabIndex        =   107
         Top             =   1087
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否提供圖檔："
         Height          =   180
         Index           =   37
         Left            =   6960
         TabIndex        =   106
         Top             =   1087
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   32
         Left            =   3375
         TabIndex        =   105
         Top             =   2610
         Width           =   330
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "582;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "加乘註記/基數修改理由："
         Height          =   180
         Left            =   1776
         TabIndex        =   104
         Top             =   4080
         Width           =   2028
      End
      Begin VB.Label Label7 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "加乘註記："
         Height          =   180
         Left            =   2160
         TabIndex        =   103
         Top             =   3810
         Width           =   900
      End
      Begin VB.Label Label6 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "基數："
         Height          =   180
         Left            =   2745
         TabIndex        =   102
         Top             =   2610
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "收卷註記：               (Y：收到卷宗)"
         Height          =   180
         Index           =   36
         Left            =   75
         TabIndex        =   100
         Top             =   420
         Width           =   2760
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦人輸入本所期限："
         Height          =   180
         Index           =   35
         Left            =   3180
         TabIndex        =   99
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "顏色說明："
         Height          =   225
         Left            =   -71160
         TabIndex        =   97
         Top             =   432
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "承辦人： "
         Height          =   180
         Index           =   0
         Left            =   -74904
         TabIndex        =   96
         Top             =   432
         Width           =   792
      End
      Begin VB.Label lblClose 
         Caption         =   "lblClose"
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
         Left            =   2970
         TabIndex        =   94
         Top             =   1620
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "支援時數："
         Height          =   180
         Index           =   28
         Left            =   4035
         TabIndex        =   93
         Top             =   3180
         Width           =   960
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   31
         Left            =   5010
         TabIndex        =   92
         Top             =   3180
         Width           =   585
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1032;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   255
         Index           =   12
         Left            =   75
         TabIndex        =   91
         Top             =   3848
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "(N:  不通知)"
         Height          =   210
         Index           =   22
         Left            =   5475
         TabIndex        =   90
         ToolTipText     =   "(N:  不通知, 自動內部收文)"
         Top             =   4830
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "是否通知客戶："
         Height          =   180
         Index           =   3
         Left            =   3675
         TabIndex        =   89
         Top             =   4830
         Width           =   1350
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   990
         TabIndex        =   88
         Top             =   3840
         Width           =   795
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1402;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "目次："
         Height          =   255
         Index           =   1
         Left            =   75
         TabIndex        =   87
         Top             =   718
         Width           =   540
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   2010
         TabIndex        =   85
         Top             =   718
         Width           =   810
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1429;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   645
         TabIndex        =   84
         Top             =   718
         Width           =   630
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1111;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦人："
         Height          =   255
         Index           =   8
         Left            =   1275
         TabIndex        =   83
         Top             =   718
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "總收文號："
         Height          =   255
         Index           =   21
         Left            =   75
         TabIndex        =   82
         Top             =   1031
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "收文日："
         Height          =   255
         Index           =   20
         Left            =   75
         TabIndex        =   81
         Top             =   1344
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   19
         Left            =   75
         TabIndex        =   80
         Top             =   1657
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   255
         Index           =   18
         Left            =   75
         TabIndex        =   79
         Top             =   1970
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "是否算案件數："
         Height          =   255
         Index           =   17
         Left            =   75
         TabIndex        =   78
         Top             =   2283
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "專利/商標種類："
         Height          =   255
         Index           =   16
         Left            =   75
         TabIndex        =   77
         Top             =   2596
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   255
         Index           =   15
         Left            =   75
         TabIndex        =   76
         Top             =   2909
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限："
         Height          =   255
         Index           =   14
         Left            =   75
         TabIndex        =   75
         Top             =   3222
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         Height          =   255
         Index           =   13
         Left            =   75
         TabIndex        =   74
         Top             =   3535
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "點數："
         Height          =   255
         Index           =   11
         Left            =   75
         TabIndex        =   73
         Top             =   4170
         Width           =   540
      End
      Begin MSForms.Label lbl1 
         Height          =   495
         Index           =   29
         Left            =   6615
         TabIndex        =   71
         Top             =   7587
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2408;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   225
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y/N)"
         Height          =   180
         Index           =   30
         Left            =   5655
         TabIndex        =   70
         Top             =   2010
         Width           =   405
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   5550
         TabIndex        =   69
         Top             =   1020
         Visible         =   0   'False
         Width           =   1110
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1958;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   270
         Index           =   6
         Left            =   5040
         TabIndex        =   68
         Top             =   1970
         Visible         =   0   'False
         Width           =   440
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "776;476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   225
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   240
         Index           =   8
         Left            =   5040
         TabIndex        =   67
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1508;423"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   180
         Index           =   10
         Left            =   5040
         TabIndex        =   66
         Top             =   1680
         Visible         =   0   'False
         Width           =   1005
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1773;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   8310
         TabIndex        =   65
         Top             =   2010
         Width           =   885
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1561;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   240
         Index           =   16
         Left            =   8070
         TabIndex        =   64
         Top             =   2280
         Width           =   1050
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1852;423"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   18
         Left            =   5280
         TabIndex        =   63
         Top             =   3540
         Visible         =   0   'False
         Width           =   1440
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2540;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   26
         Left            =   5010
         TabIndex        =   62
         Top             =   2910
         Width           =   585
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
         Index           =   28
         Left            =   5010
         TabIndex        =   61
         Top             =   5100
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   495
         Index           =   30
         Left            =   6270
         TabIndex        =   60
         Top             =   7437
         Visible         =   0   'False
         Width           =   1600
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2822;873"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   225
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "(N：不算)"
         Height          =   180
         Index           =   32
         Left            =   2200
         TabIndex        =   59
         Top             =   2283
         Width           =   1068
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   1020
         TabIndex        =   58
         Top             =   1031
         Width           =   1170
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2064;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   1020
         TabIndex        =   57
         Top             =   1344
         Width           =   1170
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2064;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   1020
         TabIndex        =   56
         Top             =   1657
         Width           =   1740
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3069;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   1020
         TabIndex        =   55
         Top             =   1970
         Width           =   2925
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5159;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   1470
         TabIndex        =   54
         Top             =   2283
         Width           =   600
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1058;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   1440
         TabIndex        =   53
         Top             =   2610
         Width           =   1410
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2487;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   990
         TabIndex        =   52
         Top             =   2895
         Width           =   1200
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   17
         Left            =   996
         TabIndex        =   51
         Top             =   3252
         Width           =   840
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1482;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   19
         Left            =   990
         TabIndex        =   50
         Top             =   3600
         Width           =   1170
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2064;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   720
         TabIndex        =   49
         Top             =   4170
         Width           =   915
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1614;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "繪圖人員："
         Height          =   180
         Index           =   4
         Left            =   4035
         TabIndex        =   48
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "齊備日："
         Height          =   180
         Index           =   23
         Left            =   4260
         TabIndex        =   47
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "會稿日："
         Height          =   180
         Index           =   24
         Left            =   4260
         TabIndex        =   46
         Top             =   2655
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "完稿日："
         Height          =   180
         Index           =   26
         Left            =   4260
         TabIndex        =   45
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "是否會稿："
         Height          =   180
         Index           =   27
         Left            =   4035
         TabIndex        =   44
         Top             =   2010
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "會稿完成日："
         Height          =   180
         Index           =   2
         Left            =   3885
         TabIndex        =   43
         Top             =   4005
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "核稿人："
         Height          =   180
         Index           =   5
         Left            =   6570
         TabIndex        =   42
         Top             =   2010
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文日："
         Height          =   180
         Index           =   6
         Left            =   4260
         TabIndex        =   41
         Top             =   4275
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "英文核稿人："
         Height          =   180
         Index           =   25
         Left            =   6225
         TabIndex        =   40
         Top             =   2340
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "取消收文日："
         Height          =   180
         Index           =   29
         Left            =   3615
         TabIndex        =   39
         Top             =   5100
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "承辦備註："
         Height          =   180
         Index           =   31
         Left            =   4095
         TabIndex        =   38
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦天數："
         Height          =   180
         Index           =   33
         Left            =   4035
         TabIndex        =   37
         Top             =   2910
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "請點選""確定""按鈕存檔!!"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5970
         TabIndex        =   36
         Top             =   30
         Width           =   3225
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦期限："
         Height          =   180
         Index           =   7
         Left            =   4035
         TabIndex        =   86
         Top             =   765
         Width           =   960
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Left            =   -74010
         TabIndex        =   151
         Top             =   390
         Width           =   2445
         VariousPropertyBits=   27
         Size            =   "4313;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10272
      TabIndex        =   143
      Text            =   "Text1"
      Top             =   1416
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ListBox lstNameAgent 
      Height          =   220
      Left            =   10224
      TabIndex        =   147
      Top             =   1104
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblCal1 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   3060
      TabIndex        =   126
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label14 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   1  '單線固定
      Caption         =   "累計會稿量"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1950
      TabIndex        =   125
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "此四項數據僅算到昨日"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   96
      TabIndex        =   124
      Top             =   36
      Width           =   1800
   End
   Begin VB.Label lblCal1 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   5280
      TabIndex        =   118
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label lblCal1 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   3060
      TabIndex        =   117
      Top             =   30
      Width           =   1110
   End
   Begin VB.Label Label12 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   1  '單線固定
      Caption         =   "累計達成比例"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4170
      TabIndex        =   116
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label11 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   1  '單線固定
      Caption         =   "目前進度"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4170
      TabIndex        =   115
      Top             =   30
      Width           =   1110
   End
   Begin VB.Label Label10 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      BorderStyle     =   1  '單線固定
      Caption         =   "累計完稿量"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1950
      TabIndex        =   114
      Top             =   30
      Width           =   1110
   End
   Begin VB.Label Label4 
      Caption         =   "申請國家："
      Enabled         =   0   'False
      Height          =   180
      Left            =   2715
      TabIndex        =   72
      Top             =   495
      Width           =   900
   End
   Begin VB.Label lblCal1 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   5280
      TabIndex        =   119
      Top             =   30
      Width           =   1110
   End
End
Attribute VB_Name = "frm090201_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/28 改成Form2.0 ; grd1改字型=新細明體-ExtB、grd2改字型=新細明體-ExtB、Combo1、Combo2、Combo4、Combo6、lbl1(index)、txt1(10)改為txtEP12、txt1(22)改為txtCP144、txt1(16)改為txtCP99、txtCP64
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Public TextOk As Boolean
'92.6.26 ADD BY SONIA
Public Combo1_String As String
'92.6.26 END
Public strEP07Tag As String 'Add By Sindy 2022/10/31
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String, SWPColor As String, SWPColor2 As String, SWPRow As String, SWPRow2 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 26) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String, Tmp001 As String, Tmp002 As String, Tmp003 As String, Tmp004 As String, k As Integer
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, SeekTmpBk As String, ChkData As Boolean
Dim strCP10 As String, AdoRs As ADODB.Recordset, StrNewSQL As String, Txt090201 As TextBox, ChkNoData As Boolean
Dim Fobj As FileSystemObject, ChkCp27 As Boolean, StrGrp090201 As String
Dim ADORECORDSET66 As New ADODB.Recordset
Dim Adorecordset99 As New ADODB.Recordset
Dim Intnick910123 As Integer
'Add By Cheng 2003/05/09
Dim m_SqlGrpStr1 As String, m_SqlGrpStr2 As String, m_SqlGrpStr3 As String, m_SqlGrpStr4 As String, m_SqlGrpStr5 As String
'Add By Cheng 2003/06/13
Dim m_blnClkSure As Boolean '判斷是否按下確定
Dim m_EP27 As String  '判斷是否修改收卷註記
'92.8.8 ADD BY SONIA
Dim m_ST03 As String
'Add By Cheng 2003/10/07
Dim m_strCP09 As String '總收文號
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_EP13 As String 'Add By Cheng 2004/04/19 記錄原繪圖人員
Dim m_CP13 As String 'add by nick 2004/09/08
'add by nick 2004/10/14 國外案
Dim F_CP14 As String
Dim F_ST02 As String
Dim F_ST03 As String
Dim F_CP01020304 As String
Dim m_CP14 As String
Dim m_CP10 As String
'add by nick 2004/11/23
Dim StrSQL61 As String
Dim StrSQL62 As String
Dim StrSQL63 As String
Dim StrSQL64 As String
'add by nickc 2006/04/07
Dim StrSPa As String
Dim StrSTM As String
Dim StrSLC As String
Dim StrSHC As String
Dim StrSSP As String
'add by nick 2005/01/27
Dim m_Country As String
Dim m_CP31 As String
'add by nickc 2005/02/22
Dim m_CaseName As String
Dim m_SaleArea As String
'add by nick 2005/03/01
Dim m_CP107 As String
'add by nickc 2005/03/04 加乘註記
Dim m_CP98 As String
Dim m_CP99 As String
'add by nickc 2006/01/23
Dim m_CuNo As String
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
'edit by nick 2006/02/27 已不用常數
Dim m_FieldList() As FIELDITEM
'add by nickc 2006/12/29   紀錄 mail 資料，在 trans 後發
Dim skMail() As SeekMails
'add by nickc 2007/08/22
Dim m_NA03 As String
Dim bolInsert As Boolean, bolUpdate As Boolean, bolDelete As Boolean, bolSelect As Boolean, bolPrint As Boolean
'Add by Morgan 2008/12/2
Dim m_CPM05 As String
Dim m_CP112 As String
Dim m_CP44 As String        '2009/3/10 add by sonia
Dim m_stNewCP112 As String
'Dim m_bAlterCP112 As Boolean '提醒是否適用會稿加乘註記有無修改 'Removed by Morgan 2021/11/9 2010/10/13 已取消
Dim m_CP43 As String 'Add by Morgan 2009/12/3
Dim m_strNeedUpdateCase As String 'Add by Morgan 2010/7/20 待更新多國副案
'Add By Sindy 2013/6/7
Public m_chkcmdok1 As Boolean '記錄確定鍵是否存檔成功
Dim dblPrevRow As Double
Public intBackTab As Integer
'2013/6/7 End
Dim m_bol203Case As Boolean 'Added by Morgan 2013/8/1 是否有關聯案3日內收文主動修正(203)或修正(204)
Dim m_CPM28 As String 'Add By Sindy 2013/9/18
Dim m_CPM29 As String 'Add By Sindy 2013/9/30
Dim m_PP04 As String 'Add By Sindy 2013/10/14 預設核稿人
Dim m_PP05 As String 'Add By Sindy 2013/10/14 預設判發人
Public m_Flow As String 'Add By Sindy 2013/10/14 欲新增的下一流程
Dim m_EP33 As String 'Add By Sindy 2013/12/18
Dim m_PER04 As String, bolHadSetProofEngReader As Boolean 'Add By Sindy 2015/3/4
Dim m_EP41 As String 'Add By Sindy 2015/3/13 核稿語文
Dim lngFormWidth As Long, lngFormHeight As Long 'Added by Morgan 2016/2/18
Dim m_intRow As Integer, m_intCol As Integer 'Add By Sindy 2016/3/7
Dim m_EP39 As String '核稿完成日
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限
Dim pa() As String, cp() As String 'Add By Sindy 2018/4/10
Dim bolHadPOAeFile As Boolean 'Add By Sindy 2019/5/17
Dim m_strErrPath As String 'Add By Sindy 2022/8/18
'Added by Lydia 2025/02/05
Dim colFS_2 As Integer, colCP09_2 As Integer, colXno_2 As Integer, colEP08_2 As Integer, colNoShow_2 As Integer, colCaseNo_2 As Integer 'grd2使用的欄位
Dim colCp06_1 As Integer, colCp07_1 As Integer, colCP09_1 As Integer, colEp12_1 As Integer, colCaseNo_1 As Integer, colCaseName_1 As Integer, colCp57_1 As Integer
Dim colCp48_1 As Integer, colPv_1 As Integer, colEp06_1 As Integer, colEp34_1 As Integer, colEp07_1 As Integer, colCp27_1 As Integer, colCPM_1 As Integer
Dim colEp09_1 As Integer, colEp28_1 As Integer, colEp04_1 As Integer, colEp08_1 As Integer, colEp35_1 As Integer

''Add By Sindy 2018/4/10
''申請書
'Private Sub StartLetter2(ByVal ET01 As String, ByVal ET03 As String)
'   Dim strTxt(110) As String, strTmp As String
'   Dim ii As Integer, jj As Integer
'   Dim strInventor As String
'   Dim strTemp As String
'
'   ii = 0
'   EndLetter ET01, lbl1(3), ET03, strUserNum
'
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','本所案號','" & pa(1) & Val(pa(2)) & IIf(pa(3) <> "0" Or pa(4) <> "00", "-" & pa(3), IIf(pa(4) <> "00", "-" & pa(4), "")) & "')"
'
'   If pa(8) = "3" Then
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','設計種類','" & PUB_GetCaseAttributeName(pa(158), pa(8)) & "')"
'   End If
'
'   '申請人
'   Call PUB_GetApplPA_EData(ET01, ET03, lbl1(3), pa(), False)
'
'   '預設出名代理人
'   If cp(110) = "" Then PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10)
'   '出名代理人
'   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
'
'   '讀取發明人資料
'   If pa(8) = "1" Then
'      strExc(1) = "發明人"
'   ElseIf pa(8) = "2" Then
'      strExc(1) = "新型創作人"
'   Else
'      strExc(1) = "設計人"
'   End If
'   strInventor = ""
'   strExc(0) = " SELECT IN03,IN04,IN05,IN11,NA72" & _
'               " FROM PatentInventor,INVENTOR,NATION" & _
'               " WHERE pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
'               " AND IN01=substr(pi06,1,8) AND IN02=substr(pi06,9,2)" & _
'               " AND NA01(+)=IN11" & _
'               " order by pi05 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   '發明人TAG後面加序號,取消內縮
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      Do While Not RsTemp.EOF
'         'Modify By Sindy 2018/4/16 Mark:不要空行
'         'If strInventor <> "" Then strInventor = strInventor & vbCrLf
'         'Modify By Sindy 2018/10/25 增加英文名稱格式化 PUB_FCPIN05Format_EName
'         strInventor = strInventor & "【" & strExc(1) & intI & "】" & vbCrLf & _
'                                     "　　【國籍】　　　　　　　　　" & RsTemp("NA72") & vbCrLf & _
'                                     "　　【中文姓名】　　　　　　　" & IIf("" & RsTemp("IN11") = "000", PUB_ConvertNameFormat(ChgSQL("" & RsTemp("IN04"))), ChgSQL("" & RsTemp("IN04"))) & vbCrLf & _
'                                     IIf("" & RsTemp("IN05") = "", "", "　　【英文姓名】　　　　　　　" & ChgSQL(PUB_FCPIN05Format_EName("" & RsTemp("IN05"), RsTemp("NA72"))) & vbCrLf)
'         RsTemp.MoveNext
'         intI = intI + 1
'      Loop
'   Else
'      strInventor = "【" & strExc(1) & "1】" & vbCrLf & _
'                    "　　【國籍】　　　　　　　　　" & vbCrLf & _
'                    "　　【中文姓名】　　　　　　　" & vbCrLf
'   End If
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','發明人資料','" & strInventor & "')"
'
'   '優先權資料
'   strExc(0) = "SELECT sqldatew(pd05) pd05,na72,pd06,pd07,decode(pd08,'1','發明','2','新型','3','設計',pd08) pd08,pd09" & _
'      " FROM pridate,nation where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "' and pd04='" & pa(4) & "'" & _
'      " and na01(+)=pd07"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      jj = 0
'      Do While Not RsTemp.EOF
'         jj = jj + 1
'         If jj > 10 Then
'            MsgBox "優先權資料超過 10 筆，請自行維護！"
'            Exit Do
'         End If
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-日','" & RsTemp("pd05") & "')"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-國','" & RsTemp("na72") & "')"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-號','" & RsTemp("pd06") & "')"
'         '輸入優先權國家代碼時,代表是以電子交換檢送
'         If RsTemp("pd07") = "" & RsTemp("pd09") Then
'            '電子交換
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-碼','交換')"
'         ElseIf Not IsNull(RsTemp("pd09")) Then
'            '非電子交換
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-種類','" & IIf("" & RsTemp("pd08") = "", "♀", ChgSQL("" & RsTemp("pd08"))) & "')"
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-碼','" & ChgSQL(RsTemp("pd09")) & "')"
'         End If
'         RsTemp.MoveNext
'      Loop
'   End If
'
'   ii = ii + 1
'   strTemp = ""
'   If GetPrjPeople1(GetPrjPeopleNum1(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = GetPrjPeople1(GetPrjPeopleNum1(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   If GetPrjPeople1(GetPrjPeopleNum2(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum2(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   If GetPrjPeople1(GetPrjPeopleNum3(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum3(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   If GetPrjPeople1(GetPrjPeopleNum4(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum4(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   If GetPrjPeople1(GetPrjPeopleNum5(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum5(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & lbl1(3) & "','" & ET03 & "','" & strUserNum & "','收據抬頭','" & strTemp & "')"
'
'   If Not ClsLawExecSQL(ii, strTxt) Then
'      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
'   End If
'End Sub

'Add By Sindy 2018/11/20
'strFileKind:副檔名
'strCaseNo:本所案號(XXX-XXXXXX-X-XX)
Private Function GetCPPFileAndDownload(strFileKind As String, strCaseNo As String) As Boolean
Dim rsTmp As ADODB.Recordset
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
   GetCPPFileAndDownload = False
   strCP01 = SystemNumber(strCaseNo, 1)
   strCP02 = SystemNumber(strCaseNo, 2)
   strCP03 = SystemNumber(strCaseNo, 3)
   strCP04 = SystemNumber(strCaseNo, 4)
   strExc(0) = "select cpp01,cpp02 from casepaperpdf,caseprogress" & _
               " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
               " and cp09=cpp01(+)" & _
               " and upper(substr(cpp02,-8))=upper('." & strFileKind & ".PDF')" & _
               " order by CPP06 desc,CPP07 desc"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetCPPFileAndDownload = True
   End If
   Set rsTmp = Nothing
End Function

'Add By Sindy 2019/5/16
'檢查是否有總委任書
Private Function ChkPOAExists(bolHadPOAeFile As Boolean) As Boolean
Dim strCaseNo As String
Dim strPOAPath As String
Dim strPOAFileName As String
'Dim strPOAFileFullName As String
Dim strSubCaseNo As String
   
   ChkPOAExists = False: bolHadPOAeFile = False
   strCaseNo = Trim(pa(1)) & Val(Trim(pa(2))) & _
                 IIf(Val(Trim(pa(3))) = 0 And Val(Trim(pa(4))) = 0, "", "-" & pa(3)) & _
                 IIf(Val(Trim(pa(4))) = 0, "", "-" & Format(pa(4), "00"))
   If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
      strPOAPath = PUB_Getdesktop
   Else
      'Modify By Sindy 2022/10/25 改用常變數 str_P_台灣電子送件檔案路徑
      strPOAPath = str_P_台灣電子送件檔案路徑
   End If
   strPOAPath = strPOAPath & "\" & strCaseNo
   m_strErrPath = strPOAPath 'Add By Sindy 2022/8/18
   
   strPOAFileName = strCaseNo & ".POA.pdf"
'   strPOAFileFullName = strPOAPath & "\" & strPOAFileName
   
   '個案,POA存取...
   'If ChkPOAExists = False Or Dir(strPOAFileFullName) = "" Then
   'Modify By Sindy 2022/10/25 改用常變數 str_P_OrderPath
   m_strErrPath = str_P_OrderPath & "\POA\" & strPOAFileName 'Add By Sindy 2022/8/18
   If Dir(str_P_OrderPath & "\POA\" & strPOAFileName) <> "" Then
      ChkPOAExists = True '有委任書
      bolHadPOAeFile = True '有委任書電子檔
   End If
   
   '是否有總委任書
   If ChkPOAExists = False Then
      If pa(26) <> "" Then 'And Dir(strPOAFileFullName) = ""
         If PUB_ChkPA165IsY(pa(26), strSubCaseNo) = True Then
            ChkPOAExists = True
            bolHadPOAeFile = GetCPPFileAndDownload("POA", strSubCaseNo)
         End If
      End If
      If pa(27) <> "" Then 'And Dir(strPOAFileFullName) = ""
         If PUB_ChkPA165IsY(pa(27), strSubCaseNo) = True Then
            ChkPOAExists = True
            bolHadPOAeFile = GetCPPFileAndDownload("POA", strSubCaseNo)
         End If
      End If
      If pa(28) <> "" Then 'And Dir(strPOAFileFullName) = ""
         If PUB_ChkPA165IsY(pa(28), strSubCaseNo) = True Then
            ChkPOAExists = True
            bolHadPOAeFile = GetCPPFileAndDownload("POA", strSubCaseNo)
         End If
      End If
      If pa(29) <> "" Then 'And Dir(strPOAFileFullName) = ""
         If PUB_ChkPA165IsY(pa(29), strSubCaseNo) = True Then
            ChkPOAExists = True
            bolHadPOAeFile = GetCPPFileAndDownload("POA", strSubCaseNo)
         End If
      End If
      If pa(30) <> "" Then 'And Dir(strPOAFileFullName) = ""
         If PUB_ChkPA165IsY(pa(30), strSubCaseNo) = True Then
            ChkPOAExists = True
            bolHadPOAeFile = GetCPPFileAndDownload("POA", strSubCaseNo)
         End If
      End If
   End If
   m_strErrPath = ""
End Function

'91.08.14 nick  本來區塊 1 文件和申請書都會跑，現在只剩文件檔案才跑，因為申請書改跑定搞
Private Sub cmd_Click(Index As Integer)
Dim strTempName As String  '2009/3/10 add by sonia
'Added by Morgan 2013/4/16
Dim arrTemp
Dim bolAsk As Boolean
'end 2013/4/16
'Dim strLetterCP10 As String 'Added by Morgan 2017/1/23
Dim nFrm As Form
Dim stContent As String
Dim m_bolShowEng As Boolean
Dim bolHadPOA As Boolean
Dim bolSysLtr As Boolean 'Added by Morgan 2021/2/24
Dim strReturnSheet As String 'Added by Morgan 2023/6/21
Dim bol As Boolean 'Added by Morgan 2023/8/24

On Error GoTo ErrHnd
   
Select Case Index
'Add By Sindy 2018/4/10 +申請書
Case 0 '申請書
   'Add By Sindy 2019/8/13 + 203,204,205
   'Modify By Sindy 2019/8/29 + 107:為延期再審,出修正申請書
   If LBL1(29) = "203" Or LBL1(29) = "204" Or LBL1(29) = "205" Or LBL1(29) = "107" Then
      frm090201_2_5.m_CP09 = LBL1(3) '總收文號
      frm090201_2_5.m_CaseNo = LBL1(7) '本所案號
      frm090201_2_5.Show vbModal
   'Added by Morgan 2023/5/25 +244補中文說明書,232補優先權證明
   ElseIf LBL1(29) = "244" Or LBL1(29) = "232" Then
      Set nFrm = Forms(0).GetForm("frm04010304_1")
      If Not nFrm Is Nothing Then
         Set nFrm.oParentForm = Me
         'frm04010304_1 會用到 總收文號 Lbl1(3) 及 本所案號 Lbl1(7)
         nFrm.Show
      End If
   'end 2023/5/25
   Else
   '2019/8/13 END
      Screen.MousePointer = vbHourglass
      ReDim pa(1 To TF_PA) As String
      ReDim cp(TF_CP)
      '專利基本檔
      pa(1) = SystemNumber(LBL1(7).Caption, 1)
      pa(2) = SystemNumber(LBL1(7).Caption, 2)
      pa(3) = SystemNumber(LBL1(7).Caption, 3)
      pa(4) = SystemNumber(LBL1(7).Caption, 4)
      Call ClsPDReadPatentDatabase(pa(), 國內)
      '進度檔
      cp(9) = LBL1(3)
      Call PUB_ReadCaseProgressDatabase(cp(), 國內)
   '   '1.基本資料
   '   StartLetterPA_EData "01", "14", lbl1(3), pa, cp, True
   '   NowPrint lbl1(3), "01", "14", True, strUserNum
   '   '2.申請書
   '   StartLetter2 "01", "03"
   '   NowPrint lbl1(3), "01", "03", True, strUserNum
      
      'Modify By Sindy 2018/12/20 改與程序程式相同
      '2.申請書
   '   StartLetter2 "01", "03"
   '   NowPrint lbl1(3), "01", "03", False, strUserNum, , stContent, True, stContent
      'Modify By Sindy 2019/5/17 ChkPOAExists:檢查有無委任書
      bolHadPOA = ChkPOAExists(bolHadPOAeFile)
      'Add By Sindy 2023/5/9 + 分割
      If m_CP10 = "307" Then
         If pa(8) = "1" Then
            Pub_P_NewCaseStartLetter2 "01", "01", LBL1(3), pa, cp, IIf(lblCM10.Visible = True, True, False), bolHadPOA, m_bolShowEng, bolHadPOAeFile
            NowPrint LBL1(3), "01", "01", False, strUserNum, , stContent, True, stContent
         ElseIf pa(8) = "2" Then
            Pub_P_NewCaseStartLetter2 "01", "02", LBL1(3), pa, cp, IIf(lblCM10.Visible = True, True, False), bolHadPOA, m_bolShowEng, bolHadPOAeFile
            NowPrint LBL1(3), "01", "02", False, strUserNum, , stContent, True, stContent
         Else
            Pub_P_NewCaseStartLetter2 "01", "03", LBL1(3), pa, cp, IIf(lblCM10.Visible = True, True, False), bolHadPOA, m_bolShowEng, bolHadPOAeFile
            NowPrint LBL1(3), "01", "03", False, strUserNum, , stContent, True, stContent
         End If
      'Added by Morgan 2023/8/24
      ElseIf m_CP10 = "239" Then
         stContent = ""
         Pub_P_NewCaseStartLetter2 "01", "01", LBL1(3), pa, cp, IIf(lblCM10.Visible = True, True, False), bolHadPOA, m_bolShowEng, bolHadPOAeFile
         NowPrint LBL1(3), "01", "01", True, strUserNum
      'end 2023/8/24
      Else
      '2023/5/9 END
         Pub_P_NewCaseStartLetter2 "01", "03", LBL1(3), pa, cp, IIf(lblCM10.Visible = True, True, False), bolHadPOA, m_bolShowEng, bolHadPOAeFile
         NowPrint LBL1(3), "01", "03", False, strUserNum, , stContent, True, stContent
      End If
      'NowPrint lbl1(3), "01", "03", False, strUserNum, , , True, strExc(9)
   '   strFileName = m_strFolder & "\" & m_strCaseNo & ".data" 'm_CPM26
   '   If Dir(strFileName) <> "" Then
   '      strFileName = m_strFolder & "\" & m_strCaseNo & ".data_" & Trim(lblPA08.Caption) & "專利申請書"
   '   End If
   '   Call PUB_MakeDoc(strExc(9), strFileName)
      
      '1.基本資料
   '   StartLetterPA_EData "01", "14", lbl1(3), pa, cp, True, True
   '   NowPrint lbl1(3), "01", "14", True, strUserNum, , stContent, , , , , True, , , , , , , , True
   '不正確的檔案名稱或號碼 (錯誤 52) ==> 檢查資料夾有無權限
      If stContent <> "" Then 'Added by Morgan 2023/8/24
         StartLetterPA_EData "01", "14", LBL1(3), pa, cp, True, True, , , m_bolShowEng
         NowPrint LBL1(3), "01", "14", True, strUserNum, , stContent, , , , , True, , , , , , , , True
      End If
      'NowPrint lbl1(3), "01", "14", True, strUserNum, , strExc(9), True, strExc(10)
   '   strFileName = m_strFolder & "\" & m_strCaseNo & ".contact"
   '   'Chr(12):跳頁
   '   Call PUB_MakeDoc(strExc(9) & Chr(12) & strExc(10), strFileName, False)
      
      Screen.MousePointer = vbDefault
      MsgBox "資料已產生完畢!!!"
   End If
'Remove by Morgan 2012/4/6
'Case 0 '文件檔案
'        Dim Wxdc As New Word.Application
'       '*****************    區塊 1  start
'        strCP10 = Trim(lbl1(15).Caption)
'        If strCP10 = "修正" Then
'           Set AdoRs = New ADODB.Recordset
'           StrNewSQL = "SELECT DECODE(PA09,'000',CPM03,CPM04) FROM CASEPROGRESS,CASEPROPERTYMAP,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105') "
'           StrNewSQL = StrNewSQL & " UNION all  SELECT DECODE(TM10,'000',CPM03,CPM04) FROM CASEPROGRESS,CASEPROPERTYMAP,TRADEMARK WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105')  "
'           StrNewSQL = StrNewSQL & " UNION all  SELECT DECODE(LC15,'000',CPM03,CPM04) FROM CASEPROGRESS,CASEPROPERTYMAP,LAWCASE WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105') "
'           StrNewSQL = StrNewSQL & " UNION all  SELECT CPM03                          FROM CASEPROGRESS,CASEPROPERTYMAP,HIRECASE WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105') "
'           StrNewSQL = StrNewSQL & " UNION all  SELECT DECODE(SP09,'000',CPM03,CPM04) FROM CASEPROGRESS,CASEPROPERTYMAP,SERVICEPRACTICE WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105') "
'           AdoRs.CursorLocation = adUseClient
'           AdoRs.Open StrNewSQL, cnnConnection, adOpenStatic, adLockReadOnly
'           If AdoRs.RecordCount <> 0 Then
'              strCP10 = CheckStr(AdoRs.Fields(0))
'           Else
'              s = MsgBox("此案號無申請案之文件檔！！", , "沒有檔案！！")
'              AdoRs.Close
'              Set AdoRs = Nothing
'              Exit Sub
'           End If
'           AdoRs.Close
'           Set AdoRs = Nothing
'        End If
'        Dim DFileName As String    '應該存放檔名
'        Dim DSFileName As String   '範本檔名
'        Dim DFilePath As String    '應該存放路徑
'        DFileName = ChangeFileName(Trim(lbl1(7).Caption), Trim(lbl1(15).Caption), Trim(lbl1(3).Caption))
'        DFilePath = GetDocFilePath(DFileName) & "\"
'        DFileName = DFileName & ".doc"
'        DSFileName = SystemNumber(Trim(lbl1(7).Caption), 1) & Trim(lbl1(15).Caption) & ".doc"
'        '****************************** 區塊 1 end
'     If Len(Dir(DFilePath & DFileName)) <> 0 Then
'        Set Wxdc = CreateObject("word.application")
'        Wxdc.Visible = True
'        Wxdc.Documents.Open DFilePath & DFileName
'     Else
'        If Len(Dir(SMPPath & "\" & DSFileName)) <> 0 Then
'            Wxdc.Visible = True
'            Wxdc.Documents.Open SMPPath & "\" & DSFileName
'            Wxdc.Documents(1).SaveAs DFilePath & DFileName
'        Else
'            s = MsgBox("檔名：" & SMPPath & DSFileName & "不存在", , "錯誤發生")
'            Exit Sub
'        End If
'     End If

'Added by Morgan 2012/4/6
'Removed by Morgan2016/2/18 取消
'Case 0 '開庭/面詢紀錄上傳
'     frm090201_2_4.m_Key = LBL1(3)
'     frm090201_2_4.Show vbModal
     
'Modify By Sindy 2013/4/16
Case 1 '承辦歷程
      'NowPrint lbl1(3).Caption, "99", "00", True, strUserNum '申請書
      'Add By Sindy 2013/9/16
      If ProState = "2" Then
         If frm090614.txt1(8) = "N" Then MsgBox "不可從（不區分個人）的資料查詢中來執行承辦歷程作業！": Exit Sub
      End If
      '2013/9/16 END
      
      'Add By Sindy 2017/8/3 個人案件不可用主管權限操作
      If ProState = "2" And m_CP14 = strUserNum Then '2.主管
         MsgBox "個人案件不可用主管權限操作！", vbExclamation
         Exit Sub
      End If
      '2017/8/3 END
      
      'Add By Sindy 2015/12/3
      '重新檢查欄位有效性
      If TxtValidate = True Then
      '2015/12/3 END
         'Add By Sindy 2013/6/10
         If SetColTag(False) = False Then
            cmdOK(1).Enabled = False 'Add By Sindy 2017/9/21
            Call cmdok_Click(1)
            cmdOK(1).Enabled = True 'Add By Sindy 2017/9/21
            If m_chkcmdok1 = False Then Exit Sub
         Else
            Call Process(LBL1(3)) '要重新查詢資料 Add By Sindy 2018/10/4
         End If
         
'         'Add By Sindy 2017/9/19
'         '檢查表單是否已開啟，若是，則關閉
'         For Each nFrm In Forms
'            If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'               Unload frm090202_2
'               'Add By Sindy 2020/1/17 有資料要儲存,尚需處理...
'               If strSaveConfirm = True Then
'                  frm090202_2.ZOrder
'                  Exit Sub
'               Else
'               '2020/1/17 END
'                  Exit For
'               End If
'            End If
'         Next
'         '2017/9/19 END
         If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
         intBackTab = 1
         '2013/6/10 End
         frm090202_2.Hide
         frm090202_2.m_EEP01 = LBL1(3) '總收文號
         frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) 'Add By Sindy 2013/9/12 案件流程所屬人員
         frm090202_2.intReceiveKind = 0 '0.承辦人工作進度
         frm090202_2.SetParent Me
         If frm090202_2.QueryData = True Then
            frm090202_2.Show
            Me.Hide
         End If
      End If
'2013/4/16 End
      
'Modify By Sindy 2025/7/10 詢問雅娟經理,此功能確定已無使用,可以撤下
''add by nickc 2005/02/22  專利處加的功能
'Case 2
'      Dim CUID As String
'      Dim INV_CUID As String    '2008/7/24 ADD BY SONIA 發明人ID
'      Dim TmpFileName As String
'      'add by nickc 2007/02/27
'      Dim oTM12 As String
'      Dim oTM15 As String
'      Dim oTM2122 As String
'      'add by  nickc 2007/02/26
'      'edit by nickc 2008/02/29 TS 跑不進去
'      'If SystemNumber(lbl1(7), 1) <> "T" Then
'      If InStr(1, SystemNumber(lbl1(7), 1), "T") = 0 Then
'                                    strSql = "select cu11 from patent,customer where pa01='" & SystemNumber(lbl1(7), 1) & "' and pa02='" & SystemNumber(lbl1(7), 2) & "' and pa03='" & SystemNumber(lbl1(7), 3) & "' and pa04='" & SystemNumber(lbl1(7), 4) & "' and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from patent,customer where pa01='" & SystemNumber(lbl1(7), 1) & "' and pa02='" & SystemNumber(lbl1(7), 2) & "' and pa03='" & SystemNumber(lbl1(7), 3) & "' and pa04='" & SystemNumber(lbl1(7), 4) & "' and substr(pa27,1,8)=cu01(+) and substr(pa27,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from patent,customer where pa01='" & SystemNumber(lbl1(7), 1) & "' and pa02='" & SystemNumber(lbl1(7), 2) & "' and pa03='" & SystemNumber(lbl1(7), 3) & "' and pa04='" & SystemNumber(lbl1(7), 4) & "' and substr(pa28,1,8)=cu01(+) and substr(pa28,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from patent,customer where pa01='" & SystemNumber(lbl1(7), 1) & "' and pa02='" & SystemNumber(lbl1(7), 2) & "' and pa03='" & SystemNumber(lbl1(7), 3) & "' and pa04='" & SystemNumber(lbl1(7), 4) & "' and substr(pa29,1,8)=cu01(+) and substr(pa29,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from patent,customer where pa01='" & SystemNumber(lbl1(7), 1) & "' and pa02='" & SystemNumber(lbl1(7), 2) & "' and pa03='" & SystemNumber(lbl1(7), 3) & "' and pa04='" & SystemNumber(lbl1(7), 4) & "' and substr(pa30,1,8)=cu01(+) and substr(pa30,9,1)=cu02(+) "
'
'            'Add By Sindy 2011/2/18 增加TM78,TM79,TM80,TM81
'            strSql = strSql & " union select cu11 from trademark,customer where tm01='" & SystemNumber(lbl1(7), 1) & "' and tm02='" & SystemNumber(lbl1(7), 2) & "' and tm03='" & SystemNumber(lbl1(7), 3) & "' and tm04='" & SystemNumber(lbl1(7), 4) & "' and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from trademark,customer where tm01='" & SystemNumber(lbl1(7), 1) & "' and tm02='" & SystemNumber(lbl1(7), 2) & "' and tm03='" & SystemNumber(lbl1(7), 3) & "' and tm04='" & SystemNumber(lbl1(7), 4) & "' and substr(tm78,1,8)=cu01(+) and substr(tm78,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from trademark,customer where tm01='" & SystemNumber(lbl1(7), 1) & "' and tm02='" & SystemNumber(lbl1(7), 2) & "' and tm03='" & SystemNumber(lbl1(7), 3) & "' and tm04='" & SystemNumber(lbl1(7), 4) & "' and substr(tm79,1,8)=cu01(+) and substr(tm79,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from trademark,customer where tm01='" & SystemNumber(lbl1(7), 1) & "' and tm02='" & SystemNumber(lbl1(7), 2) & "' and tm03='" & SystemNumber(lbl1(7), 3) & "' and tm04='" & SystemNumber(lbl1(7), 4) & "' and substr(tm80,1,8)=cu01(+) and substr(tm80,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from trademark,customer where tm01='" & SystemNumber(lbl1(7), 1) & "' and tm02='" & SystemNumber(lbl1(7), 2) & "' and tm03='" & SystemNumber(lbl1(7), 3) & "' and tm04='" & SystemNumber(lbl1(7), 4) & "' and substr(tm81,1,8)=cu01(+) and substr(tm81,9,1)=cu02(+) "
'
'            'Add By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46
'            strSql = strSql & " union select cu11 from lawcase,customer where lc01='" & SystemNumber(lbl1(7), 1) & "' and lc02='" & SystemNumber(lbl1(7), 2) & "' and lc03='" & SystemNumber(lbl1(7), 3) & "' and lc04='" & SystemNumber(lbl1(7), 4) & "' and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from lawcase,customer where lc01='" & SystemNumber(lbl1(7), 1) & "' and lc02='" & SystemNumber(lbl1(7), 2) & "' and lc03='" & SystemNumber(lbl1(7), 3) & "' and lc04='" & SystemNumber(lbl1(7), 4) & "' and substr(lc43,1,8)=cu01(+) and substr(lc43,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from lawcase,customer where lc01='" & SystemNumber(lbl1(7), 1) & "' and lc02='" & SystemNumber(lbl1(7), 2) & "' and lc03='" & SystemNumber(lbl1(7), 3) & "' and lc04='" & SystemNumber(lbl1(7), 4) & "' and substr(lc44,1,8)=cu01(+) and substr(lc44,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from lawcase,customer where lc01='" & SystemNumber(lbl1(7), 1) & "' and lc02='" & SystemNumber(lbl1(7), 2) & "' and lc03='" & SystemNumber(lbl1(7), 3) & "' and lc04='" & SystemNumber(lbl1(7), 4) & "' and substr(lc45,1,8)=cu01(+) and substr(lc45,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from lawcase,customer where lc01='" & SystemNumber(lbl1(7), 1) & "' and lc02='" & SystemNumber(lbl1(7), 2) & "' and lc03='" & SystemNumber(lbl1(7), 3) & "' and lc04='" & SystemNumber(lbl1(7), 4) & "' and substr(lc46,1,8)=cu01(+) and substr(lc46,9,1)=cu02(+) "
'
'            'Add By Sindy 2011/2/18 增加HC24,HC25,HC26,HC27
'            strSql = strSql & " union select cu11 from hirecase,customer where hc01='" & SystemNumber(lbl1(7), 1) & "' and hc02='" & SystemNumber(lbl1(7), 2) & "' and hc03='" & SystemNumber(lbl1(7), 3) & "' and hc04='" & SystemNumber(lbl1(7), 4) & "' and substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from hirecase,customer where hc01='" & SystemNumber(lbl1(7), 1) & "' and hc02='" & SystemNumber(lbl1(7), 2) & "' and hc03='" & SystemNumber(lbl1(7), 3) & "' and hc04='" & SystemNumber(lbl1(7), 4) & "' and substr(hc24,1,8)=cu01(+) and substr(hc24,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from hirecase,customer where hc01='" & SystemNumber(lbl1(7), 1) & "' and hc02='" & SystemNumber(lbl1(7), 2) & "' and hc03='" & SystemNumber(lbl1(7), 3) & "' and hc04='" & SystemNumber(lbl1(7), 4) & "' and substr(hc25,1,8)=cu01(+) and substr(hc25,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from hirecase,customer where hc01='" & SystemNumber(lbl1(7), 1) & "' and hc02='" & SystemNumber(lbl1(7), 2) & "' and hc03='" & SystemNumber(lbl1(7), 3) & "' and hc04='" & SystemNumber(lbl1(7), 4) & "' and substr(hc26,1,8)=cu01(+) and substr(hc26,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from hirecase,customer where hc01='" & SystemNumber(lbl1(7), 1) & "' and hc02='" & SystemNumber(lbl1(7), 2) & "' and hc03='" & SystemNumber(lbl1(7), 3) & "' and hc04='" & SystemNumber(lbl1(7), 4) & "' and substr(hc27,1,8)=cu01(+) and substr(hc27,9,1)=cu02(+) "
'
'            'Add By Sindy 2011/2/18 增加SP65,SP66
'            strSql = strSql & " union select cu11 from servicepractice,customer where sp01='" & SystemNumber(lbl1(7), 1) & "' and sp02='" & SystemNumber(lbl1(7), 2) & "' and sp03='" & SystemNumber(lbl1(7), 3) & "' and sp04='" & SystemNumber(lbl1(7), 4) & "' and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from servicepractice,customer where sp01='" & SystemNumber(lbl1(7), 1) & "' and sp02='" & SystemNumber(lbl1(7), 2) & "' and sp03='" & SystemNumber(lbl1(7), 3) & "' and sp04='" & SystemNumber(lbl1(7), 4) & "' and substr(sp58,1,8)=cu01(+) and substr(sp58,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from servicepractice,customer where sp01='" & SystemNumber(lbl1(7), 1) & "' and sp02='" & SystemNumber(lbl1(7), 2) & "' and sp03='" & SystemNumber(lbl1(7), 3) & "' and sp04='" & SystemNumber(lbl1(7), 4) & "' and substr(sp59,1,8)=cu01(+) and substr(sp59,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from servicepractice,customer where sp01='" & SystemNumber(lbl1(7), 1) & "' and sp02='" & SystemNumber(lbl1(7), 2) & "' and sp03='" & SystemNumber(lbl1(7), 3) & "' and sp04='" & SystemNumber(lbl1(7), 4) & "' and substr(sp65,1,8)=cu01(+) and substr(sp65,9,1)=cu02(+) "
'            strSql = strSql & " union select cu11 from servicepractice,customer where sp01='" & SystemNumber(lbl1(7), 1) & "' and sp02='" & SystemNumber(lbl1(7), 2) & "' and sp03='" & SystemNumber(lbl1(7), 3) & "' and sp04='" & SystemNumber(lbl1(7), 4) & "' and substr(sp66,1,8)=cu01(+) and substr(sp66,9,1)=cu02(+) "
'            CheckOC3
'            AdoRecordSet3.CursorLocation = adUseClient
'            AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            CUID = ""
'            If AdoRecordSet3.RecordCount <> 0 Then
'               AdoRecordSet3.MoveFirst
'               Do While Not AdoRecordSet3.EOF
'                  If Not IsNull(AdoRecordSet3.Fields(0).Value) Then
'                     CUID = CUID & "" & AdoRecordSet3.Fields(0).Value & ","
'                  End If
'                  AdoRecordSet3.MoveNext
'               Loop
'            End If
'            CheckOC3
'            'add by nickc 2005/07/06 當新申請案，受文者及副本收受者不可修改，主旨不變
'            If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "105" _
'               Or m_CP10 = "109" Or m_CP10 = "110" Or m_CP10 = "112" Or m_CP10 = "113" _
'               Or m_CP10 = "114" Or m_CP10 = "115" Or m_CP10 = "118" Or m_CP10 = "301" _
'               Or m_CP10 = "302" Or m_CP10 = "303" Or m_CP10 = "304" Or m_CP10 = "305" _
'               Or m_CP10 = "306" Or m_CP10 = "307" Or m_CP10 = "803" Then
'               'frm090201_2_2.oStrA01 = IIf(m_country = "000", "智慧局", GetNationName(m_country, 0) & " 代理人")
'               frm090201_2_2.txt1(5).Text = IIf(m_Country = "000", "智慧局", GetNationName(m_Country, 0) & " 代理人")
'               frm090201_2_2.txt1(5).Enabled = False
'               'frm090201_2_2.oStrA04 = "北所、" & IIf(m_SaleArea = "1", "", IIf(m_SaleArea = "2", "中所、", IIf(m_SaleArea = "3", "南所、", IIf(m_SaleArea = "4", "高所、", "")))) & "客戶"
'               frm090201_2_2.txt1(6).Text = "北所、" & IIf(m_SaleArea = "1", "", IIf(m_SaleArea = "2", "中所、", IIf(m_SaleArea = "3", "南所、", IIf(m_SaleArea = "4", "高所、", "")))) & "客戶"
'               frm090201_2_2.txt1(6).Enabled = False
'               'edit by nickc 2005/03/15 郭說不印專利種類，改印案件性質
'               'frm090201_2_2.txt1(0) = "為「" & m_CaseName & "」" & GetNationName(m_country, 0) & lbl1(13) & "專利案提出申請。"
'               frm090201_2_2.txt1(0) = "為「" & m_CaseName & "」" & GetNationName(m_Country, 0) & lbl1(15) & "專利案提出申請。"
'               'add by nickc 2007/12/03 郭加入台灣發明新型在備註欄放入提示
'               If (m_CP10 = "101" Or m_CP10 = "102") And m_Country = "000" Then
'                    frm090201_2_2.txt1(4).Text = "請注意：本案若同時或隨後可能申請大陸專利" & vbCrLf & "　　　　，請留意是否有超頁超項問題：" & vbCrLf & "1.專利說明書(含申請專利範圍、圖式)以30頁" & vbCrLf & "　為限，每增加1頁加收新台幣500元。" & vbCrLf & "2.申請專利範圍以10項為限，每增加1項加收" & vbCrLf & "　新台幣1000元。 "
'               End If
'            Else
'               frm090201_2_2.txt1(5).Text = ""
'               frm090201_2_2.txt1(6).Text = ""
'               'edit by nickc 2005/07/22 郭說後面加專利種類   專利之 案件性質
'               'frm090201_2_2.txt1(0).Text = "為「" & m_CaseName & "」" & GetNationName(m_country, 0)
'               frm090201_2_2.txt1(0).Text = "「" & m_CaseName & "」" & GetNationName(m_Country, 0) & lbl1(13) & "專利之" & lbl1(15)
'            End If
'            If Right(CUID, 1) = "," Then CUID = Left(CUID, Len(CUID) - 1)
'
'            'Modify By Sindy 2014/11/6 抓發明人ID字串
'            If strSrvDate(1) >= 專利發明人檔啟用日 Then
'               strSql = "select IN03 from patentInventor,Inventor where pi01='" & SystemNumber(lbl1(7), 1) & "' and pi02='" & SystemNumber(lbl1(7), 2) & "' and pi03='" & SystemNumber(lbl1(7), 3) & "' and pi04='" & SystemNumber(lbl1(7), 4) & "' and substr(pi06,1,8)=IN01(+) and substr(pi06,9,2)=IN02(+)"
'            Else
'            '2014/11/6 END
'               '2008/7/24 ADD BY SONIA 抓發明人ID字串
'               'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
'            End If
'            CheckOC3
'            AdoRecordSet3.CursorLocation = adUseClient
'            AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            INV_CUID = ""
'            If AdoRecordSet3.RecordCount <> 0 Then
'               AdoRecordSet3.MoveFirst
'               Do While Not AdoRecordSet3.EOF
'                  If Not IsNull(AdoRecordSet3.Fields(0).Value) Then
'                     INV_CUID = INV_CUID & "" & AdoRecordSet3.Fields(0).Value & ","
'                  End If
'                  AdoRecordSet3.MoveNext
'               Loop
'            End If
'            CheckOC3
'            If Right(INV_CUID, 1) = "," Then INV_CUID = Left(INV_CUID, Len(INV_CUID) - 1)
'            '2008/7/24 END
'
'            frm090201_2_2.oStrA02 = lbl1(7)
'            If lbl1(17) <> "" Then
'               frm090201_2_2.oStrA05 = "   " & Mid(Replace(lbl1(17), "/", ""), 1, Len(Replace(lbl1(17), "/", "")) - 4) & "年  " & Left(Right(Replace(lbl1(17), "/", ""), 4), 2) & "月  " & Right(Replace(lbl1(17), "/", ""), 2) & "日"
'            Else
'               frm090201_2_2.oStrA05 = "     年    月    日"
'            End If
'            If lbl1(19) <> "" Then
'               frm090201_2_2.oStrA06 = "   " & Mid(Replace(lbl1(19), "/", ""), 1, Len(Replace(lbl1(19), "/", "")) - 4) & "年  " & Left(Right(Replace(lbl1(19), "/", ""), 4), 2) & "月  " & Right(Replace(lbl1(19), "/", ""), 2) & "日"
'            Else
'               frm090201_2_2.oStrA06 = "     年    月    日"
'            End If
'            frm090201_2_2.txt1(1) = CUID
'            frm090201_2_2.txt1(2) = INV_CUID   '2008/7/24 ADD BY SONIA
'            TmpFileName = ""
'            If SystemNumber(lbl1(7), 1) = "P" Then
'               Select Case m_CP10
'               Case "101", "301"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".inv"
'               Case "102", "104", "302", "304"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".utl"
'               Case "103", "105", "303", "305"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".des"
'               Case "107"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".re"
'               Case "203", "204"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".fix"
'               Case "205"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".ex"
'               Case "206"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".add"
'               Case "501"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".app"
'               Case "503"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".exa"
'               Case "505"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".bpp"
'               Case "801"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".opp"
'               Case "802"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".oas"
'               Case "803"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".rev"
'               Case "804"
'                   TmpFileName = SystemNumber(lbl1(7), 1) & Trim(Val(SystemNumber(lbl1(7), 2))) & ".ras"
'               Case Else
'               End Select
'            End If
'            frm090201_2_2.oStrA08 = TmpFileName
'            frm090201_2_2.oStrA09 = lbl1(23)
'            frm090201_2_2.oStrA10 = lbl1(26) & "  天"
'            '2010/5/24 modify by sonia 智權人員前加所別
'            'frm090201_2_2.oStrA11 = lbl1(21)
'            frm090201_2_2.oStrA11 = ""
'            Select Case PUB_GetST06(m_CP13)
'               Case "2"
'                  frm090201_2_2.oStrA11 = "中所"
'               Case "3"
'                  frm090201_2_2.oStrA11 = "南所"
'               Case "4"
'                  frm090201_2_2.oStrA11 = "高所"
'            End Select
'            frm090201_2_2.oStrA11 = frm090201_2_2.oStrA11 & lbl1(21)
'            '2010/5/24 end
'            If txt1(2) <> "" Then
'               frm090201_2_2.oStrA12 = " " & Mid(txt1(2), 1, Len(txt1(2)) - 4) & "年" & Left(Right(txt1(2), 4), 2) & "月" & Right(txt1(2), 2) & "日"
'            Else
'               frm090201_2_2.oStrA12 = "   年  月  日"
'            End If
'            If txt1(3) <> "" Then
'               frm090201_2_2.oStrA13 = " " & Mid(txt1(3), 1, Len(txt1(3)) - 4) & "年" & Left(Right(txt1(3), 4), 2) & "月" & Right(txt1(3), 2) & "日"
'            Else
'               frm090201_2_2.oStrA13 = "   年  月  日"
'            End If
'            frm090201_2_2.oStrA14 = PUB_GetST07(m_CP14)
'
'            'Added by Morgan 2013/4/16 美專新申請案輸入完稿日列印承辦單前選擇
'            '條件：1. 主張優先權日是在2013年3月16日(不含)之前的美專新申請案,2. CIP,分割案之須判斷原案之申請日或主張優先權日是在2013年3月16日(不含)之前;
'            If SystemNumber(lbl1(7), 1) = "CFP" And m_Country = "101" And txt1(3) <> "" And InStr("101,113,307", m_CP10) > 0 Then
'               arrTemp = Split(Me.lbl1(7).Caption, "-")
'               bolAsk = False
'               If m_CP10 = "307" Then
'                  strExc(0) = "select nvl(pd05,pa10) from divisioncase,patent,pridate where dc01='" & arrTemp(0) & "' and dc02='" & arrTemp(1) & "' and dc03='" & arrTemp(2) & "' and dc04='" & arrTemp(3) & "'" & _
'                     " and pa01(+)=dc05 and pa02(+)=dc06 and pa03(+)=dc07 and pa04(+)=dc08 and pd01(+)=dc05 and pd02(+)=dc06 and pd03(+)=dc07 and pd04(+)=dc08 and nvl(pd05,pa10)<20130316 and rownum<2"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     bolAsk = True
'                  End If
'               ElseIf m_CP10 = "113" Then
'                  strExc(0) = "select nvl(pd05,pa10) from patent,pridate where pa01='" & arrTemp(0) & "' and pa02='" & arrTemp(1) & "' and pa03='0' and pa04='" & arrTemp(3) & "'" & _
'                     " and pd01(+)=pa01 and pd02(+)=pa02 and pd03(+)=pa03 and pd04(+)=pa04 and nvl(pd05,pa10)<20130316 and rownum<2"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     bolAsk = True
'                  End If
'               Else
'                  strExc(0) = "select pd05 from patent,pridate where pa01='" & arrTemp(0) & "' and pa02='" & arrTemp(1) & "' and pa03='" & arrTemp(2) & "' and pa04='" & arrTemp(3) & "'" & _
'                     " and pd01(+)=pa01 and pd02(+)=pa02 and pd03(+)=pa03 and pd04(+)=pa04 and pd05<20130316 and rownum<2"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     bolAsk = True
'                  End If
'               End If
'               If bolAsk = True Then
'                  strExc(0) = "This application (1) claims priority to or the benefit of an" & vbCrLf & _
'                     " application filed before March 16, 2013 and (2) also contains," & vbCrLf & _
'                     " or contained at any time, a claim to a claimed invention that" & vbCrLf & _
'                     " has an effective filing date on or after March16, 2013."
'
'                  If MsgBox(strExc(0), vbYesNo, "美專案件FITF之控制！( 請選擇是或否 )") = vbYes Then
'                     strExc(0) = Replace(strExc(0), vbCrLf, "") & vbCrLf & "Y▓ N□"
'                  Else
'                     strExc(0) = Replace(strExc(0), vbCrLf, "") & vbCrLf & "Y□ N▓"
'                  End If
'                  frm090201_2_2.txt1(3) = strExc(0) & vbCrLf & frm090201_2_2.txt1(3)
'               End If
'            End If
'            'end 2013/4/16
'      Else
'            'add by nickc 2007/02/26
'            '抓商品類別
'            oTM12 = ""
'            oTM15 = "'"
'            strSql = "select tm09,tm12,tm15," & SQLDate("tm21") & "||'-'||" & SQLDate("tm22") & " from trademark,customer where tm01='" & SystemNumber(lbl1(7), 1) & "' and tm02='" & SystemNumber(lbl1(7), 2) & "' and tm03='" & SystemNumber(lbl1(7), 3) & "' and tm04='" & SystemNumber(lbl1(7), 4) & "' and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
'            strSql = strSql & " union select sp73,'',''," & SQLDate("sp20") & "||'-'||" & SQLDate("sp21") & " from servicepractice,customer where sp01='" & SystemNumber(lbl1(7), 1) & "' and sp02='" & SystemNumber(lbl1(7), 2) & "' and sp03='" & SystemNumber(lbl1(7), 3) & "' and sp04='" & SystemNumber(lbl1(7), 4) & "' and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) "
'            CheckOC3
'            AdoRecordSet3.CursorLocation = adUseClient
'            AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            CUID = ""
'            If AdoRecordSet3.RecordCount <> 0 Then
'               AdoRecordSet3.MoveFirst
'               CUID = CheckStr(AdoRecordSet3.Fields(0).Value)
'               oTM12 = CheckStr(AdoRecordSet3.Fields(1).Value)
'               oTM15 = CheckStr(AdoRecordSet3.Fields(2).Value)
'               oTM2122 = CheckStr(AdoRecordSet3.Fields(3).Value)
'            End If
'            CheckOC3
'            'edit by nickc 2007/12/24 若申請國家非台灣，則空白
'            If m_Country = "000" Then
'                '商標 add by Toni 2008/10/23
'                frm090201_2_3.txt1(5).Text = "經濟部智慧財產局"
'                '2009/1/16 add by sonia 改抓案件國家收費表,預設值原為智慧局同時改為全名
'                strSql = "select cf10 from casefee where cf01='" & SystemNumber(lbl1(7), 1) & "' and cf02='000' and cf03='" & m_CP10 & "'"
'                AdoRecordSet3.CursorLocation = adUseClient
'                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                If AdoRecordSet3.RecordCount <> 0 Then
'                   If "" & CheckStr(AdoRecordSet3.Fields(0).Value) <> "" Then
'                      frm090201_2_3.txt1(5).Text = CheckStr(AdoRecordSet3.Fields(0).Value)
'                   End If
'                End If
'                CheckOC3
'                '2009/1/16 end
'                frm090201_2_3.txt1(5).Enabled = False
'                'end 2008/10/23
'            Else
'                '商標 add by Toni 2008/10/23
'                frm090201_2_3.txt1(5).Text = ""
'                frm090201_2_3.txt1(5).Enabled = True
'                'end 2008/10/23
'                '2009/3/10 add by sonia 林副理說大陸案加預設代理人名稱
'                If m_CP44 <> "" Then
'                   If PUB_GetAgentName(SystemNumber(lbl1(7), 1), m_CP44, strTempName) Then
'                      frm090201_2_3.txt1(5).Text = strTempName
'                   Else
'                      frm090201_2_3.txt1(5).Text = ""
'                   End If
'                   frm090201_2_3.txt1(5).Enabled = False
'                   frm090201_2_3.txt1(8).Text = "掛號"  '2009/5/8 add by sonia 林副理需求
'                End If
'                '2009/3/10 end
'            End If
'            '商標 add by toni 2008/10/23
'            frm090201_2_3.txt1(6).Text = "北所、" & IIf(m_SaleArea = "1", "", IIf(m_SaleArea = "2", "中所、", IIf(m_SaleArea = "3", "南所、", IIf(m_SaleArea = "4", "高所、", "")))) & "客戶"
'            frm090201_2_3.txt1(6).Enabled = False
'            '2009/2/11 MODIFY BY SONIA 加審定號或申請案號
'            'frm090201_2_3.txt1(0) = "「" & m_CaseName & "」" & lbl1(15)
'            If oTM15 = "" Then
'               frm090201_2_3.txt1(0) = "「" & m_CaseName & "」" & lbl1(15) & "(申請案號 " & oTM12 & ")"
'            Else
'               frm090201_2_3.txt1(0) = "「" & m_CaseName & "」" & lbl1(15) & "(註冊號 " & oTM15 & ")"
'            End If
'            '2009/2/11 END
'            frm090201_2_3.txt1(0).Enabled = False
'            frm090201_2_3.txt1(1).Text = CUID         '商品類別
'            frm090201_2_3.txt1(1).Enabled = False
'            frm090201_2_3.Label2.Caption = "商品(服務)類別："
'
'            If m_CP10 = "301" Then
'                If oTM15 = "" Then '申請中    ◎■□『』□■○●
'                  '商標 add by toni 2008/10/23
'                  frm090201_2_3.txt1(3).Text = "◎申請第 " & oTM12 & " 號『" & m_CaseName & "』□商標" & vbCrLf & _
'                                                                   "◎變更事項：" & vbCrLf & _
'                                                                   "　□申請人名稱　□代表人或負責人　　　□代理人印鑑" & vbCrLf & _
'                                                                   "　□申請人印鑑　□代表人或負責人印鑑　□代理人地址" & vbCrLf & _
'                                                                   "　□申請人地址　□代理人"
'                'end 2008/10/13
'                Else
'                  '商標 add by toni 2008/10/23
'                  frm090201_2_3.txt1(3).Text = "◎註冊第 " & oTM15 & " 號『" & m_CaseName & "』□商標(前服務標章)" & vbCrLf & _
'                                                                   "◎變更□商標(標章)權人" & vbCrLf & _
'                                                                   "◎變更事項：" & vbCrLf & _
'                                                                   "　□申請人中文名稱　□申請人英文名稱　　□申請人印章　　　□申請人中文地址" & vbCrLf & _
'                                                                   "　□申請人英文地址　□代表人印章　　　　□代表人中文名稱　□代表人英文名稱" & vbCrLf & _
'                                                                   "　□代理人異動：□變更、□新增、□撤銷　　□變更商標/標章名稱"
'                'end 2008/10/13
'                End If
'            ElseIf m_CP10 = "101" Then
'                    '商標 add by toni 2008/10/23
'                    frm090201_2_3.txt1(3).Text = "◎中文：" & vbCrLf & _
'                                                                   "◎外文：" & vbCrLf & _
'                                                                   "　字義：" & vbCrLf & _
'                                                                   "◎圖形說明："
'                    frm090201_2_3.txt1(1).Enabled = True
'                    frm090201_2_3.txt1(0).Enabled = True
'                     'end 2008/10/23
'            ElseIf m_CP10 = "102" Then
'                  '商標 add by Toni 2008/10/23
'                                  'Modify By Sindy 2010/02/01
'                  '                  frm090201_2_3.txt1(3).Text = "◎註冊第 " & oTM15 & " 號『" & m_CaseName & "』□商標　□商標(前服務標章)" & vbCrLf & _
''                                                                   "◎原註冊期間：" & oTM2122 & vbCrLf & _
''                                                                   "◎變更事項：" & vbCrLf & _
''                                                                   "　□申請人名稱　□代表人或負責人　　　□申請人地址" & vbCrLf & _
''                                                                   "　□申請人印鑑　□代表人或負責人印鑑　□防護商標/標章變更為商標" & vbCrLf & _
''                                                                   "　□代理人異動：□變更、□新增、□撤銷　　□變更商標/標章名稱" & vbCrLf & _
''                                                                   "◎延展商品：□全部　□部份　延展"
''                  frm090201_2_3.txt1(3).Text = "◎註冊第 " & oTM15 & " 號『" & m_CaseName & "』□商標　□商標(前服務標章)" & vbCrLf & _
''                                                                   "◎原註冊期間：" & oTM2122 & vbCrLf & _
''                                                                   "◎變更事項：" & vbCrLf & _
''                                                                   "　□變更商標/標章名稱：" & vbCrLf & _
''                                                                   "　□防護商標/標章變更為商標" & vbCrLf & _
''                                                                   "　□代理人異動：□變更、□新增、□撤銷" & vbCrLf & _
''                                                                   "◎延展商品：□全部　□部份　延展"
'                  'Modify By Sindy 2012/3/8 原延展商品改為系統註記變更事項
'                  frm090201_2_3.txt1(3).Text = "◎註冊第 " & oTM15 & " 號『" & m_CaseName & "』□商標　□商標(前服務標章)" & vbCrLf & _
'                                                                   "◎原註冊期間：" & oTM2122 & vbCrLf & _
'                                                                   "◎變更事項：" & vbCrLf & _
'                                                                   "　□變更商標/標章名稱：" & vbCrLf & _
'                                                                   "　□防護商標/標章變更為商標" & vbCrLf & _
'                                                                   "　□代理人異動：□變更、□新增、□撤銷" & vbCrLf & _
'                                                                   "◎系統註記變更事項：" & vbCrLf & _
'                                                                   "　□申請人中文名稱　□申請人英文名稱　□申請人印鑑　□申請人地址" & vbCrLf & _
'                                                                   "　□代表人中文名稱　□代表人英文名稱　□代表人印鑑"
'                  'end 2008/10/23
'            Else
'                '商標 add by Toni 2008/10/23
'                frm090201_2_3.txt1(3).Text = ""
'                'end 2008/10/23
'            End If
'            '商標 add by Toni 2008/10/23
'            frm090201_2_3.oStrA02 = lbl1(7)
'            'end 2008/10/23
'            If lbl1(17) <> "" Then
'               '商標 add by Toni 2008/10/23
'               frm090201_2_3.oStrA05 = "   " & Mid(Replace(lbl1(17), "/", ""), 1, Len(Replace(lbl1(17), "/", "")) - 4) & "年  " & Left(Right(Replace(lbl1(17), "/", ""), 4), 2) & "月  " & Right(Replace(lbl1(17), "/", ""), 2) & "日"
'               'end 2008/10/23
'            Else
'               '商標 add by Toni 2008/10/23
'               frm090201_2_3.oStrA05 = "     年    月    日"
'               'end 2008/10/23
'            End If
'            If lbl1(19) <> "" Then
'               '商標 add by Toni 2008/10/23
'              frm090201_2_3.oStrA06 = "   " & Mid(Replace(lbl1(19), "/", ""), 1, Len(Replace(lbl1(19), "/", "")) - 4) & "年  " & Left(Right(Replace(lbl1(19), "/", ""), 4), 2) & "月  " & Right(Replace(lbl1(19), "/", ""), 2) & "日"
'              'end 2008/10/23
'            Else
'               '商標 add by Toni 2008/10/26
'               frm090201_2_3.oStrA06 = "     年    月    日"
'               'end 2008/10/26
'            End If
'            '商標 add by Toni 2008/10/23
'            frm090201_2_3.oStrA08 = lbl1(3) '收文號
'            frm090201_2_3.oStrA09 = lbl1(23)
'            frm090201_2_3.txt1(4).Text = "附委任狀正本" & vbCrLf & "附委任狀影本" & vbCrLf & "正本參     卷"
'            frm090201_2_3.oStrA10 = ""
'            frm090201_2_3.oStrA11 = ""
'            frm090201_2_3.oStrA12 = ""
'            frm090201_2_3.oStrA13 = ""
'            frm090201_2_3.oStrA14 = PUB_GetST07(m_CP14)
'            'end  2008/10/23
'      End If
'      If InStr(1, SystemNumber(lbl1(7), 1), "T") = 0 Then
'         frm090201_2_2.Show vbModal
'      Else
'         '商標 add by toni 2008/10/23
'         '2012/5/15 MODIFY BY SONIA
'         'If frm090201_2.txt1(1) = "" Then
'         If frm090201_2.txt1(1) = "" Or frm090201_2.txt1(1) = "Y" Then
'         '2012/5/15 EMD
'            '2010/5/24 modify by sonia 智權人員前加所別
'            'frm090201_2_3.txt1(9) = frm090201_2.lbl1(21)
'            frm090201_2_3.txt1(9) = ""
'            Select Case PUB_GetST06(m_CP13)
'               Case "2"
'                  frm090201_2_3.txt1(9) = "中所　"
'               Case "3"
'                  frm090201_2_3.txt1(9) = "南所　"
'               Case "4"
'                  frm090201_2_3.txt1(9) = "高所　"
'            End Select
'            frm090201_2_3.txt1(9) = frm090201_2_3.txt1(9) & frm090201_2.lbl1(21)
'            '2010/5/24 end
'            frm090201_2_3.oStrA15 = frm090201_2_3.txt1(9)
'         End If
'         frm090201_2_3.Show vbModal
'         'end 2008/10/23
'      End If
      
'2008/7/31 add by sonia C類來函才可用撰寫信函按鈕
Case 3  '撰寫信函

   'Added by Morgan 2021/2/3
   ReDim pa(1 To TF_PA) As String
   '專利基本檔
   pa(1) = SystemNumber(LBL1(7).Caption, 1)
   pa(2) = SystemNumber(LBL1(7).Caption, 2)
   pa(3) = SystemNumber(LBL1(7).Caption, 3)
   pa(4) = SystemNumber(LBL1(7).Caption, 4)
   Call ClsPDReadPatentDatabase(pa(), 國內)
   'end 2021/2/3
   
   
   'Added by Morgan 2021/2/24
   bolSysLtr = False
   'Modified by Morgan 2023/5/10 +1201 通知修正,1002 核駁 --郭
   'Removed by Morgan 2023/6/21 無特殊定稿時才抓程序定稿 Ex:P-127508(核駁-舉發)
   'If pa(1) = "P" And (m_CP10 = "1002" Or m_CP10 = "1201" Or m_CP10 = "1202") And pa(75) <> "Y55435" Then
   '   strExc(0) = "select * From letterdemand where ld18='" & lbl1(3) & "'"
   '   intI = 1
   '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '   If intI = 1 Then
   '      With RsTemp
   '      g_UserNum = strUserNum
   '      strUserNum = .Fields("LD01")
   '      NowPrint .Fields("LD04"), .Fields("LD10"), .Fields("LD11"), True, strUserNum, , , , , , , , , False, , , , .Fields("LD18"), , , , Me.Name
   '      strUserNum = strUser1Num
   '      g_UserNum = ""
   '      End With
   '      bolSysLtr = True
   '   End If
   'End If
   'If bolSysLtr = False Then
   'end 2021/2/24
   
      'Add by Morgan 2009/12/3 大陸檢索報告核准要用1209抓
      strExc(1) = m_CP10
      If m_CP10 = "1001" And m_CP43 <> "" Then
         strExc(0) = "select 1 from caseprogress where cp09='" & m_CP43 & "' and cp10='421'" & _
            " and exists(select * from patent where pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and pa09='020')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "1209"
         End If
      End If
      'strLetterCP10 = strExc(1) 'Added by Morgan 2017/1/23
      
      Call Forms(0).SetTmpfrm090401 'Add By Sindy 2015/7/15
      'Add by Morgan 2008/8/6
      'Modify By Sindy 2015/7/15
      'With frm090401
      With Tmpfrm090401
      '2015/7/15 END
         .Hide
         .Text1 = SystemNumber(LBL1(7).Caption, 1)
         .Text2 = SystemNumber(LBL1(7).Caption, 2)
         .Text3 = SystemNumber(LBL1(7).Caption, 3)
         .Text4 = SystemNumber(LBL1(7).Caption, 4)
         
         'Added by Morgan 2021/2/3寶齡富錦 Y55435 特殊控制
         If pa(75) = "Y55435" Then
            .Option1(1).Value = True '點選英文
            .Option2.Value = True '讀取案件資料
            .Option2.Value = True '點選FC代理人
            .OutCallCP09 = LBL1(3)
            .Command1.Value = True
         Else
         'end 2021/2/3
         
            .Option1(0).Value = True '點選中文
            .Option3.Value = True '讀取案件資料
            .Option3.Value = True '點選申請人
            For intI = 0 To .Combo8.ListCount
               'Modified by Morgan 2017/1/23 strExc(1)會被使用而改變
               If InStr(.Combo8.List(intI), strExc(1)) > 0 Then
               'If InStr(.Combo8.List(intI), strLetterCP10) > 0 Then
               'end 2017/1/23
                  .Combo8.ListIndex = intI
                  Exit For
               End If
            Next
         
            'Added by Morgan 2016/8/30
            '有特殊定稿但案件性質沒有對到時帶特殊定稿--游經理
            If .Combo8.ListIndex = 0 And .Combo8.ListCount > 1 Then
               .Combo8.ListIndex = 1
            End If
            'end 2016/8/30
            
            'Modified by Morgan 2023/6/21 若無特殊定稿時才抓程序定稿
            '.Command1.Value = True
            bolSysLtr = False
            If pa(1) = "P" And (m_CP10 = "1002" Or m_CP10 = "1201" Or m_CP10 = "1202") Then
               If .Combo8.ListIndex = 0 Then
                  strExc(0) = "select * From letterdemand where ld18='" & LBL1(3) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     With RsTemp
                     g_UserNum = strUserNum
                     strUserNum = .Fields("LD01")
                     NowPrint .Fields("LD04"), .Fields("LD10"), .Fields("LD11"), True, strUserNum, , , , , , , , , False, , , , .Fields("LD18"), , , , Me.Name
                     strUserNum = strUser1Num
                     g_UserNum = ""
                     End With
                     bolSysLtr = True
                  End If
               End If
               '沒有程序定稿時將案件回覆單附加到撰寫信函的定稿後面
               If bolSysLtr = False Then
                  strExc(0) = "select * From nextprogress where np01='" & LBL1(3) & "' and np06 is null"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     Call g_PrtForm001.PrintReturnSheet(RsTemp("np01"), RsTemp("np07"), RsTemp("np09"), , True, strReturnSheet, , RsTemp("np02") & RsTemp("np03") & RsTemp("np04") & RsTemp("np05"), , , , Me)
                     .strReturnSheet = strReturnSheet
                  End If
               End If
            End If
            If bolSysLtr = False Then
               .Command1.Value = True
            End If
            'end 2023/6/21
         End If
         
         .Command2.Value = True
      End With
      Set Tmpfrm090401 = Nothing 'Add By Sindy 2015/7/15
      'end 2008/8/6
   
   'End If 'Added by Morgan 2021/2/24 'Removed by Morgan 2023/6/21
   
'2008/7/31 END
'Add By Sindy 2020/3/17
Case 4 '未完稿暫存區
   'Call PUB_ChkFormIsClose("frm100101_M")
   frm100101_M.m_strKey = LBL1(3).Caption '總收文號
   frm100101_M.SetParent Me
   If frm100101_M.QueryData = True Then
      frm100101_M.Show
      Me.Hide
   End If
'2020/3/17 END

'Added by Morgan 2020/12/28
Case 5 '美專IDS清單
   Set nFrm = Forms(0).GetForm("frm090401_1")
   If Not nFrm Is Nothing Then
      nFrm.m_CP09 = m_strCP09
      nFrm.m_bMan = IIf(ProState = "2", True, False)
      nFrm.Show vbModal
   End If
'end 2020/12/28

Case Else
End Select

Exit Sub

ErrHnd:
   If Err.Number <> 0 Then
      If Err.Number = 52 Then '不正確的檔案名稱或號碼 (錯誤 52)
         MsgBox "請檢查是否無資料夾權限或網路不通!!!" & vbCrLf & vbCrLf & "(" & Err.Description & ")" & _
                IIf(m_strErrPath <> "", vbCrLf & vbCrLf & m_strErrPath, ""), vbCritical
      Else
         MsgBox Err.Description, vbCritical
      End If
   End If
End Sub

'Modify By Sindy 2013/6/10
'Sub Process(strText As String)
Public Sub Process(strText As String)
'2013/6/10 End
Dim arrCaseNo '本所案號
Dim stVTB As String
Dim oLbl As Object
Dim oTxt1 As Object
Dim strPA158 As String
Dim strEP38 As String 'Add By Sindy 2013/8/12
Dim strRefEEP02 As String 'Add By Sindy 2013/9/18
Dim objText As TextBox 'Add By Sindy 2015/5/21
Dim strEP40 As String 'Add By Sindy 2015/5/21
Dim ii As Integer 'Add By Sindy 2015/5/22
Dim blnMatch As Boolean 'Add By Sindy 2015/5/22
Dim strST16 As String
   
   'add by nickc 2007/11/29
   Chk1.Value = vbUnchecked
   cmd(0).Visible = False 'Add By Sindy 2018/4/10
   m_bol203Case = False 'Added by Morgan 2013/8/1
   
   cmdFAmend.Visible = False 'Added by Lydia 2016/12/30
   
   'Modify by Morgan 2008/10/27 +CPM05
   '2009/3/10 modify by sonia 加cp44欄
   '2009/9/18 MODIFY BY SONIA DECODE(PA09,'000',PTM03,PTM04) 改為 DECODE(PA09,'020',PTM04,PTM03)
   'Modify by Morgan 2009/11/12 合併其他語法
   '2009/11/18 MODIFY BY SONIA 因核准要抓相關總文號的案件性質故加CP43 (P-083933)
   'Modified by Morgan 2012/7/23 +CP147
   'Modified by Morgan 2012/8/24 +PA158
   'Modified by Morgan 2013/5/7 +CP118
   'Modified by Sindy 2015/1/20 +PA08
   'Modified by Sindy 2016/3/9 +cp144
   'Modified by Lydia 2025/02/05 +cp142
   stVTB = " SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
            ",NVL(PA05,NVL(PA06,PA07)) C10,DECODE(PA09,'020',PTM04,PTM03) C14,'' C26,'' C28,PA57,'*' C33,pa09 as m_country,pa26 as cuno,CP43,CP147,PA158,CP118,PA08,cp144,pa27 as cuno2,pa28 as cuno3,pa29 as cuno4,pa30 as cuno5,cp142" & _
            " FROM CASEPROGRESS,PATENT,PATENTTRADEMARKMAP WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr1 & ") AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PTM01(+)='1' AND PTM02(+)=PA08"
   stVTB = stVTB + " UNION all SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
            ",NVL(SP05,NVL(SP06,SP07)) C10,'' C14,'' C26,'' C28,SP15,'*' C33,sp09 as m_country,sp08 as cuno,CP43,CP147,'' pa158,CP118,'' PA08,cp144,sp58 as cuno2,sp59 as cuno3,sp65 as cuno4,sp66 as cuno5,cp142" & _
            " FROM CASEPROGRESS,SERVICEPRACTICE WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr5 & ") AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04"
'   stVTB = stVTB + " UNION all SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
'            ",NVL(TM05,NVL(TM06,TM07)) C10,decode(tm10,'000',ptm03,ptm04) C14,'' C26,'' C28,TM29,'*' C33,tm10 as m_country,tm23 as cuno,CP43,CP147,'' PA158,CP118,TM08" & _
'            " FROM CASEPROGRESS,TRADEMARK,PATENTTRADEMARKMAP WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr2 & ") AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND PTM01(+)='2' AND PTM02(+)=TM08"
'   stVTB = stVTB + " UNION all SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
'            ",NVL(LC05,NVL(LC06,LC07)) C10,'' C14,'' C26,'' C28,LC08,'*' C33,lc15 as m_country,lc11 as cuno,CP43,CP147,'' pa158,CP118,'' PA08" & _
'            " FROM CASEPROGRESS,LAWCASE WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr3 & ") AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04"
'   stVTB = stVTB + " UNION all SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
'            ",HC06 C10,'' C14,'' C26,'' C28,HC09,'*' C33,'000' as m_country,hc05 as cuno,CP43,CP147,'' pa158,CP118,'' PA08" & _
'            " FROM CASEPROGRESS,HIRECASE WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr4 & ") AND HC01(+)=CP01 AND HC02(+)=CP02 AND HC03(+)=CP03 AND HC04(+)=CP04"
      
   '2009/11/18 MODIFY BY SONIA 因核准要抓相關總文號的案件性質故加CP43 (P-083933)
   'Modify By Sindy 2013/8/12 +,EP38
   'Modify By Sindy 2013/9/3 +,pp05,EP40
   'Modify By Sindy 2013/9/18 +cpm28,cpm29
   'Modify By Sindy 2015/1/20 +PA08,cpm23
   'Modify By Sindy 2015/3/13 +EP41
   'Modified by Sindy 2016/3/9 +cp144
   'Modify By Sindy 2017/7/10 +,EP39
   'Modified by Lydia 2025/02/05 +cp142
   strSql = "SELECT EP01,S1.ST02 C2,sqldateT(CP48) C3,CP09,EP13,sqldateT(cp05) C6,EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04 C8" & _
      ",EP06,C10,EP09,CP26,EP07,C14,EP04,decode(na01,'000',cpm03,cpm04) C16,EP03,sqldateT(CP06) C18,EP08,sqldateT(CP07) C20,CP27" & _
      ",S5.ST02 C22,EP11,CP18,EP12,C26,Nvl(EP35,0) C27,C28,sqldateT(CP57) C29,CP10,CP15,PA57,C33,EP27,EP31,cp13,ep05,m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99" & _
      ",cp106,cuno,cp111,cp112,ep28,ep32,ep33,na03,cp64,cpm05,cp44,ibf01,S3.ST02 EP04N,pp04,s6.st02 pp04N,s2.st02 EP13N,s4.st02 EP03N" & _
      ",NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CuName,CP43,CP147,pa158,CP118,EP38,EP39,pp05,EP40,cpm28,cpm29,PA08,cpm23,EP41,cp144,cp142" & _
      " from (" & stVTB & ") X,ENGINEERPROGRESS,CASEPROPERTYMAP,nation" & _
      ",STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,customer,imgbytefile,promoterproofreader,staff S6" & _
      " where EP02(+)=CP09 and cpm01(+)=CP01 and cpm02(+)=CP10 AND na01(+)=m_country" & _
      " AND S1.ST01(+)=EP05 AND S2.ST01(+)=EP13 AND S3.ST01(+)=EP04 AND S4.ST01(+)=EP03 AND S5.ST01(+)=CP13" & _
      " and cu01(+)=substr(cuno,1,8) and cu02(+)=substr(cuno,9) and pp01(+)=cp01 and pp02(+)=cp14 and pp03(+)=cp10 and s6.st01(+)=pp04" & _
      " and ibf01(+)=cp01 and ibf02(+)=cp02 and ibf03(+)=cp03 and ibf04(+)=cp04 and ibf05(+)='1'"
      
   CheckOC
   With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   '***** 清除欄位值 *****
   'Modify by Morgan 2010/10/11
   'For i = 0 To 29
   '    If i <> 12 And i <> 20 And i <> 22 And i <> 24 Then
   '       lbl1(i) = ""
   '    End If
   'Next i
   ''Add By Cheng 2002/03/27
   'Me.lbl1(31).Caption = ""
   For Each oLbl In LBL1
      oLbl.Caption = ""
      'Debug.Print olbl.Index
   Next
   'Add By Cheng 2002/04/29
   Me.lblClose.Caption = ""
   'For i = 0 To 14
   For Each oTxt1 In txt1
      oTxt1.Text = ""
   Next
   txtCP64.Text = ""
   '***** END *****
   
   strEP38 = "" 'Add By Sindy 2013/8/12
   m_CPM28 = "" 'Add By Sindy 2013/9/18
   m_CPM29 = "" 'Add By Sindy 2013/9/30
   m_EP39 = "" 'Add By Sindy 2017/7/10 '核稿完成日
   If .RecordCount <> 0 And .RecordCount > 0 Then
      .MoveFirst
      m_CP43 = "" & .Fields("CP43") 'Added by Morgan 2009/12/3
      'Modified by Morgan 2016/7/7 代主任(52)以上可設定複雜案件
      'lbl1(24) = "" & .Fields("CP147") 'Added by Morgan 2012/7/23
      txt1(11) = "" & .Fields("CP147")
      'end 2016/7/7
      strPA158 = "" & .Fields("PA158") 'Added by Morgan 2012/8/24
      strEP38 = "" & .Fields("EP38") 'Add By Sindy 2013/8/12
      m_EP39 = "" & .Fields("EP39") 'Add By Sindy 2017/7/10 '核稿完成日
      m_CPM28 = "" & .Fields("CPM28") 'Add By Sindy 2013/9/18
      m_CPM29 = "" & .Fields("CPM29") 'Add By Sindy 2013/9/30
      
      txtCP144.Text = "" & .Fields("cp144") 'Add By Sindy 2016/3/9 +報價備註
      
      'Added by Morgan 2013/5/7 電子送件
      If Not IsNull(.Fields("CP118")) Then
         lblEApp.Visible = True
      Else
         lblEApp.Visible = False
      End If
      'end 2013/5/7
      LBL1(12).Caption = ChangeWStringToTDateString("" & .Fields("CP142")) 'Added by Lydia 2025/02/05
      
      For i = 0 To 29
         'Modify by Morgan 2008/10/13 原來值由lablel改為放tag或text
         '會稿日
         If i = 12 Then
            txt1(4).Tag = ChangeWStringToTString(CheckStr(.Fields(i)))
            txt1(4).Text = txt1(4).Tag
         '會稿完成日
         ElseIf i = 18 Then
            txt1(7).Tag = ChangeWStringToTString(CheckStr(.Fields(i)))
            txt1(7).Text = txt1(7).Tag
         '是否暫停核稿
         ElseIf i = 20 Then
            txt1(8).Text = ChangeWStringToTString(CheckStr(.Fields(i)))
         '是否通知客戶
         ElseIf i = 22 Then
            txt1(9).Text = CheckStr(.Fields(i))
         '承辦備註
         ElseIf i = 24 Then
            txtEP12.Text = CheckStr(.Fields(i))
         'end 2008/10/13
         
         'Add by Morgan 2010/10/7
         '承辦期限
         ElseIf i = 2 Then
            txt1(12).Text = ChangeTDateStringToTString(CheckStr(.Fields(i)))
            
         '2009/11/18 ADD BY SONIA
         ElseIf i = 15 Then
            If Not IsNull(.Fields("CP43")) Then
               LBL1(i) = CheckStr(.Fields(i)) & PUB_GetRelateCasePropertyName(strText, "1")
            Else
               LBL1(i) = CheckStr(.Fields(i))
            End If
         '2009/11/18 END
         
         Else
            If i <> 25 And i <> 27 Then
               LBL1(i) = CheckStr(.Fields(i))
            End If
         End If
      Next i
      
      'Add By Sindy 2015/3/13 一案兩請
      strExc(1) = SystemNumber(LBL1(7).Caption, 1)
      strExc(2) = SystemNumber(LBL1(7).Caption, 2)
      strExc(3) = SystemNumber(LBL1(7).Caption, 3)
      strExc(4) = SystemNumber(LBL1(7).Caption, 4)
      strSql = "select * from casemap where cm01='" & strExc(1) & "' and cm02='" & strExc(2) & "' and cm03='" & strExc(3) & "' and cm04='" & strExc(4) & "' and cm10='3'" & _
               " Union select * from casemap where cm05='" & strExc(1) & "' and cm06='" & strExc(2) & "' and cm07='" & strExc(3) & "' and cm08='" & strExc(4) & "' and cm10='3'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         lblCM10.Visible = True
      Else
         lblCM10.Visible = False
      End If
      '2015/3/13 END
      
      'Added by Lydia 2016/06/14 +台灣大陸案件提示
      lblCMboth.Caption = ""
      If (strExc(1) = "P" Or strExc(1) = "FCP") And "" & .Fields("m_country") = 台灣國家代號 Then
         If PUB_GetRefCaseChk(strExc(1), strExc(2), strExc(3), strExc(4), "CASEMAP", "0", "A", 大陸國家代號) Then
            lblCMboth.Caption = "有大陸案"
         End If
      ElseIf strExc(1) = "P" And "" & .Fields("m_country") = 大陸國家代號 Then
         If PUB_GetRefCaseChk(strExc(1), strExc(2), strExc(3), strExc(4), "CASEMAP", "0", "A", 台灣國家代號) Then
            lblCMboth.Caption = "有台灣案"
         End If
      End If
      'end 2016/06/14

      'Add By Sindy 2018/4/10
      'Modify By Sindy 2019/8/13 + 203,204,205
      'Modify By Sindy 2019/8/29 + 107:為延期再審,出修正申請書
      'Modify By Sindy 2023/5/9 + 307:分割
      'Modified by Morgan 2023/5/25 +244補中文說明書,232補優先權證明
      'Modified by Morgan 2023/8/24 +239擇一申復
      If (LBL1(29) = "101" Or LBL1(29) = "102" Or LBL1(29) = "103" Or _
          LBL1(29) = "125" Or LBL1(29) = "203" Or LBL1(29) = "204" Or _
          LBL1(29) = "205" Or LBL1(29) = "107" Or LBL1(29) = "307" Or _
          LBL1(29) = "244" Or LBL1(29) = "232" Or LBL1(29) = "239") And _
         "" & .Fields("m_country") = 台灣國家代號 Then
'      If (lbl1(29) = "101" Or lbl1(29) = "102" Or lbl1(29) = "103" Or _
'          lbl1(29) = "125") And _
'         "" & .Fields("m_country") = 台灣國家代號 Then
         cmd(0).Visible = True
      End If
      '2018/4/10 END
      
      'add by nickc 2006/01/23
      'Modify by Morgan 2009/11/12 改一次抓
      'm_CuNo = GetCustomerName(CheckStr(.Fields("cuno").Value))
      m_CuNo = "" & .Fields("CuName")
      
      'add by nick 2004/09/08 '紀錄智權人員
      m_CP13 = "" & .Fields("cp13").Value
      'add by nick 2004/10/21
      m_CP14 = "" & .Fields("ep05").Value
      m_CP10 = "" & .Fields("cp10").Value
      'add by nickc 2007/08/22
      m_NA03 = "" & .Fields("NA03").Value
      'Add by Morgan 2008/12/2
      m_CPM05 = "" & .Fields("cpm05")
      m_CP112 = "" & .Fields("cp112")
      'end 2008/12/2
      m_CP44 = "" & .Fields("cp44")     '2009/3/10 add by sonia
        
      'add by nickc 2007/01/17
      txtCP64 = CheckStr(.Fields("cp64"))
      'Modify by Morgan 2008/10/13 原預訂會稿日由label改為放在Tag
      txt1(18).Tag = ""
      'end 2008/10/13
      
      LBL1(34) = ""

'add by nickc 2006/02/07 發明 新型 的 P & CFP 才適用
'Modify by Morgan 2008/10/27 改判斷CPM05
         
         '2008/10/31 ADD BY SONIA 從上面搬下來
         'Modify by Morgan 2008/12/3 輸入時已有控制此處不必再判斷，若遇不該有卻有日期時應找出原因來修正。
         'If UCase(CheckStr(.Fields("cp112"))) = "Y" Then
'                'Modify by Morgan 2008/10/13 改放在Tag
'                'lbl1(33) = ChangeWStringToTString(CheckStr(.Fields("ep28")))
      txt1(18).Tag = ChangeWStringToTString(CheckStr(.Fields("ep28")))
'                'end 2008/10/13
      LBL1(34) = CheckStr(.Fields("cp111"))
         'End If
         '2008/10/31 END
         
'Remove by Morgan 2010/10/13
'      Select Case UCase(CheckStr(.Fields("cp112")))
'         Case "Y"
'             opCP112(0).Value = True
'         Case "N"
'             opCP112(1).Value = True
'         Case Else
'             opCP112(0).Value = False
'             opCP112(1).Value = False
'      End Select
'
'
'      '2008/10/31 MODIFY BY SONIA
'      'If ProState = "2" And "" & .Fields("cpm05") <> "" Then
'      'Modify by Morgan 2008/12/2 改判斷變數值(前次修改原因已不可考)
'      'If ProState = "2" And "" & Not IsNull(.Fields("cpm05")) Then
'      If ProState = "2" And m_CPM05 <> "" Then
'      'end 2008/12/2
'      '2008/10/31 END
'         Frame1.Enabled = True
'      Else
'         Frame1.Enabled = False
'      End If
'end 2010/10/13

'end 2008/10/27
        
      'add by nick 2005/01/27
      m_Country = "" & .Fields("m_country").Value
      m_CP31 = "" & .Fields("cp31").Value
      'add by nickc 2005/02/22  '紀錄案件名稱
      m_CaseName = "" & .Fields(9).Value
      m_SaleArea = "" & .Fields("Area").Value
      'add by nickc 2005/03/01
      m_CP107 = "" & .Fields("cp107").Value
      'add by nickc 2005/03/04
      LBL1(32).Caption = "" & .Fields("CP97").Value
      txt1(15).Text = "" & .Fields("cp98").Value
      m_CP98 = "" & .Fields("cp98").Value
      txtCP99.Text = "" & .Fields("cp99").Value
      m_CP99 = "" & .Fields("cp99").Value
      'add by nickc 2005/04/04
      txt1(17).Text = "" & .Fields("cp106").Value

      'Add By Cheng 2003/10/07
      '記錄收文號
      'Begin
      m_strCP09 = Me.LBL1(3).Caption
      'End
      If Len(Trim(CheckStr(.Fields(20)))) <> 0 Then
          ChkCp27 = False
      Else
          ChkCp27 = True
      End If
      
      'Add By Cheng 2003/05/09
      ChkCp27 = True
      'Add By Cheng 2002/04/29
      If IsNull(.Fields(31).Value) <> 0 Then
          Me.lblClose.Caption = ""
      Else
          Me.lblClose.Caption = "已閉卷"
      End If
      
      '92.7.1 ADD BY SONIA
      'Modify By Sindy 2019/12/3 專利程序承辦的案件,不管制收卷註記和齊備日
      'If IsNull(.Fields("EP27")) Then
      If IsNull(.Fields("EP27")) And PUB_GetST03(m_CP14) <> "P12" Then
      '2019/12/3 END
         Me.cmd(1).Enabled = False 'Add By Sindy 2013/6/7 沒有收卷註記不可做電子簽核
         m_ST03 = Pub_StrUserSt03 'Add by Morgan 2009/11/12
         If m_ST03 <> "P12" Then 'Add By Sindy 2013/9/13 程序不鎖
            Me.cmd(2).Enabled = False 'Modify By Sindy 2013/9/11
         Else
            Me.cmd(2).Enabled = True
         End If
         txt1(14) = ""
         m_EP27 = ""
         '92.8.8 add by sonia
         'For i = 0 To 20 'edit by nickc 2007/08/21    12
         For Each oTxt1 In txt1
            'm_ST03 = GetStaffDepartment(strUserNum) 'Remove by Morgan 2009/11/12 只要設一次移到上面做
            oTxt1.Enabled = False
         Next
         txt1(14).Enabled = True
         Combo4.Enabled = False
         '92.8.8 end
      Else
         Me.cmd(1).Enabled = True 'Add By Sindy 2013/6/7 有收卷註記才可做電子簽核
         Me.cmd(2).Enabled = True 'Modify By Sindy 2013/9/11
         'Add By Sindy 2019/12/3
         If Not IsNull(.Fields("EP27")) Then
         '2019/12/3 END
            txt1(14) = "Y"
            m_EP27 = .Fields("EP27")
         End If
         '92.8.8 add by sonia
         'For i = 0 To 20 'edit by nickc 2007/08/21    12
         For Each oTxt1 In txt1
            oTxt1.Enabled = True
         Next
'         'Modify By Sindy 2015/4/13 當從主管身份進入此作業時,檢查是否有”英文核稿人欄修改”的權限
'         If ProState = "2" Then '主管
'            strSql = "select EPA01 from EP14Authority where EPA01='" & strUserNum & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               Combo4.Enabled = True
'            Else
'               Combo4.Enabled = False
'            End If
'         Else
'         '2015/4/13 END
            Combo4.Enabled = True
'         End If
         '92.8.8 end
      End If
      
      If IsNull(.Fields("EP31")) Then
         txt1(13) = ""
      Else
         txt1(13) = ChangeWStringToTString(.Fields("EP31"))
      End If
      'add by nickc 2007/08/21
      m_EP33 = "" 'Add By Sindy 2013/12/18
      If IsNull(.Fields("EP33")) Then
         txt1(19) = ""
      Else
         txt1(19) = ChangeWStringToTString(.Fields("EP33"))
         m_EP33 = .Fields("EP33") 'Add By Sindy 2013/12/18
      End If
      If IsNull(.Fields("EP32")) Then
         txt1(20) = ""
      Else
         txt1(20) = .Fields("EP32")
      End If
      txt1(20).Tag = txt1(20).Text
      
'Remove by Morgan 2009/11/12 已經沒有顯示
'      '92.7.1 END
'      '**************************  end
'      '92.04.03 nick add left join
'      'strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01 AND CP02=CM02 AND CP03=CM03 AND CP04=CM04 AND CP14=ST01(+) AND CP31='Y' and cp09='" & StrText & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
'      strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01(+) AND CP02=CM02(+) AND CP03=CM03(+) AND CP04=CM04(+) AND CP14=ST01(+) AND CP31='Y' and cp09='" & strText & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
'      CheckOC2
'      adoRecordset1.CursorLocation = adUseClient
'      adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'          lbl1(25) = CheckStr(adoRecordset1.Fields(0))
'          lbl1(27) = CheckStr(adoRecordset1.Fields(1))
'      Else
'          lbl1(25) = ""
'          lbl1(27) = ""
'      End If
'end 2009/11/12

      'Modify by Morgan 2009/11/12 合併到最上面的語法
      'CheckOC2
      ''add by nickc 2005/12/14 檢查有無代表圖
      'strSQL = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(lbl1(7), 1) & "' and ibf02='" & SystemNumber(lbl1(7), 2) & "' and ibf03='" & SystemNumber(lbl1(7), 3) & "' and ibf04='" & SystemNumber(lbl1(7), 4) & "' and ibf05='1' "
      'CheckOC2
      'adoRecordset1.CursorLocation = adUseClient
      'adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
      'If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
      If Not IsNull(.Fields("ibf01")) Then
      'end 2009/11/12
          cmdPic.Caption = "已設定代表圖(&I)"
          cmdPic.BackColor = &HC0FFC0
          'add by nickc 2007/11/29 加入無圖式
          Chk1.Enabled = False
      Else
          cmdPic.Caption = "未設定代表圖(&I)"
          cmdPic.BackColor = &HC0C0FF
          'add by nickc 2007/11/29 加入無圖式
          Chk1.Enabled = True
      End If
      CheckOC2
        
      '計算承辦天數
      'Modify by Morgan 2011/8/1 支援時數,修改時數,衍生時數都要抓各自的Table
      'Me.lbl1(31).Caption = IIf(IsNull(.Fields("CP15")), "0", "" & .Fields("CP15"))
      'Modified by Morgan 2024/6/5 改與查詢共用
      'SettHour
      PUB_SettHour LBL1(3).Caption, strExc(1), strExc(2)
      LBL1(31) = strExc(1)
      LBL1(2) = strExc(2)
      'end 2024/6/5
      'end 2011/8/1
      
      'Add By Sindy 2015/3/13 顯示核稿語文
      m_EP41 = "" & .Fields("EP41")
      If "" & .Fields("EP41") = "" And PUB_ChkEmpFlowExists(LBL1(3), EMP_判發) = False Then '尚未判發時才預設核稿語文
         strSql = "select st61 from staff where st01='" & m_CP14 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            '英文核稿:非承辦日文案件的工程師，不管申請國家，都預設為英文稿，不可修改
            If "" & RsTemp.Fields("st61") = "" Then
               txt1(23) = "1"
               'Modify By Sindy 2019/5/29
               '主管權限進入且未會稿日,未有送英核流程者, 主管可修改
               If ProState = "2" And Val(txt1(4)) = 0 And PUB_ChkEmpFlowExists(LBL1(3), EMP_送英核) = False Then
                  txt1(23).Enabled = True
               Else
               '2019/5/29 END
                  txt1(23).Enabled = False
               End If
            '日文核稿:可承辦日文案件的工程師，若申請國家非日本，則預設為英文稿，不可修改
            Else
               If m_Country <> "011" Then
                  txt1(23) = "1"
                  'Modify By Sindy 2019/5/29
                  '主管權限進入且未會稿日,未有送英核流程者, 主管可修改
                  If ProState = "2" And Val(txt1(4)) = 0 And PUB_ChkEmpFlowExists(LBL1(3), EMP_送英核) = False Then
                     txt1(23).Enabled = True
                  Else
                  '2019/5/29 END
                     txt1(23).Enabled = False
                  End If
               Else
                  '申請國家為日本者，預設為日文稿，除902.回覆代理人,936.回覆委任代理人外，其他案件性質不可修改；
                  txt1(23) = "2"
                  If SystemNumber(Me.LBL1(7).Caption, 1) = "CFP" And (m_CP10 = "902" Or m_CP10 = "936") Then
                     txt1(23).Enabled = True
'                     '提醒工程師以免忘記修改，但仍可選擇為日文稿
'                     If MsgBox(m_CP10 & lbl1(15) & "，是否確定仍為日文稿？", vbYesNo + vbDefaultButton2) = vbNo Then
'                        txt1(23) = "1"
'                     End If
                  Else
                     'Modify By Sindy 2019/5/29
                     '主管權限進入且未會稿日,未有送英核流程者, 主管可修改
                     If ProState = "2" And Val(txt1(4)) = 0 And PUB_ChkEmpFlowExists(LBL1(3), EMP_送英核) = False Then
                        txt1(23).Enabled = True
                     Else
                     '2019/5/29 END
                        txt1(23).Enabled = False
                     End If
                  End If
               End If
            End If
         End If
      Else
         txt1(23) = "" & .Fields("EP41")
         'Modify By Sindy 2019/5/29
         '主管權限進入且未會稿日,未有送英核流程者, 主管可修改
         If ProState = "2" And _
            Val(txt1(4)) = 0 And _
            PUB_ChkEmpFlowExists(LBL1(3), EMP_送英核) = False And _
            PUB_ChkEmpFlowExists(LBL1(3), EMP_判發) = False Then
            txt1(23).Enabled = True
         Else
         '2019/5/29 END
            txt1(23).Enabled = False
         End If
         strSql = "select st61 from staff where st01='" & m_CP14 & "' and ST61='Y'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If SystemNumber(Me.LBL1(7).Caption, 1) = "CFP" And _
               m_Country = "011" And (m_CP10 = "902" Or m_CP10 = "936") And _
               PUB_ChkEmpFlowExists(LBL1(3), EMP_判發) = False Then
               txt1(23).Enabled = True
            End If
         End If
      End If
      SetEngChecker '設定核稿人選單
      '2015/3/13 END
      
      m_PP04 = "" & .Fields("pp04") 'Add By Sindy 2013/10/14 核判表設定的核稿人
      m_PP05 = "" & .Fields("pp05") 'Add By Sindy 2013/10/14 核判表設定的判發人
      'If m_pp04 = strUserNum Then m_pp04 = "" 'Add By Sindy 2014/4/11 為自行核稿,不需再將自己ID放入核稿人欄位
      If m_PP04 = Trim(Left("" & Combo1.Text, 6)) Then m_PP04 = "" 'Add By Sindy 2014/4/11 為自行核稿,不需再將自己ID放入核稿人欄位
      'If m_pp05 = strUserNum Then m_pp05 = "" 'Add By Sindy 2014/4/11 為自行判發,不需再將自己ID放入判發人欄位
      If m_PP05 = Trim(Left("" & Combo1.Text, 6)) Then m_PP05 = "" 'Add By Sindy 2014/4/11 為自行判發,不需再將自己ID放入判發人欄位
      
      'Add By Sindy 2015/3/4 送英核權限
      bolHadSetProofEngReader = PUB_ChkIsSetProofEngReader(m_CP14, _
                                SystemNumber(Me.LBL1(7).Caption, 1), SystemNumber(Me.LBL1(7).Caption, 2), _
                                SystemNumber(Me.LBL1(7).Caption, 3), SystemNumber(Me.LBL1(7).Caption, 4), _
                                m_CP10, m_PER04)
      If m_PER04 <> "" And m_PER04 = m_CP14 Then m_PER04 = "" '=自己,無須英核
      'Add By Sindy 2013/9/3 增加判發人
      Combo6.Clear: strEP40 = ""
      Combo6.AddItem "", 0
      'Modify By Sindy 2016/5/24 不用完稿日為預設判發的基準點,改用檢查有無送判或判發歷程
      'If Val("" & .Fields("EP09")) <= 0 And Len("" & .Fields("EP40")) = 0 Then 'Modify By Sindy 2013/10/4 +if 有完稿日時,則不用再預設判發人
      If PUB_ChkEmpFlowExists(LBL1(3), EMP_送判) = False And _
         PUB_ChkEmpFlowExists(LBL1(3), EMP_判發) = False And _
         Len("" & .Fields("EP40")) = 0 Then
      '2016/5/24 END
'         If "" & .Fields("pp05") <> "" Then
'            strEP40 = .Fields("pp05")
'         End If
         If m_PP05 <> "" Then
            strEP40 = m_PP05
         End If
         If SystemNumber(Me.LBL1(7).Caption, 1) = "P" Then
            '品薇的P-大陸案,必須給書慈核稿送判
            'Modify By Sindy 2014/4/17 + 056.PCT
'            If "" & .Fields("ep05") = "98012" And _
'               ("" & .Fields("m_country") = "013" Or _
'                "" & .Fields("m_country") = "020" Or _
'                "" & .Fields("m_country") = "056") Then
            'Modify By Sindy 2014/12/9 排除013香港
            'Modified by Morgan 2025/1/22 P程序人員工作改依智權區域分配，改判斷程序部們，若特定人需要判發時自行輸入--郭
            'If "" & .Fields("ep05") = "98012" And _
               ("" & .Fields("m_country") = "020" Or _
                "" & .Fields("m_country") = "056") Then
            If PUB_GetST03("" & .Fields("ep05")) = "P12" And _
               ("" & .Fields("m_country") = "020" Or _
                "" & .Fields("m_country") = "056") Then
               'Modify By Sindy 2014/5/20 將品薇承辦的大陸案之判發人改為99050李柏翰
               'strEP40 = "93003"
               'Modify By Sindy 2016/6/3 品薇承辦之大陸案件,無須預設判發人-李柏翰,由其自行判發。
               'strEP40 = "99050"
            'Modify By Sindy 2014/12/9 013香港,判發人郭雅娟
            'Modify By Sindy 2018/10/4 98012改判斷是P12專利處程序
'            ElseIf "" & .Fields("ep05") = "98012" And _
'               "" & .Fields("m_country") = "013" Then
            ElseIf PUB_GetST03("" & .Fields("ep05")) = "P12" And _
               "" & .Fields("m_country") = "013" Then
               strEP40 = "79075"
            '2014/12/9 END
            'P案非台灣承辦人為非程序人員者,判發人為游經理
            ElseIf GetStaffDepartment("" & .Fields("ep05")) <> "P12" And _
               "" & .Fields("m_country") <> "000" Then
               'Modify By Sindy 2016/6/3 工程師承辦之大陸申請案,案件性質101、102、103及936無須預設判發人為游經理,則是回歸核判表所設定之判發人。
               'Modify By Sindy 2025/5/26 增加958代理人撰稿，無須預設判發人為林協理，回歸核判表所設定之判發人
               'Modify By Sindy 2025/7/31 取消2025/5/26
               'If InStr("101,102,103,936,958", m_CP10) > 0 And Len(m_CP10) = 3 Then
               If InStr("101,102,103,936", m_CP10) > 0 And Len(m_CP10) = 3 Then
               '2025/7/31 END
                  strEP40 = m_PP05
               Else
               '2016/6/3 END
                  'Modified by Morgan 2025/2/21 73022->Left(pub_PMan, 5)
                  pub_PMan = Pub_GetSpecMan("專利處特定編號")
                  strEP40 = Left(pub_PMan, 5)
                  'end 2025/2/21
               End If
            'Added by Morgan 2021/2/2
            'P案C類來函承辦人為工程師的案件判發人預設為(73022)游經理--玲玲
            ElseIf Left(LBL1(3), 1) = "C" And GetStaffDepartment("" & .Fields("ep05")) <> "P12" Then
               
               'Modified by Morgan 2025/2/21 73022->Left(pub_PMan, 5)
               pub_PMan = Pub_GetSpecMan("專利處特定編號")
               strEP40 = Left(pub_PMan, 5)
               'end 2025/2/21
               
            'end 2021/2/2
            End If
         ElseIf SystemNumber(Me.LBL1(7).Caption, 1) = "CFP" Then
            'CFP專利外翻判發人
            If m_CP14 <> "" And Left(m_CP14, 1) = "F" Then
               'Modify By Sindy 2025/9/19 mark; ex:CFP-35421
'               'Add By Sindy 2017/11/27
'               If m_PP05 = "" Then
'               '2017/11/27 END
               '2025/9/19 END
                  'Add By Sindy 2025/4/7 品薇提,承辦人編號為F開頭,申請國家為日本
                  If m_Country = "011" Then
                     'Modify By Sindy 2025/5/22
                     'strEP40 = "99050" '判發主管為李柏翰經理
                     strEP40 = "B4009" '判發主管改為吳嘉智
                     '2025/5/22 END
                  Else
                  '2025/4/7 END
                     strEP40 = Pub_GetSpecMan("H2")
                  End If
'               End If
            End If
         End If
         'Modify By Sindy 2013/9/11 若承辦人等於預設判發人,則為自行判發,判發人應為空白
         'If Trim(m_pp05) = "" Then
            If Trim(Left("" & Combo1.Text, 6)) = Trim(strEP40) Then
               strEP40 = ""
            End If
         'End If
         '2013/9/11 END
      Else
         strEP40 = "" & .Fields("EP40")
      End If
      If strEP40 <> "" Then
         Combo6.AddItem UCase(strEP40) & " ==> " & GetPrjSalesNM(strEP40), 1
      End If
      '2013/9/3 END
      'Add By Sindy 2015/5/21 503.行政訴訟,506.參加訴訟,211.準備程序,212.言詞辯論,206.補充說明時,判發人要加入律師
      'Modify By Sindy 2020/10/29 + 243.補書狀
      If InStr("503,506,211,212,206,243", m_CP10) > 0 And Len(m_CP10) = 3 Then
         'modify by sonia 2019/2/13 +ST20='11',因桂所長的部門由L01改管理部
         'Modify By Sindy 2022/7/18 + 15.名譽所長
         strExc(0) = "SELECT st01||' ==> '||st02 FROM STAFF WHERE (ST03='L01' OR ST20='13' OR ST20='11' OR ST20='15') AND ST04='1' ORDER BY ST01"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            .MoveFirst
            Do While .EOF = False
               blnMatch = False
               For ii = 0 To Me.Combo6.ListCount - 1
                   blnMatch = False
                   If Trim(Left(Me.Combo6.List(ii), 6)) = Left(.Fields(0), 5) Then
                       Me.Combo6.ListIndex = ii
                       blnMatch = True
                       Exit For
                   End If
               Next ii
               If blnMatch = False Then
                  intI = Combo6.ListCount
                  Combo6.AddItem "" & .Fields(0), intI
               End If
               .MoveNext
            Loop
            End With
         End If
      End If
      blnMatch = False
      For ii = 0 To Me.Combo6.ListCount - 1
         If Trim(Left(Me.Combo6.List(ii), 6)) = UCase(strEP40) Then
             Me.Combo6.ListIndex = ii
             blnMatch = True
             Exit For
         End If
      Next ii
      If blnMatch = False Then Me.Combo6.ListIndex = 0
      Combo6.Tag = Combo6.Text
      '2015/5/21 END
      
      'Add by Morgan 2009/11/12
      If Val("" & .Fields("EP09")) <= 0 And Len("" & .Fields("EP04")) = 0 Then 'Modify By Sindy 2013/10/4 +if 有完稿日時,則不用再預設核稿人
         txt1(5) = m_PP04 '"" & .Fields("pp04")
         'lbl1(14).Caption = "" & .Fields("pp04N")
         'Add By Sindy 2013/8/27 CFP專利外翻核稿人
         If SystemNumber(Me.LBL1(7).Caption, 1) = "CFP" And m_CP14 <> "" Then
            If Left(m_CP14, 1) = "F" Then
               'Add By Sindy 2017/11/27
               If m_PP04 = "" Then
               '2017/11/27 END
                  'Add By Sindy 2025/4/7 品薇提,承辦人編號為F開頭,申請國家為日本
                  If m_Country = "011" Then
                     'Modify By Sindy 2025/5/22
                     'txt1(5) = "99050" '核稿主管為李柏翰經理
                     Call GetPrjSalesNM(m_CP14, , , "st16", strST16) '取得人員的組別
                     If Not (PUB_GetST03(m_CP14) = "F52" And strST16 = "3") Then '外專內翻且組別為3=日文組的不用核稿
                        txt1(5) = "B4009" '核稿主管改為吳嘉智
                     End If
                     '2025/5/22 END
                  Else
                  '2025/4/7 END
                     txt1(5) = Pub_GetSpecMan("H2")
                  End If
               End If
            End If
         End If
         '2013/8/27 END
         
         'Add By Sindy 2015/1/20
         'Modified by Morgan 2021/2/24 審查意見通知函1202除外--玲玲
         If SystemNumber(Me.LBL1(7).Caption, 1) = "P" And Len(m_CP10) = 4 And "" & .Fields("cpm23") = "9" And m_CP10 <> "1202" Then
            If .Fields("pa08") = "1" Then
               'Modify By Sindy 2024/6/26 +m_Country
               Call PUB_ChkIsSetPromoterReader(m_CP14, SystemNumber(Me.LBL1(7).Caption, 1), "101", strExc(0), , , m_Country)
            ElseIf .Fields("pa08") = "2" Then
               'Modify By Sindy 2024/6/26 +m_Country
               Call PUB_ChkIsSetPromoterReader(m_CP14, SystemNumber(Me.LBL1(7).Caption, 1), "102", strExc(0), , , m_Country)
            ElseIf .Fields("pa08") = "3" Then
               'Modify By Sindy 2024/6/26 +m_Country
               Call PUB_ChkIsSetPromoterReader(m_CP14, SystemNumber(Me.LBL1(7).Caption, 1), "103", strExc(0), , , m_Country)
            End If
            If strExc(0) <> "" Then
               txt1(5) = strExc(0)
            End If
         End If
         '2015/1/20 END
         
         'Modify By Sindy 2013/9/11 若承辦人等於預設核稿人,則為自行核稿,核稿人應為空白
         'If Trim(m_pp04) = "" Then
            If Trim(Left("" & Combo1.Text, 6)) = Trim(txt1(5)) Then
               txt1(5) = ""
            End If
         'End If
         '2013/9/11 END
         LBL1(14).Caption = GetPrjSalesNM(txt1(5))
      Else
         txt1(5).Text = LBL1(14).Caption
         LBL1(14).Caption = "" & .Fields("ep04N")
      End If
      txt1(5).Tag = txt1(5) 'Add By Sindy 2013/10/16
      
      'Modify By Sindy 2015/3/4 有完稿日時,則不用再預設英文核稿人
      If Val("" & .Fields("EP09")) <= 0 And Len("" & .Fields("EP03")) = 0 And txt1(23) = "1" Then
         If m_PER04 <> "" And m_PER04 <> m_CP14 Then
            txt1(6).Text = m_PER04
         End If
      Else
         txt1(6).Text = LBL1(16).Caption
         'lbl1(16).Caption = "" & .Fields("ep03N")
      End If
      LBL1(16).Caption = GetPrjSalesNM(txt1(6))
      '2015/3/4 END
      
      txt1(0).Text = LBL1(4).Caption
      LBL1(4).Caption = "" & .Fields("ep13N")
      'end 2009/11/12
      
      'Add By Sindy 2019/12/23 檢查是否有設定核判表,若無,代表無辦理此案件性質;因此必須鎖定要做核判歷程
      If InStr("P10,P11", GetST15(m_CP14)) > 0 Then
         'Modify By Sindy 2024/6/26 +m_Country
         If PUB_ChkIsSetPromoterReader(m_CP14, SystemNumber(Me.LBL1(7).Caption, 1), m_CP10, , , , m_Country) = False Then
            'Modify By Sindy 2024/12/30 目的僅為鎖住核稿人和判發人不可空白
            txt1(5).Enabled = True
            m_PP04 = "需設核稿人"
            m_PP05 = "需設判發人"
'            If SystemNumber(Me.LBL1(7).Caption, 1) = "P" Then
'               txt1(5).Enabled = True
'               m_PP04 = "73022"
'               m_PP05 = "73022"
'            Else
'               'Added by Lydia 2023/04/24 修改王副總退休之相關控制
'               If strSrvDate(1) >= "20230511" Then
'                  txt1(5).Enabled = True
'                  m_PP04 = "99050"
'                  m_PP05 = "99050"
'               Else
'               'end 2023/04/24
'                  txt1(5).Enabled = True
'                  m_PP04 = "71011"
'                  m_PP05 = "71011"
'               End If 'Added by Lydia 2023/04/24
'            End If
            '2024/12/30 END
         End If
      End If
      '2019/12/23 END
   End If
   End With
   
   CheckOC
   'add by nickc 2006/02/27
   'add by nickc 2006/03/14
   InitialField
   Dim rsTmp As New ADODB.Recordset
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & LBL1(3).Caption & "' "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
       UpdateFieldOldData rsTmp
   End If
      
   'Added by Lydia 2016/12/30 新增無法收費修正案由工程師輸入事由,由專利處王副總及游經理維護免費事由
   'Modified by Lydia 2018/05/24 排除無資料的情況
   'If ProState = "1" And cmdOK(1).Visible = True And cmdOK(1).Enabled = True Then
   If ProState = "1" And cmdOK(1).Visible = True And cmdOK(1).Enabled = True And rsTmp.RecordCount > 0 Then
     'P或CFP之主動修正(203)或修正(204)，若未收費(nvl(cp16,0)=0)未取消收文(cp159=0)時，若該案件進度檔沒有未發文未取消收文之A類申復(205)時(即該案件若有A類申復尚未發文則不一定要輸入承辦備註)，承辦備註欄增加下拉選單，預設免費修正事由檔的選項供使用者點選，亦可由使用者自行輸入，但不可空白；
     '非上述情形使用者可自行決定是否輸入承辦備註；
     If Val("" & rsTmp.Fields("CP16")) = 0 And Val("" & rsTmp.Fields("CP16")) = 0 And (rsTmp.Fields("CP01") = "P" Or rsTmp.Fields("CP01") = "CFP") And (rsTmp.Fields("CP10") = "203" Or rsTmp.Fields("CP10") = "204") Then
        strSql = "select cp09 from caseprogress where cp01=" & CNULL(rsTmp.Fields("CP01")) & " and cp02=" & CNULL(rsTmp.Fields("CP02")) & " and cp03=" & CNULL(rsTmp.Fields("CP03")) & " and cp04=" & CNULL(rsTmp.Fields("CP04")) & " and cp10='205' and cp158=0 and cp159=0 "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 0 Then
           cmdFAmend.Visible = True
        End If
     End If
   End If
   'end 2016/12/30
   
   If rsTmp.State = 1 Then rsTmp.Close
   
   'Remove by Morgan 2009/11/12 移到上面減少語法
   'txt1(0).Text = lbl1(4).Caption
   'txt1(6).Text = lbl1(16).Caption
   'end 2009/11/12
   
   Dim tmpInti As Integer

   'edit by nick 2005/03/01  改工程師都可以選，但是若繪圖主管已經確認，就不能再修改了
   '若繪圖主管已經確認則不能修改
   If m_CP107 = "" Then
      Combo2.Enabled = True
      '2009/4/20 ADD BY SONIA 游經理說實審之繪圖欄鎖住,否則工程師經常會誤選
      '2009/9/2 MODIFY BY SONIA 個人使用時非新申請案都要鎖住,但主管不限制
      'If m_CP10 = "416" Then
      'modify by sonia 2017/3/28 郭副理要求開放210撰稿可點選繪圖人員
      If ProState = "1" And InStr(NewCasePtyList, m_CP10) = 0 And m_CP10 <> "210" Then
      '2009/9/2 END
         Combo2.Enabled = False
      End If
      '2009/4/20 END
   Else
      Combo2.Enabled = False
   End If

   For tmpInti = 0 To Combo2.ListCount - 1
       If Trim(txt1(0).Text) = Trim(Mid(Combo2.List(tmpInti), 1, InStr(1, Combo2.List(tmpInti), "=") - IIf(InStr(1, Combo2.List(tmpInti), "=") = 0, 0, 1))) Then
           Combo2.Text = Combo2.List(tmpInti)
           '記錄原繪圖人員
           m_EP13 = Trim(Left(Me.Combo2.Text, 6))
           'End
           Exit For
       End If
   Next tmpInti
   
   'Add by Morgan 2011/9/20 若繪圖已離職也要帶出
   If txt1(0).Text <> "" And Combo2.Text = "" Then
      Combo2.Text = txt1(0).Text & " ==> " & LBL1(4)
   End If
   
   'add by nickc 2007/08/21
   For tmpInti = 0 To Combo4.ListCount - 1
       If Trim(txt1(6).Text) = Trim(Mid(Combo4.List(tmpInti), 1, InStr(1, Combo4.List(tmpInti), "=") - IIf(InStr(1, Combo4.List(tmpInti), "=") = 0, 0, 1))) Then
           Combo4.Text = Combo4.List(tmpInti)
       End If
   Next tmpInti
   If Trim(Combo4.Text) = "" Then Combo4.Text = txt1(6).Text 'Add By Sindy 2024/9/9
   Combo4.Tag = Combo4.Text 'Add By Sindy 2015/3/4 英文核稿人下拉選單無此人時,會帶不出資料
   '                                               但為以免誤認為無欄位值,將其員編顯示出來
   If txt1(20) = "Y" Then
      txt1(19).Enabled = False
   End If
   
   'Remove by Morgan 2009/11/12 移到上面減少語法
   'lbl1(4).Caption = GetPrjSalesNM(txt1(0).Text)
   'end 2009/11/2
   txt1(2).Text = ChangeWStringToTString(LBL1(8).Caption)
   txt1(3).Text = ChangeWStringToTString(LBL1(10).Caption)
   
   'Remove by Morgan 2008/10/13 改讀資料時就設定
   'txt1(4).Text = ChangeWStringToTString(lbl1(12).Caption)
   
'Remove by Morgan 2009/11/12 併入上面程式
'   If Len(lbl1(14).Caption) = 0 Then
'      strSQL = "select pp04,st02 from caseprogress,promoterproofreader,staff where cp09='" & lbl1(3) & "' and cp01=pp01(+) and cp14=pp02(+) and cp10=pp03(+) and st01(+)=pp04"
'      CheckOC2
'      adoRecordset1.CursorLocation = adUseClient
'      adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If adoRecordset1.RecordCount <> 0 Then
'         txt1(5) = CheckStr(adoRecordset1.Fields(0))
'         lbl1(14).Caption = "" & adoRecordset1.Fields(1)
'      End If
'   Else
'      txt1(5).Text = lbl1(14).Caption
'   End If
'lbl1(14).Caption = GetPrjSalesNM(txt1(5).Text)
'end 2009/11/12


   'lbl1(16).Caption = GetPrjSalesNM(txt1(6).Text) 'Remove by Morgan 2009/11/12 移到上面減少語法
   
   '會稿完成日
   'Remove by Morgan 2008/12/8 改讀資料時就設定
   'txt1(7).Text = ChangeWStringToTString(lbl1(18).Caption)
   
   
   'Remove by Morgan 2008/10/3 改成直接寫,避免畫面上太多不用的物件佔位置
   'txt1(8).Text = ChangeWStringToTString(lbl1(20).Caption)
   'txt1(9).Text = lbl1(22).Caption
   'txtep12.Text = lbl1(24).Caption
   
   '承辦期限
   'Me.txt1(12).Text = ChangeTDateStringToTString(Me.lbl1(2).Caption) 'Remove by Morgan 2010/10/7 已改成直接寫
   '92.6.13 add by sonia
   Me.txt1(12).Enabled = True
   
   'Add by Morgan 2010/9/28 專利新規則P的非FMP案及CFP案承辦期限不可修改
   If bolNewPromoterRule Then
      If (m_FieldList(0).fiNewData = "P" And Left(m_FieldList(11).fiNewData, 1) <> "F") Or m_FieldList(0).fiNewData = "CFP" Then
         txt1(12).Enabled = False
      End If
   End If
   'end 2010/9/28
   
   'add by nickc 2006/05/08
   'Modify by Morgan 2008/10/13
   'txt1(18) = lbl1(33)
   txt1(18) = txt1(18).Tag
   'end 2008/10/13

   '若為個人工作管理
   If ProState = "1" Then
      '92.7.1 ADD BY SONIA
      'Modify By Sindy 2014/1/13 因人員休假職代關係時,開放收卷註記及本所期限可以輸入
      If Trim(Left("" & Combo1.Text, 6)) <> "" And Trim(Left("" & Combo1.Text, 6)) <> strUserNum Then
         Label1(36).Visible = True
         Label1(35).Visible = True
         'Modify by Amy 2014/09/22 取消工程師輸入本所期限
         'txt1(13).Visible = True
         'txt1(13).Enabled = True
         txt1(14).Visible = True
         txt1(14).Enabled = True
      Else
      '2014/1/13 END
         Label1(36).Visible = False
         Label1(35).Visible = False
         'Modify by Amy 2014/09/22 取消工程師輸入本所期限
         'txt1(13).Visible = False
         'txt1(13).Enabled = False
         txt1(14).Visible = False
         txt1(14).Enabled = False
      End If
      '92.7.1 END
       If Me.LBL1(7).Caption <> "" Then
           arrCaseNo = Split(Me.LBL1(7).Caption, "-")
           If arrCaseNo(0) = "P" Or arrCaseNo(0) = "CFP" Then
               '承辦期限
               If Me.txt1(12).Text <> "" Then Me.txt1(12).Enabled = False
           End If
       End If
       'add by nickc 2007/09/03
      txt1(15).Enabled = False
      txtCP99.Enabled = False
      'Frame1.Enabledd = False Remove by Morgan 2010/10/13
      txt1(18).Enabled = False
      txt1(20).Enabled = False
   ElseIf ProState = "2" Then
       frm090614.TextOk = True
       'Frame1.Enabled = True        '2008/10/31 CANCEL BY SONIA 否則上面設成FALSE此處又改為TRUE
       If Trim(txt1(14)) <> "" Then
           'Modify by Morgan 2010/10/29 預定會稿日已改上齊備日後由系統預設(隔日凌晨),此處改為可修改
           ''Modify by Morgan 2008/12/2 改判斷有設定適用規則能改(原來判斷適用的才能改)
           ''If opCP112(0).Value = True Then
           'If m_CPM05 <> "" Then
           ''end 2008/12/2
           '    txt1(18).Enabled = True
           'Else
           '    txt1(18).Enabled = False
           'End If
           If txt1(18) <> "" Then txt1(18).Enabled = True
           'end 2010/10/29
           txt1(20).Enabled = True
       End If
   Else
      Label1(36).Visible = True
      Label1(35).Visible = True
      'Modify by Amy 2014/09/22 取消工程師輸入本所期限
      'txt1(13).Visible = True
      'txt1(13).Enabled = True
      txt1(14).Visible = True
      txt1(14).Enabled = True
   '92.7.1 END
   End If

   '92.6.13 end
   'Modify By Cheng 2003/06/13
   If m_blnClkSure = False Then
       If Me.txt1(12).Text = "" And Me.txt1(2).Text <> "" Then txt1_LostFocus (2)
   End If

   'Modify By Cheng 2003/11/05
   'C類中P,CFP不可輸發文日
   If Left(UCase(strText), 1) = "C" Then
       If SystemNumber(Me.LBL1(7).Caption, 1) = "P" Or SystemNumber(Me.LBL1(7).Caption, 1) = "CFP" Then
           Label1(3).Visible = False
           txt1(9).Visible = False
           Label1(22).Visible = False
           txt1(8).Enabled = False
           txt1(8).TabStop = False
           'txt1(8).Visible = False
       Else
          Label1(3).Visible = True
          txt1(9).Visible = True
          Label1(22).Visible = True
          txt1(8).Enabled = True
          txt1(8).TabStop = True
          'txt1(8).Visible = True
       End If
   Else
      Label1(3).Visible = False
      txt1(9).Visible = False
      Label1(22).Visible = False
      txt1(8).Enabled = False
      txt1(8).TabStop = False
      'txt1(8).Visible = False
   End If

   'Add By Cheng 2003/05/14
   'Modify By Sindy 2013/9/11 前面程式段有判斷鎖及開放條件了
'   Me.txt1(2).Enabled = True
'   Me.txt1(3).Enabled = True
'   Me.txt1(4).Enabled = True
'   Me.txt1(7).Enabled = True
   '若為個人工作管理
   If ProState = "1" Then
       txt1(18).Enabled = False
       If Me.LBL1(7).Caption <> "" Then
           arrCaseNo = Split(Me.LBL1(7).Caption, "-")
           'Modify By Sindy 2019/12/3 + ) And PUB_GetST03(m_CP14) <> "P12"
           '                排除專利程序承辦的案件
           If (arrCaseNo(0) = "P" Or arrCaseNo(0) = "CFP") And PUB_GetST03(m_CP14) <> "P12" Then
               m_bol203Case = Chk203Case(CStr(arrCaseNo(0)), CStr(arrCaseNo(1)), CStr(arrCaseNo(2)), CStr(arrCaseNo(3)))  'Added by Morgan 2013/8/1
               If m_bol203Case = False Then 'Added by Morgan 2013/8/1
                  '齊備日
                  If Me.txt1(2).Text <> "" Then Me.txt1(2).Enabled = False
                  '完稿日
                  If Me.txt1(3).Text <> "" Then Me.txt1(3).Enabled = False
                  '會稿日
                  If Me.txt1(4).Text <> "" Then Me.txt1(4).Enabled = False
               End If 'Added by Morgan 2013/8/1
               
               '會稿完成日
               If Me.txt1(7).Text <> "" Then
                  Me.txt1(7).Enabled = False
'               'Add By Sindy 2013/9/17
'               Else
'                  '會稿完成日空白時,若已(送會)但沒有上(不自動更新會完日)記錄時,不可手動輸入會完日
'                  If PUB_ChkEmpFlowExists(lbl1(3), EMP_送會, , strRefEEP02) = True Then
'                     If PUB_ChkEmpFlowExists(lbl1(3), EMP_不自動更新會完日, strRefEEP02) = False Then
'                        Me.txt1(7).Enabled = False
'                     End If
'                  End If
               End If
'               '2013/9/17 END

'               'Modify By Sindy 2013/9/5
'               If Me.txt1(7).Text <> "" Then '有工程師會稿完成日
'                  If Val(strEP38) = 0 Then '無業務員會稿完成日
'                     Me.txt1(7).Enabled = False '鎖住
'                  Else
'                     If Val(DBDATE(Me.txt1(7).Text)) <> Val(strEP38) Then '二個會稿完成日不同時，鎖住，只有一次修改機會
'                        Me.txt1(7).Enabled = False
'                     End If
'                  End If
'               End If
'               '2013/8/12 END
               
               'add by nickc 2006/09/08 個人，在承辦期限前一天可以輸入，一次，輸過不能再修改
               If txt1(12) <> "" Then
                   'Modify by Morgan 2010/10/29 預定會稿日已改上齊備日後由系統預設(隔日凌晨),此處改為可修改
                   'If strSrvDate(1) >= CompWorkDay(2, DBDATE(txt1(12)), 1) And Trim(txt1(18)) = "" And Trim(Me.txt1(4).Text) = "" Then
                   '    'Modify by Morgan 2008/12/2 改判斷有設定適用規則能改(原來判斷適用的才能改)
                   '    'If opCP112(0).Value = True Then
                   '    If m_CPM05 <> "" Then
                   '    'end 2008/12/2
                   '        txt1(18).Enabled = True
                   '    Else
                   '        txt1(18).Enabled = False
                   '    End If
                   'End If
                   'If txt1(18) <> "" Then txt1(18).Enabled = True 'Removed by Morgan 2012/7/20 個人不可修改預定會稿日
                   'end 2010/10/29
               End If
               'add by nickc 2007/08/21 有輸入過，就鎖住
               txt1(20).Enabled = False
'               If Trim(Combo4.Text) <> "" Then Combo4.Enabled = False Else Combo4.Enabled = True
               If Trim(txt1(19).Text) <> "" Then
                   txt1(19).Enabled = False
               Else
                   If Trim(txt1(20)) = "" Then
                       txt1(19).Enabled = True
                   Else
                       txt1(19).Enabled = False
                   End If
               End If
           End If
       End If
       
      'Add By Sindy 2019/12/3 專利程序承辦的案件,若齊備日空白時直接帶入系統日
      If PUB_GetST03(m_CP14) = "P12" And Val(txt1(2).Text) = 0 And txt1(2).Enabled = True Then
         'Modify By Sindy 2021/7/29 品薇排除,因為她是大陸案的承辦人(工程師角色)，與其他程序人員不同。
         'Modified by Morgan 2025/1/22 P程序人員工作改依智權區域分配，改判斷新申請案程序不預設--郭
         ''If Pub_GetSpecMan("PS2") <> m_CP14 Then
         If InStr(NewCasePtyList, LBL1(29)) = 0 Then
         'end 2025/1/22
            txt1(2).Text = strSrvDate(2)
         End If
      End If
      '2019/12/3 END
   End If
   'Add By Sindy 2014/8/27
   If m_CPM29 = "N" Then
      Label18.Visible = False 'Modify By Sindy 2014/8/28 先不顯示
   Else
      Label18.Visible = False
   End If
   '2014/8/27 END
   'Add By Sindy 2013/9/30
   '若為個人工作管理及承辦人下拉選單為操作者
   If ProState = "1" Or Trim(Left("" & Combo1.Text, 6)) = strUserNum Then
'      'Add By Sindy 2013/10/4 有英文核稿人,或核稿人時一定要走電子簽核
'      If (Combo4.Text <> "" And Left(Trim(Combo4.Text), 5) <> m_CP14) Or _
'         (txt1(5).Text <> "" And txt1(5).Text <> m_CP14) Or _
'         m_CPM29 = "" Then '要電子簽核的案件性質
      
      If m_CPM29 = "" Then '要電子簽核的案件性質
         '英文核完日
         Me.txt1(19).Enabled = False
         'Add By Sindy 2013/11/20
         If Me.LBL1(7).Caption <> "" Then
            arrCaseNo = Split(Me.LBL1(7).Caption, "-")
            m_bol203Case = Chk203Case(CStr(arrCaseNo(0)), CStr(arrCaseNo(1)), CStr(arrCaseNo(2)), CStr(arrCaseNo(3)))
         End If
         If m_bol203Case = True Then
            '齊備日
            Me.txt1(2).Enabled = True
            '完稿日
            Me.txt1(3).Enabled = True
            '會稿日
            Me.txt1(4).Enabled = True
         Else
         '2013/11/20 END
            '完稿日
            Me.txt1(3).Enabled = False
            '會稿日
            Me.txt1(4).Enabled = False
         End If
         '會稿完成日
         Me.txt1(7).Enabled = False
         '不自動更新會完日時,則開放可以自行輸入會稿完成日
         If Me.txt1(7).Text = "" Then
            If PUB_ChkEmpFlowExists(LBL1(3), EMP_送會, , strRefEEP02) = True Then
               If PUB_ChkEmpFlowExists(LBL1(3), EMP_不自動更新會完日, strRefEEP02) = True Then
                  Me.txt1(7).Enabled = True
               End If
            End If
         End If
      End If
   End If
   '2013/9/30 END
   
   'Modify By Sindy 2013/9/13 開放英文核稿人
'   '非 CFP 新申請案，鎖住英文核稿相關欄位
'   If InStr(1, CaseMapOut & ",301,302,303,304,305,306,307", m_CP10) <> 0 And SystemNumber(lbl1(7).Caption, 1) = "CFP" Then
'   Else
'       Combo4.Enabled = False
'       txt1(20).Enabled = False
'       txt1(19).Enabled = False
'   End If
   
   '2008/7/31 add by sonia C類來函未發文才顯示撰寫信函按鈕
   'Modify By Sindy 2015/3/13 +B類941.分析也要顯示撰寫信函按鈕
   'Modify by Morgan 2016/5/3 +A類941
   If (Mid(LBL1(3).Caption, 1, 1) = "C" And txt1(8) = Empty) Or _
      m_CP10 = "941" Then
      cmd(3).Enabled = True
      cmd(3).Visible = True
   Else
      cmd(3).Enabled = False
      cmd(3).Visible = False
   End If
   '2008/7/31 END
   
   'Added by Morgan 2020/12/28 美專IDS清單
   If m_CP10 = "214" Then
      cmd(5).Visible = True
   Else
      cmd(5).Visible = False
   End If
   'end 2020/12/28
   
'*********************************************************
' 是否會稿
'*********************************************************
   txt1(1).Text = LBL1(6).Caption
   'Add By Sindy 2013/9/18 不會稿案件性質,則預設為N
   If txt1(1).Text = "" Then
      'Added by Morgan 2019/4/25
      '承辦人若是掛程序,會稿一律帶N--郭雅娟
      If GetStaffDepartment(Trim(Left("" & Combo1.Text, 6))) = "P12" Then
         txt1(1).Text = "N"
      Else
      'end 2019/4/25
      
         txt1(1).Text = m_CPM28
         'Add By Sindy 2014/1/10 雅娟提:B類的936回覆委任代理人預設不會稿
         If Left(Trim(LBL1(3)), 1) = "B" And m_CP10 = "936" Then
            txt1(1).Text = "N"
         End If
         '2014/1/10 END
         'Add By Sindy 2020/2/15 玲玲提:將A類收文(203)主動修正，(204)修正，會稿(Y/N)預設為”Y”
         If Mid(LBL1(3).Caption, 1, 1) = "A" And _
            (m_CP10 = "203" Or m_CP10 = "204") Then
            txt1(1).Text = "Y"
         End If
         '2020/2/15 END
      End If 'Added by Morgan 2019//4/25
   End If
   '2013/9/18 END
   'Add By Sindy 2017/2/17
   txt1(1).Enabled = True 'Add Sindy 2018/9/18
   If txt1(1) = "Y" And PUB_ChkEmpFlowExists(LBL1(3), EMP_送會) = True Then
      txt1(1).Enabled = False
   End If
   '2017/2/17 END
   'Add By Sindy 2016/2/4 特定案件性質無須上齊備,完稿等日期
   '106.主張國際優先權 121.主張國內優先權
   '123.主張優惠期     416.實體審查
   '909.後金           911.補收款
   '917.超頁、超項費   920.急件費
   '938.超頁費         939.超項費
   '943.繪圖超時       947.超圖費
   'Add By Sindy 2016/2/24 + 944會稿修改
   'Add By Sindy 2016/3/23 +來電聯絡單945, 去電聯絡單955
   'Modify By Sindy 2016/4/21 雅娟:有關特定案件性質無須上相關日期之控管,其中實體審查請排除CFP案件
   'Modify By Sindy 2019/3/6 +414.恢復權利 ex:P111266
   'Modify By Sindy 2023/6/8 +233.製作序列表
   'Modify By Sindy 2024/12/24 +126.期末拋棄
   'Modify By Sindy 2025/1/14 +CFP請願=234
   'Modify By Sindy 2025/2/11 +CFP代理人通知修正=1224
   If m_CP10 = "106" Or m_CP10 = "121" Or _
      m_CP10 = "123" Or _
      (SystemNumber(LBL1(7).Caption, 1) = "P" And m_CP10 = "416") Or _
      m_CP10 = "909" Or m_CP10 = "911" Or _
      m_CP10 = "917" Or m_CP10 = "920" Or _
      m_CP10 = "938" Or m_CP10 = "939" Or _
      m_CP10 = "943" Or m_CP10 = "947" Or _
      m_CP10 = "944" Or m_CP10 = "945" Or _
      m_CP10 = "955" Or m_CP10 = "414" Or _
      m_CP10 = "233" Or m_CP10 = "126" Or _
      (SystemNumber(LBL1(7).Caption, 1) = "CFP" And (m_CP10 = "234" Or m_CP10 = "1224")) Then
      '是否會稿-不會稿
      If txt1(1).Text = "" Then txt1(1).Text = "N"
      If ProState = "1" Then '個人
         txt1(1).Enabled = False
         txt1(2).Enabled = False '齊備日
         txt1(3).Enabled = False '完稿日
         txt1(4).Enabled = False '會稿日
         txt1(7).Enabled = False '會稿完成日
      End If
   End If
   '2016/2/4 END
'***** END ***********************************************

   'Added by Morgan 2012/8/24
   '台灣發明生醫案可讓工程師V選是否他所寄存
   txt1(21) = ""
   txt1(21).Enabled = False
   If SystemNumber(LBL1(7).Caption, 1) = "P" And m_Country = "000" And m_CP10 = "101" And strPA158 = "3" Then
      txt1(21).Visible = True
      Label1(45).Visible = True
      strExc(1) = SystemNumber(LBL1(7).Caption, 1)
      strExc(2) = SystemNumber(LBL1(7).Caption, 2)
      strExc(3) = SystemNumber(LBL1(7).Caption, 3)
      strExc(4) = SystemNumber(LBL1(7).Caption, 4)
      '未收文申請寄存才可設定
      If PUB_ChkCPExist(strExc, "108") = False Then
         '已收文寄存證明不可修改
         If PUB_ChkCPExist(strExc, "231") = True Then
            txt1(21) = "Y"
         Else
            txt1(21).Enabled = True
         End If
      End If
   Else
      txt1(21).Visible = False
      Label1(45).Visible = False
   End If
   txt1(21).Tag = txt1(21)
   'end 2012/8/24
   
   'Add By Sindy 2015/5/21
   If ProState = "2" Then
      'Modified by Morgan 2017/3/27 ST14可能放一個以上的員工編號
      'Modify By Sindy 2019/10/29 + 工作進度資料維護增加第二級期限管制人也有修改權限 => PUB_GetST52(Trim(Left("" & Combo1.Text, 6)), strUserNum) = True
      '  內外翻譯人員的等級主管增加 A2096潘韻丞(虛) 韻丞就可以幫品薇操作承辦人工作進度維護作業等欄位的輸入
      If InStr(Pub_GetSpecMan("承辦人工作管理可修改資料人員"), strUserNum) > 0 Or _
         InStr(PUB_GetST14(Trim(Left("" & Combo1.Text, 6))), strUserNum) > 0 Or _
         Pub_StrUserSt03 = "M51" Or _
         PUB_GetST52(Trim(Left("" & Combo1.Text, 6)), strUserNum) = True Then
         '有修改權限
      Else
         For Each objText In Me.txt1
            'If objText.Index <> 0 And objText.Index <> 10 Then
               objText.Enabled = False
            'End If
         Next
         Me.Combo2.Enabled = False '繪圖人員
         Me.Combo6.Enabled = False '判發人
         Me.Chk1.Enabled = False '無圖式
      End If
      'Modify By Sindy 2015/4/13 當從主管身份進入此作業時,檢查是否有”英文核稿人欄修改”的權限
      strSql = "select EPA01 from EP14Authority where EPA01='" & strUserNum & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Or Pub_StrUserSt03 = "M51" Then
         Combo4.Enabled = True
      Else
         Combo4.Enabled = False
      End If
      '2015/4/13 END
   End If
   '2015/5/21 END
   
   Call SetColTag(True) 'Add By Sindy 2013/6/10
   
   'Added by Morgan 2016/7/7 代主任(52)以上可設定複雜案件
   txt1(11).Enabled = False
   If ProState = "2" Then
      'Add By Sindy 2017/2/13 韻如:勾選複雜案件一事，增加控管”會稿前才能勾選”的限制
      'Modified by Morgan 2019/11/12 因有可能會稿後才要設定，改判斷會稿日(主管先取消會稿日,設定完後再補回) Ex:P-123813--游經理
      'If PUB_ChkEmpFlowExists(lbl1(3), EMP_送會) = False Then
      If txt1(4) = "" And txt1(4).Text = txt1(4).Tag Then
      'end 2019/11/12
      '2017/2/13 END
         strSql = "select st20 from staff where st01='" & strUserNum & "' and st20<='52'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            txt1(11).Enabled = True
         End If
      End If
   End If
   'end 2016/7/7
End Sub

'Add By Sindy 2013/6/10
'bolSetTag=true : 將輸入欄位值記錄至.tag裡面
'bolSetTag=false : 比較輸入欄位值.Tag與畫面上資料是否一致
Private Function SetColTag(bolSetTag As Boolean) As Boolean
   If bolSetTag = True Then
      txt1(14).Tag = txt1(14)
      txt1(15).Tag = txt1(15)
      txtCP99.Tag = txtCP99
      txt1(13).Tag = txt1(13)
      txt1(12).Tag = txt1(12)
      Combo2.Tag = Combo2.Text
      txt1(2).Tag = txt1(2)
      txt1(3).Tag = txt1(3)
      txt1(1).Tag = txt1(1)
      txt1(18).Tag = txt1(18)
      txt1(4).Tag = txt1(4)
      txtEP12.Tag = txtEP12
      txt1(7).Tag = txt1(7)
      txt1(8).Tag = txt1(8)
      txt1(9).Tag = txt1(9)
      txtCP64.Tag = txtCP64
      txt1(17).Tag = txt1(17)
      txt1(5).Tag = txt1(5)
      Combo6.Tag = Combo6.Text 'Add By Sindy 2013/9/4 判發人
      Combo4.Tag = Combo4.Text
      txt1(20).Tag = txt1(20)
      txt1(19).Tag = txt1(19)
      txt1(21).Tag = txt1(21)
      Chk1.Tag = Chk1.Value
   Else
      SetColTag = True
      If txt1(1) = "" Then SetColTag = False: Exit Function '是否會稿欄位空白時,確定鍵會update成Y
      If txt1(14).Tag <> txt1(14) Then SetColTag = False: Exit Function
      If txt1(15).Tag <> txt1(15) Then SetColTag = False: Exit Function
      If txtCP99.Tag <> txtCP99 Then SetColTag = False: Exit Function
      If txt1(13).Tag <> txt1(13) Then SetColTag = False: Exit Function
      If txt1(12).Tag <> txt1(12) Then SetColTag = False: Exit Function
      If Left(Combo2.Tag, 5) <> Left(Combo2.Text, 5) Then SetColTag = False: Exit Function
      If txt1(2).Tag <> txt1(2) Then SetColTag = False: Exit Function
      If txt1(3).Tag <> txt1(3) Then SetColTag = False: Exit Function
      If txt1(1).Tag <> txt1(1) Then SetColTag = False: Exit Function
      If txt1(18).Tag <> txt1(18) Then SetColTag = False: Exit Function
      If txt1(4).Tag <> txt1(4) Then SetColTag = False: Exit Function
      If txtEP12.Tag <> txtEP12 Then SetColTag = False: Exit Function
      If txt1(7).Tag <> txt1(7) Then SetColTag = False: Exit Function
      If txt1(8).Tag <> txt1(8) Then SetColTag = False: Exit Function
      If txt1(9).Tag <> txt1(9) Then SetColTag = False: Exit Function
      If txtCP64.Tag <> txtCP64 Then SetColTag = False: Exit Function
      If txt1(17).Tag <> txt1(17) Then SetColTag = False: Exit Function
      If txt1(5).Tag <> txt1(5) Then SetColTag = False: Exit Function
      If Left(Combo6.Tag, 5) <> Left(Combo6.Text, 5) Then SetColTag = False: Exit Function 'Add By Sindy 2013/9/4 判發人
      If Left(Combo4.Tag, 5) <> Left(Combo4.Text, 5) Then SetColTag = False: Exit Function
      If txt1(20).Tag <> txt1(20) Then SetColTag = False: Exit Function
      If txt1(19).Tag <> txt1(19) Then SetColTag = False: Exit Function
      If txt1(21).Tag <> txt1(21) Then SetColTag = False: Exit Function
      If Chk1.Tag <> Chk1.Value Then SetColTag = False: Exit Function
   End If
End Function

'91.08.14  nick  加畫面顯示其他國外案
Private Sub Cmd1_Click()
Dim iMouse As Integer
iMouse = Screen.MousePointer

Me.Hide
Screen.MousePointer = vbHourglass
frm090201_2_1.SetParent Me 'Add By Sindy 2014/1/14
frm090201_2_1.Show
frm090201_2_1.StrMenu (LBL1(7).Caption)
'Modify by Morgan 2009/11/12
'Screen.MousePointer = vbDefault
Screen.MousePointer = iMouse
End Sub

'Add By Sindy 2013/8/19
Private Sub cmdDetail_Click()
   Call grd2_DblClick
End Sub

'Modify By Sindy 2013/5/21
'Private Sub cmdOK_Click(index As Integer)
Public Sub cmdok_Click(Index As Integer)
'2013/5/21 End
   '***2008/11/21 加註BY SONIA 按確定後很快按結束會因為DoEvents造成錯誤,因使用者未反應故暫不取消DoEvents
   Dim iMouse As Integer
'   Dim bolUpdDate As Boolean 'Add By Sindy 2013/10/4
   
   iMouse = Screen.MousePointer
   
   Select Case Index
   Case 0 '本月統計
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      CALCUTE_090201 Trim(Left("" & Combo1.Text, 6)), Text1.Text
      Me.Enabled = True
      
      'Modify by Morgan 2009/11/12
      'Screen.MousePointer = vbDefault
      Screen.MousePointer = iMouse
      
      'Modified by Lydia 2025/02/05 改用變數
      'grd1.col = 3
      GRD1.col = colCaseNo_1
      Me.Hide
      '若為個人管理
      If ProState = "1" Then
          frm090201_3_2.m_strYear = Left(Me.Text1.Text, 4) - 1911
          frm090201_3_2.m_strMonth = Right(Me.Text1.Text, 2)
      '若為工作維護
      Else
          frm090201_3_2.m_strYear = frm090614.txt1(3).Text
          frm090201_3_2.m_strMonth = frm090614.txt1(4).Text
      End If
      frm090201_3_2.Show
        
   Case 1 '確定
         'Add By Sindy 2017/8/3 個人案件不可用主管權限操作
         If ProState = "2" And m_CP14 = strUserNum Then '2.主管
            MsgBox "個人案件不可用主管權限操作！", vbExclamation
            Exit Sub
         End If
         '2017/8/3 END
         
         Select Case ProState
         Case "1", "2"
            m_chkcmdok1 = False 'Add By Sindy 2013/6/7 進入承辦歷程時會先執行一次確定鍵,因有可能已在此畫面先修改資料,且有些日期檢查條件須先執行
            
            'add by nickc 2007/12/28 加入修正
            If SSTab1.Tab = 0 Then Exit Sub '*****
            
            'Added by Morgan 2019/4/25
            '承辦人若是掛程序,會稿一律帶N--郭雅娟
            If txt1(1) = "" Then
               If GetStaffDepartment(Trim(Left("" & Combo1.Text, 6))) = "P12" Then
                  txt1(1).Text = "N"
               End If
            End If
            'end 2019/4/25
      
            'add by nickc 2007/11/28 協理說，有完稿日才可以輸入不會稿
            'If txt1(3) = "" Then txt1(1) = "" 'Modify By Sindy 2013/9/6 開放
            If txt1(1) = "" Then txt1(1) = "Y"    '2012/5/15 ADD BY SONIA
            
            Call ChkEP34ToEP07EP08 'Modify By Sindy 2016/5/20 抽出來變函數
            
            Screen.MousePointer = vbHourglass
            ChkCp27 = True
            'Modify By Sindy 2017/9/15
            'If SSTab1.Tab = 1 Then
            'If SSTab1.Tab = 1 Or (SSTab1.Tab = 2 And Me.m_Flow <> "") Then
            If SSTab1.Tab = 1 Or Me.m_Flow <> "" Then 'Modify By Sindy 2017/9/20
            '2017/9/15 END
               If ChkNoData = False Then
                  '重新檢查欄位有效性
                  If TxtValidate = True Then
                     DoEvents
                     Me.Enabled = False
                     If FormSave = True Then
                        PUB_AskUpdateRelationCase LBL1(3) 'Added by Morgan 2015/5/25
                        
                        'Add by Morgan 2008/12/5 檢查若[適不適用會稿加乘註記]有變更但存檔後值與原畫面選項不同時提醒
                        'CP121AlterCheck 'Remove by Morgan 2010/10/14
                        
                        'add by nickc 2006/12/29 集中發信
                        'Modify By Sindy 2018/6/20 + If m_Flow = "" Then
                        If m_Flow = "" Then BatctMail
                        
                        '更新mdb暫存資料及第一畫面的Grid內容
                        UpdEngMdb
                        TextOk = False
                        'add by nickc 2005/07/11 讓存完檔的變色正確
                        'ChgGrdColor 'Remove by Morgan 2009/11/12 UpdEngMdb內做就好
                        Call SetColTag(True) 'Add By Sindy 2013/6/10
                        m_chkcmdok1 = True 'Add By Sindy 2013/6/7
'                     Else
'                        MsgBox "存檔失敗!" & Err.Number & Err.Description
                     End If
                     Me.Enabled = True
                     'Add By Sindy 2017/9/21
                     If cmdOK(1).Enabled = True Then
                     '2017/9/21 END
                        SSTab1.Tab = 0
                     End If
'                     'Modify By Sindy 2013/6/7
'                     If intBackTab = 2 Then
'                        Call QueryData(True)
'                     End If
'                     SSTab1.Tab = intBackTab
'                     intBackTab = 0
'                     '2013/6/7 End
                  End If
               End If
            Else
               SSTab1.Tab = 1
            End If
            'Modify by Morgan 2009/11/12
            'Screen.MousePointer = vbDefault
            Me.m_Flow = "" 'Add By Sindy 2017/8/28
            Screen.MousePointer = iMouse
         Case Else
         End Select
         
   Case 2 '結束
        Select Case ProState
        Case "1"
            Unload Me
            Exit Sub
        Case "2"
             frm090614.Show
             Unload Me
             Exit Sub
        Case "3"
        Case Else
        End Select
   'Added by Morgan 2024/3/19
   Case 3 '工程師達成情況(&Word)
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      ExportWord
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   Case Else
   End Select
End Sub

'Add By Sindy 2016/5/20 抽出來變函數
Private Sub ChkEP34ToEP07EP08()
Dim bolChkEmp As Boolean 'Add By Sindy 2013/10/4
   
   'add by nickc 2006/09/26 若是輸入不會稿，直接按存檔，他不會自動代
   'If txt1(1) = "N" Then txt1(4) = txt1(3): txt1(7) = txt1(3)
   'Add By Sindy 2013/10/4 ex.P-103923 檢查是否有送核流程,若有,要檢查有核稿完成日,方可上會稿日及會稿完成日
   If txt1(1) = "N" Then
      'bolUpdDate = True
      bolChkEmp = False
      '要電子簽核的案件或有電子歷程的案件
      'Modify By Sindy 2016/5/19 增加檢查是否有送核中的 ex:P-113197
      If m_CPM29 = "" Or _
         m_Flow = EMP_送核 Or _
         m_Flow = EMP_送英核 Or _
         ((PUB_ChkEmpFlowExists(LBL1(3), EMP_送核) = True Or PUB_ChkEmpFlowExists(LBL1(3), EMP_送英核) = True) And PUB_ChkEmpFlowExists(LBL1(3), EMP_核完) = False) Then
         '有英文核稿人
         If Combo4.Text <> "" And Left(Trim(Combo4.Text), 5) <> m_CP14 Then
            bolChkEmp = True
'                     If Val(txt1(19)) <= 0 Then
'                        bolUpdDate = False
'                     End If
         End If
         '有核稿人
         If txt1(5).Text <> "" And txt1(5).Text <> m_CP14 Then
            bolChkEmp = True
'                     strExc(0) = "select ep39 From engineerprogress where ep02='" & lbl1(3) & "'"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        If Val("" & RsTemp.Fields("ep39")) <= 0 Then
'                           bolUpdDate = False
'                        End If
'                     End If
         End If
      End If
      If bolChkEmp = False Then '無電子簽核或無核稿主管
         'Modify By Sindy 2016/5/19 發現核完N不會稿時,系統會上(會稿日)和(會稿完成日),所以要檢查日期是否已有值,以免重新覆蓋掉 ex:P-113197
         'txt1(4) = txt1(3): txt1(7) = txt1(3)
         If Trim(txt1(4).Text) = "" Then
            txt1(4).Text = txt1(3).Text
         End If
         If Trim(txt1(7).Text) = "" Then
            txt1(7).Text = txt1(3).Text
         End If
         '2016/5/19 END
'               Else
'                  If bolUpdDate = True Then
'                     If Trim(txt1(4)) = "" Then
'                        txt1(4) = strSrvDate(2)
'                     End If
'                     If Trim(txt1(7)) = "" Then
'                        txt1(7) = strSrvDate(2)
'                     End If
'                  End If
      End If
   End If
   '2013/10/4 END
End Sub

'Add by Morgan 2008/12/9
'判斷無國內之相同案資料
Private Function GetF_CP14x(oKey() As String) As String
   Dim stCP14 As String
   
   F_ST02 = ""
   F_CP01020304 = ""
   F_ST03 = ""
   strExc(0) = "select cp14,st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04,st03 From caserelation, CASEPROGRESS,engineerprogress , STAFF where cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and CP14=ST01(+) AND CP31='Y' and cr01='" & oKey(1) & "' and cr02='" & oKey(2) & "' and cr03='" & oKey(3) & "' and cr04='" & oKey(4) & "' and cp57 is null and cp27 is null and cp31='Y' and cp10 in (" & GetAddStr(CaseMapOut) & ") and ep02(+)=cp09 and ep06 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         stCP14 = stCP14 & CheckStr(.Fields(0).Value) & ","
         F_ST02 = F_ST02 & CheckStr(.Fields(1).Value) & ","
         F_CP01020304 = F_CP01020304 & CheckStr(.Fields(2).Value) & ","
         F_ST03 = F_ST03 & CheckStr(.Fields(3).Value) & ","
         .MoveNext
      Loop
      End With
   End If
   If Right(stCP14, 1) = "," Then stCP14 = Mid(stCP14, 1, Len(stCP14) - 1)
   If Right(F_ST02, 1) = "," Then F_ST02 = Mid(F_ST02, 1, Len(F_ST02) - 1)
   If Right(F_CP01020304, 1) = "," Then F_CP01020304 = Mid(F_CP01020304, 1, Len(F_CP01020304) - 1)
   If Right(F_ST03, 1) = "," Then F_ST03 = Mid(F_ST03, 1, Len(F_ST03) - 1)
   GetF_CP14x = stCP14
End Function

'Modify by Morgan 2008/12/3 + p_InCaseCP14:國內案承辦人,p_iOptin:1 =上會稿日,2=上會稿完成日
Function GetF_CP14(oKey As String, Optional p_InCP14 As String, Optional p_iOptin As Integer) As String
'edit by nick 2005/01/17 新案才發國外承辦人
'StrSql = " select cp14,st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04,st03 From CASEMAP, CASEPROGRESS, STAFF where cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05='" & SystemNumber(oKey, 1) & "' and cm06='" & SystemNumber(oKey, 2) & "' and cm07='" & SystemNumber(oKey, 3) & "' and cm08='" & SystemNumber(oKey, 4) & "' and cp57 is null and cp27 is null "
'edit by nickc 2005/09/30
'strSQL = " select cp14,st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04,st03 From CASEMAP, CASEPROGRESS, STAFF where cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05='" & SystemNumber(oKey, 1) & "' and cm06='" & SystemNumber(oKey, 2) & "' and cm07='" & SystemNumber(oKey, 3) & "' and cm08='" & SystemNumber(oKey, 4) & "' and cp57 is null and cp27 is null and cp31='Y' "
'edit by nickc 2007/10/24 秀玲說只要控制申請案
'strSQL = " select cp14,st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04,st03 From CASEMAP, CASEPROGRESS, STAFF where cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05='" & SystemNumber(oKey, 1) & "' and cm06='" & SystemNumber(oKey, 2) & "' and cm07='" & SystemNumber(oKey, 3) & "' and cm08='" & SystemNumber(oKey, 4) & "' and cp57 is null and cp27 is null and cp31='Y' and cm10='0' "
strSql = " select cp14,st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04,st03 From CASEMAP, CASEPROGRESS, STAFF where cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and CP14=ST01(+) AND CP31='Y' and cm05='" & SystemNumber(oKey, 1) & "' and cm06='" & SystemNumber(oKey, 2) & "' and cm07='" & SystemNumber(oKey, 3) & "' and cm08='" & SystemNumber(oKey, 4) & "' and cp57 is null and cp27 is null and cp31='Y' and cm10='0' and cp10 in (" & GetAddStr(CaseMapOut) & ") "

CheckOC
GetF_CP14 = ""
F_ST02 = ""
F_CP01020304 = ""
F_ST03 = ""
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            'If SystemNumber(UCase(CheckStr(.Fields(2).Value)), 1) <> "P" Then 'Remove by Morgan 2010/9/17 不必再限制(大陸案也要)--秀玲
               'Modify by Morgan 2008/12/8 +判斷承辦人是否相同及上會稿日或會稿完成日
               If p_InCP14 = "" Then
                  GetF_CP14 = GetF_CP14 & CheckStr(.Fields(0).Value) & ","
                  F_ST02 = F_ST02 & CheckStr(.Fields(1).Value) & ","
                  F_CP01020304 = F_CP01020304 & CheckStr(.Fields(2).Value) & ","
                  F_ST03 = F_ST03 & CheckStr(.Fields(3).Value) & ","
               ElseIf (p_iOptin = 1 And "" & .Fields("cp14") = p_InCP14) Or (p_iOptin = 2 And "" & .Fields("cp14") <> p_InCP14) Then
                  GetF_CP14 = GetF_CP14 & CheckStr(.Fields(0).Value) & ","
                  F_ST02 = F_ST02 & CheckStr(.Fields(1).Value) & ","
                  F_CP01020304 = F_CP01020304 & CheckStr(.Fields(2).Value) & ","
                  F_ST03 = F_ST03 & CheckStr(.Fields(3).Value) & ","
               End If
            'End If
            .MoveNext
        Loop
    Else
        GetF_CP14 = ""
        F_ST02 = ""
        F_CP01020304 = ""
        F_ST03 = ""
    End If
End With
If Right(GetF_CP14, 1) = "," Then GetF_CP14 = Mid(GetF_CP14, 1, Len(GetF_CP14) - 1)
If Right(F_ST02, 1) = "," Then F_ST02 = Mid(F_ST02, 1, Len(F_ST02) - 1)
If Right(F_CP01020304, 1) = "," Then F_CP01020304 = Mid(F_CP01020304, 1, Len(F_CP01020304) - 1)
If Right(F_ST03, 1) = "," Then F_ST03 = Mid(F_ST03, 1, Len(F_ST03) - 1)
CheckOC
End Function

Private Sub cmdok2_Click(Index As Integer)
Dim iMouse As Integer
iMouse = Screen.MousePointer

Screen.MousePointer = vbHourglass
GRD1.Visible = False
Select Case Index
Case 0 '當月資料
      'strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110007,R110008,R110009,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022,CP57 FROM R090614,CASEPROGRESS WHERE R110022=CP09(+) AND ID='" & strUserNum & "' AND R110001='" & Combo1.Text & "' AND (r110018 is null or decode(length(r110018),8,substr(r110018,1,5),9,substr(r110018,1,6),null,'')=decode(length(r110018),8,'" & Format(ChangeWStringToWDateString(GetTodayDate), "EE/MM") & "',9,'" & Mid(Format(ChangeWStringToTString(GetTodayDate), "0##/##/##"), 1, 6) & "',null,'')) and (cp57 is null or cp57=0 or substr(cp57,1,6)=" & Mid(GetTodayDate, 1, 6) & ") ORDER BY 1 "
      'Modify By Cheng 2002/04/16
'      strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110007,R110008,R110009,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022,CP57 FROM R090614,CASEPROGRESS WHERE R110022=CP09(+) AND ID='" & strUserNum & "' AND R110001='" & Combo1.Text & "'  ORDER BY 1 "
        'Modify By Cheng 2003/05/05
'      strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022,CP57 FROM R090614,CASEPROGRESS WHERE R110022=CP09(+) AND ID='" & strUserNum & "' AND R110001='" & Combo1.Text & "'  ORDER BY 1 desc"
        'Modify By Cheng 2003/05/08
'      strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024 FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Combo1.Text & "'  ORDER BY R110002 desc "
      'edit by nickc 2007/12/17
      'strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024 FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "'  ORDER BY R110002 desc "
      'Modify by Morgan 2009/7/14 加欄位 R110031
      'Modify by Morgan 2010/11/4 +r110025,r110030
      'Modified by Lydia 2025/02/05 +R110033
      'Modified by Morgan 2025/7/9 +R110034(收文點數)
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028)),'0.00'),R110034,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032 FROM R090614 " & _
                    " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' ORDER BY R110002 desc,R110003,R110004 "
      
      SetGrd1_New 'Added by Lydia 2025/02/05
      CheckOC
      With adoRecordset
          .CursorLocation = adUseClient
            'Modify By Cheng 2003/05/05
'          .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
          .Open strSql, adoEng, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 And .RecordCount > 0 Then
              Set GRD1.Recordset = adoRecordset
              ChgGrdColor
          'Mark by Lydia 2025/02/05
          'Else
          '   GRD1.Clear
          '   GRD1.Rows = 2
          'end 2025/02/05
          End If
      End With
      CheckOC
      SWPRow2 = 1
      SWPRow = 1 'Add By Sindy 2013/8/30
Case 1 '未發文
      'Modify By Cheng 2002/04/16
'      strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110007,R110008,R110009,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022,cp57 FROM R090614,caseprogress WHERE R110022=CP09(+) and ID='" & strUserNum & "' AND R110001='" & Combo1.Text & "' AND R110018 IS NULL and (cp57 is null or cp57=0) order by 1 "
'      strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022,cp57 FROM R090614,caseprogress WHERE R110022=CP09(+) and ID='" & strUserNum & "' AND R110001='" & Combo1.Text & "' AND R110018 IS NULL and (cp57 is null or cp57=0) order by 1 desc"
        'Modify By Cheng 2003/05/08
'      strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024 FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Combo1.Text & "' AND R110018='' and (R110024='' or R110024='0') order by R110002 desc "
      'edit by nickc 2007/12/17
      'strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024 FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' AND R110018='' and (R110024='' or R110024='0') order by R110002 desc "
      'Modify by Morgan 2009/7/14 +R110031
      'Modify by Morgan 2010/11/4 +r110025,r110030
      'Modified by Lydia 2025/02/05 +R110033
      'Modified by Morgan 2025/7/9 +R110034(收文點數)
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028)),'0.00'),R110034,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032 FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' AND R110018='' and (R110024='' or R110024='0') order by R110002 desc "
      SetGrd1_New 'Added by Lydia 2025/02/05
      CheckOC
      With adoRecordset
          .CursorLocation = adUseClient
            'Modify By Cheng 2003/05/05
'          .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
          .Open strSql, adoEng, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 And .RecordCount > 0 Then
              Set GRD1.Recordset = adoRecordset
              ChgGrdColor
          'Mark by Lydia 2025/02/05
          'Else
          '     GRD1.Clear
          '     GRD1.Rows = 2
          '     SetGrd1
          'end 2025/02/05
          End If
      End With
      CheckOC
      SWPRow2 = 1
      SWPRow = 1 'Add By Sindy 2013/8/30
Case Else
End Select
MouseClick (1)
GRD1.Visible = True
'Modify by Morgan 2009/11/12
'Screen.MousePointer = vbDefault
Screen.MousePointer = iMouse
End Sub

'illustration add by nickc 2005/10/31
Private Sub CmdPic_Click()
frmPic001.oCP01 = SystemNumber(LBL1(7), 1)
frmPic001.oCP02 = SystemNumber(LBL1(7), 2)
frmPic001.oCP03 = SystemNumber(LBL1(7), 3)
frmPic001.oCP04 = SystemNumber(LBL1(7), 4)
frmPic001.StrMenu
frmPic001.SetSeekCmdok 'Add by Amy 2018/07/19
frmPic001.Show vbModal
'add by nickc 2005/12/14 檢查有無代表圖
'Modify by Amy 2018/07/19  改寫至function
'strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(Lbl1(7), 1) & "' and ibf02='" & SystemNumber(Lbl1(7), 2) & "' and ibf03='" & SystemNumber(Lbl1(7), 3) & "' and ibf04='" & SystemNumber(Lbl1(7), 4) & "' and ibf05='1' "
'CheckOC2
'adoRecordset1.CursorLocation = adUseClient
'adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
If ChkImgByteFile(SystemNumber(LBL1(7), 1), SystemNumber(LBL1(7), 2), SystemNumber(LBL1(7), 3), SystemNumber(LBL1(7), 4)) = True Then
    cmdPic.Caption = "已設定代表圖(&I)"
    cmdPic.BackColor = &HC0FFC0
    'add by nickc 2007/11/29 加入無圖式的格式
    Chk1.Value = vbUnchecked
    Chk1.Enabled = False
Else
    cmdPic.Caption = "未設定代表圖(&I)"
    cmdPic.BackColor = &HC0C0FF
    'add by nickc 2007/11/29 加入無圖式的格式
    Chk1.Value = vbUnchecked
    Chk1.Enabled = True
End If
'CheckOC2
'end 2018/07/16
End Sub

'Modify By Sindy 2014/1/16
'Private Sub Combo1_Click()
'Modified by Lydia 2021/12/28 Form2.0點選同一人不會觸發Click事件，改用DropButtonClick事件但要控制第2次才執行
'Public Sub Combo1_Click()
''2014/1/16 END
Public Sub Combo1_DropButtonClick()
   Static bClick As Boolean
   If bClick = False Then
      bClick = True
      Exit Sub
   End If
   bClick = False
'end 2021/12/28
   
   Me.Enabled = False 'Add By Sindy 2024/3/14
   
   Dim iMouse As Integer
   iMouse = Screen.MousePointer
      
   Me.GRD1.Visible = False
   Screen.MousePointer = vbHourglass
   Me.MousePointer = vbHourglass
   GRD1.MousePointer = flexArrowHourGlass
   Me.Enabled = False
   Combo1.Enabled = False
   Label3.Caption = GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)))
   'Add By Sindy 2013/9/16 在切換承辦人時,會出現”陣列索引超出範圍”
   If Combo1.Tag <> Combo1.Text Then
      SWPRow = 0
      dblPrevRow = 0
      Combo1.Tag = Combo1.Text
      StrMenu1
      StrMenu
   End If
   '2013/9/16
   
   'Modify By Sindy 2025/5/19
'   If ChkNoData = True Then
'      'Modified by Lydia 2021/12/28 10=>9
'      For s = 0 To 9
'         txt1(s).Enabled = False
'      Next s
'   Else
'      'Modified by Lydia 2021/12/28 10=>9
'      For s = 0 To 9
'         txt1(s).Enabled = True
'      Next s
'   End If
   '2025/5/19 END
   
   'SetGrd1 'Mark by Lydia 2025/02/05
   DoEvents
   'cmdok2(0).SetFocus
   'Modify By Sindy 2025/5/19
'   If cmdok2(0).Visible = True And cmdok2(0).Enabled = True Then cmdok2(0).SetFocus
   '2025/5/19 END
   Combo1.Enabled = True
   Me.Enabled = True
   GRD1.MousePointer = flexDefault
   Me.MousePointer = vbDefault
   'Modify by Morgan 2009/11/12
   'Screen.MousePointer = vbDefault
   Screen.MousePointer = iMouse
   
   Me.GRD1.Visible = True
   
   Me.Enabled = True 'Add By Sindy 2024/3/14
End Sub

Private Sub Combo2_Click()
   txt1(0).Text = Trim(Left(Me.Combo2.Text, 6))
End Sub

Private Sub Combo2_LostFocus()
   Dim ii As Integer
   Dim blnMatch As Boolean

   For ii = 0 To Me.Combo2.ListCount - 1
       blnMatch = False
       If Trim(Left(Me.Combo2.List(ii), 6)) = Trim(Left(Me.Combo2.Text, 6)) Then
           Me.Combo2.ListIndex = ii
           blnMatch = True
           Exit For
       End If
   Next ii
   If blnMatch = False Then Me.Combo2.ListIndex = 0
End Sub

Private Sub Combo4_Click()
   txt1(6).Text = Trim(Left(Me.Combo4.Text, 6))
End Sub

Private Sub Combo4_LostFocus()
   Dim ii As Integer
   Dim blnMatch As Boolean
   
   For ii = 0 To Me.Combo4.ListCount - 1
       blnMatch = False
       If Trim(Left(Me.Combo4.List(ii), 6)) = Trim(Left(Me.Combo4.Text, 6)) Then
           Me.Combo4.ListIndex = ii
           blnMatch = True
           Exit For
       End If
   Next ii
   If blnMatch = False Then Me.Combo4.ListIndex = 0
End Sub

Private Sub Combo5_Click()
   If Me.Visible = True Then
      If QueryData(True) = False Then ShowNoData 'Add By Sindy 2023/4/12
   End If
End Sub

'Add By Sindy 2015/5/21 判發人
Private Sub Combo6_LostFocus()
   Dim ii As Integer
   Dim blnMatch As Boolean
   
   For ii = 0 To Me.Combo6.ListCount - 1
       blnMatch = False
       If Trim(Left(Me.Combo6.List(ii), 6)) = Trim(Left(Me.Combo6.Text, 6)) Then
           Me.Combo6.ListIndex = ii
           blnMatch = True
           Exit For
       End If
   Next ii
   If blnMatch = False Then
      If Trim(Left(Me.Combo6.Text, 6)) = "" Then
         Me.Combo6.ListIndex = 0
      Else
         If Len(GetPrjSalesNM(Trim(Left(Me.Combo6.Text, 6)))) = 0 Then
            'Modify By Sindy 2022/12/6
            'Call ShowStaffErr(Trim(Left(Me.Combo6.Text, 6)))
            Call PUB_GetStaffNameDept(Trim(Left(Me.Combo6.Text, 6)), strExc(10), strExc(0), True, False)
            '2022/12/6 END
            Me.Combo6.SetFocus
            Exit Sub
         Else
            Combo6.Text = UCase(Trim(Left(Me.Combo6.Text, 6))) & " ==> " & GetPrjSalesNM(Trim(Left(Me.Combo6.Text, 6)))
         End If
      End If
   End If
End Sub
'2015/5/21 END

Private Sub Form_Activate()
'Dim nFrm As Form
   ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
   If PUB_ChkFormIsClose("frm090202_2", "承辦") = False Then Exit Sub 'Add By Sindy 2020/1/21
'   'Add By Sindy 2017/8/30
'   '檢查表單是否已開啟，若是，則關閉
'   If Me.Visible = True Then
'      For Each nFrm In Forms
'         If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'            'If frm090202_2.intReceiveKind = 0 Then '0.承辦人工作進度
'            'If frm090202_2.lblCP09.Caption = "" Then Unload frm090202_2: Exit For
'            If UCase(frm090202_2.m_PrevForm.Name) <> UCase(Me.Name) Then Exit For
'            If Not (frm090202_2.cmdAdd.Visible = False And frm090202_2.cmdSend.Enabled = False) Then
'               Unload frm090202_2
'            End If
'         End If
'      Next
'   End If
'   '2017/8/30 END
End Sub

Private Sub Form_Initialize()
   'add by nick 2006/02/27 重新定義
'   ReDim m_FieldList(TF_CP)
End Sub

Private Sub Form_Load()
Dim iMouse As Integer
Dim nFrm As Form 'Add By Sindy 2018/1/24
   
   iMouse = Screen.MousePointer
   
'   'Add By Sindy 2018/1/24
'   '檢查表單是否已開啟，若是，則關閉
'   For Each nFrm In Forms
'      If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'         'Modify By Sindy 2018/10/12 + if
'         '0.承辦人工作進度:又重新登入此作業需要結束上一個已開啟的歷程作業, 因歷程存檔時會使用到此畫面「詳細資料」
'         If frm090202_2.intReceiveKind = 0 Then
'         '2018/10/12 END
'            Unload frm090202_2
'         End If
'         Exit For
'      End If
'   Next
'   '2018/1/24 END
   If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/21
   
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
   
   'add by nick 2006/02/27 重新定義
   ReDim m_FieldList(TF_CP)
   InitialField
   
   'Added by Morgan 2016/2/18
   '配合主畫面,調整表單起始大小( 預設大小 >= 起始大小 <=1024 * Screen.TwipsPerPixelX )
   lngFormWidth = 9435
   lngFormHeight = 6950 'Modified by Lydia 2021/12/28 Height 6120=> 6950
   If Forms(0).Width >= 1024 * Screen.TwipsPerPixelX Then
      lngFormWidth = 1024 * Screen.TwipsPerPixelX - 200
   ElseIf Forms(0).Width >= Me.Width Then
      lngFormWidth = Forms(0).Width - 200
   End If
   Me.Width = lngFormWidth
   Me.Height = lngFormHeight
   'end 2016/2/18
   
   MoveFormToCenter Me
   
   '讀取各基本檔可用系統別
   m_SqlGrpStr1 = SQLGrpStr("", 1)
   m_SqlGrpStr2 = SQLGrpStr("", 2)
   m_SqlGrpStr3 = SQLGrpStr("", 3)
   m_SqlGrpStr4 = SQLGrpStr("", 4)
   m_SqlGrpStr5 = SQLGrpStr("", 5)

   'add by nickc 2006/12/29
   ReDim skMail(0) As SeekMails
   
   Combo5.Text = Combo5.List(3) 'Add By Sindy 2013/9/17
   
   Select Case ProState
   Case "1" '個人
      'add by nickc 2007/12/14
      '讀取使用權限
      Me.Caption = "工作進度資料維護 (個人)" 'Add By Sindy 2024/2/23
      bolInsert = IsUserHasRightOfFunction("frm090201_4", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090201_4", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090201_4", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090201_4", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090201_4", strPrint, False)
      
      TextOk = True
      '統計年月(個人抓系統日的年月)
      Text1.Text = Mid(strSrvDate(1), 1, 6)
      '加乘註記
      'add by nickc 2005/03/04 個人只能看
      txt1(15).Enabled = False
      '加乘註記修改理由
      txtCP99.Enabled = False
      'add by nickc 2006/02/07
      'Frame1.Enabled = False 'Remove by Morgan 2010/10/13
      
      'add by nickc 2006/03/07
      '預定會稿日
      txt1(18).Enabled = False
      'add by nickc 2007/08/21
      '是否暫停核稿
      txt1(20).Enabled = False
      
   Case "2" '主管 承辦人管理工作進度資料查詢
      'add by nickc 2007/12/14
      Me.Caption = "工作進度資料維護 (主管)" 'Add By Sindy 2024/2/23
      bolInsert = IsUserHasRightOfFunction("frm090614", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090614", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090614", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090614", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090614", strPrint, False)
      
      frm090614.TextOk = True
      cmdOK(2).Caption = "回前畫面"
      '統計年月(管理抓查詢畫面輸入的年月)
      Text1.Text = Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2))
      'add by  nickc 2006/02/07
      'Frame1.Enabled = True 'Remove by Morgan 2010/10/13
      
      'Modify by Morgan 2010/10/29 預定會稿日已改上齊備日後由系統預設(隔日凌晨),此處改為可修改
      ''add by nickc 2006/03/07
      ''Modify by Morgan 2008/12/2 改判斷有設定適用規則能改(原來判斷適用的才能改)--此處設定應該無用，因為點選後還會再次設定
      ''If opCP112(0).Value = True Then
      'If m_CPM05 <> "" Then
      '    txt1(18).Enabled = True
      'Else
      '    txt1(18).Enabled = False
      'End If
      If txt1(18) <> "" Then txt1(18).Enabled = True
      'end 2010/10/29
           
      'add by nickc 2007/08/21
      txt1(20).Enabled = True
         
   Case "3" '分所
      'add by nickc 2007/12/14
      bolInsert = IsUserHasRightOfFunction("frm090201_4", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090201_4", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090201_4", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090201_4", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090201_4", strPrint, False)
      
      'add by nickc 2005/03/04 個人只能看
      txt1(15).Enabled = False
      txtCP99.Enabled = False
      'add by nickc 2006/02/07
      'Frame1.Enabled = False 'Remove by Morgan 2010/10/13
      'add by nickc 2006/03/07
      txt1(18).Enabled = False
      'add by nickc 2007/08/21
      txt1(20).Enabled = False
         
   Case Else
      'add by nickc 2006/02/07
      'Frame1.Enabled = False 'Remove by Morgan 2010/10/13
      'add by nickc 2007/08/21
      txt1(20).Enabled = False
      
   End Select
   
   Screen.MousePointer = vbHourglass
   Select Case ProState
   Case "1"
         'Add By Sindy 2013/9/17
         Combo1.AddItem strUserNum & " " & "(" & strUserName & ")", 0
         Combo1.Text = Combo1.List(0)
         '2013/9/17 END
         'StrMenu1 'Modify By Sindy 2016/9/6 因前句Combo1就會run 到 StrMenu1
         StrMenu1  'Added by Lydia 2021/12/28 因為Combo1改成Form 2.0不使用Combo1_Click，所以預設先執行承辦人選單
         
         SetEngineer '設定承辦人選單
         'Add By Sindy 2013/9/16 檢查當時是否需要為他人職代
         Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
         '2013/9/16 END
   Case "2" '承辦人管理工作進度資料查詢
         frm090614.Process1
         StrMenu1
   Case "3"
   Case Else
   End Select
   Label3.Caption = GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)))
      
   SetDrawer '設定繪圖人員選單
   SetEngChecker '設定英文核稿人選單

   StrMenu
   
   Select Case ProState
   Case "1"
      If TextOk = False Then Screen.MousePointer = iMouse: GoTo EXITSUB
      'Add By Sindy 2013/9/16
      'Combo1.Enabled = False
      Combo1.Enabled = True
      '2013/9/16 END
   Case "2"
      If frm090614.TextOk = False Then Screen.MousePointer = iMouse: TextOk = True: GoTo EXITSUB
      Combo1.Enabled = True
      If Left(Pub_StrUserSt03, 2) = "P1" Then cmdOK(3).Visible = True 'Added by Morgan 2024/3/19
   Case "3"
   Case Else
   End Select
   
   'Add by Amy 2014/09/22 取消工程師輸入本所期限
   txt1(13).Visible = False
   Label1(35).Visible = False
   'end 2014/09/22

   'SetGrd1 'Mark by Lydia 2025/02/05

   Call MouseClick(1)
   Screen.MousePointer = iMouse
   SSTab1.Tab = 0
   Me.Combo3.ListIndex = 0
   'Add By Sindy 2013/5/16
'   If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      Me.cmd(1).Visible = True '承辦歷程
      'Me.cmd(2).Visible = False '取消承辦單列印
      'Add By Sindy 2013/6/7
'      If ProState = "1" Then '個人
         Me.SSTab1.TabVisible(2) = True '待辦歷程
         SSTab1.Tab = 2
         If QueryData(True) = False Then
            SSTab1.Tab = 0
         End If
'      Else
'         Me.SSTab1.TabVisible(2) = False
'      End If
      '2013/6/7 End
'   Else
'      Me.cmd(1).Visible = False
'      Me.cmd(2).Visible = True
'      Me.SSTab1.TabVisible(2) = False
'   End If
   '2013/5/16 End
   If bolUpdate = False Then
      cmdOK(1).Visible = False
   End If
   ChkAskList 'Added by Morgan 2015/5/26
   
   Exit Sub

EXITSUB:
   Me.Hide
   Select Case ProState
   Case "1"
        Me.Hide
   Case "2"
        frm090614.Show
        Me.Hide
   Case "3"
   Case Else
   End Select
End Sub
'Added by Morgan 2015/5/25
Private Sub ChkAskList()
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim strNoList As String, arrNo() As String
   
   stSQL = "select al01 from asklist where al02='" & strUserNum & "'"
   
   Pub_SetForOthersEmpCombo strUserNum, , , strNoList
   If strNoList <> "" Then
      arrNo = Split(strNoList, ";")
      For intR = LBound(arrNo) To UBound(arrNo)
         If arrNo(intR) <> "" Then
            stSQL = stSQL & " union select al01 from asklist where al02='" & arrNo(intR) & "'"
         End If
      Next
   End If
   
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With rsQuery
      Do While Not .EOF
         PUB_AskUpdateRelationCase .Fields("al01")
         .MoveNext
      Loop
      End With
   End If
   Set rsQuery = Nothing
End Sub

'Add By Sindy 2013/6/7
Private Sub cmdQuery_Click()
   If QueryData(True) = False Then ShowNoData
End Sub

'Add By Sindy 2013/6/7
Public Function QueryData(bolFirst As Boolean) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
Dim strQuyDate As String 'Add By Sindy 2013/9/17
Dim strVal As String 'Add By Sindy 2013/10/22
   
   m_blnColOrderAsc = True
   QueryData = True
   
   'Add By Sindy 2013/9/17
   If Combo5.ListIndex = 0 Then
      strQuyDate = CompWorkDay(3, strSrvDate(1), 1) '不含當天,3個工作天
   ElseIf Combo5.ListIndex = 1 Then
      strQuyDate = CompWorkDay(5, strSrvDate(1), 1) '不含當天,5個工作天
   ElseIf Combo5.ListIndex = 2 Then
      strQuyDate = CompWorkDay(7, strSrvDate(1), 1) '不含當天,7個工作天
   Else
      '全部
   End If
   '2013/9/17 END
   
   grd2.Clear
   SetGrd2
   
   Screen.MousePointer = vbHourglass
   
   'Modify By Sindy 2013/10/22
''   strVal = "(select * from EmpElectronProcess where eep01||eep02 in(select eep01||max(eep02) from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and (EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ") or (EEP04='" & EMP_聯絡 & "'" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & ")) group by eep01) and EEP09 is null and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
''            " union select EmpElectronProcess.* from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and eep04='" & EMP_退件 & "' and eep09='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
''            ") EmpElectronProcess,"
'   strVal = "(select * from EmpElectronProcess where eep01||eep02 in(select eep01||max(eep02) from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ") group by eep01) and EEP09 is null and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            " union select e1.* from EmpElectronProcess e1,caseprogress where e1.eep01=cp09(+) and cp27 is null and cp57 is null and e1.EEP02 in (select max(eep02) from EmpElectronProcess where eep01=e1.eep01) and e1.EEP05='" & Trim(Left(Combo1.text, 6)) & "' And e1.EEP04 in('" & EMP_聯絡 & "')" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "") & _
'            " union select EmpElectronProcess.* from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and eep04='" & EMP_退件 & "' and eep09='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
'            ") EmpElectronProcess,"
   '2013/10/22 END
   'Modify By Sindy 2016/3/3 取消此句,因退件不會上待回覆Y " union select EmpElectronProcess.* from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and eep04='" & EMP_退件 & "' and eep09='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   '                         增加EEP13='Y'
   strVal = "select eep01,eep02 from EmpElectronProcess where EEP13='Y' and EEP09 is null and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ",'" & EMP_判發 & "')" & _
            " union select eep01,eep02 from EmpElectronProcess where EEP13='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and EEP04='" & EMP_聯絡 & "'" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "")
   '2016/3/3 END
   'Modify By Sindy 2015/9/30 +IIf(Pub_StrUserSt15 = "P12", " And EEP04 not in('" & EMP_判發 & "','" & EMP_退件重送 & "')", "")
   'Modify By Sindy 2016/3/3 +不顯示
   'Modify By Sindy 2016/9/2 And cp27 is null And cp57 is null -> and cp158=0 and cp159=0
   strSql = "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,EP08,EP38,' ' as 不顯示,EEP06 a,EEP07 b" & _
            " From EmpElectronProcess,CaseProgress,EngineerProgress,Patent," & _
            "staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " Where (EEP01,EEP02) in(" & strVal & ")" & _
            " and EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And PA09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)"
   'If ProState = "1" Then '個人
      strSql = strSql & " And CP14='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   'End If
   'Add By Sindy 2018/4/17
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05||SP06||SP07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(SP09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,EP08,EP38,' ' as 不顯示,EEP06 a,EEP07 b" & _
            " From EmpElectronProcess,CaseProgress,EngineerProgress,servicepractice," & _
            "staff s1,staff s2,nation,CasePropertyMap,allcode" & _
            " Where (EEP01,EEP02) in(" & strVal & ")" & _
            " and EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And SP09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)" & _
            IIf(Pub_StrUserSt15 = "P12", " And EEP04 not in('" & EMP_判發 & "','" & EMP_退件重送 & "')", "")
   'If ProState = "1" Then '個人
      strSql = strSql & " And CP14='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   'End If
   'Modify By Sindy 2013/11/21
   'strSql = strSql & " order by EP01 desc"
   '2018/4/17 END
   strSql = strSql & " order by a desc,b desc"
   '2013/11/21 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd2.Recordset = rsTmp
      
      'Add By Sindy 2013/10/18
      For i = 1 To grd2.Rows - 1
         Call SetColColor(i)
      Next i
      '2013/10/18 END
      
      'Add By Sindy 2013/8/20
      '檢查是否有會完流程,但尚未上會稿完成日的資料,若有,則更新EP08會稿完成日=EP38業務會稿完成日
'      For i = 1 To grd2.Rows - 1
'         If Val(grd2.TextMatrix(i, 16)) > 0 And Val(grd2.TextMatrix(i, 15)) = 0 Then '有業務會稿完成日 且 無會稿完成日
''            strExc(0) = "SELECT count(*) FROM EmpElectronProcess WHERE EEP01='" & grd2.TextMatrix(i, 13) & "' AND EEP04='" & EMP_會完 & "'"
''            intI = 1
''            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''            If intI = 1 Then
''               If RsTemp.Fields(0) > 0 Then '有會完流程
'               If grd2.TextMatrix(i, 12) = "會完" Then
'                  If lbl1(3) <> grd2.TextMatrix(i, 13) Then
'                     Call Process(grd2.TextMatrix(i, 13))
'                  End If
'                  Me.SSTab1.Tab = 1
'                  Me.txt1(7) = Val(grd2.TextMatrix(i, 16)) - 19110000 '更新EP08會稿完成日=EP38業務會稿完成日
'                  Call cmdOK_Click(1) '存檔
'                  If m_chkcmdok1 = False Then GoTo ExitQuery
'               End If
''            End If
'         End If
'      Next i
      '2013/8/20 END
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
      
ExitQuery:
   '若有資料時游標停在第一筆
   If bolFirst = True Then
      grd2.Visible = False
      grd2.col = 0
      grd2.row = 1
      If rsTmp.RecordCount > 0 Then
         dblPrevRow = grd2.row
         grd2.Text = "V"
         m_intRow = 1: m_intCol = 0 'Add By Sindy 2016/3/10
         For i = 0 To grd2.Cols - 1
            grd2.col = i
            'Modify By Sindy 2013/10/29
            If grd2.CellBackColor <> &H8080FF Then
               grd2.CellBackColor = &HFFC0C0
            End If
         Next i
      End If
      grd2.Visible = True
   End If
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   'Me.SSTab1.Tab = 2 'Modify By Sindy 2017/8/21 Mark
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2013/10/18
'已確認過會稿完成日時以紅色標註...2013/10/22取消顏色
Private Sub SetColColor(intRow As Integer)
Dim i As Integer
   
   grd2.row = intRow
   'Modified by Lydia 2025/02/05 改用變數
'   If grd2.TextMatrix(intRow, 12) = "會完" Then
'      If PUB_ChkEmpFlowExists(grd2.TextMatrix(intRow, 13), EMP_不自動更新會完日, grd2.TextMatrix(intRow, 14)) = True Then
'         grd2.TextMatrix(intRow, 12) = "N" & grd2.TextMatrix(intRow, 12)
'         'Modify By Sindy 2013/10/29
'         grd2.col = 12
'         grd2.CellBackColor = &H8080FF
'      ElseIf Val(grd2.TextMatrix(intRow, 15)) > 0 Then
'         grd2.TextMatrix(intRow, 12) = "Y" & grd2.TextMatrix(intRow, 12)
'      End If
''      '已確認過會稿完成日
''      If Val(grd2.TextMatrix(intRow, 15)) > 0 Or _
''         PUB_ChkEmpFlowExists(grd2.TextMatrix(intRow, 13), EMP_不自動更新會完日, grd2.TextMatrix(intRow, 14)) = True Then
''         grd2.TextMatrix(intRow, 12) = "●" & grd2.TextMatrix(intRow, 12)
'         'Modify By Sindy 2013/10/22 取消顏色
''         For i = 0 To grd2.Cols - 1
''            grd2.col = i
''            grd2.CellBackColor = &H8080FF
''         Next i
''      End If
'   'Add By Sindy 2016/3/7 柏翰:繪圖判發跟退件要淺紅色表示
'   ElseIf grd2.TextMatrix(intRow, 12) = "繪圖判發" Or _
'          grd2.TextMatrix(intRow, 12) = "退件" Then
'      grd2.col = 12
'      grd2.CellBackColor = &HC0C0FF
'   End If
   If grd2.TextMatrix(intRow, colFS_2) = "會完" Then
      If PUB_ChkEmpFlowExists(grd2.TextMatrix(intRow, colCP09_2), EMP_不自動更新會完日, grd2.TextMatrix(intRow, colXno_2)) = True Then
         grd2.TextMatrix(intRow, colFS_2) = "N" & grd2.TextMatrix(intRow, colFS_2)
         'Modify By Sindy 20013/10/29
         grd2.col = colFS_2
         grd2.CellBackColor = &H8080FF
      ElseIf Val(grd2.TextMatrix(intRow, colEP08_2)) > 0 Then
         grd2.TextMatrix(intRow, colFS_2) = "Y" & grd2.TextMatrix(intRow, colFS_2)
      End If
   'Add By Sindy 2016/3/7 柏翰:繪圖判發跟退件要淺紅色表示
   ElseIf grd2.TextMatrix(intRow, colFS_2) = "繪圖判發" Or _
          grd2.TextMatrix(intRow, colFS_2) = "退件" Then
      grd2.col = colFS_2
      grd2.CellBackColor = &HC0C0FF
   End If
   'end 2025/02/05
End Sub

'Add By Sindy 2013/6/7
Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2016/3/3 +不顯示
   arrGridHeadText = Array("V", "目次", "流程日期", "本所案號", "案件名稱", _
                           "國家", "種類", "案件性質", "本所期限", "承辦人", _
                           "承辦期限", "智權人員", "目前流程狀態", _
                           "總收文號", "序號", "EP08", "EP38", "不顯示", "EEP06 a", "EEP07 b")
   arrGridHeadWidth = Array(200, 400, 800, 1400, 1000, _
                            700, 450, 900, 800, 600, _
                            800, 600, 600, _
                            0, 0, 0, 0, 600, 0, 0)
   grd2.Visible = False
   grd2.Cols = UBound(arrGridHeadText) + 1
   grd2.Rows = 2
   For iRow = 0 To grd2.Cols - 1
      grd2.row = 0
      grd2.col = iRow
      grd2.Text = arrGridHeadText(iRow)
      grd2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      If iRow = 11 Or iRow = 12 Then
         grd2.CellAlignment = flexAlignLeftCenter
      Else
         grd2.CellAlignment = flexAlignCenterCenter
      End If
   Next
   'Added by Lydia 2025/02/05
   If colFS_2 = 0 Then
      colFS_2 = PUB_MGridGetId("目前流程狀態", grd2)
      colCP09_2 = PUB_MGridGetId("總收文號", grd2)
      colXno_2 = PUB_MGridGetId("序號", grd2)
      colEP08_2 = PUB_MGridGetId("EP08", grd2)
      colNoShow_2 = PUB_MGridGetId("不顯示", grd2)
      colCaseNo_2 = PUB_MGridGetId("本所案號", grd2)
   End If
   'end 2025/02/05
   grd2.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   'Add By Sindy 2020/3/19
'   If MsgBox("未完成稿件是否已上傳暫存區？", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption & " 重要訊息！") = vbNo Then
'      Cancel = True
'      Exit Sub
'   End If
'   '2020/3/19 END
End Sub

'Added by Morgan 2016/2/18
Private Sub Form_Resize()
   'Modified by Morgan 2016/3/11 不可最大化
   If Me.WindowState = 2 Then
      Me.WindowState = 0
   'end 2016/3/11
   ElseIf Me.WindowState = 0 Then
      If Me.Width < lngFormWidth Then Me.Width = lngFormWidth
      Me.Height = lngFormHeight '高度固定
      Me.SSTab1.Width = Me.Width - 200
      Me.GRD1.Width = Me.SSTab1.Width - 200
   End If
End Sub

'Add By Sindy 2016/3/3 增加不顯示功能
Private Sub grd2_Click()
   m_intRow = grd2.MouseRow
   m_intCol = grd2.MouseCol
   If m_intRow <> 0 Then
      'Modified by Lydia 2025/02/05 改用變數
'      If m_intCol = 17 Then '不顯示
'         If grd2.TextMatrix(m_intRow, 13) <> "" And _
'            grd2.TextMatrix(m_intRow, 12) <> "核修" And _
'            grd2.TextMatrix(m_intRow, 12) <> "核完" And _
'            grd2.TextMatrix(m_intRow, 12) <> "會修" And _
'            InStr(grd2.TextMatrix(m_intRow, 12), "會完") = 0 And _
'            grd2.TextMatrix(m_intRow, 12) <> "繪圖判發" And _
'            grd2.TextMatrix(m_intRow, 12) <> "判發" And _
'            grd2.TextMatrix(m_intRow, 12) <> "退回" And _
'            grd2.TextMatrix(m_intRow, 12) <> "退件" And _
'            grd2.TextMatrix(m_intRow, 12) <> "圖修" And _
'            InStr(grd2.TextMatrix(m_intRow, 12), "圖完") = 0 Then
'            grd2.TextMatrix(m_intRow, 17) = "V"
'            If MsgBox("請再次確定不顯示 " & vbCrLf & grd2.TextMatrix(m_intRow, 3) & " " & grd2.TextMatrix(m_intRow, 12) & " 嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'               grd2.TextMatrix(m_intRow, 17) = ""
'            Else
'               strExc(0) = "update EmpElectronProcess set eep13=null" & _
'                           " where eep01='" & grd2.TextMatrix(m_intRow, 13) & "'" & _
'                             " and eep02=" & grd2.TextMatrix(m_intRow, 14)
'               Pub_SeekTbLog strExc(0) 'Add By Sindy 2018/8/27
'               cnnConnection.Execute strExc(0)
'               grd2.RowHeight(m_intRow) = 0
'            End If
'         End If
'      End If
      If m_intCol = colNoShow_2 Then '不顯示
         If grd2.TextMatrix(m_intRow, colCP09_2) <> "" And _
            grd2.TextMatrix(m_intRow, colFS_2) <> "核修" And _
            grd2.TextMatrix(m_intRow, colFS_2) <> "核完" And _
            grd2.TextMatrix(m_intRow, colFS_2) <> "會修" And _
            InStr(grd2.TextMatrix(m_intRow, colFS_2), "會完") = 0 And _
            grd2.TextMatrix(m_intRow, colFS_2) <> "繪圖判發" And _
            grd2.TextMatrix(m_intRow, colFS_2) <> "判發" And _
            grd2.TextMatrix(m_intRow, colFS_2) <> "退回" And _
            grd2.TextMatrix(m_intRow, colFS_2) <> "退件" And _
            grd2.TextMatrix(m_intRow, colFS_2) <> "圖修" And _
            InStr(grd2.TextMatrix(m_intRow, colFS_2), "圖完") = 0 Then
            grd2.TextMatrix(m_intRow, colNoShow_2) = "V"
            If MsgBox("請再次確定不顯示 " & vbCrLf & grd2.TextMatrix(m_intRow, colCaseNo_2) & " " & grd2.TextMatrix(m_intRow, colFS_2) & " 嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               grd2.TextMatrix(m_intRow, colNoShow_2) = ""
            Else
               strExc(0) = "update EmpElectronProcess set eep13=null" & _
                           " where eep01='" & grd2.TextMatrix(m_intRow, colCP09_2) & "'" & _
                             " and eep02=" & grd2.TextMatrix(m_intRow, colXno_2)
               Pub_SeekTbLog strExc(0) 'Add By Sindy 2018/8/27
               cnnConnection.Execute strExc(0)
               grd2.RowHeight(m_intRow) = 0
            End If
         End If
      End If  '---'不顯示
      'end 2025/02/05
   End If
End Sub

'Add By Sindy 2013/6/7
Private Sub grd2_DblClick()
Dim nFrm As Form
   
   'Modify By Sindy 2016/3/3
   If m_intRow <> 0 Then
      If m_intCol <> 17 Then
   '2016/3/3 END
         'Add By Sindy 2013/9/16
         If ProState = "2" Then
            If frm090614.txt1(8) = "N" Then MsgBox "不可從（不區分個人）的資料查詢中來執行承辦歷程作業！": Exit Sub
         End If
         '2013/9/16 END
         
         'Add By Sindy 2018/1/2 個人案件不可用主管權限操作
         If ProState = "2" And m_CP14 = strUserNum Then '2.主管
            MsgBox "個人案件不可用主管權限操作！", vbExclamation
            Exit Sub
         End If
         '2018/1/2 END
         
         'For i = 1 To grd2.Rows - 1
            'Add By Sindy 2017/11/9
            If dblPrevRow = 0 Then
               MsgBox "請點選一筆資料列!", vbExclamation
               Exit Sub
            End If
            '2017/11/9 END
            If grd2.TextMatrix(dblPrevRow, 0) = "V" Then
      '         If lbl1(3) <> grd2.TextMatrix(dblPrevRow, 13) Then
                  'Modified by Lydia 2025/02/05 改用變數
                  'Call Process(grd2.TextMatrix(dblPrevRow, 13)) 'Modify By Sindy 2013/10/28 要重新查詢資料,因核稿人及判發人有預設問題 ex.P106408品薇在新增下一流程會變自行判發
                  Call Process(grd2.TextMatrix(dblPrevRow, colCP09_2))
      '         Else
'Modify By Sindy 2017/9/15 Mark
'                  If Me.cmd(1).Enabled = True Then
'                     If SetColTag(False) = False Then
'                        Call cmdok_Click(1)
'                        If m_chkcmdok1 = False Then Exit Sub
'                     End If
'                  End If
      '         End If
               If Me.cmd(1).Enabled = True Then
                  'Add By Sindy 2015/12/3
                  '重新檢查欄位有效性
                  If TxtValidate = True Then
                  '2015/12/3 END
                     
'                     'Add By Sindy 2017/9/19
'                     '檢查表單是否已開啟，若是，則關閉
'                     For Each nFrm In Forms
'                        If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'                           Unload frm090202_2
'                           Exit For
'                        End If
'                     Next
'                     '2017/9/19 END
                     If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
                     intBackTab = 2
                     frm090202_2.Hide
                     'Modified by Lydia 2025/02/05 改用變數
                     'frm090202_2.m_EEP01 = grd2.TextMatrix(dblPrevRow, 13) '總收文號
                     frm090202_2.m_EEP01 = grd2.TextMatrix(dblPrevRow, colCP09_2) '總收文號
                     frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) 'Add By Sindy 2013/9/12 案件流程所屬人員
                     frm090202_2.intReceiveKind = 0 '0.承辦人工作進度
                     frm090202_2.SetParent Me
                     If frm090202_2.QueryData = True Then
                        frm090202_2.Show
                        Me.Hide
                     End If
                     'Exit For
                  End If
               Else
                  Me.SSTab1.Tab = 1
               End If
            End If
         'Next i
      End If
   End If
End Sub

'Add By Sindy 2013/6/7
Private Sub GRD2_SelChange()
Dim j As Integer 'Add By Sindy 2016/3/4

grd2.Visible = False
'Add By Sindy 2016/3/4
If grd2.MouseRow = 0 Then
   '已選取的資料列清除反白
   For j = 1 To grd2.Rows - 1
      If grd2.TextMatrix(j, 0) = "V" Then
         grd2.col = 0
         grd2.row = j
         grd2.Text = ""
         For i = 0 To grd2.Cols - 1
            grd2.col = i
            grd2.CellBackColor = QBColor(15)
         Next i
         Call SetColColor(j)
         Exit For
      End If
   Next j
Else
'2016/3/4 END
   '上一筆資料列清除反白
   'Modify By Sindy 2016/5/9
   'If dblPrevRow > 0 Then
   If dblPrevRow > 0 And dblPrevRow <= (grd2.Rows - 1) Then
   '2016/5/9 END
      grd2.col = 0
      grd2.row = dblPrevRow
      grd2.Text = ""
      For i = 0 To grd2.Cols - 1
         grd2.col = i
         'Modify By Sindy 2013/10/29
         If grd2.CellBackColor <> &H8080FF Then
         '2013/10/29 END
            grd2.CellBackColor = QBColor(15)
         End If
      Next i
      Call SetColColor(CStr(dblPrevRow))
   End If
   '目前資料列反白
   grd2.col = 0
   grd2.row = grd2.MouseRow
   dblPrevRow = grd2.row
'   If GRD2.Text = "V" Then
'      GRD2.Text = ""
'      For i = 0 To GRD2.Cols - 1
'         GRD2.col = i
'         GRD2.CellBackColor = QBColor(15)
'      Next i
'   Else
      If grd2.TextMatrix(grd2.row, 1) <> "" Then
         grd2.Text = "V"
         For i = 0 To grd2.Cols - 1
            grd2.col = i
            'Modify By Sindy 2013/10/29
            If grd2.CellBackColor <> &H8080FF Then
            '2013/10/29 END
               grd2.CellBackColor = &HFFC0C0
            End If
         Next i
      End If
'   End If
End If
grd2.Visible = True
End Sub

'Add By Sindy 2013/6/7
Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grd2, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'grd2.col = nCol
   grd2.row = nRow
   If Me.grd2.row < 1 And Me.grd2.Text <> "V" Then
      If Me.grd2.Text = "目次" Then
         If m_blnColOrderAsc = True Then
            Me.grd2.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd2.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grd2.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd2.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Dir(DocTempPath & "\*.doc") <> "" Then 'Added by Lydia 2018/05/24 deubg
    Set Fobj = New FileSystemObject
    Fobj.DeleteFile DocTempPath & "\*.doc", False
End If
'add by nickc 2006/02/27
ClearFieldList
Set Fobj = Nothing
Set frm090201_2 = Nothing
End Sub

Sub StrMenu1()
Me.Enabled = False
'grd1.MousePointer = vbHourglass
DoEvents
'Modify By Cheng 2003/04/29
'cnnConnection.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "
'edit by nickc 2005/05/04 加欄位
'adoEng.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "
On Error GoTo ErrHnd 'Add By Sindy 2024/3/14
adoEng.Execute "drop table R090614 "
'edit by nickc 2007/11/27
'adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text,R110006 text,R110007 text,R110008 text,R110009 text,R110010 text,R110011 text,R110012 text,R110013 text,R110014 text,R110015 text,R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo,R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text)"
'Modify by Morgan 2009/7/14 +R110031
'Modify by Morgan 2011/8/3 +R110032(支援+修改+衍生基數)
RunCreateTable: 'Add By Sindy 2024/3/14
'Modify By Sindy 2024/3/29 為配合外專此暫存檔也增加 ,R110033 text,R110034 text,R110035 text 欄位
adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text" & _
               ",R110006 text,R110007 text,R110008 text,R110009 text,R110010 text" & _
               ",R110011 text,R110012 text,R110013 text,R110014 text,R110015 text" & _
               ",R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo" & _
               ",R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text" & _
               ",R110026 double,R110027 double,R110028 double,R110029 text,R110030 text" & _
               ",R110031 text,R110032 double,R110033 text,R110034 text,R110035 text)"
On Error GoTo 0 'Add by Sindy 2024/3/14 還原錯誤控制

Select Case ProState
Case "1" '承辦人個人工作進度資料維護
      StrGrp090201 = ""
      StrSQL6 = ""
      strSQL1 = ""
      strSQL2 = ""
      'add by nick 2004/11/23
      StrSQL61 = ""
      StrSQL62 = ""
      StrSQL63 = ""
      StrSQL64 = ""
      'add by nickc 2006/04/07
      StrSPa = ""
      StrSTM = ""
      StrSLC = ""
      StrSHC = ""
      StrSSP = ""
        'Modify By Cheng 2003/03/25
        '含齊備日為當月, 發文日為19221111的資料
'      StrSQL6 = StrSQL6 + " and CP14='" & strUserNum & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(GetTodayDate, 1, 6) & "01 AND CP27<=" & Mid(GetTodayDate, 1, 6) & "31 and cp57 is null ) or (CP57>=" & Mid(GetTodayDate, 1, 6) & "01 AND CP57<=" & Mid(GetTodayDate, 1, 6) & "31 and cp27 is null))) and cp05>=19980101"
        'Modify By Cheng 2003/04/28
'      StrSQL6 = StrSQL6 + " and CP14='" & strUserNum & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(GetTodayDate, 1, 6) & "01 AND CP27<=" & Mid(GetTodayDate, 1, 6) & "31 and cp57 is null ) or (CP57>=" & Mid(GetTodayDate, 1, 6) & "01 AND CP57<=" & Mid(GetTodayDate, 1, 6) & "31 and cp27 is null))) and cp05>=19980101"
'      strSQL1 = strSQL1 & " and CP14='" & strUserNum & "' AND (SUBSTR(CP27,1,6)=" & Mid(GetTodayDate, 1, 6) & " or SUBSTR(CP57,1,6)=" & Mid(GetTodayDate, 1, 6) & ") and cp05>=19980101"
        'Modify By Cheng 2003/05/26
        '加收文日為當月, 發文日小於收文日的資料
'      StrSQL6 = StrSQL6 + " and CP14='" & strUserNum & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null ) or (CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null))) and cp05>=19980101"
      'edit by nick 2004/11/23
      'StrSQL6 = StrSQL6 + " and CP14='" & strUserNum & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null ) or (CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null) " & _
                        " Or (CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP05<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null And CP05>CP27 ))) and cp05>=19980101"
       'Modify By Sindy 2013/9/17
       'StrSQL6 = StrSQL6 & " and CP14='" & strUserNum & "' and cp05>=19980101 "
       StrSQL6 = StrSQL6 & " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' and cp05>=19980101 "
       '2013/9/17 END
       'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
       StrSQL61 = StrSQL61 & " and cp158=0 and cp159=0 "
       'edit by nickc 2005/05/13
       'StrSQL62 = StrSQL62 & " and CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null  "
       StrSQL62 = StrSQL62 & " and CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 "
       StrSQL63 = StrSQL63 & " and CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null "
       StrSQL64 = StrSQL64 & " and CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP05<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null And CP05>CP27  "
       'add by nickc 2006/04/07
       'Modified by Morgan 2015/8/28 有發文日者除外
       'Mofified by Morgan 2016/8/4 C類未發文的不管閉不閉卷承辦人及判發人都要能在系統中作業 --玲玲 Ex.P-74098
       StrSPa = StrSPa & " and ((pa58>=" & Mid(strSrvDate(1), 1, 6) & "01 AND pa58<=" & Mid(strSrvDate(1), 1, 6) & "31) or pa58 is null or cp27>0 or (cp01='P' and cp09>'C')) "
       StrSTM = StrSTM & " and ((tm30>=" & Mid(strSrvDate(1), 1, 6) & "01 AND tm30<=" & Mid(strSrvDate(1), 1, 6) & "31) or tm30 is null or cp27>0) "
       StrSLC = StrSLC & " and ((lc09>=" & Mid(strSrvDate(1), 1, 6) & "01 AND lc09<=" & Mid(strSrvDate(1), 1, 6) & "31) or lc09 is null or cp27>0) "
       StrSHC = StrSHC & " and ((hc10>=" & Mid(strSrvDate(1), 1, 6) & "01 AND hc10<=" & Mid(strSrvDate(1), 1, 6) & "31) or hc10 is null or cp27>0) "
       StrSSP = StrSSP & " and ((sp16>=" & Mid(strSrvDate(1), 1, 6) & "01 AND sp16<=" & Mid(strSrvDate(1), 1, 6) & "31) or sp16 is null or cp27>0) "

Case "2" '承辦人管理工作進度資料查詢
      StrGrp090201 = frm090614.ManaGrp
      'Modify By Cheng 2002/04/22
      '改成收文日要小於等於發文年月當月的最後一天
      StrSQL6 = " and cp05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 "
      '92.6.26 add by sonia
      If Len(Trim(frm090614.txt1(6))) <> 0 Then
         StrSQL6 = StrSQL6 & " AND S1.ST03>='" & frm090614.txt1(6) & "' "
      End If
      If Len(Trim(frm090614.txt1(7))) <> 0 Then
         StrSQL6 = StrSQL6 & " AND S1.ST03<='" & frm090614.txt1(7) & "' "
      End If
      '92.6.26 end
      strSQL1 = ""
      strSQL2 = ""
      'add by nick 2004/11/23
      StrSQL61 = ""
      StrSQL62 = ""
      StrSQL63 = ""
      StrSQL64 = ""
      'add by nickc 2006/04/07
      'Modified by Morgan 2015/8/28 有發文日者除外
      'Mofified by Morgan 2016/8/4 C類未發文的不管閉不閉卷承辦人及判發人都要能在系統中作業 --玲玲 Ex.P-74098
      StrSPa = " and ((pa58>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and pa58<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or pa58 is null or cp27>0 or (cp01='P' and cp09>'C')) "
      StrSTM = " and ((tm30>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and tm30<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or tm30 is null or cp27>0) "
      StrSLC = " and ((lc09>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and lc09<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or lc09 is null or cp27>0) "
      StrSHC = " and ((hc10>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and hc10<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or hc10 is null or cp27>0) "
      StrSSP = " and ((sp16>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and sp16<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or sp16 is null or cp27>0) "
'      StrSQL6 = StrSQL6 + " and CP14='" & Trim(Combo1.Text) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp27 is null))) and cp05>=19980101"
'      strSQL1 = strSQL1 & " and CP14='" & Trim(Combo1.Text) & "' AND ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 )) and cp05>=19980101"
'      StrSQL6 = StrSQL6 + " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp27 is null))) and cp05>=19980101"
        'Modify By Cheng 2003/05/26
        '加收文日為當月, 發文日小於收文日的資料
'      StrSQL6 = StrSQL6 + " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp27 is null))) and cp05>=19980101"
      '92.6.26 modify by sonia
      'StrSQL6 = StrSQL6 + " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp27 is null) " & _
      '                  " Or (CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27))) and cp05>=19980101"
      If frm090614.txt1(8) = "N" Then
        'Modify By Cheng 2003/07/18
        '不限制發文日止日及取消收文日止日
'         StrSQL6 = StrSQL6 + " and CP14 IN (" & Combo1_String & ") AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp27 is null) " & _
'                           " Or (CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27))) and cp05>=19980101"
         'edit by nick 2004/11/23
         'StrSQL6 = StrSQL6 + " and CP14 IN (" & Combo1_String & ") AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27 is null) " & _
                           " Or (CP05>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27))) and cp05>=19980101"
         StrSQL6 = StrSQL6 & " and CP14 IN (" & Combo1_String & ")  and cp05>=19980101 "
         'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
         StrSQL61 = StrSQL61 & " and cp158=0 and cp159=0 "
         'edit by nickc 2005/05/13
         'StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp57 is null "
         StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 "
         StrSQL63 = StrSQL63 & " and CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27 is null "
         StrSQL64 = StrSQL64 & " and CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27 "
      Else
        'Modify By Cheng 2003/07/18
        '不限制發文日止日及取消收文日止日
'         StrSQL6 = StrSQL6 + " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp27 is null) " & _
'                           " Or (CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27))) and cp05>=19980101"
         'edit by nick 2004/11/23
         'StrSQL6 = StrSQL6 + " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57 is null) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27 is null) " & _
                           " Or (CP05>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27))) and cp05>=19980101"
         StrSQL6 = StrSQL6 & " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "'  and cp05>=19980101 "
         'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
         StrSQL61 = StrSQL61 & " and cp158=0 and cp159=0 "
         'edit by nickc 2005/05/13
         'StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp57 is null "
         StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 "
         StrSQL63 = StrSQL63 & " and CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and cp27 is null "
         StrSQL64 = StrSQL64 & " and CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp57 is null And CP05>CP27 "
      End If
      '92.6.26 end
'      strSQL1 = strSQL1 & " and CP14='" & Trim(Left("" & Combo1.Text, 6)) & "' AND ((CP27>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp27<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31) or (CP57>=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "01 and cp57<=" & Trim((Val(frm090614.Txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.Txt1(4)), 2)) & "31 )) and cp05>=19980101"
Case Else
End Select

CheckOC

'Modify by Morgan 2009/7/14 +ep28

'第一次
'Modify by Morgan 2011/1/3 修正日期欄位排序問題(小於100年的前面補空白)
'Modify By Sindy 2024/3/29 為配合外專此暫存檔也增加 ,R110033 text,R110034 text,R110035 text 欄位
'Modified by Lydia 2025/02/05 內專工程師增加顯示「指定日期」CP142=R110033>> '','',''取代為,SQLDATET2(CP142) AS CP142,'',''
'Modified by Morgan 2025/7/9 +CP18
strSql = "SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'020',PTM04,PTM03),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),PGMID.CheckCuInAD(PA01,PA09,PA26,PA27,PA28,PA29,PA30)||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,PA58)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
               " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & StrSQL61 & StrSPa & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL61 & StrSTM & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") "
strSql = strSql + " UNION   SELECT CP14,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL61 & StrSLC & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL61 & StrSHC & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL61 & StrSSP & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") "
'AddToMdb (strSQL)
'第二次
'Modify by Morgan 2011/1/3 修正日期欄位排序問題(小於100年的前面補空白)
'Modified by Lydia 2025/02/05 內專工程師增加顯示「指定日期」CP142=R110033>> '','',''取代為,SQLDATET2(CP142) AS CP142,'',''
'Modified by Morgan 2025/7/9 +CP18
strSql = strSql + " union SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'020',PTM04,PTM03),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),PGMID.CheckCuInAD(PA01,PA09,PA26,PA27,PA28,PA29,PA30)||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,PA58)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
               " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & StrSQL62 & StrSPa & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL62 & StrSTM & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") "
strSql = strSql + " UNION   SELECT CP14,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL62 & StrSLC & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL62 & StrSHC & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL62 & StrSSP & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") "
'AddToMdb (strSQL)
'第三次
'Modify by Morgan 2011/1/3 修正日期欄位排序問題(小於100年的前面補空白)
'Modified by Lydia 2025/02/05 內專工程師增加顯示「指定日期」CP142=R110033>> '','',''取代為,SQLDATET2(CP142) AS CP142,'',''
'Modified by Morgan 2025/7/9 +CP18
strSql = strSql + " union SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'020',PTM04,PTM03),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),PGMID.CheckCuInAD(PA01,PA09,PA26,PA27,PA28,PA29,PA30)||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,PA58)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
               " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & StrSQL63 & StrSPa & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL63 & StrSTM & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") "
strSql = strSql + " UNION   SELECT CP14,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL63 & StrSLC & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL63 & StrSHC & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL63 & StrSSP & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") "
'AddToMdb (strSQL)
'第四次
'Modify by Morgan 2011/1/3 修正日期欄位排序問題(小於100年的前面補空白)
'Modified by Lydia 2025/02/05 內專工程師增加顯示「指定日期」CP142=R110033>> '','',''取代為,SQLDATET2(CP142) AS CP142,'',''
'Modified by Morgan 2025/7/9 +CP18
strSql = strSql + " union SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'020',PTM04,PTM03),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),PGMID.CheckCuInAD(PA01,PA09,PA26,PA27,PA28,PA29,PA30)||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,PA58)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
               " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & StrSQL64 & StrSPa & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL64 & StrSTM & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") "
strSql = strSql + " UNION   SELECT CP14,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL64 & StrSLC & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL64 & StrSHC & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION   SELECT CP14,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,SQLDATET2(CP142) AS CP142,cp18,''" & _
               " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
               " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL64 & StrSSP & _
               " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") "
'add by nickc 2005/04/27
AddToMdb (strSql)

'add by nickc 2008/01/03 加入累計
Dim theWorkDay As Integer    '截到目前工作天
Dim TheNowWorks As String
Dim StrSQLa As String
Select Case ProState
Case "1" '承辦人個人工作進度資料維護
      StrSQL6 = ""
      'Modify By Sindy 2013/9/17
      'StrSQL6 = StrSQL6 & " ma01='" & strUserNum & "' "
      StrSQL6 = StrSQL6 & " ma01='" & Trim(Left("" & Combo1.Text, 6)) & "' "
      '2013/9/17 END
      StrSQL6 = StrSQL6 & " and ma02='" & Mid(strSrvDate(1), 1, 6) & "' and ma03='1' "
      StrSQLa = "Select Count(*) From WorkDay Where WD01>='" & Mid(strSrvDate(1), 1, 6) & "01' And WD01<='" & CompWorkDay(2, strSrvDate(1), 1) & "' "
      
Case "2" '承辦人管理工作進度資料查詢
      StrSQL6 = ""
      StrSQL6 = StrSQL6 & " ma01='" & Trim(Left("" & Combo1.Text, 6)) & "'  "
      StrSQL6 = StrSQL6 & " and ma02='" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "' and ma03='1' "
      StrSQLa = "Select Count(*) From WorkDay Where WD01>='" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01' And WD01<='" & IIf(Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) = Mid(strSrvDate(1), 1, 6), CompWorkDay(2, strSrvDate(1), 1), Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31") & "' "
Case Else
End Select
'Modify by Morgan 2009/7/21 目前進度改以件數來表示(原來用%)
strSql = "select * from monthassess where " & StrSQL6
'Modify by Morgan 2010/12/30 件->基數
Me.lblCal1(0).Caption = "0.00 基數"
Me.lblCal1(1).Caption = "0.00 %"
Me.lblCal1(2).Caption = "0.00 基數"
Me.lblCal1(3).Caption = "0.00 基數"

theWorkDay = 0
With adoRecordset
    CheckOC
    .CursorLocation = adUseClient
    .Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        theWorkDay = Val(CheckStr(.Fields(0)))
        CheckOC
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 And theWorkDay > 0 Then
            'Modify by Morgan 2010/3/4 不必判斷都顯示
            'If Val(CheckStr(.Fields("ma34"))) <> 0 And Val(CheckStr(.Fields("ma33"))) <> 0 Then
                'Modify by Morgan 2010/12/30 件->基數
                lblCal1(0).Caption = CheckStr(.Fields("ma33")) & " 基數"
                lblCal1(1).Caption = CheckStr(.Fields("ma34")) & " %"
                lblCal1(3).Caption = CheckStr(.Fields("ma54")) & " 基數"
            'End If
            '應完成工作量
            If Val(CheckStr(.Fields("ma04"))) <> 0 Then
                'Modify by Morgan 2009/7/21
                'lblCal1(2).Caption = Format((((Val(CheckStr(.Fields("ma33"))) / (Val(CheckStr(.Fields("ma04"))) / Val(CheckStr(.Fields("ma05"))) * theWorkDay))) * 100) - 100, "0.00") & " %"
                strExc(0) = Format(Val(CheckStr(.Fields("ma33"))) - Val(CheckStr(.Fields("ma04"))) / Val(CheckStr(.Fields("ma05"))) * theWorkDay, "0.00")
                'Modify by Morgan 2010/12/30 件->基數
                lblCal1(2).Caption = IIf(Val(strExc(0)) >= 0, "+", "") & strExc(0) & " 基數"
            Else
                lblCal1(2).Caption = "尚無目標"
            End If
        End If
    End If
End With

'Added by Morgan 2025/7/18
If ProState = "2" Then
   Picture1.Visible = True
   SetPoint
Else
   Picture1.Visible = False
End If
'end 2025/7/18

Me.Enabled = True
'Add By Sindy 2024/3/14
Exit Sub

ErrHnd:
   GoTo RunCreateTable
'2024/3/14 END
End Sub

Sub AddToMdb(oStrSQL As String)
Dim strCP09s As String

CheckOC
With adoRecordset
    .CursorLocation = adUseClient
   .Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        k = 0
        'DoEvents
        strCP09s = "''"
        Do While .EOF = False
            strCP09s = strCP09s & ",'" & .Fields("cp09") & "'" 'Add by Morgan 2011/8/3
            For i = 0 To 26  'edit by nickc 2007/11/27  21
                strTemp(i) = CheckStr(.Fields(i))
                'Modify by Morgan 2011/1/3 修正日期欄位排序問題(小於100年的前面補空白)
                If Len(strTemp(i)) = 8 Then
                  If Mid(strTemp(i), 3, 1) = "/" And Mid(strTemp(i), 6, 1) = "/" Then
                     strTemp(i) = " " & strTemp(i)
                  End If
                End If
            Next i
            'edit by nickc 2007/11/27
            'strSQL = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "' ) "
            'Modify by Morgan 2009/7/14 +EP28
            'Modify By Sindy 2024/3/29 為配合外專此暫存檔也增加 ,R110033 text,R110034 text,R110035 text 欄位
            'Modified by Lydia 2025/02/05 內專工程師增加顯示「指定日期」CP142=R110033
            'strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "'," & Val("" & .Fields("cp97")) & "," & Val("" & .Fields("cp98")) & "," & Val("" & .Fields("cp111")) & ",'" & "" & .Fields("ep34").Value & "','" & "" & .Fields("cp112").Value & "','" & .Fields("ep28").Value & "',0,'','','') "
            'Modified by Morgan 2025/7/9 +CP18=R110034
            strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "'," & Val("" & .Fields("cp97")) & "," & Val("" & .Fields("cp98")) & "," & Val("" & .Fields("cp111")) & ",'" & "" & .Fields("ep34").Value & "','" & "" & .Fields("cp112").Value & "','" & .Fields("ep28").Value & "',0,'" & "" & .Fields("CP142").Value & "','" & Format(Val("" & .Fields("cp18")), "0.00") & "','') "
            adoEng.Execute strSql
            .MoveNext
            'DoEvents
        Loop
        'Add By Sindy 2024/6/20 P12專利處程序排除
        If Pub_StrUserSt03 <> "P12" Then
        '2024/6/20 END
         'Add by Morgan 2011/8/3
         '更新支援+修改+衍生基數
         'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
         'strExc(0) = "select SH12,sum(pp) pp from (" & _
          " select SH12,Round(Decode(SH06, 'CFP', Nvl(SH05,0)/3, Nvl(SH05,0)/4) ,2) pp from supporthour where SH12 in (" & strCP09s & ") And SH11='V' and SH05>0" & _
          " Union All Select MH12,Round(Nvl(MH05,0)*0.2 ,2) pp From ModifyHour Where MH12 in (" & strCP09s & ") And MH11='V' and MH05>0" & _
          " Union All Select EH12,Round(Nvl(EH05,0)*0.2 ,2) pp From ExtendHour Where EH12 in (" & strCP09s & ") And EH11='V'and EH05>0) X group by SH12"
         'Modified by Morgan 2019/4/9 108考核支援時數轉換要除組別參數
         'Modified by Morgan 2023/9/23 支援人員=承辦人才要計算Ex:P-132159(102)--游經理
         'Modified by Morgan 2025/7/9 +cp18,a1u07(更新有銷帳的收文點數)
         'Modified by Morgan 2025/7/10 還原,可不必減銷帳--柏翰
         'strExc(0) = "select SH12,sum(pp) pp,sum(cp18) cp18,sum(a1u07) a1u07 from (" & _
          " select SH12,Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2) pp,0 cp18,0 a1u07 from supporthour,staff where st01(+)=sh02 and SH12 in (" & strCP09s & ") And SH11='V' and SH05>0 and exists(select * from caseprogress where cp09=sh12 and cp14=sh02)" & _
          " Union All Select MH12,Round(Nvl(MH05,0)*0.2 ,2) pp,0 cp18,0 a1u07 From ModifyHour Where MH12 in (" & strCP09s & ") And MH11='V' and MH05>0 and exists(select * from caseprogress where cp09=MH12 and cp14=mh02)" & _
          " Union All Select EH12,Round(Nvl(EH05,0)*0.2 ,2) pp,0 cp18,0 a1u07 From ExtendHour Where EH12 in (" & strCP09s & ") And EH11='V'and EH05>0 and exists(select * from caseprogress where cp09=EH12 and cp14=eh02)" & _
          " union All select a1u03,0,max(cp18) cp18,sum(a1u07)/1000 a1u07 from acc1u0,caseprogress where a1u03 in (" & strCP09s & ") and a1u07>0 and cp09(+)=a1u03 group by a1u03" & _
          ") X group by SH12"
         strExc(0) = "select SH12,sum(pp) pp from (" & _
          " select SH12,Round(" & Sh2EPtCode & " / GetDivNum(st70,sh01) ,2) pp from supporthour,staff where st01(+)=sh02 and SH12 in (" & strCP09s & ") And SH11='V' and SH05>0 and exists(select * from caseprogress where cp09=sh12 and cp14=sh02)" & _
          " Union All Select MH12,Round(Nvl(MH05,0)*0.2 ,2) pp From ModifyHour Where MH12 in (" & strCP09s & ") And MH11='V' and MH05>0 and exists(select * from caseprogress where cp09=MH12 and cp14=mh02)" & _
          " Union All Select EH12,Round(Nvl(EH05,0)*0.2 ,2) pp From ExtendHour Where EH12 in (" & strCP09s & ") And EH11='V'and EH05>0 and exists(select * from caseprogress where cp09=EH12 and cp14=eh02)" & _
          ") X group by SH12"
          'end 2014/3/20
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
                'Modified by Morgan 2025/7/9
                'Modified by Morgan 2025/7/10 還原,可不必減銷帳--柏翰
                'strSql = "update R090614 set R110032=" & Val("" & .Fields("pp")) & IIf(Val("" & .Fields("a1u07")) <> 0, ",R110034='" & Format(Val("" & .Fields("cp18")) - Val("" & .Fields("a1u07")), "0.00") & "'", "") & " where R110022='" & .Fields("SH12") & "'"
                strSql = "update R090614 set R110032=" & .Fields("pp") & " where R110022='" & .Fields("SH12") & "'"
                adoEng.Execute strSql, intI
                .MoveNext
            Loop
            End With
         End If
         'end 2011/8/3
        End If
    End If
End With
CheckOC
End Sub

'Modify By Sindy 2021/3/15 + , Optional bolRun21 As Boolean = True
Sub ChgGrdColor(Optional iRow As Integer = -1, Optional bolRun21 As Boolean = True)
'Add By Cheng 2002/09/19
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim arrCaseNo
Dim ii As Integer
Dim tmpNat As String
'add by nick 2004/11/23
Dim ColorFlag As String
'Add by Morgan 2009/11/12
Dim iStart As Integer, iEnd As Integer
Dim i As Integer 'Add By Sindy 2024/3/14

With GRD1
   'Add by Morgan 2009/11/12
   If iRow >= 0 Then
      iStart = iRow
      iEnd = iRow
   Else
      iStart = 1
      iEnd = .Rows - 1
   End If

   '.Visible = False
   'Modify by Morgan 2009/11/12
   'For i = 1 To .Rows - 1
   For i = iStart To iEnd
      DoEvents
      'Add By Sindy 2024/3/14
      If i > GRD1.Rows - 1 Or i < 0 Then
         MsgBox "不正確的資料列值。( .Row=" & i & " )"
         Exit Sub
      ElseIf i = 1 And iEnd = 1 Then
         .row = i
         If .Text = "" Then
            Exit Sub
         End If
      End If
      '2024/3/14 END
      .row = i

              'edit by nick 2004/11/23 start
              'add by nick 2004/07/16 秀玲說每一筆資料都要判斷所有的客戶，若有其中一個客戶沒在 applicantdiscount 的話，秀淺綠色
              'edit by nick 只做 P 台灣案
        '      .col = 5
        '      tmpNat = .Text
        '      .col = 3
        '      If SystemNumber(Trim(.Text), 1) = "P" And Trim(tmpNat) = "台灣" Then
        '            If CheckCuInAD(Trim(.Text)) = False Then
        '                  .col = 4
        '                  .CellBackColor = QBColor(10)
        '            End If
        '       End If
              
              'Modify by Morgan 2009/7/13 加欄位:預會日 15
              '.col = 20 'edit by nickc 2007/11/27   加欄位後退   19
              'Modify By Sindy 2021/3/15
              If bolRun21 = True Then
              '2021/3/15 END
               'Modified by Lydia 2025/02/05 改用變數
               '.col = 21
               .col = colEp12_1
               
               ColorFlag = Mid(.Text, 1, 1)
               .Text = Mid(.Text, 2)
               '2009/10/23 modify by sonia 商標處人員不顯示綠色
               'If ColorFlag = "1" Then
               'Modify by Morgan 2009/11/12 改用全域變數判斷
               'If ColorFlag = "1" And Mid(PUB_GetStaffST15(strUserNum, 1), 1, 2) <> "P2" Then
               If ColorFlag = "1" Then
                   'Modified by Lydia 2025/02/05 改用變數
                   '.col = 4
                   .col = colCaseName_1
                   .CellBackColor = QBColor(10)
               End If
               'edit by nick 2004/11/23 end
              End If
              '911025 nick 薛說已經取消收文就不能有在顯示其他顏色
               
               'Modify by Morgan 2009/7/13 加欄位:預會日 15
               '.col = 23 'edit by nickc 2007/11/27   加欄位後退    22
               'Modified by Lydia 2025/02/05 改用變數
               '.col = 24
               .col = colCp57_1
               
                Tmp003 = Trim(.Text)
                 '若有取消收文日期
                If Tmp003 <> "" Then
                    '灰色
                     'Modified by Lydia 2025/02/05 改用變數
                     '.col = 3
                     .col = colCaseNo_1
                     .CellBackColor = QBColor(8)
                     'Modified by Lydia 2025/02/05 改用變數
                     '.col = 10
                     .col = colCp48_1
                     .CellBackColor = QBColor(8)
                     'Modified by Lydia 2025/02/05 改用變數
                     '.col = 11
                     .col = colPv_1
                     .CellBackColor = QBColor(8)
                     'Modified by Lydia 2025/02/05 改用變數
                     '.col = 13
                     .col = colEp06_1
                     .CellBackColor = QBColor(8)
                Else
                    'add by nickc 2006/07/25 郭說C類P 不要有顏色
                    'edit by nickc 2007/11/27 若是不會稿，就不管
                    'If .TextMatrix(i, 1) & SystemNumber(.TextMatrix(i, 3), 1) <> "CP" Then
                    'Modify by Morgan 2009/7/14 加欄位:預會日 15
                    'If .TextMatrix(i, 1) & SystemNumber(.TextMatrix(i, 3), 1) <> "CP" And .TextMatrix(i, 24) <> "N" Then
                    'Modified by Lydia 2025/02/05 改用變數
                    'If .TextMatrix(i, 1) & SystemNumber(.TextMatrix(i, 3), 1) <> "CP" And .TextMatrix(i, 25) <> "N" Then
                    If .TextMatrix(i, 1) & SystemNumber(.TextMatrix(i, colCaseNo_1), 1) <> "CP" And .TextMatrix(i, colEp34_1) <> "N" Then
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 10 'edit by nickc 2007/11/27 欄位對調   9
                           .col = colCp48_1
                           Tmp001 = Trim(.Text)
                           '會稿日
                           'Modify by Morgan 2009/7/13 加欄位:預會日 15
                           '.col = 15 'edit by nickc 2007/11/27   加欄位後退   14
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 16
                           .col = colEp07_1
                           
                           
                           Tmp002 = Trim(.Text)
                           '取消收文日期
                           'Modify by Morgan 2009/7/13 加欄位:預會日 15
                           '.col = 23 'edit by nickc 2007/11/27   加欄位後退   22
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 24
                           .col = colCp57_1
                           
                           Tmp003 = Trim(.Text)
                           '若有承辦期限, 無會稿日及取消收文日期
                           If Tmp001 <> "" And Tmp002 = "" And Tmp003 = "" Then
                               '若承辦期限小於等於系統日(逾承辦期限未會稿)
                                'Modify By Cheng 2003/04/28
        '                       If ChangeTStringToWString(ChangeTDateStringToTString(Tmp001)) <= GetTodayDate Then
                               'edit by nickc 2005/06/03
                               'If ChangeTStringToWString(ChangeTDateStringToTString(Tmp001)) <= strSrvDate(1) Then
                               If Val(ChangeTStringToWString(ChangeTDateStringToTString(Tmp001))) < Val(strSrvDate(1)) Then
                                   'If GetWorkDay(GetTodayDate, Tmp001) > 8 Then
                                   
                                   'Modify By Sindy 2017/11/14 P107826(目次78)-分析:此分析函於作業前，智權同仁告知客戶不續辦，因此未作業，但因程序人員直接上發文日，而沒有完稿日，故於承辦人系統中呈現逾期未辦理的黃色狀態
                                   '因此,黃色狀態資料加控制無發文日條件
                                   'Modified by Lydia 2025/02/05 改用變數
                                   '.col = 19
                                   .col = colCp27_1
                                   '若無發文日
                                   If .Text = "" Then
                                   '2017/11/14 END
                                       '黃色
                                       'Modified by Lydia 2025/02/05 改用變數
                                       '.col = 3
                                       .col = colCaseNo_1
                                       .CellBackColor = &H80FFFF
                                       'Modified by Lydia 2025/02/05 改用變數
                                       '.col = 10
                                       .col = colCp48_1
                                       .CellBackColor = &H80FFFF
                                       'Modified by Lydia 2025/02/05 改用變數
                                       '.col = 11
                                       .col = colPv_1
                                       .CellBackColor = &H80FFFF
                                       'Modified by Lydia 2025/02/05 改用變數
                                       '.col = 13
                                       .col = colEp06_1
                                       .CellBackColor = &H80FFFF
                                   End If
                                   'add by nickc 2008/02/01  若是當日為該月第一個工作天，則，寄件值 * 加乘註記就好
                                   If strSrvDate(1) <> GetMonthStdDay(Mid(strSrvDate(1), 1, 6)) Then
                                        'add by nickc  2008/01/03 計算黃色預估考核值
                                         'Modify by Morgan 2009/7/14 加欄位:預會日 15
                                         '.TextMatrix(i, 11) = StrMenu6(.TextMatrix(i, 11), .TextMatrix(i, 22))
                                         'Remove by Morgan 2010/11/3 改變規則移到外層(已與累計會稿量無關)
                                         '.TextMatrix(i, 11) = StrMenu6(.TextMatrix(i, 11), .TextMatrix(i, 23))
                                    End If
                                    'Removed by Morgan 2021/11/9 108考核已取消會稿加乘
                                    '.TextMatrix(i, 11) = StrMenu20(i) 'Add by Morgan 2010/11/3
                                    'end 2021/11/9
                                End If
                           Else
                                 'Modify By Sindy 2016/3/7 唐韻如:取消顯示淺黃色
'                                'add by nickc 2005/05/26 若是有會稿日，且過承辦期限，給淡黃色
'                                If Tmp001 <> "" And Tmp002 <> "" And Tmp003 = "" Then
'                                       'edit by nickc 2005/06/03
'                                       'If ChangeTStringToWString(ChangeTDateStringToTString(Tmp001)) <= ChangeTStringToWString(ChangeTDateStringToTString(Tmp002)) Then
'                                       If Val(ChangeTStringToWString(ChangeTDateStringToTString(Tmp001))) < Val(ChangeTStringToWString(ChangeTDateStringToTString(Tmp002))) Then
'                                                '淡黃色
'                                             .col = 3
'                                             .CellBackColor = &HC0FFFF
'                                             .col = 10
'                                             .CellBackColor = &HC0FFFF
'                                             .col = 11
'                                             .CellBackColor = &HC0FFFF
'                                             .col = 13
'                                             .CellBackColor = &HC0FFFF
'                                           'End If
'                                        End If
'                                End If
                           End If
                    End If
            
                  'Modify by Morgan 2009/7/13 加欄位:預會日 15
                  '.col = 18 'edit by nickc 2007/11/27   加欄位後退   17
                  'Modified by Lydia 2025/02/05 改用變數
                  '.col = 19
                  .col = colCp27_1
                  
                  '若無發文日
                  If .Text = "" Then
                     'Modified by Lydia 2025/02/05 改用變數
                     '.col = 9 'edit by nickc 2007/11/27 欄位對調  10
                     .col = colCp06_1
                     '若系統日大於等於本所期限且本所期限有值(逾本所期限未發文)
                          'edit by nickc 2006/04/27 要用數字比對
                          'If Trim(.Text) <= ChangeTStringToTDateString(strSrvDate(2)) And Trim(.Text) <> "" Then
                          If Val(ChangeTStringToWString(ChangeTDateStringToTString(Trim(.Text)))) <= Val(strSrvDate(1)) And Trim(.Text) <> "" Then
                                     '淺紅色
                                    'Modified by Lydia 2025/02/05 改用變數
                                    '.col = 3
                                    .col = colCaseNo_1
                                    .CellBackColor = &HC0C0FF    'edit by nickc 2006/04/27 在淺一點 &H8080FF
                                    'Modified by Lydia 2025/02/05 改用變數
                                    '.col = 10
                                    .col = colCp48_1
                                    .CellBackColor = &HC0C0FF     'edit by nickc 2006/04/27 在淺一點 &H8080FF
                                    'Modified by Lydia 2025/02/05 改用變數
                                    '.col = 11
                                    .col = colPv_1
                                    .CellBackColor = &HC0C0FF     'edit by nickc 2006/04/27 在淺一點 &H8080FF
                                    'Modified by Lydia 2025/02/05 改用變數
                                    '.col = 13
                                    .col = colEp06_1
                                    .CellBackColor = &HC0C0FF     'edit by nickc 2006/04/27 在淺一點 &H8080FF
                          'add by nickc 2005/05/30 下面的不管有無發文
                          End If
        '          End If   '2005/6/2 ADD BY SONIA
                     '若無本所期限或本所期限大於系統日
                     'Else  '2005/6/2 CANCEL BY SONIA
                           '承辦期限
            '               .Col = 8
        '                   .col = 9
        '                   Tmp001 = Trim(.Text)
        '                   '會稿日
        '    '               .Col = 13
        '                   .col = 14
        '                   Tmp002 = Trim(.Text)
        '                   '取消收文日期
        '    '               .Col = 21
        '                   .col = 22
        '                   Tmp003 = Trim(.Text)
        '                   '若有承辦期限, 無會稿日及取消收文日期
        '                   If Tmp001 <> "" And Tmp002 = "" And Tmp003 = "" Then
        '                       '若承辦期限小於等於系統日(逾承辦期限未會稿)
        '                        'Modify By Cheng 2003/04/28
        ''                       If ChangeTStringToWString(ChangeTDateStringToTString(Tmp001)) <= GetTodayDate Then
        '                       'edit by nickc 2005/06/03
        '                       'If ChangeTStringToWString(ChangeTDateStringToTString(Tmp001)) <= strSrvDate(1) Then
        '                       If ChangeTStringToWString(ChangeTDateStringToTString(Tmp001)) < strSrvDate(1) Then
        '                           'If GetWorkDay(GetTodayDate, Tmp001) > 8 Then
        '                                '黃色
        '                             .col = 3
        '                             .CellBackColor = &H80FFFF
        '    '                         .Col = 9
        '                             .col = 10
        '                             .CellBackColor = &H80FFFF
        '    '                         .Col = 10
        '                             .col = 11
        '                             .CellBackColor = &H80FFFF
        '    '                         .Col = 12
        '                             .col = 13
        '                             .CellBackColor = &H80FFFF
        '                           'End If
        '                        End If
        '                   Else
        '                        'add by nickc 2005/05/26 若是有會稿日，且過承辦期限，給淡黃色
        '                        If Tmp001 <> "" And Tmp002 <> "" And Tmp003 = "" Then
        '                               'edit by nickc 2005/06/03
        '                               'If ChangeTStringToWString(ChangeTDateStringToTString(Tmp001)) <= ChangeTStringToWString(ChangeTDateStringToTString(Tmp002)) Then
        '                               If ChangeTStringToWString(ChangeTDateStringToTString(Tmp001)) < ChangeTStringToWString(ChangeTDateStringToTString(Tmp002)) Then
        '                                        '淡黃色
        '                                     .col = 3
        '                                     .CellBackColor = &HC0FFFF
        '                                     .col = 10
        '                                     .CellBackColor = &HC0FFFF
        '                                     .col = 11
        '                                     .CellBackColor = &HC0FFFF
        '                                     .col = 13
        '                                     .CellBackColor = &HC0FFFF
        '                                   'End If
        '                                End If
        '                        End If
        '    '                 .Col = 21
        '                   '911025 nick 薛說已經取消收文就不能有在顯示其他顏色
        '                   '  .col = 22
        '                   '  Tmp003 = Trim(.Text)
        '                   '   '若有取消收文日期
        '                   '  If Tmp003 <> "" Then
        '                   '       .col = 3
        '                   '       .CellBackColor = QBColor(8)
        '    '              '        .Col = 9
        '                   '       .col = 10
        '                   '       .CellBackColor = QBColor(8)
        '    '              '        .Col = 10
        '                   '       .col = 11
        '                   '       .CellBackColor = QBColor(8)
        '    '              '        .Col = 12
        '                   '       .col = 13
        '                   '       .CellBackColor = QBColor(8)
        '                   '  End If
        '                   End If
        'edit by nickc 2005/05/30
        '             End If
                  '若有發文日
        'edit by nickc 2005/05/30
                  Else
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 13 'edit by nickc 2007/11/27   加欄位後退   12
                           .col = colEp06_1
                           Tmp001 = Trim(.Text)
                           
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 14 'edit by nickc 2007/11/27   加欄位後退   13
                           .col = colEp09_1
                           Tmp002 = Trim(.Text)
                           
                           'Modify by Morgan 2009/7/13 加欄位:預會日 15
                           '.col = 15 'edit by nickc 2007/11/27   加欄位後退   14
                            'Modified by Lydia 2025/02/05 改用變數
                           '.col = 16
                           .col = colEp07_1
                           
                           Tmp003 = Trim(.Text)
                           
                           'Modify by Morgan 2009/7/13 加欄位:預會日 15
                           '.col = 17 'edit by nickc 2007/11/27   加欄位後退   16
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 18
                           .col = colEp08_1
                           
                           Tmp004 = Trim(.Text)
                           
                           'Modify by Morgan 2009/7/13 加欄位:預會日 15
                           '.col = 23 'edit by nickc 2007/11/27   加欄位後退   22
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 24
                           .col = colCp57_1
                           
                           If (Tmp001 = "" Or Tmp002 = "" Or Tmp003 = "" Or Tmp004 = "") Then
                              
                              'Modify by Morgan 2009/7/13 加欄位:預會日 15
                              '.col = 17 'edit by nickc 2007/11/27   加欄位後退   16
                              'Modified by Lydia 2025/02/05 改用變數
                              '.col = 18
                              .col = colEp08_1
                              'Modify by Morgan 2011/1/4 修正日期排序問題前面補空白
                              '.Text = "******"
                              .Text = " ******"
                           End If
                  End If   '2005/6/2 CANCEL BY SONIA
                  'add by nickc 2005/06/16 從上面搬來 不管發文日
        
                  'Add By Cheng 2002/09/19
                  '若系統類別為"P", 且案件性質為"申請優先權證明書"(405)時, 若其專利基本檔的"申請案號"欄有值但該收文號尚無發文日時, 以紅色表示(申請優先權未發文)
                    'Modify By Cheng 2003/04/28
        '          If .TextMatrix(i, 3) <> "" Then
                  'Modified by Lydia 2025/02/05 改用變數
                  'If "" & .TextMatrix(i, 3) <> "" And "" & .TextMatrix(i, 7) = "申請優先權證明書" Then
                  '   arrCaseNo = Split(.TextMatrix(i, 3), "-")
                  If "" & .TextMatrix(i, colCaseNo_1) <> "" And "" & .TextMatrix(i, colCPM_1) = "申請優先權證明書" Then
                     arrCaseNo = Split(.TextMatrix(i, colCaseNo_1), "-")
                  'end 2025/02/05
                     If arrCaseNo(0) = "P" Then
                        'Modify by Morgan 2009/7/13 加欄位:預會日 15
                        'StrSQLa = " Select * From CaseProgress,Patent Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP10='405' AND PA11 IS NOT NULL AND CP27 IS NULL AND CP09='" & .TextMatrix(i, 22) & "'"
                        'Modified by Lydia 2025/02/05 改用變數
                        'StrSQLa = " Select * From CaseProgress,Patent Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP10='405' AND PA11 IS NOT NULL and cp27 is null AND CP09='" & .TextMatrix(i, 23) & "'"
                        StrSQLa = " Select * From CaseProgress,Patent Where CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP10='405' AND PA11 IS NOT NULL and cp27 is null AND CP09='" & .TextMatrix(i, colCP09_1) & "'"
                        rsA.CursorLocation = adUseClient
                        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                        If rsA.RecordCount > 0 Then
                           '紅色
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 3
                           .col = colCaseNo_1
                           .CellBackColor = vbRed
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 10
                           .col = colCp48_1
                           .CellBackColor = vbRed
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 11
                           .col = colPv_1
                           .CellBackColor = vbRed
                           'Modified by Lydia 2025/02/05 改用變數
                           '.col = 13
                           .col = colEp06_1
                           .CellBackColor = vbRed
                        End If
                        If rsA.State <> adStateClosed Then rsA.Close
                        Set rsA = Nothing
                     End If
                  End If
                '911025 nick 薛說已經取消收文就不能有在顯示其他顏色
                End If
   Next i
   'Modify By Sindy 2024/11/7 mark
'   'Add By Cheng 2002/10/23
'   '預設目前在第一筆的位置
'   With Me.GRD1
'      .row = 1
'      .col = 0
'      .CellBackColor = &HFFC0C0
'      .col = 12
'      .CellBackColor = &HFFC0C0
'      SWPColor2 = SWPColor
'      SWPRow2 = .row
'   End With
'   '.Visible = True
   
   'Modified by Lydia 2025/02/05
   'SetGrd1
   Call SetGrd1_New(False)
End With
End Sub

Sub StrMenu()
Dim iMouse As Integer
iMouse = Screen.MousePointer
'Modify by Morgan 2011/8/3 +R110032(支援+修改+衍生基數)
Select Case ProState
Case "1"
      'Modify by Morgan 2009/7/14 加欄位 R110031
      'Modify by Morgan 2010/11/4 +r110025,r110030
      'Modified by Lydia 2025/02/05 +R110033
      'Modified by Morgan 2025/7/9 +R110034(收文點數)
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028)  or r110028=0,1,r110028))+R110032,'0.00'),R110034,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032 FROM R090614 " & _
                    " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "'"
'      'Modify By Sindy 2013/6/7 依會稿完成日大至小+目次大至小排序
'      If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
'         strSql = strSql & " ORDER BY R110017 desc,R110002 desc,R110003,R110004"
'      Else
         strSql = strSql & " ORDER BY R110002 desc,R110003,R110004"
'      End If
'      '2013/6/7 End
Case "2"
      If frm090614.txt1(8) = "N" Then
         'Modify by Morgan 2009/7/14 加欄位 R110031
         'Modify by Morgan 2010/11/4 +r110025,r110030
         'Modified by Lydia 2025/02/05 +R110033
         'Modified by Morgan 2025/7/9 +R110034(收文點數)
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028)  or r110028=0,1,r110028))+R110032,'0.00'),R110034,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032 FROM R090614 WHERE ID='" & strUserNum & "' AND R110001 IN (" & Combo1_String & ") ORDER BY R110005,R110002 desc,  R110004 "
      Else
         'Modify by Morgan 2009/7/14 加欄位 R110031
         'Modify by Morgan 2010/11/4 +r110025,r110030
         'Modified by Lydia 2025/02/05 +R110033
         'Modified by Morgan 2025/7/9 +R110034(收文點數)
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110033,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028))+R110032,'0.00'),R110034,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032 FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' ORDER BY R110002 desc, R110003, R110004 "
      End If
Case "3"
Case Else
End Select

SetGrd1_New 'Added by Lydia 2025/02/05

CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        If ProState = "2" Then
            InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        End If
        Set GRD1.Recordset = adoRecordset
        ChkNoData = False
    Else
        If ProState = "2" Then
            InsertQueryLog (0) 'Add By Sindy 2010/12/17
        End If
        ChkNoData = True
        'Mark by Lydia 2025/02/05
        'GRD1.Clear
        'GRD1.Rows = 2
        'end 2025/02/05
        'Modify by Morgan 2009/11/12
        'Screen.MousePointer = vbDefault
        Screen.MousePointer = iMouse
        Exit Sub
    End If
End With
CheckOC
ChgGrdColor
SWPRow2 = "1"
GRD1.row = Val(SWPRow2)
If ChkNoData = False Then GRD1.col = 1

End Sub

'Modify by Morgan 2010/12/31 配合百年修改日期欄位寬度
'Mark by Lydia 2025/02/05 改寫法 >> SetGrd1_New
Private Sub SetGrd1()

With GRD1
    .Visible = False
    'Modify by Morgan 2009/7/13 加欄位:預會日 15
    '.Cols = 27 'edit by nickc 2007/11/27 加欄位 25
    .Cols = 29

    .row = 0
    .col = 0:   .Text = "目次"
    .ColWidth(0) = 350
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "收文類別"
    .ColWidth(1) = 200
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "收文日"
    .ColWidth(2) = 795
    .ColAlignment(2) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "本所案號"
    .ColWidth(3) = 1005
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "案件名稱"
    .ColWidth(4) = 1155
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "國家"
    .ColWidth(5) = 450
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "種類"
    .ColWidth(6) = 450
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "案件性質"
    .ColWidth(7) = 795
    .CellAlignment = flexAlignCenterCenter
    'Add By Cheng 2002/04/16
    .col = 8:   .Text = "Y/N"
    .ColWidth(8) = 285
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "本所期限" 'edit by nickc 2007/11/27  對調 "承辦期限"
    .ColWidth(9) = 795
    .ColAlignment(9) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    '.col = 10:   .Text = "收卷"
    '.ColWidth(10) = 300
    '.CellAlignment = flexAlignCenterCenter
    .col = 10:   .Text = "承辦期限"  'edit by nickc 2007/11/27  對調 "本所期限"
    .ColWidth(10) = 795
    .ColAlignment(10) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
'edit by nickc 2007/11/27 加欄位
    .col = 11:  .Text = "考核值"
    .ColWidth(11) = 435
    .ColAlignment(11) = flexAlignRightCenter
    .CellAlignment = flexAlignLeftCenter
    '.col = 12:   .Text = "確認"
    '.ColWidth(12) = 300
    '.CellAlignment = flexAlignCenterCenter
    'edit by nickc 2007/11/27 加欄位  以下都往後退
    .col = 12:  .Text = "法定期限"
    .ColWidth(12) = 0
    .ColAlignment(12) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "齊備日"
    .ColWidth(13) = 795
    .ColAlignment(13) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 14:  .Text = "完稿日"
    .ColWidth(14) = 795
    .ColAlignment(14) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    'Modify by Morgan 2009/7/13 加欄位:預會日 15
    .col = 15:  .Text = "預會日"
    .ColWidth(15) = 795
    .ColAlignment(15) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "會稿日"
    .ColWidth(16) = 795
    .ColAlignment(16) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "核稿人"
    .ColWidth(17) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 18:  .Text = "會稿完成日"
    .ColWidth(18) = 795
    .ColAlignment(18) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "發文日"
    .ColWidth(19) = 795
    .ColAlignment(19) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = "承辦天數"
    .ColWidth(20) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 21:  .Text = "備註"
    .ColWidth(21) = 2000
    .CellAlignment = flexAlignCenterCenter
    .col = 22:  .Text = "智權人員"
    .ColWidth(22) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 23:  .Text = ""
    .ColWidth(23) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 24:  .Text = ""
    .ColWidth(24) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 25:  .Text = ""
    .ColWidth(25) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 26:  .Text = ""
    .ColWidth(26) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 27:  .Text = ""
    .ColWidth(27) = 0
    .CellAlignment = flexAlignCenterCenter

    'Add by Morgan 2011/8/15
    For intI = 28 To .Cols - 1
      .ColWidth(intI) = 0
    Next

    .Visible = True

End With
'Modify By Sindy 2024/11/7 mark
'   'Add By Cheng 2002/10/23
'   '預設目前在第一筆的位置
'   With Me.GRD1
'      .row = 1
'      .col = 0
'      .CellBackColor = &HFFC0C0
'      .col = 12
'      .CellBackColor = &HFFC0C0
'      SWPColor2 = SWPColor
'      SWPRow2 = .row
'   End With
End Sub
'end 2025/02/05

Private Sub GRD1_DblClick()
    'Modify By Cheng 2004/03/08
    If Me.GRD1.MouseRow > 0 Then
        'Add By Cheng 2003/04/28
        '若有資料
        If Me.GRD1.Rows > 1 Then
            SWPRow = str(GRD1.MouseRow)
            'Modify By Cheng 2003/05/05
            '若點選的那筆無資料, 則退出函式
            If Me.GRD1.TextMatrix(SWPRow, 1) = "" Then Exit Sub
    '        MouseClick Val(SWPRow)
            SSTab1.Tab = 1
        End If
    End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Add By Cheng 2003/04/28
Dim Strindex As Integer
Dim iMouse As Integer
iMouse = Screen.MousePointer

'Add By Cheng 2004/03/08
If Me.GRD1.MouseRow <= 0 Then Exit Sub
'End
If Button = 1 Then
    Screen.MousePointer = vbHourglass
    SWPRow = str(GRD1.MouseRow)
    'Add By Cheng 2003/04/28
    Strindex = SWPRow
    With GRD1
        DoEvents
        .Visible = False
        If SWPRow2 <> "" Then
           .row = SWPRow2
           .col = 0
           .CellBackColor = QBColor(15)
           'Modified by Lydia 2025/02/05 改用變數
           '.col = 12
           .col = colCp07_1
           .CellBackColor = QBColor(15)
        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Then
            .row = 1
        End If
        .col = 0
        .CellBackColor = &HFFC0C0
        'Modified by Lydia 2025/02/05 改用變數
        '.col = 12
        .col = colCp07_1
        .CellBackColor = &HFFC0C0
        SWPColor2 = SWPColor
        SWPRow2 = .row
        .Visible = True
    End With
    'Modify by Morgan 2009/11/12
    'Screen.MousePointer = vbDefault
    Screen.MousePointer = iMouse
End If
End Sub

Sub MouseClick(Optional Strindex As Integer = 0)
    Dim iMouse As Integer
    
    iMouse = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    With GRD1
        DoEvents
        .Visible = False
        If SWPRow2 <> "" Then
           .row = SWPRow2
           .col = 0
           .CellBackColor = QBColor(15)
           'Modified by Lydia 2025/02/05 改用變數
           '.col = 12
           .col = colCp07_1
           .CellBackColor = QBColor(15)
        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Then
            .row = 1
        End If
        'Modify by Morgan 2009/7/13 加欄位:預會日 15
        '.col = 22 'edit by nickc 2007/11/27 加欄位  以下都往後退21
        'Modified by Lydia 2025/02/05 改用變數
        '.col = 23
        .col = colCP09_1
        Call Process(.Text)
        
        .col = 0
        .CellBackColor = &HFFC0C0
        'Modified by Lydia 2025/02/05 改用變數
        '.col = 12
        .col = colCp07_1
        .CellBackColor = &HFFC0C0
        SWPColor2 = SWPColor
        SWPRow2 = .row
        .Visible = True
    End With
    'Modify by Morgan 2009/11/12
    'Screen.MousePointer = vbDefault
    Screen.MousePointer = iMouse
End Sub

'Add By Cheng 2003/09/18
'存檔時使用
Sub MouseClick_1(Optional Strindex As Integer = 0)
    Dim iMouse As Integer
    iMouse = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    With GRD1
        DoEvents
        .Visible = False
        If SWPRow2 <> "" Then
           .row = SWPRow2
           .col = 0
           .CellBackColor = QBColor(15)
           'Modified by Lydia 2025/02/05 改用變數
           '.col = 12
           .col = colCp07_1
           .CellBackColor = QBColor(15)
        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Then
            .row = 1
        End If
        .col = 0
        .CellBackColor = &HFFC0C0
        'Modified by Lydia 2025/02/05 改用變數
        '.col = 12
        .col = colCp07_1
        .CellBackColor = &HFFC0C0
        SWPColor2 = SWPColor
        SWPRow2 = .row
        .Visible = True
    End With
    
    'Modify by Morgan 2009/11/12
    'Screen.MousePointer = vbDefault
    Screen.MousePointer = iMouse
End Sub

Private Sub grd1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Add By Cheng 2004/03/08
    If Me.GRD1.MouseRow < 1 Then
        Select Case Me.GRD1.MouseCol
        Case 0
            If m_blnColOrderAsc = True Then
                Me.GRD1.Sort = 3 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.GRD1.Sort = 4 '降冪
                m_blnColOrderAsc = True
            End If
        Case Else
            If m_blnColOrderAsc = True Then
                Me.GRD1.Sort = 5 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.GRD1.Sort = 6 '降冪
                m_blnColOrderAsc = True
            End If
        End Select
    End If
    'End
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Dim ii As Integer
   'Add By Sindy 2013/6/7
   'Modify By Sindy 2017/8/21
'   If SSTab1.Tab = 2 And Me.cmd(1).Visible = True Then
'      Call QueryData(False)
'   End If
   If SSTab1.Tab = 2 Then
      Call QueryData(True)
   Else
      Call QueryData(False)
   End If
   '2017/8/21 END
   
   cmdFAmend.Visible = False 'Added by Lydia 2016/12/30
   
   If PreviousTab = 2 Then
      '若有資料
      If (Me.grd2.Rows - 1) < dblPrevRow Then dblPrevRow = 0 'Add By Sindy 2024/7/10
      If Me.grd2.Rows > 1 And dblPrevRow > 0 Then
         If Me.grd2.TextMatrix(dblPrevRow, 1) <> "" Then
            For i = 1 To Me.GRD1.Rows - 1
               If Me.grd2.TextMatrix(dblPrevRow, 1) = Me.GRD1.TextMatrix(i, 0) Then
                  SWPRow = i
                  Exit For
               End If
            Next i
            MouseClick Val(SWPRow)
            If SSTab1.Tab = 1 Then
               SSTab1.Tab = 1
               '游標預設在繪圖人員欄
               'Modify By Sindy 2013/9/13 Mark Run歷程送出時,會出現執行階段錯誤5:程序呼叫或引數不正確
               'If Me.Combo2.Enabled = True Then Me.Combo2.SetFocus
               If Me.Combo2.Enabled = True And Combo2.Visible = True Then Me.Combo2.SetFocus
               '2013/9/13 END
            End If
         End If
      End If
   End If
   '2013/6/7 End
   'Add By Cheng 2003/04/28
   If PreviousTab = 0 Or PreviousTab = 1 Then
      '若有資料
      If Me.GRD1.Rows > 1 Then
         'Modify By Cheng 2003/05/05
         '若點選的那筆無資料, 則退出函式
         If Me.GRD1.TextMatrix(Val("0" & SWPRow), 1) = "" Then SSTab1.Tab = 0: Exit Sub
         'Add by Sindy 2013/6/7
         If Val(SWPRow) > 0 Then
            '上一筆資料列清除反白
            'Modify By Sindy 2016/5/9
            'If dblPrevRow > 0 Then
            If dblPrevRow > 0 And dblPrevRow <= (grd2.Rows - 1) Then
            '2016/5/9 END
               grd2.col = 0
               grd2.row = dblPrevRow
               grd2.Text = ""
               For ii = 0 To grd2.Cols - 1
                  grd2.col = ii
                  'Modify By Sindy 2013/10/29
                  If grd2.CellBackColor <> &H8080FF Then
                  '2013/10/29 END
                     grd2.CellBackColor = QBColor(15)
                  End If
               Next ii
               dblPrevRow = 0
               Call SetColColor(CStr(dblPrevRow))
            End If
            For i = 1 To Me.grd2.Rows - 1
               If Me.grd2.TextMatrix(i, 1) = Me.GRD1.TextMatrix(Val("0" & SWPRow), 0) Then
                  '目前資料列反白
                  dblPrevRow = i
                  grd2.col = 0
                  grd2.row = dblPrevRow
                  If grd2.TextMatrix(grd2.row, 1) <> "" Then
                     grd2.Text = "V"
                     For ii = 0 To grd2.Cols - 1
                        grd2.col = ii
                        'Modify By Sindy 2013/10/29
                        If grd2.CellBackColor <> &H8080FF Then
                           grd2.CellBackColor = &HFFC0C0
                        End If
                     Next ii
                  End If
                  Exit For
               End If
            Next i
         End If
         '2013/6/7 End
         If PreviousTab = 0 Then
            MouseClick Val(SWPRow)
            'Add by Sindy 2013/6/7
            If SSTab1.Tab = 1 Then
            '2013/6/7 End
               SSTab1.Tab = 1
               'Add By Cheng 2003/05/09
               '游標預設在繪圖人員欄
               'Modify By Sindy 2013/9/13 Mark Run歷程送出時,會出現執行階段錯誤5:程序呼叫或引數不正確
               'If Me.Combo2.Enabled = True Then Me.Combo2.SetFocus
               If Me.Combo2.Enabled = True And Combo2.Visible = True Then Me.Combo2.SetFocus
               '2013/9/13 END
            End If
         End If
      End If
   End If
End Sub

Private Sub txt1_Change(Index As Integer)
   Select Case Index
      'Add By Sindy 2015/3/13
      Case 23 '核稿語文
         If txt1(Index) = "2" Then
            Label1(25).Caption = "日文核稿人："
            Label1(41).Caption = "日文核完日："
         Else
            Label1(25).Caption = "英文核稿人："
            Label1(41).Caption = "英文核完日："
         End If
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'add by nickc 2007/11/28 協理說，有輸入完稿日才可以輸入不會稿
Select Case Index
   Case 1
      If txt1(3) = "" Then
'         KeyAscii = 0 'Sindy 2013/9/6 開放
      'Add by Morgan 2011/3/8 改控制只能輸入 N
      'Modify By Sindy 2013/9/24 也可輸入 Y
      ElseIf KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
           KeyAscii = 0
           Beep
      End If
   'Added by Morgan 2012/8/24
   'Modify By Sindy 2015/3/13 +17
   'Modified by Morgan 2016/7/7 +11
   Case 21, 17, 11
      If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
   'Add By Sindy 2015/3/13
   Case 23
      If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
         KeyAscii = 0
         Beep
      End If
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   'Add by Morgan 2010/9/6 若回第一頁籤時不檢查,否則若有錯誤時會無窮回圈
   If Me.SSTab1.Tab = 0 Then Exit Sub

'add by nickc 2005/04/12 暫存該所繪圖人員
Dim tmpEp13BySt06 As String
Dim tmpInti As Integer
'If SSTab1.Tab = 1 Then
Select Case Index
Case 0
     LBL1(4) = GetPrjSalesNM(txt1(0))
     If Len(LBL1(4)) = 0 And Len(txt1(Index)) <> 0 Then
         'Modify By Sindy 2022/12/6
         'Call ShowStaffErr(txt1(0))
         Call PUB_GetStaffNameDept(txt1(0), strExc(10), strExc(0), True, False)
         '2022/12/6 END
         txt1(0).SetFocus
         txt1_GotFocus (0)
         Exit Sub
     End If

Case 1 '是否會稿
     Select Case Trim(txt1(1))
     Case "Y", ""
     Case "N"
'          txt1(4) = txt1(3)
'          txt1(7) = txt1(3)
          Call ChkEP34ToEP07EP08 'Modify By Sindy 2016/5/20 抽出來變函數
          txt1_LostFocus (4)
     Case Else
          s = MsgBox("是否會稿只能輸入 Y 或 N !!", , "USER 輸入錯誤")
          txt1(1).SetFocus
          txt1(1).SelStart = 0
          txt1(1).SelLength = Len(txt1(1))
          Exit Sub
     End Select
     
Case 2 '齊備日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
         'add by nickc 2005/04/12 加入繪圖人員預設值
         tmpEp13BySt06 = ""
         '判斷所別，鎖按鈕
         'add by nickc 2005/04/19 加判斷專利處的才做
         If UCase(Mid(PUB_GetST03(strUserNum), 1, 2)) = "P1" Then
               'edit by nickc 2006/03/16 加入其他，給中所用，因為他們會有寫稿繪圖但客戶自己辦的案子
               'If Trim(Combo2.Text) = "" And (m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "104" Or m_CP10 = "105" Or m_CP10 = "109" Or m_CP10 = "110" Or m_CP10 = "112" Or m_CP10 = "113" Or m_CP10 = "114" Or m_CP10 = "115") Then
               If Trim(Combo2.Text) = "" And (m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "104" Or m_CP10 = "105" Or m_CP10 = "109" Or m_CP10 = "110" Or m_CP10 = "112" Or m_CP10 = "113" Or m_CP10 = "114" Or m_CP10 = "115" Or m_CP10 = "910") Then
                     If PUB_GetST06(strUserNum) = "1" Then
                        'eidt by nickc 2005/05/04 取消北所控制
                        'tmpEp13BySt06 = "72006"
                     ElseIf PUB_GetST06(strUserNum) = "2" Then
                        tmpEp13BySt06 = "82018"
                     'edit by nickc 2006/03/16
                     'ElseIf PUB_GetST06(strUserNum) = "3" Or PUB_GetST06(strUserNum) = "4" Then
                     ElseIf (PUB_GetST06(strUserNum) = "3" Or PUB_GetST06(strUserNum) = "4") And m_CP10 <> "910" Then
                        tmpEp13BySt06 = "78007"
                     'edit by nickc 2006/03/16
                     'Else  '其他算北所
                     ElseIf m_CP10 <> "910" Then '其他算北所
                        tmpEp13BySt06 = "87025" '72006 Modify By Sindy 2016/3/3 72006張瓊玉退休,改87025陳翔龍
                     End If
                     'add by nickc 2006/03/16
                     If tmpEp13BySt06 <> "" Then
                        For tmpInti = 0 To Combo2.ListCount - 1
                            If tmpEp13BySt06 = Trim(Mid(Combo2.List(tmpInti), 1, InStr(1, Combo2.List(tmpInti), "=") - IIf(InStr(1, Combo2.List(tmpInti), "=") = 0, 0, 1))) Then
                                Combo2.Text = Combo2.List(tmpInti)
                                txt1(0).Text = Trim(Left(Me.Combo2.Text, 6))
                            End If
                        Next tmpInti
                    End If
               End If
         End If
         
         If (SystemNumber(Trim(LBL1(7).Caption), 1) = "P" Or SystemNumber(Trim(LBL1(7).Caption), 1) = "CFP") Then
            'Modify by Sindy 2013/8/12
            'If Not (bolNewPromoterRule And txt1(12).Locked) Then 'Add by Morgan 2010/9/28
            If Not (bolNewPromoterRule) And txt1(12).Enabled = False Then 'Add by Morgan 2010/9/28
            '2013/8/12 END
               Me.txt1(12).Text = ChangeWStringToTString(Pub_GetHandleDay(SystemNumber(Trim(LBL1(7).Caption), 1), m_Country, m_CP10, ChangeTStringToWString(txt1(Index)), ChangeTStringToWString(Replace(LBL1(17).Caption, "/", "")), LBL1(3)))
               'Me.lbl1(2).Caption = ChangeTStringToTDateString(Me.txt1(12).Text) 'Remove by Morgan 2010/10/7
            End If
            
         'Add by Morgan 2008/9/3
         'FCP的檢視中說209上齊備日時要計算承辦期限
         ElseIf SystemNumber(LBL1(7).Caption, 1) = "FCP" And m_CP10 = "209" Then
            txt1(12).Text = TransDate(Pub_GetHandleDay("FCP", "000", m_CP10, DBDATE(txt1(2)), DBDATE(Replace(LBL1(17).Caption, "/", ""))), 1)
            'lbl1(2).Caption = ChangeTStringToTDateString(Me.txt1(12).Text) 'Remove by Morgan 2010/10/7
         End If
         
    'Add By Cheng 2003/05/09
     Else
        '若未輸入齊備日則清空承辦期限
        'Me.lbl1(2).Caption = "" 'Remove by Morgan 2010/10/7
        Me.txt1(12).Text = ""
     End If
     
Case 3 '完稿日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
     
Case 4 '會稿日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
     
Case 5 '核稿人
     LBL1(14) = GetPrjSalesNM(txt1(5))
     If Len(LBL1(14)) = 0 And Len(txt1(Index)) <> 0 Then
         'Modify By Sindy 2022/12/6
         'Call ShowStaffErr(txt1(5))
         Call PUB_GetStaffNameDept(txt1(5), strExc(10), strExc(0), True, False)
         '2022/12/6 END
         txt1(5).SetFocus
         txt1_GotFocus (5)
         Exit Sub
     End If

''Add By Sindy 2013/9/4
'Case 22 '判發人
'     lbl1(12) = GetPrjSalesNM(txtCP144)
'     If Len(lbl1(12)) = 0 And Len(txt1(Index)) <> 0 Then
'         ShowStaffErr
'         txtCP144.SetFocus
'         Call txt1_GotFocus(22)
'         Exit Sub
'     End If
     
Case 6 '校稿人
     LBL1(16) = GetPrjSalesNM(txt1(6))
     If Len(LBL1(16)) = 0 And Len(txt1(Index)) <> 0 Then
         'Modify By Sindy 2022/12/6
         'Call ShowStaffErr(txt1(6))
         Call PUB_GetStaffNameDept(txt1(6), strExc(10), strExc(0), True, False)
         '2022/12/6 END
         Combo4.SetFocus
         'txt1(6).SetFocus
         'txt1_GotFocus(6)
         Exit Sub
     End If
     
Case 7 '會稿完成日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
     
Case 8 '發文日
     If Len(txt1(Index)) <> 0 Then
        'Add By Cheng 2003/01/27
        '若發文日為111111則不檢查是否為工作日
        If Me.txt1(Index).Text <> "111111" Then
            If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
               ShowDateErr
               txt1(Index).SetFocus
               txt1(Index).SelLength = Len(txt1(Index))
               Exit Sub
            End If
        End If
     End If
     
Case 9 '是否通知客戶
     Select Case Trim(txt1(9))
     Case "Y", "N", ""
     Case Else
          s = MsgBox("是否通知客戶只能輸入 Y 或 N !!", , "USER 輸入錯誤")
          txt1(9).SetFocus
          txt1(9).SelStart = 0
          txt1(9).SelLength = Len(txt1(9))
          Exit Sub
     End Select
     
'Add By Cheng 2003/06/12
Case 12 '承辦期期
    '若有承辦期限
    If Me.txt1(12).Text <> "" And Me.LBL1(17).Caption <> "" Then
        '若承辦期限大於本所期限
        'Modify by Morgan 2009/8/3 改以數字比較(100年問題)
        'If Me.txt1(12).Text > Replace(lbl1(17).Caption, "/", "") Then
        If Val(txt1(12).Text) > Val(Replace(LBL1(17).Caption, "/", "")) Then
            Me.txt1(12).Text = Replace(LBL1(17).Caption, "/", "")
        End If
    End If
    'Me.lbl1(2).Caption = ChangeTStringToTDateString(Me.txt1(12).Text) 'Remove by Morgan 2010/10/7
    
Case 14 '收卷註記
   Select Case Trim(txt1(14))
   Case "Y", ""
   Case Else
      s = MsgBox("收卷註記只能輸入 Y !!", , "USER 輸入錯誤")
      txt1(14).SetFocus
      txt1(14).SelStart = 0
      txt1(14).SelLength = Len(txt1(14))
      Exit Sub
   End Select
   
Case 13 '工程師輸入本所期限
   If Len(txt1(Index)) <> 0 Then
      If CheckIsTaiwanDate(txt1(Index).Text) = False Then
         txt1(Index).SetFocus
         txt1(Index).SelLength = Len(txt1(Index))
         Exit Sub
      End If
      If txt1(Index).Text <> Replace(LBL1(17).Caption, "/", "") Then
         s = MsgBox("輸入本所期限與程序輸入本所期限不同!!", , "USER 輸入錯誤")
         txt1(Index).SetFocus
         txt1(Index).SelLength = Len(txt1(Index))
         Exit Sub
      End If
      txt1(14) = "Y"
   End If
Case Else
End Select

'If Index = 1 And Index = 3 Then
If Index = 1 Or Index = 3 Then
   Select Case Trim(txt1(1))
   Case "N"
'        txt1(4) = txt1(3)
'        txt1(7) = txt1(3)
         Call ChkEP34ToEP07EP08 'Modify By Sindy 2016/5/20 抽出來變函數
   Case Else
   End Select
End If
End Sub

'Modify By Cheng 2003/05/06
'Sub ChkTxt()
Sub ChkTxt(Strindex As String)
'If ChkCp27 = True Then
    ChkData = False
    '齊備日
    If Strindex = "2" Or Strindex = "" Then
         If Len(txt1(2)) = 0 And ChkCp27 = True Then
             If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Or Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
                 ShowDateRanErr
                 txt1(2).SetFocus
                 txt1_GotFocus (2)
                 Exit Sub
             End If
         End If
        
         If (SystemNumber(Trim(LBL1(7).Caption), 1) = "P" Or SystemNumber(Trim(LBL1(7).Caption), 1) = "CFP") And Trim(ChangeWStringToTString(LBL1(8).Caption)) <> Trim(txt1(2).Text) Then
            'add by nick 2004/09/24
            Select Case ProState
            Case "1" '承辦人
'edit by nick 2004/11/04 秀玲說可以小於
'                    If ChangeTStringToWString(Txt1(2)) < SystemDate Then
'                        s = MsgBox("文件齊備日不可小於系統日!!!")
'                        Txt1(2).SetFocus
'                        txt1_GotFocus (2)
'                        Exit Sub
'                    End If
                  'Add By Sindy 2009/10/29
                  If Len(txt1(2)) <> 0 And ChkCp27 = True Then
                     If m_bol203Case = False Then 'Added by Morgan 2013/8/1
                        If CompWorkDay(2, strSrvDate(1), 1) > ChangeTStringToWString(txt1(2)) Then
                           'Modified by Morgan 2022/6/14 修改訊息--王副總
                           's = MsgBox("齊備日須 >= 系統日 - 2 工作天(含當天)", , "日期錯誤!!")
                           s = MsgBox("齊備日須 >= 系統日 - 1 工作天", , "日期錯誤!!")
                           'end 2022/6/14
                           txt1(2).SetFocus
                           txt1_GotFocus (2)
                           Exit Sub
                        End If
                     End If 'Added by Morgan 2013/8/1
                     
                     If CompWorkDay(2, strSrvDate(1), 0) < ChangeTStringToWString(txt1(2)) Then
                        'Modified by Morgan 2022/6/14 修改訊息--王副總
                        's = MsgBox("齊備日須 <= 系統日 + 2 工作天(含當天)", , "日期錯誤!!")
                        s = MsgBox("齊備日須 <= 系統日 + 1 工作天", , "日期錯誤!!")
                        'end 2022/6/14
                        txt1(2).SetFocus
                        txt1_GotFocus (2)
                        Exit Sub
                     End If
                  End If
                  '2009/10/29 End
            Case "2" '主管
            Case Else
            End Select
         End If
        
         If Not ChkDateRanPro(ChangeTDateStringToTString(LBL1(5)), txt1(2), 1) And Len(txt1(2)) <> 0 And ChkCp27 = True Then
             txt1(2).SetFocus
             txt1_GotFocus (2)
             Exit Sub
         End If
    End If
    
    '完稿日
    If Strindex = "3" Or Strindex = "" Then
        If Len(txt1(3)) = 0 And ChkCp27 = True Then
            If Len(txt1(4)) <> 0 Or Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
                ShowDateRanErr
                txt1(3).SetFocus
                Exit Sub
            End If
        End If
        If Not ChkDateRanPro(txt1(2), txt1(3), 2) And Len(txt1(3)) <> 0 And ChkCp27 = True Then
            txt1(3).SetFocus
            txt1_GotFocus (3)
            Exit Sub
        End If
    End If
    
    '會稿日
    If Strindex = "4" Or Strindex = "" Then
        If Len(txt1(4)) = 0 And ChkCp27 = True Then
            If Len(txt1(7)) <> 0 Or Len(txt1(8)) <> 0 Then
                ShowDateRanErr
                txt1(4).SetFocus
                Exit Sub
            End If
        End If
        If Not ChkDateRanPro(txt1(3), txt1(4), 3) And Len(txt1(4)) <> 0 And ChkCp27 = True Then
            txt1(4).SetFocus
            txt1_GotFocus (4)
            Exit Sub
        End If
    End If
    
    '會稿完成日
    If Strindex = "7" Or Strindex = "" Then
        If Len(txt1(7)) = 0 And ChkCp27 = True Then
            If Len(txt1(8)) <> 0 Then
                ShowDateRanErr
                txt1(7).SetFocus
                Exit Sub
            End If
        End If
        If Not ChkDateRanPro(txt1(4), txt1(7), 4) And Len(txt1(7)) <> 0 And ChkCp27 = True Then
            txt1(7).SetFocus
            txt1_GotFocus (7)
            Exit Sub
        End If
    End If
    
    'Modify By Cheng 2003/02/05
    '若發文日為111111則不檢查
    If Strindex = "8" Or Strindex = "" Then
        If Me.txt1(8).Text <> "111111" Then
            If Not ChkDateRanPro(txt1(7), txt1(8), 5) And Len(txt1(8)) <> 0 And ChkCp27 = True Then
                'Modify By Cheng 2003/01/03
                If Me.txt1(8).Enabled And Me.txt1(8).Visible Then
                    '游標設在發文日
                    txt1(8).SetFocus
                    txt1_GotFocus (8)
                    Exit Sub
                Else
                    '游標設在完稿日
                    txt1(3).SetFocus
                    txt1_GotFocus (3)
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '是否通知客戶
    If Strindex = "9" Or Strindex = "" Then
        If Not CheckLengthIsOK(txt1(9), 1) And ChkCp27 = True Then
            txt1(9).SetFocus
            txt1_GotFocus (9)
            Exit Sub
        End If
    End If
    
    '備註
    If Strindex = "10" Or Strindex = "" Then
        If Not CheckLengthIsOK(txtEP12, 2000) And ChkCp27 = True Then
            txtEP12.SetFocus
            txtEP12_GotFocus
            Exit Sub
        End If
    End If
    
    '承辦期限
    If Strindex = "12" Or Strindex = "" Then
        If CheckIsTaiwanDate(Me.txt1(12).Text) = False Then
            MsgBox "承辦期限輸入錯誤！", vbExclamation
            Me.txt1(12).SetFocus
            txt1_GotFocus 12
            Exit Sub
        End If
    End If
    
    'add by nickc 2007/08/16
    If Strindex = "19" Or Strindex = "" Then
        If CheckIsTaiwanDate(Me.txt1(19).Text) = False Then
            MsgBox "英文核完日輸入錯誤！", vbExclamation
            Me.txt1(19).SetFocus
            txt1_GotFocus 19
            Exit Sub
        End If
    End If
    
    If Strindex = "20" Or Strindex = "" Then
        If Me.txt1(20).Text <> "Y" And txt1(20).Text <> "" Then
            MsgBox "是否暫停核稿輸入錯誤！", vbExclamation
            Me.txt1(20).SetFocus
            txt1_GotFocus 20
            Exit Sub
        End If
    End If
    ChkData = True
'End If
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
'add by nickc 2006/03/07
Dim P_DateLine As String
Dim CFP_DateLine As String
'Add By Cheng 2003/05/08
'If Index <> 2 And Index <> 3 And Index <> 4 And Index <> 7 And Index <> 8 And Index <> 9 And Index <> 10 Then Exit Sub
Select Case Index
'Modified by Lydia 2021/12/28 去掉txt1(10)
Case 2, 3, 4, 7, 8, 9, 12
    'Add By Sindy 2022/6/23 王錦寬.副總:凡C類來函，承辦人為工程師者，該道程序之齊備日僅能更改，請禁止取消。
    If Index = 2 Then '齊備日
      If Val(txt1(Index).Tag) > 0 And Left(m_strCP09, 1) = "C" And (PUB_GetST03(m_CP14) = "P10" Or PUB_GetST03(m_CP14) = "P11") Then
         If Val(txt1(Index).Text) = 0 Then
            MsgBox "C類來函齊備日僅能更改，禁止取消！", , "錯誤！"
            txt1(Index).Text = txt1(Index).Tag
            Cancel = True
            Exit Sub
         End If
      End If
    End If
    '2022/6/23 END
    
    'Add By Cheng 20036/04/28
    '若欄位無資料則不檢查
    If Me.txt1(Index).Text = "" Then Exit Sub
    'Modify By Cheng 2003/05/06
    'ChkTxt
    'Modify By Cheng 2004/02/03
'    ChkTxt "" & Index: DoEvents
    ChkTxt "" & Index
    'End
    If ChkData = False Then
        Cancel = True
        Exit Sub
    End If
    
'add by nickc 2005/03/04
Case 15
      Dim iMax As Integer
      iMax = 3
      
      'Add by Morgan 2010/9/27
      If bolNewPromoterRule Then
         iMax = 9
      End If
      'end 2010/9/27
       If txt1(Index).Enabled = True Then
           If IsNumeric(txt1(Index)) = False Then
               MsgBox "請輸入數字！", , "錯誤！"
               Cancel = True
               Exit Sub
           Else
               If Val(txt1(Index)) > iMax Then
                  MsgBox "上限為 " & iMax & " ，請重新輸入！", , "輸入錯誤！"
                  Cancel = True
                  Exit Sub
               End If
               If Val(txt1(Index)) < 0 Then
                  MsgBox "下限為 0 ，請重新輸入！", , "輸入錯誤！"
                  Cancel = True
                  Exit Sub
               End If
           End If
           If Val(txt1(Index)) <> Val(m_CP98) Then
               If Trim(txtCP99) = Trim(m_CP99) Then
                  txtCP99.Text = ""
               End If
           Else
               txtCP99.Text = m_CP99
           End If
       End If
Case 16
      If txt1(Index).Enabled = True Then
         If CheckLengthIsOK(txt1(Index), 100) = False Then
             MsgBox "最長為 50 個中文字！", , " 輸入錯誤！"
             Cancel = True
             Exit Sub
         End If
      End If
Case 17 '是否提供圖檔
     Select Case Trim(txt1(17))
     Case "Y", ""
     Case Else
          s = MsgBox("是否會稿只能輸入 Y 或 空白 !!", , "USER 輸入錯誤")
          txt1(17).SetFocus
          txt1(17).SelStart = 0
          txt1(17).SelLength = Len(txt1(17))
          Cancel = True
          Exit Sub
     End Select
'add by nickc 2006/03/07
Case 18
     If txt1(Index).Enabled = True And Trim(txt1(Index).Text) <> "" Then
      
         If Trim(txt1(12).Text) = "" Then
            MsgBox "要輸入預定會稿日，請先輸入〔齊備日〕產生〔承辦期限〕！", vbExclamation, "操作錯誤！"
            Cancel = True 'Add by Morgan 2010/9/27
            txt1(2).SetFocus
            Exit Sub
         End If
         
         '2010/5/19 MODIFY BY SONIA 發現與frm090201_8不同故改為一致
         'P_DateLine = CompWorkDay(6, ChangeTStringToWString(txt1(12)), 0)
         'CFP_DateLine = CompWorkDay(11, ChangeTStringToWString(txt1(12)), 0)
         P_DateLine = CompWorkDay(5, ChangeTStringToWString(txt1(12)), 0)
         CFP_DateLine = CompWorkDay(10, ChangeTStringToWString(txt1(12)), 0)
         '2010/5/19 END
         If ChkWork(ChangeTStringToWString(txt1(Index))) = False Then
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Cancel = True
            Exit Sub
         End If
         
'Remove by Morgan 2010/10/29 改個人也可修改,主管不受限
'         If (SystemNumber(lbl1(7), 1) = "P" And ChangeTStringToWString(Txt1(Index)) > P_DateLine) Or (SystemNumber(lbl1(7), 1) = "CFP" And ChangeTStringToWString(Txt1(Index)) > CFP_DateLine) Then
'            '2008/8/25 modify by sonia 王協理操作不檢查 2009/11/17 郭雅娟也不檢查
'            If strUserNum <> "71011" And strUserNum <> "79075" Then
'               MsgBox "P 案上限 5 工作天，CFP 案上限 10 作天！", vbCritical, "錯誤！"
'               Txt1(Index).SetFocus
'               Txt1(Index).SelStart = 0
'               Txt1(Index).SelLength = Len(Txt1(Index))
'               Cancel = True
'               Exit Sub
'            End If
'         End If
         If CheckIsTaiwanDate(txt1(Index).Text) = False Then
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Cancel = True
            Exit Sub
         End If
         
         'Add by Morgan 2009/8/12
         If txt1(Index) <> txt1(Index).Tag Then
            If Val(DBDATE(txt1(Index))) < Val(strSrvDate(1)) Then
               MsgBox "預定會稿日不可早於系統日！"
               Cancel = True
               Exit Sub
            'Modify by Morgan 2010/11/2 改為齊備日
            'ElseIf Val(DBDATE(Txt1(Index))) <= Val(DBDATE(Txt1(12))) Then
            '   MsgBox "預定會稿日必須晚於承辦期限！"
            ElseIf Val(DBDATE(txt1(Index))) < Val(DBDATE(txt1(2))) Then
               MsgBox "預定會稿日不可早於齊備日！"
               Cancel = True
               Exit Sub
            End If
            
         End If
     End If
'add by nickc 2007/08/16
Case 19   '英文核完日
      If txt1(Index).Enabled = True And Trim(txt1(Index).Text) <> "" And txt1(Index).Enabled = True Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Cancel = True
            Exit Sub
         End If
    End If
Case 20
    If txt1(Index).Enabled = True And Trim(txt1(Index).Text) <> "" And txt1(Index).Enabled = True Then
        Select Case Trim(txt1(20))
        Case "Y", ""
        Case Else
           s = MsgBox("是否暫停核稿只能輸入 Y 或空白!!", , "USER 輸入錯誤")
           txt1(20).SetFocus
           txt1(20).SelStart = 0
           txt1(20).SelLength = Len(txt1(20))
           Cancel = True
           Exit Sub
        End Select
    End If
Case Else
End Select
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add By Cheng 2002/11/19
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
'Add By Cheng 2003/01/07
Dim rsB As New ADODB.Recordset
Dim StrSqlB As String
'Add by Morgan 2010/10/29
Dim strMsg As String
Dim strDate As String
Dim MsgText As String
Dim MsgParam As VbMsgBoxStyle
Dim arrCaseNo '本所案號
Dim tmpInti As Integer 'Add By Sindy 2015/3/13
Dim blnMatch As Boolean 'Add By Sindy 2015/5/22

TxtValidate = False

'Added by Morgan 2012/4/13
'案件性質211、212、226、408，無附件不可輸入完稿日
'Mofified by Morgan 2012/6/4 +213現場勘察,408面詢
'Removed by Morgan 2013/10/8 取消,由電子承辦單流程上傳--專利處
'If cmd(0).Enabled = True And (lbl1(29) = "211" Or lbl1(29) = "212" Or lbl1(29) = "226" Or lbl1(29) = "213" Or lbl1(29) = "408") And txt1(3) <> "" And lbl1(10) = "" Then
'   strExc(0) = "select 1 from caseprogressappendix where cpa01='" & lbl1(3) & "' and rownum<2"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 0 Then
'      MsgText = "請先完成開庭/面詢記錄的電子檔並存入系統，再輸入完稿日！"
'      If strGroup = "71" Then
'         MsgText = MsgText & vbCrLf & vbCrLf & "是否仍然要存檔?"
'         MsgParam = vbExclamation + vbYesNo + vbDefaultButton2
'      Else
'         MsgParam = vbExclamation
'      End If
'      If MsgBox(MsgText, MsgParam) <> vbYes Then
'         Exit Function
'      End If
'   End If
'End If
'end 2013/10/8
'end 2012/4/13

'Add By Sindy 2015/3/13 提醒工程師以免忘記修改，但仍可選擇為日文稿
If m_EP41 = "" And txt1(23).Enabled = True And txt1(23).Text = "2" Then
   If MsgBox(m_CP10 & LBL1(15) & "，是否確定仍為日文稿？", vbYesNo + vbDefaultButton2) = vbNo Then
      txt1(23) = "1"
      SetEngChecker '設定核稿人選單
      '預設英文核稿人
      If m_PER04 <> "" And m_PER04 <> m_CP14 Then
         txt1(6).Text = m_PER04
      End If
      LBL1(16).Caption = GetPrjSalesNM(txt1(6))
      For tmpInti = 0 To Combo4.ListCount - 1
         If Trim(txt1(6).Text) = Trim(Mid(Combo4.List(tmpInti), 1, InStr(1, Combo4.List(tmpInti), "=") - IIf(InStr(1, Combo4.List(tmpInti), "=") = 0, 0, 1))) Then
            Combo4.Text = Combo4.List(tmpInti)
         End If
      Next tmpInti
      Combo4.Tag = Combo4.Text
   End If
End If
'2015/3/13 END

'add by nickc 2005/03/04 判斷是否有修改過，有要輸入理由
If txt1(15).Enabled = True Then
   If Val(Me.txt1(15).Text) <> Val(m_CP98) Then
         If Trim(txtCP99.Text) = "" Then
            MsgBox "修改過加乘註記，請輸入理由!!!", vbExclamation + vbOKOnly
            SSTab1.Tab = 1 'Add By Sindy 2024/3/15
            Exit Function
         End If
   End If
End If
'add by nickc 2007/03/06 核稿人不可與承辦人相同
If txt1(5).Enabled = True Then
   'Add By Sindy 2013/10/14 若核判表有設定核稿人時只可以修改但不可以空白
   'Modify By Sindy 2014/3/31 解開Mark
   If Trim(m_PP04) <> "" And Trim(txt1(5)) = "" Then
      MsgBox "核稿人不可空白!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1 'Add By Sindy 2024/3/15
      txt1(5).SetFocus
      Exit Function
   End If
   '2013/10/14 END
   If Trim(txt1(5)) <> "" Then
      'Add By Sindy 2017/7/10 增加檢查核稿人是否離職
      If PUB_ChkEmpFlowExists(LBL1(3), EMP_送核) = True And m_EP39 = "" Then
         If ChkStaffST04(txt1(5)) = True Then
            SSTab1.Tab = 1
            txt1(5).SetFocus
            Exit Function
         End If
      End If
      '2017/7/10 END
      'If Trim(Left("" & Combo1.Text, 6)) = Trim(txt1(5)) Then
      If UCase(Trim(Left("" & Combo1.Text, 6))) = UCase(Trim(txt1(5))) Then
         MsgBox "核稿人不可與承辦人相同!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         txt1(5).SetFocus
         Exit Function
      End If
      'Add By Sindy 2014/6/30 只要非系統設定的人員均要檢查權限
      'Add By Sindy 2014/12/9 承辦人非程序人員時,才需檢查核判權限
      'If GetStaffDepartment(Trim(Left("" & Combo1.Text, 6))) <> "P12" Then
      If GetStaffDepartment(Trim(Left("" & Combo1.Text, 6))) <> "P12" Then
      '2014/12/9 END
         If txt1(5).Tag <> txt1(5) And ProState = "1" Then 'Add By Sindy 2014/7/2 +if
            arrCaseNo = Split(Me.LBL1(7).Caption, "-")
            If Trim(m_PP04) <> Trim(txt1(5)) Then
               'Modify By Sindy 2024/6/26 +m_Country
               If PUB_ChkPromoterReader(arrCaseNo(0), m_CP10, "1", Trim(txt1(5)), , m_Country) = False Then
                  MsgBox "此人無核稿權限，請重新輸入！"
                  SSTab1.Tab = 1 'Add By Sindy 2024/3/15
                  txt1(5).SetFocus
                  Exit Function
               End If
            End If
         End If
         '2014/6/30 END
      End If
   Else
      'Add By Sindy 2013/9/10
      If PUB_GetST07(Trim(Left("" & Combo1.Text, 6))) = "專99" Then
         MsgBox "[專99] 核稿人不可空白!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         txt1(5).SetFocus
         Exit Function
      End If
   End If
End If
'Add By Sindy 2013/9/4
If Combo6.Enabled = True Then
   'Add By Sindy 2013/10/14 若核判表有設定判發人時只可以修改但不可以空白
   'Modify By Sindy 2014/3/31 解開Mark
   If Trim(m_PP05) <> "" And Trim(Left(Combo6.Text, 6)) = "" Then
      MsgBox "判發人不可空白!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1 'Add By Sindy 2024/3/15
      Combo6.SetFocus
      Exit Function
   End If
   '2013/10/14 END
   If Trim(Left(Combo6.Text, 6)) <> "" Then
      'Add By Sindy 2017/7/10 增加檢查判發人是否離職
      If PUB_ChkEmpFlowExists(LBL1(3), EMP_送判) = True And _
         PUB_ChkEmpFlowExists(LBL1(3), EMP_判發) = False Then
         If ChkStaffST04(Trim(Left(Combo6.Text, 6))) = True Then
            SSTab1.Tab = 1
            Combo6.SetFocus
            Exit Function
         End If
      End If
      '2017/7/10 END
      If UCase(Trim(Left(Combo1.Text, 6))) = UCase(Trim(Left(Combo6.Text, 6))) Then
         MsgBox "判發人不可與承辦人相同!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         Combo6.SetFocus
         Exit Function
      End If
   'Modify By Sindy 2014/4/3 Mark 因志建分析案有核稿人但自行判發(ex.P-099638 AA3013462)
'   ElseIf Trim(txtCP144) = "" And Trim(txt1(5)) <> "" Then
'       MsgBox "有核稿人時,判發人不可空白!!!", vbExclamation + vbOKOnly
'       txtCP144.SetFocus
'       Exit Function

      'Add By Sindy 2013/11/26 當代理狀況時，檢查輸入的判發人是否有判發權限
      'Add By Sindy 2014/12/9 承辦人非程序人員時,才需檢查核判權限
      If GetStaffDepartment(Trim(Left(Combo1.Text, 6))) <> "P12" Then
      '2014/12/9 END
         If Combo6.Tag <> Combo6.Text And ProState = "1" Then 'Add By Sindy 2014/7/2 +if
            arrCaseNo = Split(Me.LBL1(7).Caption, "-")
            'Modify By Sindy 2014/6/30 只要非系統設定的人員均要檢查權限
            If Trim(m_PP05) <> Trim(Left(Combo6.Text, 6)) Then
            'If Trim(Left("" & Combo1.Text, 6)) <> strUserNum And Trim(txtCP144) <> "" Then
            '2014/6/30 END
               'Modify By Sindy 2024/6/26 +m_Country
               If PUB_ChkPromoterReader(arrCaseNo(0), m_CP10, "2", Trim(Left(Combo6.Text, 6)), , m_Country) = False Then
                  'Add By Sindy 2015/5/22
                  For ii = 0 To Me.Combo6.ListCount - 1
                      blnMatch = False
                      If Trim(Left(Me.Combo6.List(ii), 6)) = Trim(Left(Me.Combo6.Text, 6)) Then
                          Me.Combo6.ListIndex = ii
                          blnMatch = True
                          Exit For
                      End If
                  Next ii
                  If blnMatch = False Then
                     MsgBox "此人無判發權限，請重新輸入！"
                     SSTab1.Tab = 1 'Add By Sindy 2024/3/15
                     Combo6.SetFocus
                     Exit Function
                  End If
                  '2015/5/22 END
               End If
            End If
         End If
         '2013/11/26 END
      End If
   Else
     'Add By Sindy 2013/9/10
     If PUB_GetST07(Trim(Left("" & Combo1.Text, 6))) = "專99" Then
        MsgBox "[專99] 判發人不可空白!!!", vbExclamation + vbOKOnly
        SSTab1.Tab = 1 'Add By Sindy 2024/3/15
        Combo6.SetFocus
        Exit Function
     End If
   End If
End If
'2013/9/4 END

'Add By Sindy 2015/3/4 檢查英文核稿人
If Combo4.Enabled = True Then
   '不可任意清除英文核稿人,主管除外
   If Combo4.Tag <> Combo4.Text And Trim(Combo4.Text) = "" Then
'      'Add By Sindy 2015/7/22 有”可取消英核之人員”的權限,且無須英核時,可在個人工作管理拿掉英文核稿人員
'      If Not (ProState = "1" And _
'              InStr(Pub_GetSpecMan("可取消英核之人員"), strUserNum) > 0 And _
'              m_CP14 = strUserNum And _
'              (m_PER04 = "" Or m_PER04 = m_CP14)) Then
'      '2015/7/22 END
      'Modify By Sindy 2016/3/17
      If Not (ProState = "1" And _
              InStr(Pub_GetSpecMan("可取消英核之人員"), strUserNum) > 0 And _
              m_CP14 = strUserNum) Then
      '2016/3/17 END
         If ProState <> "2" Then
            'Add By Sindy 2017/6/1 + if 無須送英核
            If Not (bolHadSetProofEngReader = False And m_PER04 = "") Then
            '2017/6/1 END
               If txt1(23) = "1" Then
                  MsgBox "英文核稿人不可空白!!!", vbExclamation + vbOKOnly
                  SSTab1.Tab = 1 'Add By Sindy 2024/3/15
                  Combo4.SetFocus
                  Exit Function
               End If
            End If
         Else
            'Add By Sindy 2017/6/1 + if 無須送英核
            If Not (bolHadSetProofEngReader = False And m_PER04 = "") Then
            '2017/6/1 END
               If Val(txt1(3)) <= 0 And txt1(23) <> "2" Then '無完稿日且非日核
                  'Modify By Sindy 2015/12/8
   '               If MsgBox("因無完稿日當工程師再進入此筆資料維護時，系統仍然會預設英文核稿人" & vbCrLf & _
   '                         "，確定現在取消嗎？", vbYesNo + vbDefaultButton2) = vbNo Then
   '                  Exit Function
   '               End If
                  MsgBox "因無完稿日當工程師再進入此筆資料維護時，系統仍然會預設英文核稿人，無法取消英核!!!", vbExclamation + vbOKOnly
                  Combo4.Text = Combo4.Tag
                  SSTab1.Tab = 1 'Add By Sindy 2024/3/15
                  Combo4.SetFocus
                  Exit Function
                  '2015/12/8 END
               End If
            End If
         End If
      End If
   End If
End If
'2015/3/4 END

'檢查承辦期限
If Me.txt1(12).Text <> "" And txt1(12).Enabled = True Then
    If Me.txt1(2).Text = "" Then
        MsgBox "無齊備日不可輸入承辦期限!!!", vbExclamation + vbOKOnly
        SSTab1.Tab = 1 'Add By Sindy 2024/3/15
        Exit Function
    End If
End If

'add by nickc 2006/10/23 有齊備日的，承辦期限若是空白再抓一次
If txt1(2).Text <> "" And txt1(12) = "" Then
    txt1_LostFocus 2
End If

'add by nickc 2006/10/12 加入日期檢查  edit by nickc 2006/10/16 不管  預定會稿日
If Trim(txt1(2)) = "" And Trim(txt1(3)) & Trim(txt1(4)) & Trim(txt1(7)) & Trim(txt1(8)) <> "" And txt1(2).Enabled = True And txt1(2).Enabled = True Then
    MsgBox "有下列日期，齊備日不能空白！" & vbCrLf & "完稿日、會稿日、會稿完成日、發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(2).SetFocus
    Exit Function
End If
If Trim(txt1(3)) = "" And Trim(txt1(4)) & Trim(txt1(7)) & Trim(txt1(8)) <> "" And txt1(3).Enabled = True And txt1(3).Enabled = True Then
    MsgBox "有下列日期，完稿日不能空白！" & vbCrLf & "會稿日、會稿完成日、發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(3).SetFocus
    Exit Function
End If
If Trim(txt1(4)) = "" And Trim(txt1(7)) & Trim(txt1(8)) <> "" And txt1(4).Enabled = True And txt1(4).Enabled = True Then
    MsgBox "有下列日期，會稿日不能空白！" & vbCrLf & "會稿完成日、發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(4).SetFocus
    Exit Function
End If
If Trim(txt1(7)) = "" And Trim(txt1(8)) <> "" And txt1(7).Enabled = True And txt1(7).Enabled = True Then
    MsgBox "有下列日期，會稿完成日不能空白！" & vbCrLf & "發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(7).SetFocus
    Exit Function
End If
'add by nickc 2008/01/24 若是專利案件，申請案需要有繪圖，避免日期代不過去
'2008/1/28 modify by sonia 瓊玉說只要控制新申請案即可
'If ((SystemNumber(lbl1(7).Caption, 1) = "P" And InStr(1, CaseMapIn & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Or (SystemNumber(lbl1(7).Caption, 1) = "CFP") And InStr(1, CaseMapOut & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Then
'2009/9/11 MODIFY BY SONIA P,CFP案用NewCasePtyList判斷新申請案
'If ((SystemNumber(lbl1(7).Caption, 1) = "P" And InStr(1, CaseMapIn, m_CP10) <> 0) Or (SystemNumber(lbl1(7).Caption, 1) = "CFP") And InStr(1, CaseMapOut, m_CP10) <> 0) And PUB_GetST03(Trim(Left("" & Combo1.Text, 6))) <> "P12" Then
If txt1(2) <> "" Then 'Added by Morgan 2013/10/9 +判斷有齊備日才檢查否則只是要新增連絡會無法作業--薛經理
   If ((SystemNumber(LBL1(7).Caption, 1) = "P" Or SystemNumber(LBL1(7).Caption, 1) = "CFP") And InStr(1, NewCasePtyList, m_CP10) <> 0) And PUB_GetST03(Trim(Left("" & Combo1.Text, 6))) <> "P12" Then
   '2009/9/11 END
       If Trim(Combo2.Text) = "" Then
           MsgBox "專利申請案，繪圖人員不可空白，若不清楚給哪位繪圖人員，請選擇該區繪圖主管！", vbExclamation + vbOKOnly, "操作錯誤！"
           SSTab1.Tab = 1 'Add By Sindy 2024/3/15
           Combo2.SetFocus
           Exit Function
       End If
   End If
End If 'Added by Morgan 2013/10/9 +判斷有齊備日才檢查否則只是要新增連絡會無法作業--薛經理

'Remove by Morgan 2010/10/13
''Add by Morgan 2008/10/23
''會稿日大於本所期限會稿加乘註記需為不適用
'If SystemNumber(lbl1(7).Caption, 1) = "P" Or SystemNumber(lbl1(7).Caption, 1) = "CFP" Then
'   If Txt1(4).Enabled = True And Txt1(4).Enabled = true And lbl1(17) <> "" And opCP112(1).Value = False Then
'      If Val(Txt1(4)) > Val(Replace(lbl1(17).Caption, "/", "")) Then
'          MsgBox "會稿日大於本所期限會稿加乘註記需為不適用!!"
'          Txt1(4).SetFocus
'          txt1_GotFocus (4)
'          Exit Function
'      End If
'   End If
'End If

'Add by Morgan 2010/9/8 預定會稿日不可晚於本所期限 --王副總
If txt1(18).Enabled = True Then
   If Trim(txt1(18).Text) <> "" And txt1(18) <> txt1(18).Tag Then
      If LBL1(17).Caption <> "" Then
         strDate = Replace(LBL1(17).Caption, "/", "")
         'Modify by Morgan 2010/10/29 改為P案為所限-1個工作天;CFP案為所限-2個工作天
         strMsg = ""
         If m_FieldList(0).fiNewData = "P" Or InStr(",103,105,901,902,1002,1006,1201,1205,1206,1209,", "," & m_CP10 & ",") > 0 Then
            strDate = CompWorkDay(1, CompDate(2, -1, strDate), 1)
            strMsg = "前1個工作天"
         ElseIf m_FieldList(0).fiNewData = "CFP" Then
            strDate = CompWorkDay(2, CompDate(2, -1, strDate), 1)
            strMsg = "前2個工作天"
         End If
         strDate = TransDate(strDate, 1)
         'end 2010/10/29
         
         If Val(txt1(18)) > Val(strDate) Then
            MsgBox "預定會稿日不可晚於〔本所期限" & strMsg & "〕！", vbExclamation, "操作錯誤！"
            SSTab1.Tab = 1 'Add By Sindy 2024/3/15
            txt1(18).SetFocus
            txt1_GotFocus (18)
            Exit Function
         ElseIf Val(txt1(18)) < Val(txt1(2)) Then
            MsgBox "預定會稿日不可早於齊備日！", vbExclamation, "操作錯誤！"
            SSTab1.Tab = 1 'Add By Sindy 2024/3/15
            txt1(18).SetFocus
            txt1_GotFocus (18)
            Exit Function
         End If
      End If
      
   '有承辦期限時預定會稿日不可清除
   ElseIf txt1(12) <> "" And txt1(18) = "" And txt1(18).Tag <> "" Then
      If (m_FieldList(0).fiNewData = "P" And Left(m_FieldList(11).fiNewData, 1) <> "F") Or m_FieldList(0).fiNewData = "CFP" Then
         MsgBox "預定會稿日不可空白！", vbExclamation, "操作錯誤！"
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         txt1(18).SetFocus
         txt1_GotFocus (18)
         Exit Function
      End If
   End If
    
'Removed by Morgan 2012/7/20 個人不可修改預定會稿日
'   'Add by Morgan 2010/10/29
'   If ProState <> "2" And txt1(12) <> "" And txt1(18) <> txt1(18).Tag Then
'      strDate = ""
'      strMsg = ""
'      If m_FieldList(0).fiNewData = "P" Or InStr(",103,105,901,902,1002,1006,1201,1205,1206,1209,", "," & m_CP10 & ",") > 0 Then
'         strDate = CompWorkDay(5, CompDate(2, 1, txt1(12)))
'         strMsg = "+5個工作天"
'      ElseIf m_FieldList(0).fiNewData = "CFP" Then
'         strDate = CompWorkDay(10, CompDate(2, 1, txt1(12)))
'         strMsg = "+10個工作天"
'      End If
'      If strDate <> "" Then
'         strDate = TransDate(strDate, 1)
'         If Val(txt1(18)) > Val(strDate) Then
'            MsgBox "預定會稿日不可晚於〔承辦期限" & strMsg & "〕！", vbExclamation, "操作錯誤！"
'            txt1(18).SetFocus
'            txt1_GotFocus (18)
'            Exit Function
'         End If
'      End If
'   End If
'   'end 2010/10/29
End If

'Removed by Morgan 2015/5/25 改存檔後呼叫共用函數
'ChkRefCaseDate 'Add by Morgan 2010/7/20

'Added by Lydia 2016/12/30 P或CFP之主動修正(203)或修正(204)，若未收費未取消收文(cp159=0)時，若該案件進度檔沒有未發文未取消收文之A類申復(205)時，承辦備註欄不可空白；
If cmdFAmend.Visible = True And Trim(Replace(Replace(PUB_RepToOneSpace(txtEP12.Text), Chr(13), ""), Chr(10), "")) = "" Then
   'Add By Sindy 2020/10/6 增加歷程時間點判斷
   If ((txt1(5) <> Trim(Left("" & Combo1.Text, 6)) And Val(txt1(3)) > 0) Or _
       txt1(5) = "") And _
      ((txt1(1) = "Y" And Val(txt1(7)) > 0) Or _
       txt1(1) = "N") Then
   '2020/10/6 END
      MsgBox LBL1(15).Caption & "的承辦備註欄不可空白!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1 'Add By Sindy 2024/3/15
      If txtEP12.Enabled = True Then 'Added by Lydia 2018/05/11 判斷可執行
          txtEP12.SetFocus
      End If
      Exit Function
   End If
End If
'end 2016/12/30

'Added by Lydia 2021/12/28 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If txtEP12 <> "" Or txtCP99 <> "" Or txtCP144 <> "" Then
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
          Exit Function
    End If
End If
'end 2021/12/29

TxtValidate = True
End Function

'Add By Cheng 2003/04/28
Sub StrMenuOneRec(Optional ByVal Strindex As Integer = 1)
Dim ii As Integer
    For ii = 1 To Me.GRD1.Rows - 1
        '若目次相同, 收文號也相同
        'edit by nickc 2007/12/18
        'If Me.grd1.TextMatrix(ii, 0) = Me.lbl1(0).Caption And Me.grd1.TextMatrix(ii, 21) = m_strCP09 Then
        'Modify by Morgan 2009/7/14 加欄位:預會日 15(所有15以後欄位序號+1)
        'If Me.grd1.TextMatrix(ii, 0) = Me.lbl1(0).Caption And Me.grd1.TextMatrix(ii, 22) = m_strCP09 Then
        'Modified by Lydia 2025/02/05 改用變數
        'If Me.grd1.TextMatrix(ii, 0) = Me.lbl1(0).Caption And Me.grd1.TextMatrix(ii, 23) = m_strCP09 Then
        If Me.GRD1.TextMatrix(ii, 0) = Me.LBL1(0).Caption And Me.GRD1.TextMatrix(ii, colCP09_1) = m_strCP09 Then
            '承辦期限
            'Modify by Morgan 2010/10/7
            'Me.grd1.TextMatrix(ii, 10) = Me.lbl1(2).Caption
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 10) = ChangeTStringToTDateString(Me.txt1(12).Text)
            Me.GRD1.TextMatrix(ii, colCp48_1) = ChangeTStringToTDateString(Me.txt1(12).Text)
            'end 2010/10/7
            '齊備日
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 13) = ChangeTStringToTDateString(Me.txt1(2).Text)
            Me.GRD1.TextMatrix(ii, colEp06_1) = ChangeTStringToTDateString(Me.txt1(2).Text)
            '完稿日
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 14) = ChangeTStringToTDateString(Me.txt1(3).Text)
            Me.GRD1.TextMatrix(ii, colEp09_1) = ChangeTStringToTDateString(Me.txt1(3).Text)
            '預會日
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 15) = ChangeTStringToTDateString(Me.txt1(18).Text)
            Me.GRD1.TextMatrix(ii, colEp28_1) = ChangeTStringToTDateString(Me.txt1(18).Text)
            '會稿日
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 16) = ChangeTStringToTDateString(Me.txt1(4).Text)
            Me.GRD1.TextMatrix(ii, colEp07_1) = ChangeTStringToTDateString(Me.txt1(4).Text)
            '核稿人
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 17) = GetStaffName(Me.txt1(5).Text, True)
            Me.GRD1.TextMatrix(ii, colEp04_1) = GetStaffName(Me.txt1(5).Text, True)
            '會稿完成日
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 18) = ChangeTStringToTDateString(Me.txt1(7).Text)
            Me.GRD1.TextMatrix(ii, colEp08_1) = ChangeTStringToTDateString(Me.txt1(7).Text)
            '發文日
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 19) = ChangeTStringToTDateString(Me.txt1(8).Text)
            Me.GRD1.TextMatrix(ii, colCp27_1) = ChangeTStringToTDateString(Me.txt1(8).Text)
            '承辦天數
            '計算承辦天數
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 20) = "" & Intnick910123
             Me.GRD1.TextMatrix(ii, colEp35_1) = "" & Intnick910123
            '備註
            'Modified by Lydia 2025/02/05 改用變數
            'Me.grd1.TextMatrix(ii, 21) = Me.txtEP12.Text
            Me.GRD1.TextMatrix(ii, colEp12_1) = Me.txtEP12.Text
            
            'Modify by Morgan 2011/1/4 修正日期欄位排序問題(小於100年的前面補空白)
            For intI = 10 To 21
               If Len(GRD1.TextMatrix(ii, intI)) = 8 Then
                 If Mid(GRD1.TextMatrix(ii, intI), 3, 1) = "/" And Mid(GRD1.TextMatrix(ii, intI), 6, 1) = "/" Then
                    GRD1.TextMatrix(ii, intI) = " " & GRD1.TextMatrix(ii, intI)
                 End If
               End If
            Next
            'end 2011/1/4
            
            'Modify By Sindy 2021/3/15 + , False
            ChgGrdColor ii, False 'Add by Morgan 2009/11/12
            Exit For
        End If
    Next ii
    
    'ChgGrdColor 'Remove by Morgan 2009/11/12 移到回圈內
    
'    SWPRow2 = "1"
    SWPRow2 = Strindex
    GRD1.row = Val(SWPRow2)
    GRD1.col = 1
End Sub

'add by nick  2004/07/16
' StrIndex 本所案號
' false 為不存在，true 存在
'檢查該本所案號之所有申請人(非個人才檢查)是否有在 ApplicantDiscount 出現
Function CheckCuInAD(Strindex As String) As Boolean
If UCase(Left(Strindex, 1)) = "N" Then
    Strindex = Right(Strindex, Len(Strindex) - 1)
End If
Dim strTemp As String
strSql = "select count(*) from (select pa26 pacu  FROM PATENT ,customer " & _
           " WHERE pa09='000' and pa01='" & SystemNumber(Strindex, 1) & "'  and pa02='" & SystemNumber(Strindex, 2) & "'  and pa03='" & SystemNumber(Strindex, 3) & "'  and pa04='" & SystemNumber(Strindex, 4) & "'  " & _
           " and substr(nvl(pa26,''),1,8)=cu01(+) and cu15 not in ('0') " & _
           " Union select pa27 pacu FROM PATENT ,customer " & _
           " WHERE pa09='000' and pa01='" & SystemNumber(Strindex, 1) & "'  and pa02='" & SystemNumber(Strindex, 2) & "'  and pa03='" & SystemNumber(Strindex, 3) & "'  and pa04='" & SystemNumber(Strindex, 4) & "'  " & _
           " and substr(nvl(pa27,''),1,8)=cu01(+) and cu15 not in ('0') " & _
           " Union select pa28 pacu FROM PATENT ,customer " & _
           " WHERE pa09='000' and pa01='" & SystemNumber(Strindex, 1) & "'  and pa02='" & SystemNumber(Strindex, 2) & "'  and pa03='" & SystemNumber(Strindex, 3) & "'  and pa04='" & SystemNumber(Strindex, 4) & "'  " & _
           " and substr(nvl(pa28,''),1,8)=cu01(+) and cu15 not in ('0') " & _
           " Union select pa29 pacu FROM PATENT ,customer " & _
           " WHERE  pa09='000' and pa01='" & SystemNumber(Strindex, 1) & "'  and pa02='" & SystemNumber(Strindex, 2) & "'  and pa03='" & SystemNumber(Strindex, 3) & "'  and pa04='" & SystemNumber(Strindex, 4) & "'  " & _
           " and substr(nvl(pa29,''),1,8)=cu01(+) and cu15 not in ('0') " & _
           " Union select pa30 pacu FROM PATENT ,customer " & _
           " WHERE pa09='000' and pa01='" & SystemNumber(Strindex, 1) & "'  and pa02='" & SystemNumber(Strindex, 2) & "'  and pa03='" & SystemNumber(Strindex, 3) & "'  and pa04='" & SystemNumber(Strindex, 4) & "'  and substr(nvl(pa30,''),1,8)=cu01(+) and cu15 not in ('0') ) NewCU  where substr(newcu.pacu,1,8) not in (select ad01 from applicantdiscount where ad02='000') "
CheckOC3
AdoRecordSet3.CursorLocation = adUseClient
AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
    If AdoRecordSet3.Fields(0).Value <> 0 Then
        CheckCuInAD = False
    Else
        CheckCuInAD = True
    End If
End If
CheckOC3
End Function

'add by nick 2005/01/31
'檢查國內案件
Function CheckCaseMapAndCp14() As Boolean
CheckOC3
strSql = "select C1.cp14 as A,C2.Cp14 as B from casemap,caseprogress c1,caseprogress c2 where cm01='" & SystemNumber(Trim(LBL1(7).Caption), 1) & "' and cm02='" & SystemNumber(Trim(LBL1(7).Caption), 2) & "' and cm03='" & SystemNumber(Trim(LBL1(7).Caption), 3) & "' and cm04='" & SystemNumber(Trim(LBL1(7).Caption), 4) & "' and cm10='0' " & _
             " and cm01=c1.cp01(+) and cm02=c1.cp02(+) and cm03=c1.cp03(+) and cm04=c1.cp04(+) and cm05=c2.cp01(+) and cm06=c2.cp02(+) and cm07=c2.cp03(+) and cm08=c2.cp04(+) "
With AdoRecordSet3
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .MaxRecords <> 0 Then
        If "" & .Fields("A").Value = "" & .Fields("B").Value Then
            CheckCaseMapAndCp14 = True
        Else
            CheckCaseMapAndCp14 = False
        End If
    Else
        CheckCaseMapAndCp14 = False
    End If
End With
CheckOC3
End Function
'add by nickc 2006/02/27 控制只跟 DB 溝通一次
' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To TF_CP
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CP" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      '定義數字
      Select Case nIndex
         Case 5, 6, 7, 15, 16, 17, 18, 19, 25, 27, 33, 34, 46, 47, 48, 53, 54, 57, 66, 67, 69, 70, 73, 74, 75, 76, 77, 78, 79, 82, 84, 85, 97, 98, 100, 101, 103, 104, 108, 109, 111:
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

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   For nIndex = 0 To TF_CP - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            'add by nickc 2006/03/14
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

'CANCEL BY SONIA 2014/4/9 沒使用
''add by nickc 2008/01/03 計算預估的考核值
'Function StrMenu6(oOld As String, oCp09 As String) As String
'StrMenu6 = oOld
'Dim rs1 As New ADODB.Recordset   '最外圈用
'Dim Rs2 As New ADODB.Recordset   '單一承辦人所有案子
'Dim Rs3 As New ADODB.Recordset   '單一承辦人當天所有適用的案子
'Dim Rs4 As New ADODB.Recordset   '會稿加乘註記
'Dim Rs5 As New ADODB.Recordset   '大陸案件會稿加乘檢查
'Dim tmpTarget As Double          '目標
'Dim tmpAddUpTarget As Double     '目前累計承辦量
'Dim tmpSt01 As String            '暫存承辦人
'Dim MaxWorkDay As Integer        '本月最大工作天
'Dim NowWorkDay As Integer        '月初到現在有多少工作天
'Dim OneWorkDay As Double         '一個工作天需做多少件
'Dim ThisSvrDate As String        '截止日
'Dim i As Integer                 '暫時變數
'Dim CalDay As Integer            '目前累計工作天
'Dim IsDelay As Boolean           '是否延遲
'Dim EpDay As Integer             '承辦天數
'Dim CalYM As String
'Dim oST01 As String
'Dim StrSQLa As String
'Dim StrSqlB As String
''Add by Morgan 2008/10/23
'Dim strRule1 As String           '適用規則
'Dim strRule2 As String           '次規則代碼
'Dim strRule3 As String           '是否有核稿人
'Dim strDate1 As String           '以收文日計算起算日
'Dim strDate2 As String           '以本所期限計算起算日
'
'Select Case ProState
'Case "1" '承辦人個人工作進度資料維護
'         CalYM = Mid(strSrvDate(1), 1, 6)
'         'Modify By Sindy 2013/9/17
'         'oST01 = strUserNum
'         oST01 = Trim(Left("" & Combo1.Text, 6))
'         '2013/9/17 END
'         StrSQLa = "Select Count(*) From WorkDay Where WD01>='" & Mid(strSrvDate(1), 1, 6) & "01' And WD01<='" & CompWorkDay(2, strSrvDate(1), 1) & "' "
'         StrSqlB = "Select Count(*) From WorkDay Where WD01>='" & Mid(strSrvDate(1), 1, 6) & "01' And WD01<='" & Mid(strSrvDate(1), 1, 6) & "31' "
'         ThisSvrDate = CompWorkDay(2, strSrvDate(1), 1)
'Case "2" '承辦人管理工作進度資料查詢
'         CalYM = Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2))
'         oST01 = Trim(Left("" & Combo1.Text, 6))
'         StrSQLa = "Select Count(*) From WorkDay Where WD01>='" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01' And WD01<='" & IIf(Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) = Mid(strSrvDate(1), 1, 6), CompWorkDay(2, strSrvDate(1), 1), Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31") & "' "
'         StrSqlB = "Select Count(*) From WorkDay Where WD01>='" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01' And WD01<='" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31' "
'         ThisSvrDate = IIf(Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) = Mid(strSrvDate(1), 1, 6), CompWorkDay(2, strSrvDate(1), 1), ChangeWDateStringToWString(DateAdd("d", -1, DateAdd("m", 1, ChangeWStringToWDateString(Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01")))))
'Case Else
'End Select
''先取得所有承辦人及其目標和核稿人
''跑迴圈，一次只算一個承辦人
''抓該承辦人所有案子，累計承辦量，若適用則更新會稿加乘註記，不適用則繼續累加承辦量
'
'NowWorkDay = 0
'MaxWorkDay = 0
''抓截至目前為止工作天
'Set rs1 = New ADODB.Recordset
'rs1.CursorLocation = adUseClient
'rs1.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rs1.RecordCount <> 0 Then
'    NowWorkDay = CheckStr(rs1.Fields(0))
'    If NowWorkDay = 0 Then
'        StrMenu6 = "Err."
'        GoTo ExitPort
'    End If
'Else
'    StrMenu6 = "Err."
'    GoTo ExitPort
'End If
'
''抓所有工作天
'Set rs1 = New ADODB.Recordset
'rs1.CursorLocation = adUseClient
'rs1.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
'If rs1.RecordCount <> 0 Then
'    MaxWorkDay = CheckStr(rs1.Fields(0))
'    If MaxWorkDay = 0 Then
'        StrMenu6 = "Err."
'        GoTo ExitPort
'    End If
'Else
'    StrMenu6 = "Err."
'    GoTo ExitPort
'End If
'
'
''查所有工程師和他的目標，沒目標的不出來
'strSql = "select st01,Sum(decode(pe02,'CFP',(Nvl(PE05,0)+Nvl(PE07,0)) * 2,Nvl(PE05,0)+Nvl(PE07,0))) from staff,performance where st03>='P10' and st03<='P11' and st04='1' And (ST05='72' or st05='74' Or ST05='76' Or ST05='77' Or ST05='78' Or ST05='79' or st05='87') And ST01<>'88024' And ST01<'F' and pe01=st01(+) and pe03='" & CalYM & "' and pe01='" & oST01 & "'  group by st01"
'Set rs1 = New ADODB.Recordset
'rs1.CursorLocation = adUseClient
'rs1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If rs1.RecordCount <> 0 Then
'    rs1.MoveFirst
'    Do While Not rs1.EOF
'        tmpTarget = Val(CheckStr(rs1.Fields(1))) / MaxWorkDay
'        tmpSt01 = CheckStr(rs1.Fields(0))
'        'add by nickc 2006/06/01 加入檢查滿半年沒
'        strSql = "select st13 from staff where st01='" & tmpSt01 & "' "
'        Set Rs2 = New ADODB.Recordset
'        If Rs2.State = 1 Then Rs2.Close
'        Rs2.CursorLocation = adUseClient
'        Rs2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'        If Rs2.RecordCount <> 0 And Val(CheckStr(Rs2.Fields(0))) <> 0 Then
'            If DateDiff("m", ChangeWStringToWDateString(CheckStr(Rs2.Fields("st13"))), ChangeWStringToWDateString(ThisSvrDate)) >= 7 Then
'            Else    '未滿半年不算
'                GoTo ExitPort
'            End If
'        End If
'
'                IsDelay = False
'                '查該工程師1號到該天的計件，沒計件值的不出來
'                'Modify by Morgan 2010/3/24 達成計算要考慮會稿加乘註記
'                'Modify by Morgan 2010/3/29 支援紀錄也要算
'                'Memo by Morgan 2010/3/30 是否達成目標用累計會稿量判斷(上月完稿當月才會稿的也算,可接受上月份統計數字有變動)--柄佑
'                'strSql = "select sum(cp97 * cp98) from caseprogress,engineerprogress where cp14='" & tmpSt01 & "' aND ep02=cp09(+) and ep07>='" & CalYM & "01' and ep07<='" & ThisSvrDate & "'  order by ep07"
'                'Modified by Morgan 2014/3/20 2014/4/1 起支援改每小時折計0.2基數
'                'strSql = "select sum(pp) from (select sum(cp97 * cp98 * decode(cp112,'Y',nvl(cp111,1),1)) pp from caseprogress,engineerprogress where cp14='" & tmpSt01 & "' aND ep02=cp09(+) and ep07>='" & CalYM & "01' and ep07<='" & ThisSvrDate & "'" & _
'                " union all Select Sum(Round(Decode(SH06, 'CFP', Nvl(SH05,0)/3, Nvl(SH05,0)/4) ,2)) pp From SupportHour Where SH02='" & tmpSt01 & "' And SH01>=" & (Val(txt1(0).Text) + 191100) & "01 And SH01<=" & (Val(txt1(0).Text) + 191100) & Format(i, "00") & " And SH11='V' ) X"
'                strSql = "select sum(pp) from (select sum(cp97 * cp98 * decode(cp112,'Y',nvl(cp111,1),1)) pp from caseprogress,engineerprogress where cp14='" & tmpSt01 & "' aND ep02=cp09(+) and ep07>='" & CalYM & "01' and ep07<='" & ThisSvrDate & "'" & _
'                " union all Select Sum(Round(" & Sh2EPtCode & " ,2)) pp From SupportHour Where SH02='" & tmpSt01 & "' And SH01>=" & (Val(txt1(0).Text) + 191100) & "01 And SH01<=" & (Val(txt1(0).Text) + 191100) & Format(i, "00") & " And SH11='V' ) X"
'                'end 2014/3/20
'                Set Rs2 = New ADODB.Recordset
'                If Rs2.State = 1 Then Rs2.Close
'                Rs2.CursorLocation = adUseClient
'                Rs2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                '若是有資料，且 計件值不為 0
'                If Rs2.RecordCount <> 0 Then
'                    '若是當時計件值未達當時承辦目標時
'                    If Val(CheckStr(Rs2.Fields(0))) < (tmpTarget * NowWorkDay) Then
'                        IsDelay = True
'                    End If
'                    '查出所有當天資料，且適用的
'                    'Modify by Morgan 2008/10/27 要區分A,B規則,會稿加乘註記改抓 MeetScript
'                    'strSQL = "select *  from caseprogress,engineerprogress,promoterproofreader where cp14='" & tmpSt01 & "' aND cp09=ep02(+) and cp09='" & oCp09 & "'  and cp01=pp01(+) and cp14=pp02(+) and cp10=pp03(+) and cp112='Y' and cp10 in ('101','102') order by ep07 "
'                    strSql = "select *  from caseprogress,engineerprogress,promoterproofreader,casepropertymap where cp14='" & tmpSt01 & "' aND cp09=ep02(+) and cp09='" & oCp09 & "'  and cp01=pp01(+) and cp14=pp02(+) and cp10=pp03(+) and cp112='Y' and cpm01(+)=cp01 and cpm02(+)=cp10 order by ep07 "
'                    Set Rs3 = New ADODB.Recordset
'                    If Rs3.State = 1 Then Rs3.Close
'                    Rs3.CursorLocation = adUseClient
'                    Rs3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                    If Rs3.RecordCount <> 0 Then
'                        Rs3.MoveFirst
'                        Do While Not Rs3.EOF
'                           strSql = ""
'                           strRule1 = "" & Rs3.Fields("cpm05")
'                           'Modify by Morgan 2010/2/1 +C規則(CFP案適用P案A規則者)
'                           If strRule1 = "A" Or strRule1 = "C" Then
'                              '檢查齊備到會稿幾個工作天
'                              EpDay = GetWorkDay(ThisSvrDate, CheckStr(Rs3.Fields("ep06")))
'                              '抓會稿加乘註記
'                              '優先判斷是否有預設核稿人(若沒有預設核稿人時即使個案有設也用無核稿人的加乘註記)
'                              If CheckStr(Rs3.Fields("pp04")) = "" Then
'                                   '沒核稿人
'                                   'strSQL = "select ea07 from EngAssemble where ea01='2' and ea02='" & CheckStr(rs3.Fields("cp01")) & "' and ea03<=" & EpDay & " and ea05>=" & EpDay & " " & IIf(IsDelay = False, " and ea07 >1 ", "")
'                                   strRule3 = "2"
'                              Else
'                                  'add by nickc 2007/03/06 加入 P 的不管，CFP 的，若是承辦人與核稿人相同，表示為無核稿人
'                                  If CheckStr(Rs3.Fields("cp01")) = "CFP" And CheckStr(Rs3.Fields("EP05")) = CheckStr(Rs3.Fields("EP04")) Then
'                                      'strSQL = "select ea07 from EngAssemble where ea01='2' and ea02='" & CheckStr(rs3.Fields("cp01")) & "' and ea03<=" & EpDay & " and ea05>=" & EpDay & " " & IIf(IsDelay = False, " and ea07 >1 ", "")
'                                      strRule3 = "2"
'                                  ElseIf CheckStr(Rs3.Fields("EP05")) = "" Then
'                                      'strSQL = "select ea07 from EngAssemble where ea01='2' and ea02='" & CheckStr(rs3.Fields("cp01")) & "' and ea03<=" & EpDay & " and ea05>=" & EpDay & " " & IIf(IsDelay = False, " and ea07 >1 ", "")
'                                      strRule3 = "2"
'                                  Else
'                                      'strSQL = "select ea07 from EngAssemble where ea01='1' and ea02='" & CheckStr(rs3.Fields("cp01")) & "' and ea03<=" & EpDay & " and ea05>=" & EpDay & " " & IIf(IsDelay = False, " and ea07 >1 ", "")
'                                      strRule3 = "1"
'                                  End If
'                              End If
'
'                              If strRule3 = "1" Then
'                                 strSql = "select MS06 from MeetScript where MS01='" & strRule1 & "' and MS02='1' and MS03='" & CheckStr(Rs3.Fields("cp01")) & "' and MS04<=" & EpDay & " and MS05>=" & EpDay & " " & IIf(IsDelay = False, " and MS06 >1 ", "")
'                              Else
'                                 strSql = "select MS07 from MeetScript where MS01='" & strRule1 & "' and MS02='1' and MS03='" & CheckStr(Rs3.Fields("cp01")) & "' and MS04<=" & EpDay & " and MS05>=" & EpDay & " " & IIf(IsDelay = False, " and MS06 >1 ", "")
'                              End If
'                           '規則B
'                           ElseIf "" & Rs3.Fields("cpm05") = "B" Then
'                              strRule2 = "1"
'                              If Not IsNull(Rs3.Fields("cp06")) Then
'                                 'P案若超過14個工作天或本所期限前6個工作天，以較早的為準，適用倒扣
'                                 If Rs3.Fields("cp01") = "P" Then
'                                    strDate1 = CompWorkDay(15, CheckStr(Rs3.Fields("ep06")))
'                                    strDate2 = CompWorkDay(5, Rs3.Fields("cp06"), 1)
'                                 'CFP案若超過20個工作天或本所期限前10個工作天，以較早的為準，適用倒扣
'                                 Else
'                                    strDate1 = CompWorkDay(21, CheckStr(Rs3.Fields("ep06")))
'                                    strDate2 = CompWorkDay(9, Rs3.Fields("cp06"), 1)
'                                 End If
'                                 If strDate2 <= strDate1 Then
'                                    strRule2 = "2"
'                                 End If
'                              End If
'                              '以所限計算
'                              If strRule2 = "2" Then
'                                 EpDay = GetWorkDay(Rs3.Fields("CP06"), ThisSvrDate)
'                              '以齊備日計算
'                              Else
'                                 EpDay = GetWorkDay(ThisSvrDate, Rs3.Fields("EP06"))
'                              End If
'                              strSql = "select MS06 from MeetScript where MS01='" & strRule1 & "' and MS02='" & strRule2 & "' and MS03='" & CheckStr(Rs3.Fields("cp01")) & "' and MS04<=" & EpDay & " and MS05>=" & EpDay & " " & IIf(IsDelay = False, " and MS06 >1 ", "")
'                           End If
'
'                           If strSql <> "" Then
'                              Set Rs4 = New ADODB.Recordset
'                              If Rs4.State = 1 Then Rs4.Close
'                              Rs4.CursorLocation = adUseClient
'                              Rs4.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                              If Rs4.RecordCount <> 0 Then
'                                 If Rs4.Fields(0) > 1 And Rs3("cp01") = "P" And (Rs3("cp10") = "101" Or Rs3("cp10") = "102" Or Rs3("cp10") = "103") Then
'                                    '若與台灣案同一工程師時大陸案不適用大於1的會稿加乘註記
'                                    strSql = "select 1 from patent a where pa01='" & Rs3("cp01") & "' and pa02='" & Rs3("cp02") & "' and pa03='" & Rs3("cp03") & "' and pa04='" & Rs3("cp04") & "' and pa09='020'" & _
'                                       " and exists(select * from casemap,caseprogress,patent b where cm01=a.pa01 and cm02=a.pa02 and cm03=a.pa03 and cm04=a.pa04 and cm10='0'" & _
'                                       " and cp01(+)=cm05 and cp02(+)=cm06 and cp03(+)=cm07 and cp04(+)=cm08 and cp10 in ('101','102','103') and cp14='" & Rs3("cp14") & "'" & _
'                                       " and b.pa01(+)=cp01 and b.pa02(+)=cp02 and b.pa03(+)=cp03 and b.pa04(+)=cp04 and b.pa09='000')"
'                                    If Rs5.State = 1 Then Rs5.Close
'                                    Rs5.CursorLocation = adUseClient
'                                    Rs5.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                                    If Rs5.RecordCount <> 0 Then
'                                       GoTo ExitPort
'                                    Else
'                                       StrMenu6 = Format(Trim(Val(CheckStr(Rs4.Fields(0))) * Val(CheckStr(Rs3.Fields("cp97"))) * Val(CheckStr(Rs3.Fields("cp98")))), "0.00")
'                                    End If
'                                 Else
'                                    StrMenu6 = Format(Trim(Val(CheckStr(Rs4.Fields(0))) * Val(CheckStr(Rs3.Fields("cp97"))) * Val(CheckStr(Rs3.Fields("cp98")))), "0.00")
'                                 End If
'
'                                 GoTo ExitPort
'                              Else
'                                 GoTo ExitPort
'                              End If
'                           End If
'                           Rs3.MoveNext
'                        Loop
'                    Else
'                        GoTo ExitPort
'                    End If
'                End If
'        rs1.MoveNext
'    Loop
'End If
'
'ExitPort:
'
'Set rs1 = Nothing
'Set Rs2 = Nothing
'Set Rs3 = Nothing
'Set Rs4 = Nothing
'Set Rs5 = Nothing
'End Function
'2014/4/9 END

'Add by Morgan 2008/12/3 自cmdok_Click抽出
Private Function FormSave() As Boolean
   'add by nickc 2007/11/29 若勾選無圖式，複製無圖式的圖檔給該案
   Dim BytesS() As Byte
   Dim BytesVal As String
   Dim PicRs As New ADODB.Recordset
   'Add by Morgan 2008/12/9
   Dim stRefDate As String '計算用的暫存日期
   Dim stRefDateDesc As String '計算用的暫存日期中文說明
   Dim stCP(1 To 4) As String '本所案號
   'Add by Morgan 2009/10/5
   Dim st020Msg As String '台灣案會稿大陸案加費用提醒
   Dim stVTB As String 'Add by Morgan 2009/11/3
   Dim iMouse As Integer
   Dim bol201NoSkip As Boolean
   Dim arr1, arr2 'Add by Morgan 2010/10/4
   Dim stDate(3) As String 'Added by Morgan 2012/9/6
   Dim intMaxEEP02 As Integer, strSubject As String, strContent As String 'Add By Sindy 2015/3/4
   Dim bolUpdateEP06 As Boolean 'Added by Morgan 2016/5/4
   Dim p_FileName As String, strFtpPath As String
   
   iMouse = Screen.MousePointer
   
On Error GoTo ErrHnd

'Modify By Sindy 2018/6/20 + If m_Flow = "" Then
If m_Flow = "" Then cnnConnection.BeginTrans

cnnConnection.Execute "begin user_data.user_formname:='" & Me.Name & "';end;" 'Add by Morgan 2010/10/29


   'Add by Morgan 2010/9/23 翻譯要判斷是否有收文其他案件性質
   bol201NoSkip = True
   If m_CP10 = "201" Then
      strExc(0) = "select * from caseprogress a where cp09='" & m_strCP09 & "'" & _
         " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04" & _
         " and b.cp09<>a.cp09 and b.cp10 in (" & CaseMapIn & ") and b.cp57 is null)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         bol201NoSkip = False
      End If
   End If
      

'Modified by Morgan 2015/5/25 改存檔後呼叫共用函數
'   'Add by Morgan 2010/7/21
'   '更新關聯副案的相同事項
'   'Modified by Morgan 2015/5/25 +會稿完成日也要更新
'   If m_strNeedUpdateCase <> "" And txt1(3) & txt1(18) & txt1(4) <> "" Then
'      strExc(0) = ""
'      If txt1(3) <> "" Then
'         strExc(0) = strExc(0) & IIf(strExc(0) <> "", ",", "") & "ep09=nvl(ep09," & DBDATE(txt1(3)) & ")"
'      End If
'      If txt1(18) <> "" Then
'         strExc(0) = strExc(0) & IIf(strExc(0) <> "", ",", "") & "ep28=nvl(ep28," & DBDATE(txt1(18)) & ")"
'      End If
'      If txt1(4) <> "" Then
'         strExc(0) = strExc(0) & IIf(strExc(0) <> "", ",", "") & "ep07=nvl(ep07," & DBDATE(txt1(4)) & ")"
'      End If
'      'Added by Morgan 2015/5/25
'      If txt1(7) <> "" Then
'         strExc(0) = strExc(0) & IIf(strExc(0) <> "", ",", "") & "ep08=nvl(ep08," & DBDATE(txt1(7)) & ")"
'      End If
'      'end 2015/5/25
'
'      strSql = "update engineerprogress set " & strExc(0) & " where ep02 in (" & m_strNeedUpdateCase & ") and ep06>0"
'      cnnConnection.Execute strSql, intI
'   End If
'   'end 2010/7/21
'Removed by Morgan 2016/3/9--柄佑 105/1/25 請作單
'系統在主案「會稿」及「送會」時，不詢問工程師其他多國CFP案是否同步處理。
'   If Left(lbl1(7), 3) = "CFP" And txt1(3) & txt1(18) & txt1(4) & txt1(7) <> "" Then
'      strExc(0) = "delete asklist where al01='" & lbl1(3) & "'"
'      cnnConnection.Execute strExc(0), intI
'
'      strExc(0) = "insert into asklist(al01,al02) values('" & lbl1(3) & "','" & m_CP14 & "')"
'      cnnConnection.Execute strExc(0), intI
'   End If
'end 2016/3/9
'end 2015/5/25

   stCP(1) = SystemNumber(Trim(LBL1(7).Caption), 1)
   stCP(2) = SystemNumber(Trim(LBL1(7).Caption), 2)
   stCP(3) = SystemNumber(Trim(LBL1(7).Caption), 3)
   stCP(4) = SystemNumber(Trim(LBL1(7).Caption), 4)
   
   '計算承辦天數
   If Len(txt1(4)) <> 0 And Len(txt1(2)) <> 0 And Val(txt1(4)) <> 0 And Val(txt1(2)) <> 0 Then
      Intnick910123 = GetWorkDay(ChangeWDateStringToWString(ChangeTStringToWDateString(txt1(4))), ChangeWDateStringToWString(ChangeTStringToWDateString(txt1(2))))
   Else
      If Len(txt1(3)) <> 0 And Len(txt1(2)) <> 0 And Val(txt1(3)) <> 0 And Val(txt1(2)) <> 0 Then
         Intnick910123 = GetWorkDay(ChangeWDateStringToWString(ChangeTStringToWDateString(txt1(3))), ChangeWDateStringToWString(ChangeTStringToWDateString(txt1(2))))
      Else
         Intnick910123 = 0
      End If
   End If
                   
'Remove by Morgan 2010/10/13 新規則改由 Trigger 設定(程式先改後上)
'   'Add by Morgan 2008/12/3 若 [是否適用會稿加乘註記CP112] 有修改時需更新 [加乘註記適用設定是否人工修改CP122]
'   '因為Trigger判斷要需要，本更新需放在會稿日(會觸發重設CP112的欄位)更新之前。
'   If opCP112(0).Value = True Then
'      m_stNewCP112 = "Y"
'   ElseIf opCP112(1).Value Then
'      m_stNewCP112 = "N"
'   Else
'      m_stNewCP112 = ""
'   End If
'   m_bAlterCP112 = False
'   If m_stNewCP112 <> m_CP112 Then
'      m_bAlterCP112 = True
'      '若改成不適用時記錄有人工修改,否則清除之
'      If m_stNewCP112 = "N" Then
'         strSql = "UPDATE CASEPROGRESS SET CP122='Y' WHERE CP09='" & Me.lbl1(3).Caption & "'"
'      Else
'         strSql = "UPDATE CASEPROGRESS SET CP122=NULL WHERE CP09='" & Me.lbl1(3).Caption & "'"
'      End If
'      cnnConnection.Execute strSql, intI
'      Pub_SaveLog strUserNum, "適不適用會稿加乘註記異動：[" & m_CP112 & "]==>[" & m_stNewCP112 & "]", SystemNumber(lbl1(7).Caption, 1), SystemNumber(lbl1(7).Caption, 2), SystemNumber(lbl1(7).Caption, 3), SystemNumber(lbl1(7).Caption, 4), lbl1(3).Caption
'   End If
'   'end 2008/12/3
'end 2010/10/13
                   
   'add by nickc 2007/11/29 若勾選無圖式，複製無圖式的圖檔給該案
   If Chk1.Value = vbChecked Then
      strSql = "select * from ImgByteFile where ibf01='000' and ibf02='000000' and ibf03='0' and ibf04='01'"
      Set PicRs = New ADODB.Recordset
      PicRs.CursorLocation = adUseClient
      PicRs.Open strSql, cnnConnection, adOpenStatic, adLockOptimistic
      If PicRs.RecordCount <> 0 Then
         BytesVal = PicRs.Fields("ibf13").Value
'         ReDim BytesS(Val(BytesVal))
'         BytesS() = PicRs.Fields("ibf14").GetChunk(Val(BytesVal))
         'Add By Sindy 2017/8/10 下載檔案
         p_FileName = App.path & "\TempFile"
         RidFile p_FileName
         If "" & PicRs.Fields("IBF15") <> "" Then
            If PUB_GetFtpFile(PicRs.Fields("IBF15"), p_FileName, UCase("ImgByteFile")) = False Then
               GoTo ErrHnd
            End If
         End If
         '2017/8/10 END
         PicRs.AddNew
         PicRs.Fields("ibf07").Value = strUserNum
         PicRs.Fields("ibf08").Value = Val(strSrvDate(1))
         PicRs.Fields("ibf09").Value = Val(Format(time, "HHMM"))
         PicRs.Fields("ibf01").Value = SystemNumber(LBL1(7).Caption, 1)
         PicRs.Fields("ibf02").Value = SystemNumber(LBL1(7).Caption, 2)
         PicRs.Fields("ibf03").Value = SystemNumber(LBL1(7).Caption, 3)
         PicRs.Fields("ibf04").Value = SystemNumber(LBL1(7).Caption, 4)
         PicRs.Fields("ibf05").Value = "1"
         PicRs.Fields("ibf06").Value = "6"
         PicRs.Fields("ibf13").Value = BytesVal
'         PicRs.Fields("ibf14").Value = Null
'         PicRs.Fields("ibf14").AppendChunk BytesS()
         'Modify By Sindy 2017/8/10
         '檔案改放FTP
         If FileExists(p_FileName) Then
            PUB_PutFtpFile p_FileName, SystemNumber(LBL1(7).Caption, 1) & "-" & SystemNumber(LBL1(7).Caption, 2) & "-" & SystemNumber(LBL1(7).Caption, 3) & "-" & SystemNumber(LBL1(7).Caption, 4) & "-1", SystemNumber(LBL1(7).Caption, 1) & "-" & SystemNumber(LBL1(7).Caption, 2) & "-" & SystemNumber(LBL1(7).Caption, 3) & "-" & SystemNumber(LBL1(7).Caption, 4) & "-1", strFtpPath, UCase("imgbytefile")
            If strFtpPath <> "" Then
               PicRs.Fields("ibf15") = strFtpPath
            End If
         End If
         '2017/8/10 END
         PicRs.UPDATE
      End If
   End If
                   
   SeekTmpBk = Trim(LBL1(0).Caption)
   'edit by nickc 2007/01/29 移除 ep13 下面 cp29 控制
   'edit by nickc 2007/08/22 加入是否暫停核稿、英文核稿人、英文核完日
   'Modiyf By Sindy 2013/9/4 +,EP40='" & txtCP144 & "'
   'Modified by Morgan 2013/10/8 EP08 改用 UpdateEp08 更新
   'strSql = "Update EngineerProgress Set EP06=" & IIf(ChangeTStringToWString(txt1(2)) = "", "NULL", ChangeTStringToWString(txt1(2))) & ",EP09=" & IIf(ChangeTStringToWString(txt1(3)) = "", "NULL", ChangeTStringToWString(txt1(3))) & ",EP07=" & IIf(ChangeTStringToWString(txt1(4)) = "", "NULL", ChangeTStringToWString(txt1(4))) & _
      ",EP40='" & txtCP144 & "',EP04='" & txt1(5) & "',EP03='" & txt1(6) & "',EP08=" & IIf(ChangeTStringToWString(txt1(7)) = "", "NULL", ChangeTStringToWString(txt1(7))) & ",EP31=" & IIf(ChangeTStringToWString(txt1(13)) = "", "NULL", ChangeTStringToWString(txt1(13))) & _
      ",EP11='" & txt1(9) & "',EP12='" & txtep12 & "',EP34='" & txt1(1) & "', EP35=" & IIf(Intnick910123 = 0, "Null", Intnick910123) & ",ep32=" & IIf(Trim(txt1(20)) = "", "null ", "'" & txt1(20) & "' ") & ",ep33=" & IIf(ChangeTStringToWString(txt1(19)) = "", "NULL", ChangeTStringToWString(txt1(19))) & " Where EP02='" & lbl1(3).Caption & "' "
   'Add By Sindy 2015/3/13 +核稿語文 EP41=" & CNULL(txt1(23))
   strSql = "Update EngineerProgress Set EP06=" & IIf(ChangeTStringToWString(txt1(2)) = "", "NULL", ChangeTStringToWString(txt1(2))) & ",EP40='" & Trim(Left("" & Combo6.Text, 6)) & "',EP04='" & txt1(5) & "',EP03='" & txt1(6) & "',EP31=" & IIf(ChangeTStringToWString(txt1(13)) = "", "NULL", ChangeTStringToWString(txt1(13))) & _
            ",EP41=" & CNULL(txt1(23)) & ",EP11='" & txt1(9) & "',EP12='" & txtEP12 & "',EP34='" & txt1(1) & "', EP35=" & IIf(Intnick910123 = 0, "Null", Intnick910123) & ",ep32=" & IIf(Trim(txt1(20)) = "", "null ", "'" & txt1(20) & "' ")
   'Modify By Sindy 2013/12/18 防止簽核流程已存入日期,但此處又更新到日期,如英文核完日 ex.CFP-023734
   '完稿日
   If Val(txt1(3).Tag) <> Val(txt1(3)) Then
      strSql = strSql & ",EP09=" & IIf(ChangeTStringToWString(txt1(3)) = "", "NULL", ChangeTStringToWString(txt1(3)))
   End If
   '會稿日
   strEP07Tag = Val(txt1(4).Tag) 'Add By Sindy 2022/10/31
   If Val(txt1(4).Tag) <> Val(txt1(4)) Then
      strSql = strSql & ",EP07=" & IIf(ChangeTStringToWString(txt1(4)) = "", "NULL", ChangeTStringToWString(txt1(4)))
   End If
   '英文核完日
   If Val(m_EP33) <> Val(ChangeTStringToWString(txt1(19))) Then
      strSql = strSql & ",ep33=" & IIf(ChangeTStringToWString(txt1(19)) = "", "NULL", ChangeTStringToWString(txt1(19)))
   End If
   '2013/12/18 END
   strSql = strSql & " Where EP02='" & LBL1(3).Caption & "' "
   cnnConnection.Execute strSql
   
   'add by nickc 2007/08/22 若是改變<<是否暫停核稿>>，要記錄
   If txt1(20).Tag <> txt1(20).Text Then
      strSql = "insert into EngManLog (em01,em02,em03,em04,em05) select '" & LBL1(3).Caption & "',nvl(max(em02),0)+1,to_number(to_char(sysdate,'YYYYMMDD')),'" & IIf(Trim(txt1(20)) = "", "0", "1") & "','" & strUserNum & "' from EngManLog where em01='" & LBL1(3).Caption & "' "
      cnnConnection.Execute strSql
   End If
   
'Remove by Morgan 2011/4/28 改由 Trigger 執行
'   'add by nickc 2006/03/24 若是不會稿的案子，不更新墨圖的齊備日
'   '若墨圖計件(EP29)則墨齊日(EP17)=會稿完成日(EP08)
'   '若已上墨齊日則不更新
'   'Modify by Morgan 2011/3/31 若草圖要計件但無草齊日則會稿完成時一併更新為會完日
'   strSql = "Update EngineerProgress Set EP14=nvl(ep14,decode(ep20,null,ep08))" & _
'      " ,EP17=nvl(ep17,decode(ep29,null,ep08)) Where EP02='" & lbl1(3).Caption & "' "
'   cnnConnection.Execute strSql
   
   If txt1(14) = "" Then
      strSql = "Update EngineerProgress Set EP27=NULL Where EP02='" & LBL1(3).Caption & "' "
      cnnConnection.Execute strSql
   Else
      If m_EP27 = "" Then
         strSql = "Update EngineerProgress Set EP27=" & ServerDate & " Where EP02='" & LBL1(3).Caption & "' "
         cnnConnection.Execute strSql
      End If
   End If
   

   'add by nickc 2006/03/07 主管改預定會稿日，要發 mail
   'edit by nickc 2006/07/25 協理跟秀玲說要可以拿掉沒有請作單 CFP-018708
   'edit by nickc 2006/09/08 個人可以輸一次
   'Modify by Morgan 2008/10/13 預定會稿日異動改判斷tag(本來以label放原資料)
   If Trim(txt1(18).Tag) <> Trim(txt1(18).Text) Then
      'Add by Morgan 2010/9/28 只修改(新增統一凌晨做)
      If (stCP(1) = "P" Or stCP(1) = "CFP") And Not PUB_IfSetCP48(m_strCP09) Then
         strSql = "Update EngineerProgress Set EP28=" & CNULL(ChangeTStringToWString(txt1(18))) & ",EP30=nvl(EP30,0)+1 Where EP02='" & LBL1(3).Caption & "' and EP28>0"
         Pub_SeekTbLog strSql
      Else
      'end 2010/9/28
         strSql = "Update EngineerProgress Set EP28=" & CNULL(ChangeTStringToWString(txt1(18))) & " Where EP02='" & LBL1(3).Caption & "' "
      End If
      cnnConnection.Execute strSql, intI

'Remove by Morgan 2010/11/5 取消發EMail--秀玲
'
'      '輸入預定會稿日要發 mail
'      If m_CP13 <> "" Then
'         'edit by nickc 2006/12/29 Mail改在 trans 後發
'         ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
'         skMail(UBound(skMail)).fiSender = strUserNum
'         skMail(UBound(skMail)).fiReceiver = m_CP13
'         'Modify by Morgan 2008/12/3 +案件性質(因改成不只是新案才可輸預定會稿日)
'         skMail(UBound(skMail)).fiContent = "本所案號：" & lbl1(7) & vbCrLf & "收文號：" & lbl1(3) & vbCrLf & "案件名稱：" & lbl1(9) & vbCrLf & "案件性質：" & lbl1(15) & vbCrLf & "預定會稿日：" & ChangeTStringToTDateString(txt1(18))
'         skMail(UBound(skMail)).fiSubject = lbl1(7) & "已輸入預定會稿日！"
'         skMail(UBound(skMail)).fiRecriverNo = ""
'      End If
   End If
                  
   '更新案件進度檔的繪圖人員代號欄
   SetFieldNewData "CP29", Me.txt1(0).Text
   'Remove by Morgan 2010/10/13
   'SetFieldNewData "CP112", IIf(opCP112(0).Value = False And opCP112(1).Value = False, "", IIf(opCP112(0).Value = True, "Y", "N"))
   
   'add by nickc 2005/03/04  加入儲存加乘註記及理由
   If Val(txt1(15)) <> Val(m_CP98) Then
      SetFieldNewData "CP98", Me.txt1(15).Text
      SetFieldNewData "CP99", Me.txtCP99.Text
      '增加紀錄
      strSql = "insert into flagstory (fs01,fs02,fs03,fs04,fs05,fs06,fs07,fs08) select '" & Me.LBL1(3).Caption & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MIss')),'1','" & m_CP98 & "','" & Trim(txt1(15)) & "','" & ChgSQL(Trim(txtCP99)) & "','" & strUserNum & "' from dual  "
      cnnConnection.Execute strSql
      
      'add by nickc 2005/04/13  發 mail
      'edit by nickc 2006/12/29 Mail改在 trans 後發
      ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
      skMail(UBound(skMail)).fiSender = strUserNum
      'Added by Lydia 2023/04/24 修改王副總退休之相關控制
      'Modified by Morgan 2025/2/21 73022->pub_PMan
      'If strSrvDate(1) >= "20230511" Then
      '    skMail(UBound(skMail)).fiReceiver = "73022;99050"
      'ElseIf strSrvDate(1) >= "20230501" Then
      '    skMail(UBound(skMail)).fiReceiver = "71011;73022;99050"
      'Else
      ''end 2025/2/20
      ''end 2023/04/24
      '   skMail(UBound(skMail)).fiReceiver = "71011"
      'End If 'Added by Lydia 2023/04/24
      pub_PMan = Pub_GetSpecMan("專利處特定編號")
      skMail(UBound(skMail)).fiReceiver = pub_PMan & ";99050"
      'end 2025/2/21
      skMail(UBound(skMail)).fiContent = "本所案號：" & LBL1(7).Caption & vbCrLf & "收文號：" & LBL1(3).Caption & vbCrLf & "原加乘註記：" & m_CP98 & vbCrLf & "更改後加乘註記：" & txt1(15) & vbCrLf & "理由：" & txtCP99
      skMail(UBound(skMail)).fiSubject = "更改承辦人加乘註記"
      skMail(UBound(skMail)).fiRecriverNo = Me.LBL1(3).Caption
   End If

   '是否提供圖檔
   SetFieldNewData "CP106", IIf(Len(Trim(txt1(17).Text)) = 0, "", txt1(17).Text)
   If Mid(LBL1(3).Caption, 1, 1) = "C" Then
     SetFieldNewData "CP27", IIf(ChangeTStringToWString(txt1(8)) = "", "", ChangeTStringToWString(txt1(8)))
   End If
   'Modify by Morgan 2010/10/7
   'SetFieldNewData "CP48", IIf(lbl1(2) <> "", ChangeTStringToWString(ChangeTDateStringToTString(lbl1(2))), "")
   SetFieldNewData "CP48", IIf(txt1(12) <> "", ChangeTStringToWString(txt1(12)), "")
   'end 2010/10/7
   
   'Added by Morgan 2016/7/7
   '開放可設定複雜案件
   SetFieldNewData "CP147", txt1(11)
   'end 2016/7/7
   
   'add by nick 2004/11/30 更新繪圖的草圖齊備日
   'edit by nick 2004/12/10 分所不會先上繪圖人員
   'Modify by Morgan 2011/3/31 國外主案文齊日先不上等會稿完成時和墨齊日一起上
   'cnnConnection.Execute "update  EngineerProgress set ep14=ep06 Where EP02='" & lbl1(3).Caption & "' and ep14 is null and ep20 is null "
   'Modify by Morgan 2011/4/21 國外案文齊都不必更新草齊統一等會完時一起更新--瓊玉
   'strSql = "select * from caseprogress,casemap where cp09='" & lbl1(3).Caption & "' and cp21 is null and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm10='0'"
   
'Removed by Morgan 2014/4/14 改在 Trigger (ENGINEERPROGRESS_BEFORE4) 做，否則更新多國案齊備時會沒有更新到 Ex.CFP-26745 -> CFP-26747-1-00
   'strSql = "select * from caseprogress,casemap where cp09='" & lbl1(3).Caption & "'  and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm10='0'"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   'If intI = 0 Then
   '   cnnConnection.Execute "update  EngineerProgress set ep14=ep06 Where EP02='" & lbl1(3).Caption & "' and ep14 is null and ep20 is null "
   'End If
'end 2014/4/14

   'end 2011/3/31
                      
   'Memo by Morgan 2008/12/9 上齊備日時更新相同承辦人案件的齊備日和承辦期限(不同承辦人則在上)
   
   'add by nick 2004/12/01 更新相關卷號齊備日(輸入文件齊備日時從無到有)
   'edit by nick 2005/01/27 加要更新承辦期限(每筆重算)
   Dim tmpM_cp48 As String
   Dim tmpM_cp09 As String
'Removed by Morgan 2016/3/9--柄佑 105/1/25 請作單
'在「主案」依據規則設定「齊備日」時，同一人承辦的其他多國案不設定「齊備日」
'   If lbl1(8) = "" And txt1(2) <> "" Then
'      'edit by nickc 2006/03/09 加入未發文未取消收文
'      'edit by nickc 2007/10/24 秀玲說只要控制申請案
'      strSql = "select cp06,cp09,cp01,pa09,cp10 from caseprogress,caserelation,engineerprogress,patent where cr01='" & SystemNumber(lbl1(7).Caption, 1) & "' and cr02='" & SystemNumber(lbl1(7).Caption, 2) & "' and cr03='" & SystemNumber(lbl1(7).Caption, 3) & "' and cr04='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
'         " and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and cp21='Y' and cp09=ep02(+) and ep06 is null  and ep05='" & m_CP14 & "' and cr05=pa01(+) and cr06=pa02(+) and cr07=pa03(+) and cr08=pa04(+)  and cp27 is null and cp57 is null and cp10 in (" & GetAddStr(CaseMapOut) & ") "
'      CheckOC3
'      With AdoRecordSet3
'         .CursorLocation = adUseClient
'         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If .RecordCount <> 0 Then
'            .MoveFirst
'            Do While Not .EOF
'               tmpM_cp09 = "" & .Fields("cp09").Value
'               cnnConnection.Execute "update engineerprogress set ep06=" & ChangeTStringToWString(txt1(2)) & " where ep02='" & tmpM_cp09 & "' "
'               If PUB_IfSetCP48() Then 'Add by Morgan 2010/10/7 CR只有CFP案沒有FMP問題
'                  'edit by nickc 2007/10/05  抓有時效的承辦期限
'                  tmpM_cp48 = Pub_GetHandleDay(CheckStr(.Fields("CP01")), CheckStr(.Fields("PA09")), CheckStr(.Fields("CP10")), ChangeTStringToWString(txt1(2)), CheckStr(.Fields("cp06")), CheckStr(.Fields("cp09")))
'                  If tmpM_cp48 <> "" Then
'                     cnnConnection.Execute "update caseprogress set cp48=" & tmpM_cp48 & " Where cp09='" & tmpM_cp09 & "'  "
'                  End If
'               End If 'Add by Morgan 2010/10/7
'               .MoveNext
'            Loop
'         End If
'      End With
'      CheckOC3
'   End If
'end 2016/3/9
   
'Removed by Morgan 2013/10/8 改在 UpdateEp08
'
'   'add by nickc 2006/09/14 加入上會完日的，多國案相同承辦人的資料,前面要是沒上要一併上，上過的不管
'   'edit by nickc 2008/04/24 郭雅娟請作單，限制新申請案
'   'Modify by Morgan 2010/9/23 新案翻譯201要排除有收文其他新案案件性質
'   'If ChangeTStringToWString(txt1(7)) <> "" And InStr(1, CaseMapIn, m_CP10) <> 0 And (SystemNumber(lbl1(7).Caption, 1) = "CFP" Or SystemNumber(lbl1(7).Caption, 1) = "P") Then
'   If ChangeTStringToWString(txt1(7)) <> "" And InStr(1, CaseMapIn, m_CP10) <> 0 And bol201NoSkip And (SystemNumber(lbl1(7).Caption, 1) = "CFP" Or SystemNumber(lbl1(7).Caption, 1) = "P") Then
'      '齊備
'      cnnConnection.Execute "update EngineerProgress set EP06=" & IIf(ChangeTStringToWString(txt1(2)) = "", "NULL", ChangeTStringToWString(txt1(2))) & " Where EP02 in (select cp09 from caseprogress,caserelation,engineerprogress where cr01='" & SystemNumber(lbl1(7).Caption, 1) & "' and cr02='" & SystemNumber(lbl1(7).Caption, 2) & "' and cr03='" & SystemNumber(lbl1(7).Caption, 3) & "' and cr04='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
'                            " and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and cp21='Y' and cp09=ep02(+) and ep06 is null and ep08 is null  and ep05='" & m_CP14 & "' and cp27 is null and cp57 is null ) "
'      '會稿
'      '2010/1/14 MODIFY BY SONIA 加限制日本案不更新(柄佑提因英日文同承辦人但會稿時間不同)
'      'cnnConnection.Execute "update  EngineerProgress set EP07=" & IIf(ChangeTStringToWString(txt1(4)) = "", "NULL", ChangeTStringToWString(txt1(4))) & " Where EP02 in (select cp09 from caseprogress,caserelation,engineerprogress where cr01='" & SystemNumber(lbl1(7).Caption, 1) & "' and cr02='" & SystemNumber(lbl1(7).Caption, 2) & "' and cr03='" & SystemNumber(lbl1(7).Caption, 3) & "' and cr04='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
'                                             " and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and cp21='Y' and cp09=ep02(+) and ep07 is null and ep08 is null  and ep05='" & m_CP14 & "' and cp27 is null and cp57 is null ) "
'      cnnConnection.Execute "update EngineerProgress set EP07=" & IIf(ChangeTStringToWString(txt1(4)) = "", "NULL", ChangeTStringToWString(txt1(4))) & " Where EP02 in (select cp09 from caseprogress,caserelation,engineerprogress,patent where cr01='" & SystemNumber(lbl1(7).Caption, 1) & "' and cr02='" & SystemNumber(lbl1(7).Caption, 2) & "' and cr03='" & SystemNumber(lbl1(7).Caption, 3) & "' and cr04='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
'                            " and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and cp21='Y' and cp09=ep02(+) and ep07 is null and ep08 is null  and ep05='" & m_CP14 & "' and cp27 is null and cp57 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09<>'011' ) "
'      '完稿
'      '2010/1/14 MODIFY BY SONIA 加限制日本案不更新(柄佑提因英日文同承辦人但完稿時間不同)
'      'cnnConnection.Execute "update EngineerProgress set EP09=" & IIf(ChangeTStringToWString(txt1(3)) = "", "NULL", ChangeTStringToWString(txt1(3))) & ",EP35=" & IIf(Intnick910123 = 0, "Null", Intnick910123) & " Where EP02 in (select cp09 from caseprogress,caserelation,engineerprogress where cr01='" & SystemNumber(lbl1(7).Caption, 1) & "' and cr02='" & SystemNumber(lbl1(7).Caption, 2) & "' and cr03='" & SystemNumber(lbl1(7).Caption, 3) & "' and cr04='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
'                                             " and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and cp21='Y' and cp09=ep02(+) and ep09 is null and ep08 is null  and ep05='" & m_CP14 & "' and cp27 is null and cp57 is null ) "
'      cnnConnection.Execute "update EngineerProgress set EP09=" & IIf(ChangeTStringToWString(txt1(3)) = "", "NULL", ChangeTStringToWString(txt1(3))) & ",EP35=" & IIf(Intnick910123 = 0, "Null", Intnick910123) & " Where EP02 in (select cp09 from caseprogress,caserelation,engineerprogress,patent where cr01='" & SystemNumber(lbl1(7).Caption, 1) & "' and cr02='" & SystemNumber(lbl1(7).Caption, 2) & "' and cr03='" & SystemNumber(lbl1(7).Caption, 3) & "' and cr04='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
'                            " and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and cp21='Y' and cp09=ep02(+) and ep09 is null and ep08 is null  and ep05='" & m_CP14 & "' and cp27 is null and cp57 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09<>'011' ) "
'      'add by nick 2004/12/01 更新相關卷號會稿完成日
'      'edit by nickc 2006/03/07 加入未發文未取消收文的才更新
'      '2010/1/14 MODIFY BY SONIA 加限制日本案不更新(柄佑提因英日文同承辦人但會稿完成時間不同)
'      'cnnConnection.Execute "update EngineerProgress set EP08=" & IIf(ChangeTStringToWString(txt1(7)) = "", "NULL", ChangeTStringToWString(txt1(7))) & " Where EP02 in (select cp09 from caseprogress,caserelation,engineerprogress where cr01='" & SystemNumber(lbl1(7).Caption, 1) & "' and cr02='" & SystemNumber(lbl1(7).Caption, 2) & "' and cr03='" & SystemNumber(lbl1(7).Caption, 3) & "' and cr04='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
'         " and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and cp21='Y' and cp09=ep02(+) and ep08 is null  and ep05='" & m_CP14 & "' and cp27 is null and cp57 is null ) "
'      cnnConnection.Execute "update EngineerProgress set EP08=" & IIf(ChangeTStringToWString(txt1(7)) = "", "NULL", ChangeTStringToWString(txt1(7))) & " Where EP02 in (select cp09 from caseprogress,caserelation,engineerprogress,patent where cr01='" & SystemNumber(lbl1(7).Caption, 1) & "' and cr02='" & SystemNumber(lbl1(7).Caption, 2) & "' and cr03='" & SystemNumber(lbl1(7).Caption, 3) & "' and cr04='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
'                            " and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and cp21='Y' and cp09=ep02(+) and ep08 is null  and ep05='" & m_CP14 & "' and cp27 is null and cp57 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09<>'011' ) "
'   End If
'end 2013/10/8
   
   'add by nick   2004/12/16 P,000  設定繪圖人(從無到有)時，國外案之繪圖未設定
   'edit by nickc 2005/04/11 改不管如何，都要更新
   'edit by nickc 2005/11/01 工程師的補文件和其他都不管繪圖，若有要畫圖，由繪圖人員請工程師補
   'edit by nickc 2005/11/04 加入 實體審查不更新
   '2006/1/24 MODIFY BY SONIA 郭說不限制台灣案,例P-78462及CFP-18273
   '2009/9/18 modify by sonia 改為輸在國內外案件關聯表的國內新申請案(不限P)才要更新其國外案的新申請案,國外案不可更新國內案
   'If SystemNumber(lbl1(7).Caption, 1) = "P" And m_CP10 <> "202" And m_CP10 <> "910" And m_CP10 <> "416" Then
   If InStr(LBL1(7).Caption, "P") > 0 And InStr(NewCasePtyList, m_CP10) > 0 Then
      '加入新案才更新
      '所有的P 設計才出來，其他只有台灣在出來
      'edit by nick 2005/04/11 不管如何都要更新，若有不同瓊玉會在分案改國外案的繪圖
      'edit by nickc 2005/04/12 改更新案件進度
      'edit by nickc 2005/11/23 若已經確認過分案的不更新
      'edit by nickc  2006/07/03 相關聯的大陸案不上繪圖人員
      '2009/4/29 MODIFY BY SONIA 瓊玉說取消大陸案限制 P-090796
      'cnnConnection.Execute "update  CaseProgress set cp29=" & IIf(Trim(txt1(0)) = "", "null ", " '" & txt1(0).Text & "' ") & " Where cp09 in (select cp09 from caseprogress,casemap,engineerprogress,patent where cm05='" & SystemNumber(lbl1(7).Caption, 1) & "' and cm06='" & SystemNumber(lbl1(7).Caption, 2) & "' and cm07='" & SystemNumber(lbl1(7).Caption, 3) & "' and cm08='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
         " and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and cp09=ep02(+) and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and cm10='0' and pa09<>'020' and cp107 is null and (cp01||cp10 in ('CFP101','CFP102','CFP103','CFP104','CFP109','CFP110','CFP112','CFP113','CFP114','CFP115','P103') or pa09||cp01||cp10 in ('000P101','000P102','000P104','000P105','000P109','000P110','000P112','000P113','000P114','000P115'))) "
      '2009/9/11 MODIFY BY SONIA 取消P110,加未發文條件
      'cnnConnection.Execute "update  CaseProgress set cp29=" & IIf(Trim(txt1(0)) = "", "null ", " '" & txt1(0).Text & "' ") & " Where cp09 in (select cp09 from caseprogress,casemap,engineerprogress,patent where cm05='" & SystemNumber(lbl1(7).Caption, 1) & "' and cm06='" & SystemNumber(lbl1(7).Caption, 2) & "' and cm07='" & SystemNumber(lbl1(7).Caption, 3) & "' and cm08='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
      '   " and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and cp09=ep02(+) and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and cm10='0' and cp107 is null and (cp01||cp10 in ('CFP101','CFP102','CFP103','CFP104','CFP109','CFP110','CFP112','CFP113','CFP114','CFP115','P103') or pa09||cp01||cp10 in ('000P101','000P102','000P104','000P105','000P109','000P110','000P112','000P113','000P114','000P115'))) "
      'Modified by Morgan 2013/6/26 取消分案確認條件,目前看來國外案不會與國內案不同(因為國外案有可能先設了繪圖人員導致與後設的國內案不同)--瓊玉
      'Modified by Morgan 2017/3/1 若國內案無繪圖人員但國外案已有繪圖人員時不要清除--翔龍 Ex.CFP-29257
      'cnnConnection.Execute "update CaseProgress set cp29=" & IIf(Trim(txt1(0)) = "", "null ", " '" & txt1(0).Text & "' ") & " Where cp09 in (select cp09 from caseprogress,casemap where cm05='" & SystemNumber(lbl1(7).Caption, 1) & "' and cm06='" & SystemNumber(lbl1(7).Caption, 2) & "' and cm07='" & SystemNumber(lbl1(7).Caption, 3) & "' and cm08='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
         " and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) AND CP27 IS NULL and instr('" & NewCasePtyList & "',cp10)>0 AND cp01||cp10 <>'P110') "
      cnnConnection.Execute "update CaseProgress set cp29=" & IIf(Trim(txt1(0)) = "", "cp29 ", " '" & txt1(0).Text & "' ") & " Where cp09 in (select cp09 from caseprogress,casemap where cm05='" & SystemNumber(LBL1(7).Caption, 1) & "' and cm06='" & SystemNumber(LBL1(7).Caption, 2) & "' and cm07='" & SystemNumber(LBL1(7).Caption, 3) & "' and cm08='" & SystemNumber(LBL1(7).Caption, 4) & "' " & _
         " and cm10='0' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) AND CP27 IS NULL and instr('" & NewCasePtyList & "',cp10)>0 AND cp01||cp10 <>'P110') "
      'end 2017/3/1
      '2009/9/11 END
      '2009/4/29 END
      
      'Added by Morgan 2013/5/20
      '更新一案兩請繪圖人員
      'Modified by Morgan 2013/6/26 取消分案確認條件
      cnnConnection.Execute "update CaseProgress set cp29=" & IIf(Trim(txt1(0)) = "", "null ", " '" & txt1(0).Text & "' ") & " Where cp09 in " & _
         "(select cp09 from caseprogress,casemap where cm05='" & SystemNumber(LBL1(7).Caption, 1) & "' and cm06='" & SystemNumber(LBL1(7).Caption, 2) & "' and cm07='" & SystemNumber(LBL1(7).Caption, 3) & "' and cm08='" & SystemNumber(LBL1(7).Caption, 4) & "' " & _
         " and cm10='3' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) AND CP27 IS NULL and instr('" & NewCasePtyList & "',cp10)>0 " & _
         " union all select cp09 from caseprogress,casemap where cm01='" & SystemNumber(LBL1(7).Caption, 1) & "' and cm02='" & SystemNumber(LBL1(7).Caption, 2) & "' and cm03='" & SystemNumber(LBL1(7).Caption, 3) & "' and cm04='" & SystemNumber(LBL1(7).Caption, 4) & "' " & _
         " and cm10='3' and cm05=cp01(+) and cm06=cp02(+) and cm07=cp03(+) and cm08=cp04(+) AND CP27 IS NULL and instr('" & NewCasePtyList & "',cp10)>0) "
      'end 2013/5/20
   End If
                    
   'add by nick 2004/12/16 CFP Cp21='Y' ep13=null 相關卷號的多國案
   '2009/9/18 MODIFY BY SONIA 多國案的主案且為新申請案的案件性質才更新其他多國案的新申請案,改多國案不更新其他案
   'If SystemNumber(lbl1(7).Caption, 1) = "CFP" Then
   If SystemNumber(LBL1(7).Caption, 1) = "CFP" And InStr(NewCasePtyList, m_CP10) > 0 Then
   '2009/9/18 END
      'edit by nickc 繪圖人員改更新 cp29(EP13)
      '2009/9/18 modify by sonia 多國案的主案才更新其他多國案未發文且繪圖主管未分案確認之新申請程序
      'cnnConnection.Execute "update CaseProgress set Cp29=" & IIf(m_EP13 <> Me.txt1(0).Text, CNULL(Me.txt1(0).Text), "(select EP13 from engineerprogress where ep02=cp09)") & " Where Cp09 in (select cp09 from caseprogress,caserelation,engineerprogress where cr01='" & SystemNumber(lbl1(7).Caption, 1) & "' and cr02='" & SystemNumber(lbl1(7).Caption, 2) & "' and cr03='" & SystemNumber(lbl1(7).Caption, 3) & "' and cr04='" & SystemNumber(lbl1(7).Caption, 4) & "' " & _
      '   " and cr05=cp01(+) and cr06=cp02(+) and cr07=cp03(+) and cr08=cp04(+) and cp21='Y' and cp09=ep02(+) and ep14 is null and instr('" & NewCasePtyList & "',cp10)>0 and cp27 is null and cp107 is null ) "
      cnnConnection.Execute "update CaseProgress set Cp29=" & IIf(m_EP13 <> Me.txt1(0).Text, CNULL(Me.txt1(0).Text), "(select EP13 from engineerprogress where ep02=cp09)") & " Where Cp09 in (select C1.cp09 from caseprogress C1,CASEPROGRESS C2,caserelation,engineerprogress where C2.cp09='" & LBL1(3) & "' AND C2.CP21 IS NULL AND C2.CP01=cr01 and C2.CP02=cr02 and C2.CP03=cr03 and C2.CP04=cr04 " & _
         " and cr05=C1.cp01(+) and cr06=C1.cp02(+) and cr07=C1.cp03(+) and cr08=C1.cp04(+) and C1.cp21='Y' and C1.cp09=ep02(+) and ep14 is null and instr('" & NewCasePtyList & "',C1.cp10)>0 and C1.cp27 is null and C1.cp107 is null ) "
      '2009/9/18 END
   End If
   
   'add by nickc 2006/01/11 當新申請案，輸入齊備日(從無到有)，發 mail 給智權人員
   'edit by nickc 2006/01/19 僅限專利
   'Modify by Morgan 2010/9/17 修改也要發Mail--秀玲
   'If lbl1(8) = "" And txt1(2) <> "" And ((SystemNumber(lbl1(7).Caption, 1) = "P" And InStr(1, CaseMapIn & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Or (SystemNumber(lbl1(7).Caption, 1) = "CFP") And InStr(1, CaseMapOut & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Then
   '2010/9/21 MODIFY BY SONIA 改西元日期比較
   'If lbl1(8) <> txt1(2) And ((SystemNumber(lbl1(7).Caption, 1) = "P" And InStr(1, CaseMapIn & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Or (SystemNumber(lbl1(7).Caption, 1) = "CFP") And InStr(1, CaseMapOut & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Then
   'Modify by Morgan 2010/9/23 新案翻譯201要排除有收文其他新案案件性質
   'If lbl1(8) <> ChangeTStringToWString(txt1(2)) And ((SystemNumber(lbl1(7).Caption, 1) = "P" And InStr(1, CaseMapIn & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Or (SystemNumber(lbl1(7).Caption, 1) = "CFP") And InStr(1, CaseMapOut & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Then
   
   'Modify by Morgan 2010/11/2 齊備,修改移到凌晨發
   'Modified by Morgan 2020/9/17 取消--郭雅娟
   If LBL1(8) <> ChangeTStringToWString(txt1(2)) And ((SystemNumber(LBL1(7).Caption, 1) = "P" And InStr(1, CaseMapIn & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Or (SystemNumber(LBL1(7).Caption, 1) = "CFP") And InStr(1, CaseMapOut & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) And bol201NoSkip Then
   'If txt1(2) = "" And lbl1(8) <> ChangeTStringToWString(txt1(2)) And ((SystemNumber(lbl1(7).Caption, 1) = "P" And InStr(1, CaseMapIn & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) Or (SystemNumber(lbl1(7).Caption, 1) = "CFP") And InStr(1, CaseMapOut & ",301,302,303,304,305,306,307,803", m_CP10) <> 0) And bol201NoSkip Then
   'end 2020/9/17
      'edit by nickc 2006/01/23 加案件名稱及客戶名稱
      'edit by nickc 2006/12/29 Mail改在 trans 後發
      ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
      skMail(UBound(skMail)).fiSender = strUserNum
      skMail(UBound(skMail)).fiReceiver = m_CP13
      'Modify By Sindy 2009/10/29
      'skMail(UBound(skMail)).fiContent = vbCrLf + "客戶名稱：" & m_CuNo & vbCrLf & "本所案號： " + lbl1(7) + vbCrLf + "案件名稱： " + lbl1(9).Caption + vbCrLf + "案件性質： " + lbl1(15).Caption + vbCrLf + "本所期限： " + lbl1(17).Caption + vbCrLf + "法定期限： " + lbl1(19).Caption + vbCrLf & vbCrLf & " 工程師已齊備！"
      'Modify by Morgan 2010/9/17
      'skMail(UBound(skMail)).fiContent = vbCrLf + "客戶名稱：" & m_CuNo & vbCrLf & "本所案號： " + lbl1(7) + vbCrLf + "案件名稱： " + lbl1(9).Caption + vbCrLf + "案件性質： " + lbl1(15).Caption + vbCrLf + "承  辦  人： " + lbl1(1).Caption + vbCrLf + "齊  備  日： " + Mid(txt1(2), 1, Len(txt1(2)) - 4) & "年" & Left(Right(txt1(2), 4), 2) & "月" & Right(txt1(2), 2) & "日" + vbCrLf + "本所期限： " + lbl1(17).Caption + vbCrLf + "法定期限： " + lbl1(19).Caption + vbCrLf & vbCrLf & " 工程師已齊備！"
      ''2010/4/13 modify by sonia 總收文號後加案件性質
      'skMail(UBound(skMail)).fiSubject = "[通知] " & lbl1(7) & "/" & lbl1(3) & lbl1(15) & "/" & lbl1(9) & "/" & m_CuNo & " 工程師已齊備！"
      If LBL1(8) = "" Then
         skMail(UBound(skMail)).fiContent = vbCrLf + "客戶名稱：" & m_CuNo & vbCrLf & "本所案號： " + LBL1(7) + vbCrLf + "案件名稱： " + LBL1(9).Caption + vbCrLf + "案件性質： " + LBL1(15).Caption + vbCrLf + "承  辦  人： " + LBL1(1).Caption + vbCrLf + "齊  備  日： " + Mid(txt1(2), 1, Len(txt1(2)) - 4) & "年" & Left(Right(txt1(2), 4), 2) & "月" & Right(txt1(2), 2) & "日" + vbCrLf + "本所期限： " + LBL1(17).Caption + vbCrLf + "法定期限： " + LBL1(19).Caption + vbCrLf & vbCrLf & " 工程師已齊備！"
         skMail(UBound(skMail)).fiSubject = "[通知] " & LBL1(7) & "/" & LBL1(3) & LBL1(15) & "/" & LBL1(9) & "/" & m_CuNo & " 工程師已齊備！"
      ElseIf txt1(2) = "" Then
         skMail(UBound(skMail)).fiContent = vbCrLf + "客戶名稱：" & m_CuNo & vbCrLf & "本所案號： " + LBL1(7) + vbCrLf + "案件名稱： " + LBL1(9).Caption + vbCrLf + "案件性質： " + LBL1(15).Caption + vbCrLf + "承  辦  人： " + LBL1(1).Caption + vbCrLf + "齊  備  日： (未齊備)" + vbCrLf + "本所期限： " + LBL1(17).Caption + vbCrLf + "法定期限： " + LBL1(19).Caption + vbCrLf & vbCrLf & " 齊備日已取消！"
         skMail(UBound(skMail)).fiSubject = "[通知] " & LBL1(7) & "/" & LBL1(3) & LBL1(15) & "/" & LBL1(9) & "/" & m_CuNo & " 齊備日已取消！"
      Else
         skMail(UBound(skMail)).fiContent = vbCrLf + "客戶名稱：" & m_CuNo & vbCrLf & "本所案號： " + LBL1(7) + vbCrLf + "案件名稱： " + LBL1(9).Caption + vbCrLf + "案件性質： " + LBL1(15).Caption + vbCrLf + "承  辦  人： " + LBL1(1).Caption + vbCrLf + "齊  備  日： " + Mid(txt1(2), 1, Len(txt1(2)) - 4) & "年" & Left(Right(txt1(2), 4), 2) & "月" & Right(txt1(2), 2) & "日" + vbCrLf + "本所期限： " + LBL1(17).Caption + vbCrLf + "法定期限： " + LBL1(19).Caption + vbCrLf & vbCrLf & " 齊備日已修改！"
         skMail(UBound(skMail)).fiSubject = "[通知] " & LBL1(7) & "/" & LBL1(3) & LBL1(15) & "/" & LBL1(9) & "/" & m_CuNo & " 齊備日已修改！"
      End If
      'end 2010/9/17
      
      skMail(UBound(skMail)).fiRecriverNo = ""
   End If
      
   'add by nickc 2007/08/22 修改或新輸入CFP 的新申請案的完稿日時，寄mail 通知粘竺儒(84012)
   'Modify By Sindy 2017/4/7 (CFP-29286)經雅娟同意Mark起來不須Mail通知
'   If InStr(1, CaseMapOut & ",301,302,303,304,305,306,307", m_CP10) <> 0 And SystemNumber(lbl1(7).Caption, 1) = "CFP" And txt1(3).Text <> ChangeWStringToTString(lbl1(10).Caption) Then
'      ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
'      skMail(UBound(skMail)).fiSender = strUserNum
'      skMail(UBound(skMail)).fiReceiver = "84012"
'      skMail(UBound(skMail)).fiContent = vbCrLf & "本所案號： " + lbl1(7) + vbCrLf + "案件名稱： " + lbl1(9).Caption + vbCrLf + "申請國家： " + m_NA03 + vbCrLf + "承辦人： " + Combo1.Text + vbCrLf + "智權人員： " + lbl1(21).Caption + vbCrLf & " 完稿日：" & txt1(3).Text & vbCrLf & "英文核稿人：" & Combo4.Text & vbCrLf & vbCrLf & " 完稿日已修改！"
'      '2010/4/13 modify by sonia 總收文號後加案件性質
'      skMail(UBound(skMail)).fiSubject = "[通知] " & lbl1(7) & "/" & lbl1(3) & lbl1(15) & "/" & lbl1(9) & "/" & m_CuNo & " 完稿日已修改！"
'      skMail(UBound(skMail)).fiRecriverNo = ""
'   End If
   'add by nickc 2007/10/09 P 案新申請案上齊備日，則一併上該案的新案翻譯(201)的齊備日，若CaseFee 無紀錄 則，齊備日與承辦期限同
   If LBL1(8) = "" And txt1(2) <> "" And SystemNumber(LBL1(7).Caption, 1) = "P" And InStr(1, CaseMapIn, m_CP10) <> 0 And m_CP10 <> "201" Then
      strExc(0) = "select cp09,cp12 from caseprogress,engineerprogress where cp01='" & SystemNumber(LBL1(7).Caption, 1) & "' and cp02='" & SystemNumber(LBL1(7).Caption, 2) & "' and cp03='" & SystemNumber(LBL1(7).Caption, 3) & "' and cp04='" & SystemNumber(LBL1(7).Caption, 4) & "' and cp10='201' and ep06 is null and ep02(+)=cp09"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         cnnConnection.Execute "update engineerprogress set ep06=" & ChangeTStringToWString(txt1(2)) & " where ep02='" & RsTemp(0) & "'", intI
         If Left(RsTemp(1), 1) = "F" Or PUB_IfSetCP48() Then 'Add by Morgan 2010/10/14
            tmpM_cp48 = Pub_GetHandleDay(SystemNumber(LBL1(7).Caption, 1), m_Country, "201", ChangeTStringToWString(txt1(2)))
            If tmpM_cp48 = "" Then tmpM_cp48 = ChangeTStringToWString(txt1(2))
            cnnConnection.Execute "update caseprogress set cp48=" & tmpM_cp48 & " Where cp09='" & RsTemp(0) & "'", intI
         End If
      End If
   End If
   'add by nickc 2007/10/31 加入 齊備日異動時，紀錄
   If DBDATE(Trim(LBL1(8))) <> DBDATE(Trim(txt1(2))) Then
       Pub_SaveLog strUserNum, "齊備日異動：" & DBDATE(LBL1(8)) & "==>" & DBDATE(txt1(2)) & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
   End If
   
   'add by sindy 2013/10/16 加入 核稿人異動時，紀錄
   If txt1(5).Tag <> txt1(5) Then
       Pub_SaveLog strUserNum, "核稿人異動：" & txt1(5).Tag & "==>" & txt1(5) & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
   End If
   'add by sindy 2013/10/16 加入 判發人異動時，紀錄
   If Combo6.Tag <> Combo6.Text Then
       Pub_SaveLog strUserNum, "判發人異動：" & Combo6.Tag & "==>" & Combo6.Text & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
   End If
   'Add By Sindy 2015/3/4 加入 英文核稿人異動時，紀錄
   If Combo4.Tag <> Combo4.Text Then
      If txt1(23) = "2" Then
         Pub_SaveLog strUserNum, "日文核稿人異動：" & Combo4.Tag & "==>" & Combo4.Text & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
      Else
         Pub_SaveLog strUserNum, "英文核稿人異動：" & Combo4.Tag & "==>" & Combo4.Text & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
      End If
   End If
   
   'add by nick 2005/01/27 P 台灣 設計 輸入齊備日(從無到有)時，若有國外案，更新該齊備日與畫面相同及承辦期限(重算)
   If LBL1(8) = "" And txt1(2) <> "" And SystemNumber(LBL1(7).Caption, 1) = "P" And m_CP10 = "103" And m_Country = "000" Then
      'edit by nickc 2006/03/07 加入未發文未取消收文的才更新
      'edit by nickc 2006/07/24 郭加入新規則，國外案若是大陸的不做
      'edit by nickc 2007/10/05 加入有效控制
      'edit by nickc 2007/10/24 秀玲說只要控制申請案
      'Modify by Morgan 2010/10/14 未齊備條件一併加入
      strSql = "select * from caseprogress,casemap,patent,ENGINEERPROGRESS where cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and cm10='0' " & _
         " and cm05='" & SystemNumber(LBL1(7).Caption, 1) & "' and cm06='" & SystemNumber(LBL1(7).Caption, 2) & "' and cm07='" & SystemNumber(LBL1(7).Caption, 3) & "' " & _
         " and cm08='" & SystemNumber(LBL1(7).Caption, 4) & "' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) And cp27 is null And cp57 is null and pa09<>'020' and cp10 in (" & GetAddStr(CaseMapOut) & ") AND EP02(+)=CP09 AND EP06 IS NULL"
      CheckOC3
      With AdoRecordSet3
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If .RecordCount <> 0 Then
            'Add by Morgan 2009/7/20 需考慮多筆國外案情形
            .MoveFirst
            Do While Not .EOF
            'end  2009/7/20
               tmpM_cp09 = "" & .Fields("cp09").Value
               cnnConnection.Execute "update engineerprogress set ep06=" & ChangeTStringToWString(txt1(2)) & " where ep02='" & tmpM_cp09 & "' and ep06 is null", intI
               '承辦期限
               If intI = 1 Then
                  If (.Fields("cp01") = "P" And Left(.Fields("cp12"), 1) = "F") Or PUB_IfSetCP48() Then 'Add by Morgan 2010/10/14
                     tmpM_cp48 = Pub_GetHandleDay(CheckStr(.Fields("pa01")), CheckStr(.Fields("pa09")), CheckStr(.Fields("cp10")), ChangeTStringToWString(txt1(2)), CheckStr(.Fields("cp06")), CheckStr(.Fields("cp09")))
                     If tmpM_cp48 <> "" Then
                        cnnConnection.Execute "update caseprogress set cp48=" & tmpM_cp48 & " Where cp09='" & tmpM_cp09 & "'", intI
                     End If
                  End If
               End If
               .MoveNext
            Loop
            CheckOC3
         End If
      End With
   End If

   'add by nick 2004/10/14 國外案
   F_CP14 = ""
   F_ST02 = ""
   F_ST03 = ""
   F_CP01020304 = ""
   
'*****注意!!國內外及多國的控制若有修改時，要考慮事後才建關聯的狀況，並檢查分案建多國案的程式是否也要改。*****
   
   'add by nick 2004/07/09 發 mail 給智權人員
   '當會稿完成日與原先不同時
   '93.10.1 MODIFY BY SONIA P 及 CFP 案且要會稿的才發 MAIL
   'edit by nickc 2007/10/04 本來是會稿完成日改成會稿日
   'Modify by Morgan 2008/10/13 原會稿日改放在 Tag(原來存放在label)
   'Modify by Morgan 2008/12/8
   '若CFP案與P案的承辦人 [相同] 則以P案的會稿日為CFP案的齊備日
   '若CFP案與P案的承辦人 [不同] 則以P案的會稿完成日為CFP案的齊備日
   '若CFP案無國內案則以該案的會稿完成日更新其他多國案的齊備日
   'If Trim(txt1(4).Tag) = "" And txt1(4).Text <> "" And (txt1(1) = "" Or UCase(txt1(1)) = "Y") And (SystemNumber(Trim(lbl1(7).Caption), 1) = "P" Or SystemNumber(Trim(lbl1(7).Caption), 1) = "CFP") Then
   '要會稿
   If txt1(1) = "" Or UCase(txt1(1)) = "Y" Then
      'edit by nick 2004/12/02 P 類不發
      '改在 GetF_CP14 作 edit by nick 2004/12/16
      'edit by nick 新案才做
      'edit by nickc 2007/10/04 本來是會稿完成日改成會稿日
      'Modify by Morgan 2008/12/4 上會稿日需判斷國外案是相同承辦人才要更新
      'F_CP14 = GetF_CP14(lbl1(7).Caption)
      
'Modified by Morgan 2013/10/8 會稿完成日改在 UpdateEp08
'      'P案輸入會稿日或會稿完成日
'      If stCP(1) = "P" Then
'
'         '會稿日&會稿完成日同時上時不管承辦人條件
'         If (txt1(4).Tag = "" And txt1(4).Text <> "") And (txt1(7).Tag = "" And txt1(7).Text <> "") Then
'            F_CP14 = GetF_CP14(lbl1(7).Caption)
'         '上會稿日時抓相同承辦人案件
'         ElseIf (txt1(4).Tag = "" And txt1(4).Text <> "") Then
'            F_CP14 = GetF_CP14(lbl1(7).Caption, m_CP14, 1)
'         '上會稿完成日抓不同承辦人案件
'         ElseIf (txt1(7).Tag = "" And txt1(7).Text <> "") Then
'            F_CP14 = GetF_CP14(lbl1(7).Caption, m_CP14, 2)
'         End If
'      'CFP案輸入會稿完成日(應該只會抓到不同承辦人的案件,因為同承辦人的前面已經有做了)
'      ElseIf stCP(1) = "CFP" And txt1(7).Tag = "" And txt1(7).Text <> "" Then
'         '抓多國案
'         F_CP14 = GetF_CP14x(stCP)
'      End If

      'P案輸入會稿日
      'Modified by Morgan 2025/7/17 +判斷案件性質為101、102、103--郭
      If stCP(1) = "P" And (m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103") Then
         If (txt1(4).Tag = "" And txt1(4).Text <> "") Then
            F_CP14 = GetF_CP14(LBL1(7).Caption, m_CP14, 1)
         End If
      End If
'end 2013/10/8
      
      If F_CP14 <> "" Then
         'Remove by Morgan 2008/12/9 移到Transaction外，否則會造成Table鎖定
         's = MsgBox("本案有國外案件(" & F_CP01020304 & ")一併申請，請將資料轉給國外承辦人(" & F_ST02 & ")！", , "警告！")
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
            
         'add by nick 2004/12/16  修正發 mail  每個人都要發
         Dim tmpArr931216_F_ST03 As Variant
         Dim tmpArr931216_F_CP14 As Variant
         Dim tmpArr931216_F_CP01020303 As Variant
         Dim tmpIndexI As Integer
         Dim tmpF_CP0104 As String
         Dim tmpDelayDay As Integer
         Dim tmpP_Rs As New ADODB.Recordset
            
         tmpArr931216_F_ST03 = Split(F_ST03, ",")
         tmpArr931216_F_CP14 = Split(F_CP14, ",")
         tmpArr931216_F_CP01020303 = Split(F_CP01020304, ",")
         'add by nick 2005/01/31 將本會稿完成日上到相關國外的文件齊備日，並重算承辦期限
         'Modify by Morgan 2008/12/9 改判斷相同承辦人時用會稿日，不同時用會稿完成日
         For tmpIndexI = 0 To UBound(tmpArr931216_F_ST03)
            tmpF_CP0104 = tmpArr931216_F_CP01020303(tmpIndexI)
            bolUpdateEP06 = False 'Added by Morgan 2016/5/4
            
'Modified by Morgan 2013/10/8 會稿完成日改在 UpdateEp08
'            If stCP(1) = "P" Then
'               '相同承辦人用會稿日
'               If tmpArr931216_F_CP14(tmpIndexI) = m_CP14 Then
'                  stRefDate = DBDATE(txt1(4))
'                  stRefDateDesc = "會稿日"
'               '不同承辦人用會稿完成日
'               Else
'                  stRefDate = DBDATE(txt1(7))
'                  stRefDateDesc = "會稿完成日"
'               End If
'            Else 'CFP
'               stRefDate = DBDATE(txt1(7))
'               stRefDateDesc = "會稿完成日"
'            End If
            stRefDate = DBDATE(txt1(4))
            stRefDateDesc = "會稿日"
'end 2013/10/8

            'edit by nick 2005/02/01 有些案件性質不用(只抓未齊備的)
            'Modified by Morgan 2016/5/5 +控制CFP只更新主案(判斷要計件者,因日本案可能非主案要計件但也要更新)--柄佑
            'Modified by Morgan 2016/5/17 +改判斷是主案或日本要計件案--柄佑
            strSql = "select cp06,cp09,pa01,pa09,cp10,cp12 from caseprogress,engineerprogress,patent where pa01='" & SystemNumber(tmpF_CP0104, 1) & "' and pa02='" & SystemNumber(tmpF_CP0104, 2) & "' and pa03='" & SystemNumber(tmpF_CP0104, 3) & "' and pa04='" & SystemNumber(tmpF_CP0104, 4) & "' " & _
               " and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+)  and cp10 in (" & GetAddStr(CaseMapOut) & ") and (cp01='P' or (cp01='CFP' and (cp21 is null or (pa09='011' and cp26 is null)))) and cp09=ep02(+) and ep06 is null "
            CheckOC3
            With AdoRecordSet3
               .CursorLocation = adUseClient
               .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If .RecordCount <> 0 Then
                  .MoveFirst
                  Do While Not .EOF
                     bolUpdateEP06 = True 'Added by Morgan 2016/5/4
                     tmpM_cp09 = "" & .Fields("cp09").Value
                     cnnConnection.Execute "update engineerprogress set ep06=" & stRefDate & " where ep02='" & tmpM_cp09 & "' ", intI
                     
                     If (.Fields("pa01") = "P" And Left(.Fields("cp12"), 1) = "F") Or PUB_IfSetCP48() Then 'Add by Morgan 2010/10/14

                        'edit by nickc 2005/09/30 若是相關是 CFP 則看 P 有沒延遲，有要扣 CFP 的天數，且要同一承辦人
                        tmpDelayDay = 0
                        
'Removed by Morgan 2013/10/9 20101026 起已改用專利新規則,只剩FMP(P案)依收費表設定算承辦期限,故下列可取消
'                        'Modify by Morgan 2008/12/8 應該要判斷陣列內的值否則多筆會錯
'                        'If SystemNumber(tmpF_CP0104, 1) = "CFP" And F_CP14 = Trim(Left("" & Combo1.Text, 6)) Then
'                        If stCP(1) = "P" And SystemNumber(tmpF_CP0104, 1) = "CFP" And tmpArr931216_F_CP14(tmpIndexI) = m_CP14 Then
'                           tmpDelayDay = GetWorkDay(ChangeTStringToWString(txt1(3)), ChangeTStringToWString(txt1(2)))
'                           Set tmpP_Rs = New ADODB.Recordset
'                           If tmpP_Rs.State = 1 Then tmpP_Rs.Close
'                           tmpP_Rs.CursorLocation = adUseClient
'                           tmpP_Rs.Open "select * from casefee,caseprogress,patent where cp09='" & lbl1(3) & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cf01 and cf02=pa09 and cf03=cp10 and cp14='" & Trim(Left("" & Combo1.Text, 6)) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
'                           If tmpP_Rs.RecordCount <> 0 Then
'                              If tmpDelayDay > CheckStr(tmpP_Rs.Fields("cf04").Value) Then
'                                 tmpDelayDay = tmpDelayDay - Val(CheckStr(tmpP_Rs.Fields("cf04").Value))
'                              Else
'                                 tmpDelayDay = 0
'                              End If
'                           Else
'                              tmpDelayDay = 0
'                           End If
'                        End If
'end 2013/10/9
                        'edit by nickc 2005/09/30 加入計算延遲
                        'edit by nickc 改抓有效承辦天數並將會稿完成日改成會稿日
                        'edit by nickc 2007/10/05 本來是會稿完成日更新到國外齊備日，現改成輸會稿日
                        'Modify by Morgan 2008/12/9
                        'tmpM_cp48 = Pub_GetHandleDay(CheckStr(.Fields("pa01")), CheckStr(.Fields("pa09")), CheckStr(.Fields("cp10")), ChangeTStringToWString(txt1(4)), CheckStr(.Fields("cp06")), CheckStr(.Fields("cp09")), tmpDelayDay)
                        tmpM_cp48 = Pub_GetHandleDay(CheckStr(.Fields("pa01")), CheckStr(.Fields("pa09")), CheckStr(.Fields("cp10")), stRefDate, CheckStr(.Fields("cp06")), CheckStr(.Fields("cp09")), tmpDelayDay)
                        If tmpM_cp48 <> "" Then
                           cnnConnection.Execute "update caseprogress set cp48=" & tmpM_cp48 & " Where cp09='" & tmpM_cp09 & "'  "
                        End If
                        'end 2008/12/9
                        
                     End If 'Add by Morgan 2010/10/14
                     .MoveNext
                  Loop
               End If
            End With
            CheckOC3
                    
            
            If bolUpdateEP06 Then 'Added by Morgan 2016/5/4 有更新齊備日才通知
            
               'edit by nick 2004/10/26 外翻 改發 79075
               'edit by nickc 2006/10/12 F部門的檢查改到 pub_sendmail
               'edit by nickc 2006/12/29 改在 trans 後發
               ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
               skMail(UBound(skMail)).fiSender = strUserNum
               skMail(UBound(skMail)).fiReceiver = tmpArr931216_F_CP14(tmpIndexI)
               'Modify by Morgan 2007/9/13 內容至少要有空白否則不會發信
               skMail(UBound(skMail)).fiContent = " "
               'Modify by Morgan 2008/12/9
               'skMail(UBound(skMail)).fiSubject = lbl1(7) & "已輸會稿日！(相關國外案：" + tmpArr931216_F_CP01020303(tmpIndexI) + ")"
               skMail(UBound(skMail)).fiSubject = LBL1(7) & "已輸" & stRefDateDesc & "！(相關國外案：" + tmpArr931216_F_CP01020303(tmpIndexI) + ")"
               skMail(UBound(skMail)).fiRecriverNo = ""
            
            End If 'Added by Morgan 2016/5/4 有更新齊備日才通知
            
         Next tmpIndexI
        
         'edit by nick 2005/01/17 不會稿不發
         '2008/9/15 modify by sonia 會稿完成通知移至會稿日檢查之外
         'edit by nickc 2006/12/29 改在 trans 後才發信
         Me.Enabled = True
         'Modify by Morgan 2009/11/12
         'Screen.MousePointer = vbDefault
         Screen.MousePointer = iMouse
      End If
   End If
                   
'Removed by Morgan 2013/10/8 改在 UpdateEp08
'   '2008/9/15 modify by sonia 自會稿日檢查內移出
'   If txt1(7).Text <> "" And _
'      txt1(7).Text <> txt1(7).Tag And _
'      (txt1(1) = "" Or UCase(txt1(1)) = "Y") And _
'      (SystemNumber(Trim(lbl1(7).Caption), 1) = "P" Or SystemNumber(Trim(lbl1(7).Caption), 1) = "CFP") Then
'      If GetStaffDepartment(m_CP14) <> "P12" And _
'         ((m_CP10 >= "101" And m_CP10 <= "107") Or (m_CP10 >= "201" And m_CP10 <= "206") Or (m_CP10 >= "501" And m_CP10 <= "504") Or (m_CP10 >= "801" And m_CP10 <= "804")) And _
'         txt1(1).Text <> "N" Then
'         Screen.MousePointer = vbHourglass
'         Me.Enabled = False
'         ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
'         skMail(UBound(skMail)).fiSender = strUserNum
'         skMail(UBound(skMail)).fiReceiver = m_CP13
'         skMail(UBound(skMail)).fiContent = vbCrLf + "客戶名稱：" & m_CuNo & vbCrLf & "本所案號： " + lbl1(7) + vbCrLf + "案件名稱： " + lbl1(9).Caption + vbCrLf + "案件性質： " + lbl1(15).Caption + vbCrLf + "本所期限： " + lbl1(17).Caption + vbCrLf + "法定期限： " + lbl1(19).Caption + vbCrLf + "會稿完成日：" + ChangeTStringToTDateString(txt1(7).Text) + vbCrLf + vbCrLf + "已完成會稿！"
'         skMail(UBound(skMail)).fiSubject = lbl1(7) & "已會稿完成！"
'         skMail(UBound(skMail)).fiRecriverNo = ""
'         Me.Enabled = True
'         'Modify by Morgan 2009/11/12
'         'Screen.MousePointer = vbDefault
'         Screen.MousePointer = iMouse
'      End If
'   End If
'   '2008/9/15 end
'end 2013/10/8
                                     
   'add by nickc 2006/02/27
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   strSql = " UPDATE CASEPROGRESS SET "
   bFirst = True
   bDifference = False
      
   For nIndex = 0 To TF_CP - 1
      strTmp = Empty
      If nIndex < 64 Or nIndex > 69 Then
         If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
            If m_FieldList(nIndex).fiType = 0 Then
               If m_FieldList(nIndex).fiNewData = Empty Then
                  strTmp = m_FieldList(nIndex).fiName & " = NULL "
               Else
                  strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
               End If
            Else
               If m_FieldList(nIndex).fiNewData = Empty Then
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
                  
   strSql = strSql & " " & _
      "WHERE CP09 = '" & Me.LBL1(3).Caption & "' "
     
   If bDifference = True Then
      Pub_SeekTbLog strSql 'Added by Morgan 2016/7/7
      cnnConnection.Execute strSql, intI
   End If
   
   'Add by Morgan 2009/10/5
   '台灣專利申請案上會稿日時,若該案同時有辦大陸案,則請SHOW訊息告知台灣案之工程師
   '台灣專利申請案上會稿完成日時,發E-MAIL通知大陸案之工程師
   st020Msg = ""
   'Modify by Morgan 2010/9/23 新案翻譯201要排除有收文其他新案案件性質
   'If ((txt1(4).Tag = "" And txt1(4).Text <> "") Or (txt1(7).Tag = "" And txt1(7).Text <> "")) And m_country = "000" And InStr(CaseMapIn, m_CP10) > 0 Then
   
'Modified by Morgan 2013/10/8 會稿完成日更新改在 UpdateEp08
'   If ((txt1(4).Tag = "" And txt1(4).Text <> "") Or (txt1(7).Tag = "" And txt1(7).Text <> "")) And _
'      m_country = "000" And InStr(CaseMapIn, m_CP10) > 0 And bol201NoSkip Then
   If (txt1(4).Tag = "" And txt1(4).Text <> "") And _
      m_Country = "000" And InStr(CaseMapIn, m_CP10) > 0 And bol201NoSkip Then
      '2011/6/10 modify by sonia 郭說大陸案承辦人為程序的案件才要發mail通知大陸案繪圖人員做 pdf 檔,故加承辦人部門欄
      strSql = "select cm01||'-'||cm02||decode(cm03||cm04,'000','','-'||cm03||'-'||cm04) C1,cp14,cp09,cp29,ep20,ep29,st03" & _
         " from casemap,patent,caseprogress,engineerprogress,staff where cm05='" & stCP(1) & "' and cm06='" & stCP(2) & "'" & _
         " and cm07='" & stCP(3) & "' and cm08='" & stCP(4) & "' and cm01='P' and cm10='0'" & _
         " and pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa09='020' and pa57 is null" & _
         " and cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04 and cp10 in (" & GetAddStr(CaseMapOut) & ") and ep02(+)=cp09" & _
         " and cp14=st01(+)"
      CheckOC3
      intI = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If (txt1(4).Tag = "" And txt1(4).Text <> "") Then
            st020Msg = "本案同時辦理大陸案 " & AdoRecordSet3(0) & ",若申請專利範圍有超出10項或說明書(含圖式)有超出30頁者,請一併告知智權同仁加收費用。"
         End If
         
'Removed by Morgan 2013/10/8 改在 UpdateEp08
'         If (txt1(7).Tag = "" And txt1(7).Text <> "") Then
'            strExc(1) = stCP(1) & "-" & stCP(2) & IIf(stCP(3) & stCP(4) = "000", "", "-" & stCP(3) & "-" & stCP(4))
'
'            ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
'            skMail(UBound(skMail)).fiSender = strUserNum
'            skMail(UBound(skMail)).fiReceiver = AdoRecordSet3("cp14")
'            skMail(UBound(skMail)).fiContent = vbCrLf + AdoRecordSet3(0) & " 案之台灣案 " & strExc(1) & " 已會稿完成,可先行寄委托書給大陸代理人。！"
'            skMail(UBound(skMail)).fiSubject = AdoRecordSet3(0) & " 案之台灣案已會稿完成！"
'            skMail(UBound(skMail)).fiRecriverNo = ""
'
'            'Add by Morgan 2010/5/11
'            '台灣案上會稿完成日要通知大陸案繪圖人員做 pdf 檔
'            '2011/6/10 modify by sonia 郭說大陸案承辦人為程序P12的案件才要發,工程師案件由工程師自行通知
'            'If Not IsNull(AdoRecordSet3("cp29")) Then
'            'Remove by Morgan 2011/6/14 都改在國內案發文時通知--瓊玉
'            'If Not IsNull(AdoRecordSet3("cp29")) And "" & AdoRecordSet3("st03") = "P12" Then
'            '   strSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'            '      " select '" & strUserNum & "',cp29,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            '      ",'台灣案 " & strExc(1) & " 已會稿完成,請製作大陸案 " & AdoRecordSet3(0) & " PDF 檔圖式,並以 E-MAIL 交承辦人 '||st02||'('||cp14||')。','如旨'" & _
'            '      " FROM CASEPROGRESS,STAFF WHERE CP09='" & AdoRecordSet3("cp09") & "' and st01(+)=cp14"
'            '   cnnConnection.Execute strSQL, intI
'            'End If
'            'end 2011/6/14
'
'            'Remove by Morgan 2011/6/27 都改在國內案發文時上繪圖齊備(與製作pdf通知同步)--瓊玉
'            'If AdoRecordSet3("ep20") = "N" And AdoRecordSet3("ep29") = "N" Then
'            '   strSql = "update caseprogress set cp101=0.2,cp104=0.2,cp100=0.6,cp103=0.6,cp107=decode(cp29,null,null,'Y') where cp09='" & AdoRecordSet3("cp09") & "'"
'            '   cnnConnection.Execute strSql, intI
'            '   'Modify by Morgan 2010/6/30 草齊墨齊改上系統日(因為可能工程師還沒有輸文齊日)
'            '   'strSql = "update engineerprogress set ep20=null,ep29=null,ep16=0,ep19=0,ep14=ep06,ep17=ep06 where ep02='" & AdoRecordSet3("cp09") & "'"
'            '   strSql = "update engineerprogress set ep20=null,ep29=null,ep16=0,ep19=0,ep14=nvl(ep14," & strSrvDate(1) & "),ep17=nvl(ep17," & strSrvDate(1) & ") where ep02='" & AdoRecordSet3("cp09") & "'"
'            '   cnnConnection.Execute strSql, intI
'            ''Add by Morgan 2011/4/12
'            ''新規則改要計件
'            'Else
'            '   strSql = "update engineerprogress set ep14=nvl(ep14,decode(ep20,null,to_char(sysdate,'yyyymmdd')))" & _
'            '      ",ep17=nvl(ep17,decode(ep29,null,to_char(sysdate,'yyyymmdd')))" & _
'            '      " where ep02='" & AdoRecordSet3("cp09") & "'"
'            '   cnnConnection.Execute strSql, intI
'            '
'            'End If
'            'end 2010/5/11
'            'end 2011/6/27
'         End If
'end 2013/10/8

      End If
   End If
   'end 2009/10/5
   
   
   'Removed by Morgan 2013/10/2 取消,承辦單已電子化--瓊玉
   ''Add by Morgan 2011/6/14
   ''單獨收文之大陸案於會稿完成時通知繪圖人員製作pdf檔--瓊玉
   ''Modified by Morgan 2011/10/31 +PCT--瓊玉
   'If (txt1(7).Tag = "" And txt1(7).Text <> "") And (m_country = "020" Or m_country = "056") And InStr(CaseMapIn, m_CP10) > 0 And bol201NoSkip Then
   '   strExc(0) = "select * from casemap where cm01='" & stCP(1) & "' and cm02='" & stCP(2) & "'" & _
   '      " and cm03='" & stCP(3) & "' and cm04='" & stCP(4) & "' and cm10='0'"
   '   intI = 1
   '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '   If intI = 0 Then
   '      strExc(1) = stCP(1) & "-" & stCP(2) & IIf(stCP(3) & stCP(4) = "000", "", "-" & stCP(3) & "-" & stCP(4))
   '      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
   '               " select '" & strUserNum & "',cp29,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
   '               ",'" & IIf(m_country = "056", "PCT", "大陸") & "案 " & strExc(1) & " 已會稿完成,請製作 PDF 檔圖式,並以 E-MAIL 交承辦人 '||st02||'('||cp14||')。','如旨'" & _
   '               " FROM CASEPROGRESS,STAFF WHERE CP09='" & lbl1(3).Caption & "' and cp29 is not null and st01(+)=cp14"
   '      cnnConnection.Execute strSql, intI
   '   End If
   'End If
   ''end 2011/6/14
   'end 2013/10/2
   
   'Added by Morgan 2012/9/6
   '若有勾選他所寄存則內部收文寄存證明
   If txt1(21).Tag <> txt1(21) And txt1(21) = "Y" Then
      strExc(2) = ""
      strExc(3) = PUB_GetFirstPriDate(stCP)
      '法限=最早優先權日起16個月
      If strExc(3) <> "" Then
         strExc(3) = CompDate(1, 16, strExc(3))
         stDate(0) = ""
         stDate(1) = stCP(1)
         stDate(2) = m_Country
         stDate(3) = strExc(3)
         GetCtrlDT stDate
         '本所期限
         strExc(2) = PUB_GetWorkDay1(stDate(0), True)
      End If

      strExc(1) = AutoNo("B", 6)
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07" & _
         ",CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43) " & _
         " select cp01,cp02,cp03,cp04,to_char(sysdate,'yyyymmdd')," & CNULL(strExc(2), True) & "," & CNULL(strExc(3), True) & _
         ",'" & strExc(1) & "','231','90',cp12,cp13,cp14,'N','N','N',cp09" & _
         " from caseprogress where cp09='" & LBL1(3).Caption & "'"
      cnnConnection.Execute strSql, intI
   End If
   'end 2012/9/6
   
   'Add By Sindy 2015/3/4 主管取消英/日文核稿人時,發E-Mail給原核稿人,副本71011.王副總,同時產生一道聯絡的歷程
   'If ProState = "2" And bolHadSetProofEngReader = True And m_PER04 <> m_CP14 And Combo4.Text = "" Then
   If Combo4.Tag <> Combo4.Text And Combo4.Text = "" Then
      'Modify By Sindy 2015/7/22 有”可取消英核之人員”的權限,且無須英核時,可在個人工作管理拿掉英文核稿人員
      '                          一樣要發mail要記錄
      If ProState = "2" Or _
         (ProState = "1" And _
          InStr(Pub_GetSpecMan("可取消英核之人員"), strUserNum) > 0 And _
          m_CP14 = strUserNum And _
          (m_PER04 = "" Or m_PER04 = m_CP14)) Then
      '2015/7/22 END
         '主旨
         strSubject = Replace(LBL1(7).Caption, "-0-00", "") & "之" & LBL1(15).Caption & "取消" & IIf(txt1(23) = "2", "日文", "英文") & "核稿!"
         strContent = "本所案號：" & LBL1(7).Caption & vbCrLf & _
                      "案件性質：" & LBL1(15).Caption & vbCrLf & _
                      "承 辦 人：" & LBL1(1).Caption & vbCrLf & _
                      "總收文號：" & LBL1(3).Caption & vbCrLf & _
                      "案件名稱：" & LBL1(9).Caption & vbCrLf & _
                      "申 請 人：" & GetPrjPeople1(GetPrjPeopleNum1(LBL1(7).Caption))
         '發E-Mail給原核稿人,副本71011.王副總
         ReDim Preserve skMail(UBound(skMail) + 1) As SeekMails
         skMail(UBound(skMail)).fiSender = strUserNum
         'Modified by Lydia 2023/04/24 修改王副總退休之相關控制
         'skMail(UBound(skMail)).fiReceiver = Left(Trim(Combo4.Tag), 5) & ";71011"
         If strSrvDate(1) >= "20230511" Then
             strExc(1) = "99050"
         ElseIf strSrvDate(1) >= "20230501" Then
             strExc(1) = "71011;99050"
         Else
             strExc(1) = "71011"
         End If
         skMail(UBound(skMail)).fiReceiver = Left(Trim(Combo4.Tag), 5) & ";" & strExc(1)
         'end 2023/04/24
         skMail(UBound(skMail)).fiReceiver = Left(Trim(Combo4.Tag), 5) & ";" & Pub_GetSpecMan("專利處工程師通知主管_A")
         skMail(UBound(skMail)).fiContent = strContent
         skMail(UBound(skMail)).fiSubject = strSubject
         skMail(UBound(skMail)).fiRecriverNo = Me.LBL1(3).Caption
         '內容
         '取得最大序號
         intMaxEEP02 = 0
         strSql = "select eep02 From empelectronprocess where eep01='" & LBL1(3).Caption & "' order by eep02 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            RsTemp.MoveFirst
            If RsTemp.RecordCount > 0 Then
               intMaxEEP02 = RsTemp.Fields(0)
            End If
         End If
         '新增歷程
         'Modified by Lydia 2023/04/24 修改王副總退休之相關控制;
         'strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep10) values(" & _
                  CNULL(lbl1(3).Caption) & "," & intMaxEEP02 + 1 & "," & CNULL(strUserNum) & "," & _
                  CNULL(EMP_聯絡) & "," & CNULL(Left(Trim(Combo4.Tag), 5)) & "," & strSrvDate(1) & "," & _
                  Right("000000" & ServerTime, 6) & "," & CNULL(strSubject) & "," & CNULL("71011") & ")"
         strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep10) values(" & _
                  CNULL(LBL1(3).Caption) & "," & intMaxEEP02 + 1 & "," & CNULL(strUserNum) & "," & _
                  CNULL(EMP_聯絡) & "," & CNULL(Left(Trim(Combo4.Tag), 5)) & "," & strSrvDate(1) & "," & _
                  Right("000000" & ServerTime, 6) & "," & CNULL(strSubject) & "," & CNULL(strExc(1)) & ")"
         cnnConnection.Execute strSql
         'Add By Sindy 2016/3/4 系統一併檢查是否有正在送英核的流程，若有，一併刪除該道歷程及附件檔
         '讀取送英核中的歷程
         strSql = "select eep01,eep02 From empelectronprocess where eep01='" & LBL1(3).Caption & "' and eep04='" & EMP_送英核 & "' and eep09='Y'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            '刪除送英核中的歷程
            strSql = "delete from empelectronprocess where eep01='" & RsTemp.Fields("eep01") & "' and eep02=" & RsTemp.Fields("eep02")
            cnnConnection.Execute strSql
            strSql = "delete from empelectronfile where eef01='" & RsTemp.Fields("eep01") & "' and eef02=" & RsTemp.Fields("eep02")
            cnnConnection.Execute strSql
         End If
         '2016/3/4 END
      End If
   End If
   '2015/3/4 END
   
   UpdateEp08 LBL1(3).Caption, txt1(7) 'Added by Morgan 2013/10/8
   
   cnnConnection.Execute "begin user_data.user_formname:=Null;end;" 'Add by Morgan 2010/10/29
   
   'Modify By Sindy 2018/6/20 + If m_Flow = "" Then
   If m_Flow = "" Then cnnConnection.CommitTrans
   
'Removed by Morgan 2013/10/8 電子承辦單送判時會自動新增[連絡]到 CFP 並同時複製檔案過去,此處不用在提醒
'   'Add by Morgan 2008/12/9 從Transaction內移出來，否則會造成Table鎖定
'   'Modify by Morgan 2010/10/4 P案不必提醒
'   'If F_CP14 <> "" Then
'   '   s = MsgBox("本案有國外案件(" & F_CP01020304 & ")一併申請，請將資料轉給國外承辦人(" & F_ST02 & ")！", , "警告！")
'   'End If
'   arr1 = Split(F_CP01020304, ",")
'   arr2 = Split(F_ST02, ",")
'   strExc(1) = ""
'   strExc(2) = ""
'   For intI = LBound(arr1) To UBound(arr1)
'      If Left(arr1(intI), 2) <> "P-" Then
'         strExc(1) = strExc(1) & "," & arr1(intI)
'         strExc(2) = strExc(2) & "," & arr2(intI)
'      End If
'   Next
'   If strExc(1) <> "" Then
'      strExc(1) = Mid(strExc(1), 2)
'      strExc(2) = Mid(strExc(2), 2)
'      s = MsgBox("本案有國外案件(" & strExc(1) & ")一併申請，請將資料轉給國外承辦人(" & strExc(2) & ")！", , "警告！")
'   End If
'   'end 2010/10/4
'
'   'Add by Morgan 2009/10/5
'   If st020Msg <> "" Then s = MsgBox(st020Msg, , "警告！")
'end 2013/10/8

   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.Execute "begin user_data.user_formname:=Null;end;" 'Add by Morgan 2010/10/29
   'Modify By Sindy 2018/6/20 + If m_Flow = "" Then
   If m_Flow = "" Then cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
         
End Function

'Removed by Morgan 2021/11/9 2010/10/14 已取消
''Add by Morgan 2008/12/5 檢查若[適不適用會稿加乘註記]有變更但存檔後值與原畫面選項不同時提醒
'Private Sub CP121AlterCheck()
'   'Add by Morgan 2008/12/3 若存檔之會稿加乘註記設定與畫面不同時提醒
'   If m_bAlterCP112 = True Then
'      strExc(0) = "select cp112 from caseprogress where cp09='" & Me.lbl1(3).Caption & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         strExc(1) = ""
'         strExc(2) = ""
'         If m_stNewCP112 <> "" & RsTemp.Fields(0) Then
'            Select Case "" & RsTemp.Fields(0)
'               Case "Y"
'                  strExc(2) = "適用"
'               Case "N"
'                  strExc(2) = "不適用"
'               Case Else
'                  strExc(2) = "未設定"
'            End Select
'            MsgBox "系統已依規則將此案設定為【" & strExc(2) & "】!!", vbExclamation, "適不適用會稿加乘註記更新提醒"
'         End If
'      End If
'   End If
'End Sub
'end 2021/11/9

'Add by Morgan 2008/12/5 自cmdok_Click抽出
'批次發Mail
'Modify By Sindy 2018/6/20
'Private Sub BatctMail()
Public Sub BatctMail()
'2018/6/20 END
   Dim i As Integer
   For i = 1 To UBound(skMail)
        PUB_SendMail skMail(i).fiSender, skMail(i).fiReceiver, skMail(i).fiRecriverNo, skMail(i).fiSubject, skMail(i).fiContent
   Next i
   ReDim skMail(0) As SeekMails
   'Add by Morgan 2009/11/4
   'Trigger 也會產生待發郵件
   PUB_SendMailCache
   'end 2009/11/4
End Sub

'Add by Morgan 2008/12/5 自cmdok_Click抽出
'更新mdb暫存資料及第一畫面的Grid內容
'Modify By Sindy 2013/6/10
'Private Sub UpdEngMdb()
Public Sub UpdEngMdb()
'2013/6/10 End
On Error GoTo ErrHnd
   
   'Modify by Morgan 2010/10/7 承辦期限改抓txt1(12)<--lbl1(2)
   'Modify by Morgan 2011/1/4 修正日期欄位排序問題(抓9碼不足前面補空白)
   strSql = "UPDATE R090614 SET " & _
      "R110013='" & IIf(txt1(2) = "", "", Right(" " & ChangeTStringToTDateString(txt1(2)), 9)) & "'," & _
      "R110014='" & IIf(txt1(3) = "", "", Right(" " & ChangeTStringToTDateString(txt1(3)), 9)) & "'," & _
      "R110015='" & IIf(txt1(4) = "", "", Right(" " & ChangeTStringToTDateString(txt1(4)), 9)) & "'," & _
      "R110017='" & IIf(txt1(7) = "", "", Right(" " & ChangeTStringToTDateString(txt1(7)), 9)) & "'," & _
      "R110016='" & IIf(txt1(5) = "", "", LBL1(14).Caption) & "'," & _
      "R110010='" & IIf(txt1(12) = "", "", Right(" " & ChangeTStringToTDateString(txt1(12)), 9)) & "'," & _
      "R110018='" & IIf(txt1(8) = "", "", Right(" " & ChangeTStringToTDateString(txt1(8)), 9)) & "'," & _
      "R110019=" & Intnick910123 & "," & _
      "R110020='" & txtEP12 & "' " & _
       " WHERE ID='" & strUserNum & "' AND R110022='" & LBL1(3).Caption & "' "
   adoEng.Execute strSql, intI
   
   m_blnClkSure = True
   For i = 1 To GRD1.Rows - 1
      GRD1.row = i
      GRD1.col = 0
      '若目次相同, 收文號也相同
      'Modify by Morgan 2009/7/14 加欄位:預會日 15
      'If grd1.Text = SeekTmpBk And Me.grd1.TextMatrix(i, 22) = m_strCP09 Then
      'Modified by Lydia 2025/02/05 改用變數
      'If grd1.Text = SeekTmpBk And Me.grd1.TextMatrix(i, 23) = m_strCP09 Then
      If GRD1.Text = SeekTmpBk And Me.GRD1.TextMatrix(i, colCP09_1) = m_strCP09 Then
         MouseClick_1 (i)
         StrMenuOneRec SWPRow2
         Exit For
      End If
   Next i
   m_blnClkSure = False
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
   'Resume
End Sub
'Add by Morgan 2008/12/8 從Form_Load抽出
'設定承辦人選單
Private Sub SetEngineer()
   strSql = "SELECT Distinct (R110001&' '&'(' & R110025&')') FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Trim(strUserNum) & "' ORDER BY (R110001&' '&'(' & R110025&')') "

   CheckOC
   i = 0
   Combo1.Clear
   Combo1_String = "" '92.6.26 ADD BY SONIA
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, adoEng, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
         Do While .EOF = False
           Combo1.AddItem "" & .Fields(0), i
           i = i + 1
           '92.6.26 ADD BY SONIA
           If Combo1_String = "" Then
              Combo1_String = "'" & Trim(Left("" & .Fields(0), 6)) & "'"
           Else
              Combo1_String = Combo1_String + ",'" & Trim(Left("" & .Fields(0), 6)) & "'"
           End If
           '92.6.26 END
           .MoveNext
         Loop
         Combo1.Text = Combo1.List(0)
       End If
   End With
End Sub

'Add by Morgan 2008/12/8 從Form_Load抽出並簡化
'設定繪圖人員選單
Private Sub SetDrawer()
   Combo2.Clear
   Combo2.AddItem "", 0
   strExc(0) = "select st01||' ==> '||st02 from staff where st04='1' and st05 in ('79','81','82','AC') order by Decode(ST01,'99999','00000',ST01) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      .MoveFirst
      Do While .EOF = False
          Combo2.AddItem "" & .Fields(0), intI
          intI = intI + 1
          .MoveNext
      Loop
      Combo2.Text = Combo2.List(0)
      End With
   End If
End Sub

'Add by Morgan 2008/12/8 從Form_Load抽出並簡化
'設定英文核稿人選單
Private Sub SetEngChecker()
   Combo4.Clear
   Combo4.AddItem "", 0
   'add by nickc 2007/08/21  加入英文核稿人
   '2010/8/19 MODIFY BY SONIA 改抓專利處英文顧問
   '*****改此部門條件要改四個畫面frm090201_2,frm090218,frm090218_1,frm100101_F
   'strExc(0) = "select st01||' ==> '||st02 from staff where st04='1' and st03='F62' order by Decode(ST01,'99998','00000',ST01) "
   'Modified by Morgan 2013/10/8 +日文顧問 F71,王文安 88003
   'Modify By Sindy 2015/3/13 依核稿語文帶核稿主管
   If txt1(23) = "2" Then '日核
      'Added by Morgan 2022/11/24 88003退休不必再列出
      If strSrvDate(1) >= 20221130 Then
         strExc(0) = "select st01||' ==> '||st02 from staff where st04='1' and st03='F71' order by st03 desc, Decode(ST01,'99998','00000',ST01)"
      Else
      'end 2022/11/24
         strExc(0) = "select st01||' ==> '||st02 from staff where st04='1' and (st01='88003' or st03='F71') order by st03 desc, Decode(ST01,'99998','00000',ST01)"
      End If 'Added by Morgan 2022/11/24
   Else '英核
      strExc(0) = "select st01||' ==> '||st02 from staff where st04='1' and st03='P14' order by st03 desc, Decode(ST01,'99998','00000',ST01)"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      .MoveFirst
      Do While .EOF = False
          Combo4.AddItem "" & .Fields(0), intI
          intI = intI + 1
          .MoveNext
      Loop
      End With
   End If
   Combo4.Text = Combo4.List(0)
End Sub

'Add by Morgan 2010/7/20
'檢查非日本德國案的主案若有完稿日,預定會稿日,會稿日時,是否有子案已齊備但無相關日期
Private Sub ChkRefCaseDate()
   Dim cp(1 To 4) As String '本所案號
   Dim stCon As String, stMsg As String
   cp(1) = SystemNumber(Trim(LBL1(7).Caption), 1)
   m_strNeedUpdateCase = ""
   If cp(1) = "CFP" And m_Country <> "011" And m_Country <> "231" And txt1(3) & txt1(18) & txt1(4) <> "" Then
      strExc(0) = "select cp02,cp03,cp04 from caseprogress where cp09='" & LBL1(3) & "' and cp21 is null and cp10 in (" & NewCasePtyList & ")"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      '主案
      If intI = 1 Then
         cp(2) = RsTemp("cp02")
         cp(3) = RsTemp("cp03")
         cp(4) = RsTemp("cp04")
         stCon = "": stMsg = ""
         If txt1(3) <> "" Then
            stCon = stCon & IIf(stCon <> "", " or ", "") & "ep09 is null"
            stMsg = stMsg & IIf(stMsg <> "", " 、 ", "") & "完稿日"
         End If
         If txt1(18) <> "" Then
            stCon = stCon & IIf(stCon <> "", " or ", "") & "ep28 is null"
            stMsg = stMsg & IIf(stMsg <> "", " 、 ", "") & "預定會稿日"
         End If
         If txt1(4) <> "" Then
            stCon = stCon & IIf(stCon <> "", " or ", "") & "ep07 is null"
            stMsg = stMsg & IIf(stMsg <> "", " 、 ", "") & "會稿日"
         End If
         stCon = " and (" & stCon & ")"
         
         strExc(0) = "select ''''||ep02||'''' from caserelation,caseprogress,engineerprogress" & _
            " where cr01='" & cp(1) & "' and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "'" & _
            " and cp01(+)=cr05 and cp02(+)=cr06 and cp03(+)=cr07 and cp04(+)=cr08 and cp21='Y' and cp14='" & m_CP14 & "'" & _
            " and cp10 in (" & NewCasePtyList & ") and ep02(+)=cp09 and ep06>0" & stCon
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If MsgBox("本案已輸入" & stMsg & "，是否在關聯副案的相同事項輸入相同日期？", vbYesNo + vbDefaultButton2) = vbYes Then
               RsTemp.MoveFirst
               m_strNeedUpdateCase = RsTemp.GetString(, , , ",")
               m_strNeedUpdateCase = Left(m_strNeedUpdateCase, Len(m_strNeedUpdateCase) - 1)
            End If
         End If
      End If
   End If
End Sub

'Removed by Morgan 2021/11/9 108考核已取消會稿加乘
''Add by Morgan 2010/9/16
''計算會稿加乘(未會稿預估)
'Private Function StrMenu20(idx As Integer) As String
'   Dim ThisSvrDate As String        '截止日
'   Dim stCP01 As String, stEP04 As String, stEP05 As String, stEP06 As String, stCP111 As String
'   Dim stDate1 As String, stDate2 As String
'
'   StrMenu20 = grd1.TextMatrix(idx, 11)
'
'   If grd1.TextMatrix(idx, 27) <> "Y" Then Exit Function
'
'   'Modified by Morgan 2012/4/16 預估會稿日改用當日 --柄佑
'   Select Case ProState
'   Case "1" '承辦人個人工作進度資料維護
'            'ThisSvrDate = CompWorkDay(2, strSrvDate(1), 1)
'            ThisSvrDate = strSrvDate(1)
'   Case "2" '承辦人管理工作進度資料查詢
'            'ThisSvrDate = IIf(Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) = Mid(strSrvDate(1), 1, 6), CompWorkDay(2, strSrvDate(1), 1), ChangeWDateStringToWString(DateAdd("d", -1, DateAdd("m", 1, ChangeWStringToWDateString(Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01")))))
'            ThisSvrDate = IIf(Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) = Mid(strSrvDate(1), 1, 6), strSrvDate(1), ChangeWDateStringToWString(DateAdd("d", -1, DateAdd("m", 1, ChangeWStringToWDateString(Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01")))))
'   Case Else
'   End Select
'
'   stEP06 = DBDATE(grd1.TextMatrix(idx, 13))
'   stCP01 = SystemNumber(grd1.TextMatrix(idx, 3), 1)
'   stEP04 = grd1.TextMatrix(idx, 17) '核稿人
'   stEP05 = grd1.TextMatrix(idx, 26) '承辦人
'
'   If stCP01 = "P" Then
'      stDate1 = CompWorkDay(7, stEP06)
'      stDate2 = CompWorkDay(10, stEP06)
'   Else
'      stDate1 = CompWorkDay(14, stEP06)
'      stDate2 = CompWorkDay(20, stEP06)
'   End If
'   stCP111 = 1
'   If Val(ThisSvrDate) <= stDate1 Then
'      If stEP04 = "" Or stEP04 = stEP05 Then
'         stCP111 = 1.2
'      Else
'         stCP111 = 1.4
'      End If
'   ElseIf Val(ThisSvrDate) <= stDate2 Then
'      stCP111 = 1.2
'   End If
'   'Modify by Morgan 2011/8/3 要先減(支援+修改+衍生)基數後再乘,然後再加回來
'   'StrMenu20 = Format(Val(grd1.TextMatrix(idx, 11)) * Val(stCP111), "0.00")
'   StrMenu20 = Format((Val(grd1.TextMatrix(idx, 11)) - Val(grd1.TextMatrix(idx, 28))) * Val(stCP111) + Val(grd1.TextMatrix(idx, 28)), "0.00")
'
'End Function
'end 2021/11/9

'Added by Morgan 2013/8/1
'檢查是否有關聯案3日內收文主動修正(203)或修正(204)
Private Function Chk203Case(pa01 As String, pa02 As String, pa03 As String, pa04 As String) As Boolean
   Dim stSQL As String, stVTB As String, intR As Integer, stDate As String
   Dim adoRst As ADODB.Recordset
   
   stDate = CompWorkDay(3, strSrvDate(1), 1)
   
   stVTB = "select cm01,cm02,cm03,cm04 from casemap where cm10='0' and cm05='" & pa01 & "' and cm06='" & pa02 & "' and cm07='" & pa03 & "' and cm08='" & pa04 & "'" & _
      " union select cm05,cm06,cm07,cm08 from casemap where cm10='0' and cm01='" & pa01 & "' and cm02='" & pa02 & "' and cm03='" & pa03 & "' and cm04='" & pa04 & "'" & _
      " union select cr01,cr02,cr03,cr04 from caserelation where cr05='" & pa01 & "' and cr06='" & pa02 & "' and cr07='" & pa03 & "' and cr08='" & pa04 & "'"
   stSQL = "select cp09 from (" & stVTB & "),caseprogress where cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04 and (cp10='203' or cp10='204') and cp05>=" & stDate
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      Chk203Case = True
   End If
   Set adoRst = Nothing
End Function

'Added by Lydia 2016/12/30 呼叫共用表單->免費修正事由
Private Sub cmdFAmend_Click()
Dim rsRead As New ADODB.Recordset
Dim sqlB As String
Dim intB As Integer

    Me.Tag = ""
    sqlB = "select '' V,FA02,FA01 FROM FreeAmendment order by FA01"
    intB = 0
    Set rsRead = ClsLawReadRstMsg(intB, sqlB)
    If intB = 1 Then
       Set frm880012.grdDataList.Recordset = rsRead
       Set frm880012.fmParent = Me
       frm880012.iTyp = "2"
       frm880012.Show vbModal
       If Me.Tag <> "" And txtEP12.Text <> "" Then
          intB = MsgBox("是否取代承辦備註？ (是:取代原備註, 否:新增在原備註前面, 取消: 不變更)", vbCritical + vbYesNoCancel, "免費修正事由")
          If intB = 6 Then 'Yes
             txtEP12.Text = Me.Tag
          ElseIf intB = 7 Then 'No
             txtEP12.Text = Me.Tag & ";" & txtEP12.Text
          End If
       ElseIf Me.Tag <> "" Then
          txtEP12.Text = Me.Tag
       End If
    End If
End Sub

'Added by Lydia 2021/12/28
Private Sub txtEP12_GotFocus()
    TextInverse txtEP12
End Sub

Private Sub txtCP144_GotFocus()
    TextInverse txtCP144
End Sub

Private Sub txtCP99_GotFocus()
    TextInverse txtCP99
End Sub

Private Sub txtCP64_GotFocus()
    TextInverse txtCP64
End Sub

'Added by Morgan 2024/3/19
Private Sub ExportWord()
   Dim iResumeCnt As Integer
   Dim stTmp As String
   Dim oTable As Word.Table
   Dim iCol As Integer, iRow As Integer, ii As Integer, jj As Integer
   Dim iWorkDays As Integer    '截到目前工作天
   Dim stDate1 As String, stDate2 As String, stlstWkDate As String, stMA As String
   
On Error GoTo ErrHnd
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   
   If g_WordAp.Visible And g_WordAp.Documents.Count > 0 Then
      If MsgBox("輸出資料是否附加在目前的文件後面？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
         g_WordAp.Selection.EndKey Unit:=wdStory
         g_WordAp.Selection.TypeParagraph
      Else
         g_WordAp.Documents.add
      End If
   Else
      g_WordAp.Documents.add
   End If
   
   With g_WordAp.Application
      .WindowState = wdWindowStateMaximize
      .Visible = True
      
      '邊框設單線
      With .Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
      End With
      '橫印
      .Selection.PageSetup.Orientation = wdOrientLandscape
      '邊界
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1)
      
      .Selection.Font.Name = "標楷體"
      .Selection.Font.Size = 12
      
      stTmp = "工程師達成情況："
      .Selection.TypeText Text:=stTmp
      .Selection.TypeParagraph
      
      '新增表格
      Set oTable = .Selection.Tables.add(Range:=.Selection.Range, NumRows:=1, NumColumns:=6)
      
      'oTable.AllowAutoFit = True
      .Selection.SelectRow
      With .Selection.Borders(wdBorderTop)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderLeft)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderBottom)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderRight)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      With .Selection.Borders(wdBorderHorizontal)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
      End With
      With .Selection.Borders(wdBorderVertical)
          .LineStyle = g_WordAp.Options.DefaultBorderLineStyle
          .LineWidth = g_WordAp.Options.DefaultBorderLineWidth
      End With
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
      
      ii = 1
      oTable.Columns(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.7), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(1).Select
      .Selection.TypeText "工程師"
      
      oTable.Columns(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
      oTable.Columns(3).SetWidth ColumnWidth:=.CentimetersToPoints(2), RulerStyle:=wdAdjustProportional
      oTable.Columns(4).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
      oTable.Columns(5).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
      oTable.Rows(ii).Cells(5).Select
      .Selection.TypeText "上週增加"
      
      oTable.Rows(ii).Cells(6).Select
      .Selection.TypeText "特殊未達成原因 "
      
      stDate1 = 100 * (Val(frm090614.txt1(3) + 1911)) + Val(frm090614.txt1(4)) & "01"
      stDate2 = CompDate(2, -1, CompDate(1, 1, stDate1))
      stlstWkDate = CompWorkDay(2, strSrvDate(1), 1)
      
      iWorkDays = 0
      strSql = "Select Count(*) From WorkDay Where WD01>=" & stDate1 & " and WD01<=" & stDate2 & _
         " and WD01<=" & stlstWkDate
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         iWorkDays = RsTemp(0)
      End If
      
'      intI = Weekday(ChangeWStringToWDateString(stDate1))
'      If intI < 4 Then
'         stDate1 = stDate1 + (7 - intI)
'      Else
'         stDate1 = stDate1 + 7 + (7 - intI)
'      End If
'
'      If stlstWkDate <= stDate1 Then
'         stMA = "MA08"
'      ElseIf stlstWkDate <= stDate1 + 7 Then
'         stMA = "MA13"
'      ElseIf stlstWkDate <= stDate1 + 14 Then
'         stMA = "MA21"
'      Else
'         stMA = "MA29"
'      End If
      
      
      For jj = 0 To Combo1.ListCount - 1
         strExc(1) = Trim(Left("" & Combo1.List(jj), 6))
         If Right(strExc(1), 2) < "9" Then 'Added by Morgan 2025/2/18 排除兼職工程師的編號--柏翰
            strSql = "select * from monthassess,staff where ma02=" & (100 * (Val(frm090614.txt1(3) + 1911)) + Val(frm090614.txt1(4))) & " and ma03='1' and ma01='" & strExc(1) & "' and st01(+)=ma01"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               ii = ii + 1
               oTable.Rows.add
               oTable.Rows(ii).Cells(1).Select
               .Selection.TypeText RsTemp("st02") '"工程師"
               oTable.Rows(ii).Cells(2).Select
               .Selection.TypeText RsTemp("ma33") '"累計完稿"
               oTable.Rows(ii).Cells(3).Select
               .Selection.TypeText RsTemp("ma34") & "%" '"達成比例"
               oTable.Rows(ii).Cells(4).Select
               stTmp = Format(Val(RsTemp("ma33")) - (Val(RsTemp("ma04")) / Val(RsTemp("ma05"))) * iWorkDays, "0.00")
               .Selection.TypeText IIf(Val(stTmp) >= 0, "+", "") & stTmp '"目前進度"
               '不用
               'oTable.Rows(ii).Cells(5).Select
               '.Selection.TypeText RsTemp(stMA)   '"上週增加"
            End If
         End If
      Next
      
      '小字
      oTable.Rows(1).Cells(5).Select
      .Selection.Font.Size = 9
      '欄位合併
      oTable.Rows(1).Cells(2).Select
      .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
      .Selection.Cells.Merge
      oTable.Rows(1).Cells(2).Select
      .Selection.TypeText "目前進度"
      
      .Activate
   End With
   
ErrHnd:
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤" & Err.Number & " : " & Err.Description, vbCritical
         End Select
      End If
   End If
End Sub

'Added by Lydia 2025/02/05
Private Sub SetGrd1_New(Optional ByVal pReset As Boolean = True)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
                  '參考StrMenu1和AddToMdb實際存放的欄位
                  'cmdok2_Click欄位: R110002,R110003,R110004,R110005,R110006,R110023=NVL(NA03,NA04),R110008,R110009,R110007,R110011
                  '                  R110033,R110010,format(R110026=CP97 * r110027=CP98 * iif(r110030='N',1,iif(isnull(r110028=) or r110028=0,1,r110028)),'0.00') ,R110012,R110013,R110014,R110031,R110015,R110016,R110017,
                  '                  R110018,R110019,R110020,R110021,R110022, R110024=取消收文日NVL(CP57,PA58),r110029=是否會稿EP34,r110025=承辦人,r110030=加乘註記CP112,r110032=0
                  '================================================================
                  'R110026=CP97, R110027=CP98, R110028=CP111, R110029=EP34, R11030=CP112
                  'R110031=EP28, R110032=0, R110033=null, R110034=null
   '內專工程師增加顯示「指定日期」CP142=R110033
   'Modified by Morgan 2025/7/9 +收文點數
   arrGridHeadText = Array("目次", "收文類別", "收文日", "本所案號", "案件名稱", "國家", "種類", "案件性質", "Y/N", "本所期限", _
                         "指定日期", "承辦期限", "考核值", "收文點數", "法定期限", "齊備日", "完稿日", "預會日", "會稿日", "核稿人", "會稿完成日", _
                         "發文日", "承辦天數", "備註", "智權人員", "總收文號", "取消收文日", "是否會稿", "承辦人", "加乘註記", "R110032")
   arrGridHeadWidth = Array(350, 200, 795, 1005, 1155, 450, 440, 795, 285, 795, _
                         795, 795, 435, 0, 0, 795, 795, 795, 795, 700, 795, _
                         795, 800, 2000, 800, 0, 0, 0, 0, 0, 0)
   If pReset = True Then
       GRD1.Clear
       GRD1.Rows = 2
   End If
'   GRD1.Visible = False 'Modify By Sindy 2025/3/5 mark,外層控制不然會相互影響
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      If iRow <= UBound(arrGridHeadWidth) Then
         GRD1.Text = arrGridHeadText(iRow)
         GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
        Else
         GRD1.Text = ""
         GRD1.ColWidth(iRow) = 0
      End If
   Next
   
   'Added by Morgan 2025/7/18
   '收文點數
   If ProState = "2" Then
      GRD1.ColWidth(13) = 600
      GRD1.ColAlignment(13) = 7
   End If
   'end 2025/7/18
         
   'Modify By Sindy 2025/3/5 mark,外層控制不然會相互影響
'   GRD1.Visible = True 'Added by Lydia 2025/02/10
   
   If colCP09_1 = 0 Then
      colCP09_1 = PUB_MGridGetId("總收文號", GRD1)
      colCp06_1 = PUB_MGridGetId("本所期限", GRD1)
      colCp07_1 = PUB_MGridGetId("法定期限", GRD1)
      colEp12_1 = PUB_MGridGetId("備註", GRD1)
      colCaseNo_1 = PUB_MGridGetId("本所案號", GRD1)
      colCaseName_1 = PUB_MGridGetId("案件名稱", GRD1)
      colCp57_1 = PUB_MGridGetId("取消收文日", GRD1)
      colCp48_1 = PUB_MGridGetId("承辦期限", GRD1)
      colPv_1 = PUB_MGridGetId("考核值", GRD1)
      colEp06_1 = PUB_MGridGetId("齊備日", GRD1)
      colEp09_1 = PUB_MGridGetId("完稿日", GRD1)
      colEp28_1 = PUB_MGridGetId("預會日", GRD1)
      colEp34_1 = PUB_MGridGetId("是否會稿", GRD1)
      colEp07_1 = PUB_MGridGetId("會稿日", GRD1)
      colCp27_1 = PUB_MGridGetId("發文日", GRD1)
      colCPM_1 = PUB_MGridGetId("案件性質", GRD1)
      colEp04_1 = PUB_MGridGetId("核稿人", GRD1)
      colEp08_1 = PUB_MGridGetId("會稿完成日", GRD1)
      colEp35_1 = PUB_MGridGetId("承辦天數", GRD1)
   End If

End Sub

'Added by Morgan 2025/7/8
'累計收文點數
Private Sub SetPoint()
   Dim stYr As String, stSQL As String, intQ As Integer, stVTB As String, rsQuery As ADODB.Recordset
   Dim stConCP14 As String, stCon As String
   
   stYr = Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2))
   stConCP14 = " and cp14='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   
   '累計完稿收文點數
   stCon = " and ep09>=" & stYr & "01 and ep09<=" & stYr & "31"
   'Modified by Morgan 2025/7/18 統一都不扣銷帳
   'stVTB = "SELECT A1U03,SUM(A1U07) AS A1U07 FROM ENGINEERPROGRESS,CASEPROGRESS,ACC1U0 WHERE CP09(+)=EP02 AND A1U03(+)=CP09 AND A1U07<>0 " & stCon & stConCP14 & " GROUP BY A1U03"
   'stSQL = "select nvl(sum(CP18-nvl(a1u07/1000,0)),0) from ENGINEERPROGRESS,CASEPROGRESS,(" & stVTB & ") X where CP09(+)=EP02 and A1U03(+)=CP09" & stCon & stConCP14
   stSQL = "select nvl(sum(CP18),0) from ENGINEERPROGRESS,CASEPROGRESS where CP09(+)=EP02 " & stCon & stConCP14
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      lblCal1(4) = Format(rsQuery(0), "0.00") & " 點"
   End If
   
   '累計會稿收文點數
   stCon = " and ep07>=" & stYr & "01 and ep07<=" & stYr & "31"
   'Modified by Morgan 2025/7/18 統一都不扣銷帳
   'stVTB = "SELECT A1U03,SUM(A1U07) AS A1U07 FROM ENGINEERPROGRESS,CASEPROGRESS,ACC1U0 WHERE CP09(+)=EP02 AND A1U03(+)=CP09 AND A1U07<>0 " & stCon & stConCP14 & " GROUP BY A1U03"
   'stSQL = "select nvl(sum(CP18-nvl(a1u07/1000,0)),0) from ENGINEERPROGRESS,CASEPROGRESS,(" & stVTB & ") X where CP09(+)=EP02 and A1U03(+)=CP09" & stCon & stConCP14
   stSQL = "select nvl(sum(CP18),0) from ENGINEERPROGRESS,CASEPROGRESS where CP09(+)=EP02" & stCon & stConCP14
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      lblCal1(5) = Format(rsQuery(0), "0.00") & " 點"
   End If
   
   '累計發文收文點數
   'Modified by Morgan 2025/7/18 統一都不扣銷帳
   stCon = " and cp27>=" & stYr & "01 and cp27<=" & stYr & "31"
   'stVTB = "SELECT A1U03,SUM(A1U07) AS A1U07 FROM CASEPROGRESS,ACC1U0 WHERE A1U03(+)=CP09 AND A1U07<>0 " & stCon & stConCP14 & " GROUP BY A1U03"
   'stSQL = "select nvl(sum(CP18-nvl(a1u07/1000,0)),0) from CASEPROGRESS,(" & stVTB & ") X where A1U03(+)=CP09" & stCon & stConCP14
   stSQL = "select nvl(sum(CP18),0) from CASEPROGRESS where 1=1" & stCon & stConCP14
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      lblCal1(6) = Format(rsQuery(0), "0.00") & " 點"
   End If
   
   Set rsQuery = Nothing
End Sub
