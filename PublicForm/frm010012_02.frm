VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010012_02 
   BorderStyle     =   1  '單線固定
   Caption         =   "內部收文"
   ClientHeight    =   6180
   ClientLeft      =   4440
   ClientTop       =   3936
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9144
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   350
      Left            =   3120
      TabIndex        =   62
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   350
      Left            =   4740
      TabIndex        =   63
      Top             =   60
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Left            =   6915
      TabIndex        =   64
      Top             =   60
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5970
      TabIndex        =   61
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8145
      TabIndex        =   65
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   408
      Width           =   1335
   End
   Begin VB.TextBox textTM29_2 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000FF&
      Height          =   264
      Left            =   2460
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   408
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5430
      Left            =   90
      TabIndex        =   31
      Top             =   705
      Width           =   8895
      _ExtentX        =   15706
      _ExtentY        =   9589
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm010012_02.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label39"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label38"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label28"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label27"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label26"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label25"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(8)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label9"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label23(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label15"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label16"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label13"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label11"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label12"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label14"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label33"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label34"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label37"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label36"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label41"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label23(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(39)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label42"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label43"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label44"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label45"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label46"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label47"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label40"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCP14_2"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCP13_2"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textTM05"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textTM07"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textTM06"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textTM23_2"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textSP58_2"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textSP59_2"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textCP56_2"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textTM05_1"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textTM81_2"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textCP89_2"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textCP90_2"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textCP91_2"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textCP92_2"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textTM23_3"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textCP16"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "textCP48"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "textCP14"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "textCP10"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textCP10_2"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "textCP07"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "textCP06"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "textTM10"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "textCP13"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "textTM08"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "textTM08_2"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "textTM10_2"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "textTM28"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "textCP26"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "textTM29"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "textTM23"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "textSP58"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "textSP59"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "textCP05"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "textCP56"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "textCP20"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "textCP18"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "textCP17"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "textTM81"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "textTM80_2"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "textTM80"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "textCP89"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "textCP90"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "textCP91"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "textCP92"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "chkWebApp"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "grdList"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).ControlCount=   84
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm010012_02.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label22"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label55"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label18"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label49"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label50"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label51"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label52"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label53"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label54"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label56"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label57"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label58"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label59"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label60"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label61"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label35"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Line1"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label20"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label24"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "textTM44_2"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "textTM24"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "textTM26"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "textTM25"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "textTM86"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "textTM90"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "textTM82"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "textTM87"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "textTM91"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "textTM83"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "textTM88"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "textTM92"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "textTM84"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "textTM89"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "textTM93"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "textTM85"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "textTM45"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "textTM44"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "textCP43"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "textCP01_S"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "textCP02_S"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "textCP03_S"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "textCP04_S"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "textTM34"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "textTM35"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).ControlCount=   47
      TabCaption(2)   =   "第三頁"
      TabPicture(2)   =   "frm010012_02.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "textTM09"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdPriority"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "textTM32"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "textTM58"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "textCP64"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label29"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label30"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label31"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label32"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label48"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1224
         Left            =   984
         TabIndex        =   149
         Top             =   4080
         Width           =   7752
         _ExtentX        =   13674
         _ExtentY        =   2159
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin VB.CheckBox chkWebApp 
         Caption         =   "電子送件"
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   2130
         Width           =   1050
      End
      Begin VB.TextBox textTM35 
         Height          =   270
         Left            =   -69180
         MaxLength       =   50
         TabIndex        =   55
         Top             =   5040
         Width           =   2892
      End
      Begin VB.TextBox textTM34 
         Height          =   270
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   54
         Top             =   5040
         Width           =   2805
      End
      Begin VB.TextBox textCP04_S 
         Height          =   270
         Left            =   -71370
         MaxLength       =   2
         TabIndex        =   53
         Top             =   4775
         Width           =   465
      End
      Begin VB.TextBox textCP03_S 
         Height          =   270
         Left            =   -71835
         MaxLength       =   1
         TabIndex        =   52
         Top             =   4775
         Width           =   345
      End
      Begin VB.TextBox textCP02_S 
         Height          =   270
         Left            =   -72900
         MaxLength       =   6
         TabIndex        =   51
         Top             =   4775
         Width           =   975
      End
      Begin VB.TextBox textCP01_S 
         Height          =   270
         Left            =   -73440
         MaxLength       =   1
         TabIndex        =   50
         Top             =   4775
         Width           =   465
      End
      Begin VB.TextBox textTM09 
         Height          =   270
         Left            =   -73470
         MaxLength       =   395
         TabIndex        =   57
         Top             =   630
         Width           =   7152
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "輸入(&V)"
         Height          =   252
         Left            =   -73470
         TabIndex        =   56
         Top             =   330
         Width           =   1332
      End
      Begin VB.TextBox textTM32 
         Height          =   270
         Left            =   -73470
         MaxLength       =   1500
         TabIndex        =   58
         Top             =   930
         Width           =   7152
      End
      Begin VB.TextBox textCP92 
         Height          =   270
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   14
         Top             =   2130
         Width           =   1095
      End
      Begin VB.TextBox textCP91 
         Height          =   270
         Left            =   5520
         MaxLength       =   9
         TabIndex        =   13
         Top             =   1860
         Width           =   1095
      End
      Begin VB.TextBox textCP90 
         Height          =   270
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   12
         Top             =   1860
         Width           =   1095
      End
      Begin VB.TextBox textCP89 
         Height          =   270
         Left            =   5520
         MaxLength       =   9
         TabIndex        =   11
         Top             =   1590
         Width           =   1095
      End
      Begin VB.TextBox textTM80 
         Height          =   270
         Left            =   960
         MaxLength       =   9
         TabIndex        =   29
         Top             =   4320
         Width           =   972
      End
      Begin VB.TextBox textTM80_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   4320
         Width           =   2172
      End
      Begin VB.TextBox textTM81 
         Height          =   270
         Left            =   5475
         MaxLength       =   9
         TabIndex        =   30
         Top             =   4320
         Width           =   972
      End
      Begin VB.TextBox textCP17 
         Height          =   270
         Left            =   4260
         MaxLength       =   8
         TabIndex        =   20
         Top             =   2670
         Width           =   1452
      End
      Begin VB.TextBox textCP18 
         Height          =   270
         Left            =   7260
         MaxLength       =   8
         TabIndex        =   21
         Top             =   2685
         Width           =   1095
      End
      Begin VB.TextBox textCP20 
         Height          =   270
         Left            =   7260
         MaxLength       =   1
         TabIndex        =   18
         Top             =   2400
         Width           =   372
      End
      Begin VB.TextBox textCP56 
         Height          =   270
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   10
         Top             =   1590
         Width           =   1095
      End
      Begin VB.TextBox textCP05 
         Height          =   270
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1335
         Width           =   1215
      End
      Begin VB.TextBox textSP59 
         Height          =   270
         Left            =   5472
         MaxLength       =   9
         TabIndex        =   28
         Top             =   4035
         Width           =   972
      End
      Begin VB.TextBox textSP58 
         Height          =   270
         Left            =   960
         MaxLength       =   9
         TabIndex        =   27
         Top             =   4035
         Width           =   972
      End
      Begin VB.TextBox textTM23 
         Height          =   270
         Left            =   960
         MaxLength       =   9
         TabIndex        =   26
         Top             =   3750
         Width           =   972
      End
      Begin VB.TextBox textTM29 
         Height          =   270
         Left            =   4260
         MaxLength       =   1
         TabIndex        =   17
         Top             =   2400
         Width           =   372
      End
      Begin VB.TextBox textCP26 
         Height          =   270
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   16
         Top             =   2400
         Width           =   372
      End
      Begin VB.TextBox textTM28 
         Height          =   270
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1335
         Width           =   372
      End
      Begin VB.TextBox textTM10_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   6180
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   540
         Width           =   1572
      End
      Begin VB.TextBox textTM08_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1572
      End
      Begin VB.TextBox textTM08 
         Height          =   270
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1080
         Width           =   732
      End
      Begin VB.TextBox textCP13 
         Height          =   270
         Left            =   5520
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1080
         Width           =   852
      End
      Begin VB.TextBox textTM10 
         Height          =   270
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   3
         Top             =   540
         Width           =   612
      End
      Begin VB.TextBox textCP06 
         Height          =   270
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   4
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox textCP07 
         Height          =   270
         Left            =   5520
         MaxLength       =   7
         TabIndex        =   5
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox textCP10_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   540
         Width           =   1572
      End
      Begin VB.TextBox textCP10 
         Height          =   270
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   2
         Top             =   540
         Width           =   732
      End
      Begin VB.TextBox textCP14 
         Height          =   270
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Top             =   270
         Width           =   732
      End
      Begin VB.TextBox textCP43 
         Height          =   270
         Left            =   -73440
         MaxLength       =   9
         TabIndex        =   32
         Top             =   270
         Width           =   2655
      End
      Begin VB.TextBox textTM44 
         Height          =   270
         Left            =   -69600
         MaxLength       =   8
         TabIndex        =   34
         Top             =   535
         Width           =   972
      End
      Begin VB.TextBox textTM45 
         Height          =   270
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   33
         Top             =   535
         Width           =   2655
      End
      Begin VB.TextBox textCP48 
         Height          =   270
         Left            =   5520
         MaxLength       =   8
         TabIndex        =   1
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox textCP16 
         Height          =   270
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   19
         Top             =   2670
         Width           =   1452
      End
      Begin VB.TextBox textTM23_3 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   3750
         Width           =   2895
      End
      Begin MSForms.TextBox textTM85 
         Height          =   300
         Left            =   -73440
         TabIndex        =   47
         Top             =   3975
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM93 
         Height          =   300
         Left            =   -73440
         TabIndex        =   49
         Top             =   4515
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM89 
         Height          =   300
         Left            =   -73440
         TabIndex        =   48
         Top             =   4245
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   154
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM84 
         Height          =   300
         Left            =   -73440
         TabIndex        =   44
         Top             =   3180
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM92 
         Height          =   300
         Left            =   -73440
         TabIndex        =   46
         Top             =   3720
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM88 
         Height          =   300
         Left            =   -73440
         TabIndex        =   45
         Top             =   3450
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   154
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM83 
         Height          =   300
         Left            =   -73440
         TabIndex        =   41
         Top             =   2385
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM91 
         Height          =   300
         Left            =   -73440
         TabIndex        =   43
         Top             =   2925
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM87 
         Height          =   300
         Left            =   -73440
         TabIndex        =   42
         Top             =   2655
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   154
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM82 
         Height          =   300
         Left            =   -73440
         TabIndex        =   38
         Top             =   1590
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM90 
         Height          =   300
         Left            =   -73440
         TabIndex        =   40
         Top             =   2130
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM86 
         Height          =   300
         Left            =   -73440
         TabIndex        =   39
         Top             =   1860
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   154
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM58 
         Height          =   1125
         Left            =   -73470
         TabIndex        =   59
         Top             =   1230
         Width           =   7185
         VariousPropertyBits=   -1467987941
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12674;1984"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   1455
         Left            =   -73485
         TabIndex        =   60
         Top             =   2460
         Width           =   7185
         VariousPropertyBits=   -1467987941
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12674;2566"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP92_2 
         Height          =   264
         Left            =   2340
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   2130
         Width           =   2055
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP91_2 
         Height          =   264
         Left            =   6660
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   1860
         Width           =   2055
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP90_2 
         Height          =   264
         Left            =   2340
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   1860
         Width           =   2055
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP89_2 
         Height          =   264
         Left            =   6660
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   1590
         Width           =   2055
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM81_2 
         Height          =   264
         Left            =   6510
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   4320
         Width           =   2172
         VariousPropertyBits=   671107103
         ForeColor       =   -2147483641
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05_1 
         Height          =   792
         Left            =   1440
         TabIndex        =   22
         Top             =   2940
         Width           =   7272
         VariousPropertyBits=   -1467987941
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "12356;1402"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP56_2 
         Height          =   264
         Left            =   2340
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   1590
         Width           =   2055
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP59_2 
         Height          =   264
         Left            =   6516
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   4035
         Width           =   2172
         VariousPropertyBits=   671107103
         ForeColor       =   -2147483641
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP58_2 
         Height          =   264
         Left            =   2040
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   4035
         Width           =   2172
         VariousPropertyBits=   671107103
         ForeColor       =   -2147483641
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM25 
         Height          =   300
         Left            =   -73440
         TabIndex        =   36
         Top             =   1065
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   154
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM26 
         Height          =   300
         Left            =   -73440
         TabIndex        =   37
         Top             =   1335
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM24 
         Height          =   300
         Left            =   -73440
         TabIndex        =   35
         Top             =   795
         Width           =   7155
         VariousPropertyBits=   671107097
         MaxLength       =   70
         Size            =   "12621;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2 
         Height          =   264
         Left            =   2040
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   3750
         Width           =   2172
         VariousPropertyBits=   671107103
         ForeColor       =   -2147483641
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM06 
         Height          =   264
         Left            =   1440
         TabIndex        =   24
         Top             =   3210
         Width           =   7272
         VariousPropertyBits=   671107099
         MaxLength       =   60
         Size            =   "12303;476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM07 
         Height          =   264
         Left            =   1440
         TabIndex        =   25
         Top             =   3480
         Width           =   7272
         VariousPropertyBits=   671107099
         MaxLength       =   60
         Size            =   "12303;476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05 
         Height          =   264
         Left            =   1440
         TabIndex        =   23
         Top             =   2940
         Width           =   7272
         VariousPropertyBits=   671107099
         MaxLength       =   140
         Size            =   "12303;476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP13_2 
         Height          =   264
         Left            =   6480
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2295
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   264
         Left            =   2040
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   270
         Width           =   1572
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM44_2 
         Height          =   264
         Left            =   -68520
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   535
         Width           =   2232
         VariousPropertyBits=   671107103
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label40 
         Caption         =   "本案期限："
         Height          =   255
         Left            =   45
         TabIndex        =   101
         Top             =   4110
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "客戶案件案號 :"
         Height          =   255
         Left            =   -70560
         TabIndex        =   148
         Top             =   5045
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "分所案號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   147
         Top             =   5045
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   -73110
         X2              =   -71040
         Y1              =   4890
         Y2              =   4890
      End
      Begin VB.Label Label35 
         Caption         =   "查名本所案號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   146
         Top             =   4780
         Width           =   1335
      End
      Begin VB.Label Label61 
         Caption         =   "申請地址5(日) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   145
         Top             =   4515
         Width           =   1335
      End
      Begin VB.Label Label60 
         Caption         =   "申請地址5(中) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   144
         Top             =   3985
         Width           =   1335
      End
      Begin VB.Label Label59 
         Caption         =   "申請地址5(英) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   143
         Top             =   4250
         Width           =   1215
      End
      Begin VB.Label Label58 
         Caption         =   "申請地址4(日) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   142
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label57 
         Caption         =   "申請地址4(中) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   141
         Top             =   3190
         Width           =   1335
      End
      Begin VB.Label Label56 
         Caption         =   "申請地址4(英) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   140
         Top             =   3455
         Width           =   1215
      End
      Begin VB.Label Label54 
         Caption         =   "申請地址3(日) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   139
         Top             =   2925
         Width           =   1335
      End
      Begin VB.Label Label53 
         Caption         =   "申請地址3(中) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   138
         Top             =   2395
         Width           =   1335
      End
      Begin VB.Label Label52 
         Caption         =   "申請地址3(英) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   137
         Top             =   2660
         Width           =   1215
      End
      Begin VB.Label Label51 
         Caption         =   "申請地址2(日) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   136
         Top             =   2130
         Width           =   1335
      End
      Begin VB.Label Label50 
         Caption         =   "申請地址2(中) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   135
         Top             =   1600
         Width           =   1335
      End
      Begin VB.Label Label49 
         Caption         =   "申請地址2(英) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   134
         Top             =   1865
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "申請地址1(日) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   133
         Top             =   1335
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "商品類別 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   132
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "優先權資料 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   131
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "案件備註 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   130
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   129
         Top             =   2490
         Width           =   975
      End
      Begin VB.Label Label48 
         Caption         =   "商品組群 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   128
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Label47 
         Caption         =   "移轉申請人5 :"
         Height          =   195
         Left            =   120
         TabIndex        =   127
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label46 
         Caption         =   "移轉申請人4 :"
         Height          =   195
         Left            =   4440
         TabIndex        =   125
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "移轉申請人3 :"
         Height          =   195
         Left            =   120
         TabIndex        =   123
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label44 
         Caption         =   "移轉申請人2 :"
         Height          =   195
         Left            =   4440
         TabIndex        =   121
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label43 
         Caption         =   "申請人4 :"
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label42 
         Caption         =   "申請人5 :"
         Height          =   255
         Left            =   4440
         TabIndex        =   118
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否向客戶收款:"
         Height          =   180
         Index           =   39
         Left            =   5880
         TabIndex        =   115
         Top             =   2445
         Width           =   1305
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不收)"
         Height          =   255
         Index           =   1
         Left            =   7890
         TabIndex        =   114
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label41 
         Caption         =   "案件名稱 :"
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "移轉申請人1 :"
         Height          =   195
         Left            =   120
         TabIndex        =   111
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "收文日 :"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   1335
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "申請人3 :"
         Height          =   255
         Left            =   4440
         TabIndex        =   105
         Top             =   4050
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "申請人2 :"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   4050
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "申請地址1(英) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   103
         Top             =   1070
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "申請地址1(中) :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   102
         Top             =   805
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "申請人1:"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   3765
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "案件英文名稱 :"
         Height          =   165
         Left            =   120
         TabIndex        =   99
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "案件日文名稱 :"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   3465
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "案件中文名稱 :"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "(Y:取消)"
         Height          =   255
         Left            =   4860
         TabIndex        =   96
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "是否取消閉卷 :"
         Height          =   255
         Left            =   2970
         TabIndex        =   95
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不算)"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   94
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "(1:申請 2:異議 3:評定 4:舉發)"
         Height          =   255
         Left            =   6000
         TabIndex        =   93
         Top             =   1335
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "卷宗性質 :"
         Height          =   255
         Left            =   4440
         TabIndex        =   91
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "商標種類 :"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員 :"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   89
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "申請國家 :"
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   88
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   4440
         TabIndex        =   86
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "案件性質 :"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "相關總收文號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   83
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label55 
         Caption         =   "FC代理人 :"
         Height          =   255
         Left            =   -70560
         TabIndex        =   82
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "彼所案號："
         Height          =   255
         Left            =   -74880
         TabIndex        =   81
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   255
         Left            =   4440
         TabIndex        =   80
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "費用 :"
         Height          =   180
         Left            =   840
         TabIndex        =   79
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "規費 :"
         Height          =   180
         Left            =   3690
         TabIndex        =   78
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "點數 :"
         Height          =   180
         Left            =   6720
         TabIndex        =   77
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label Label39 
         Caption         =   "申請人國籍 :"
         Height          =   255
         Left            =   4440
         TabIndex        =   76
         Top             =   3765
         Width           =   1095
      End
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "商標及查名的商品類別請輸在第二頁的""案件備註""欄內!!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4200
      TabIndex        =   112
      Top             =   450
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號 :"
      Height          =   255
      Left            =   120
      TabIndex        =   109
      Top             =   405
      Width           =   855
   End
End
Attribute VB_Name = "frm010012_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Morgan 2021/5/11 改成Form2.0 (textTM85,textTM93 ...)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/22 日期欄已修改
Option Explicit

Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String

Dim m_CPKeyList() As String
Dim m_CPKeyCount As Integer
' 收文日
Dim m_CP05 As String
' 收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 國家代碼
Dim m_TM10 As String
' 卷宗性質
Dim m_TM28 As String
' 是否閉卷
Dim m_TM29 As String
' 移轉申請人
Dim m_CP56 As String
'add by nickc 2006/12/01
Dim m_CP89 As String
Dim m_CP90 As String
Dim m_CP91 As String
Dim m_CP92 As String
Dim m_CP55 As String
Dim m_CP93 As String
Dim m_CP94 As String
Dim m_CP95 As String
Dim m_CP96 As String

' 相關總收文號
Dim m_CP43 As String

'910626 Sieg 601
' 收據編號
Dim m_CP60 As String
' 申請人
Dim m_TM23 As String
'add by nickc 2007/01/10
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存商標基本檔或服務業務基本檔檔案欄位的串列
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
' 儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
' 儲存國家的字串
Dim m_strCountry As String
'
Dim m_CurrSel As Integer
' 優先權畫面所使用的變數
Dim m_Pa(1 To 4) As String '本所案號
Dim m_Priority(1 To 6) As String
'Add By Cheng 2002/06/12
Dim m_strCP06 As String '原本所期限
Dim m_strCP07 As String '原法定期限
Dim m_TM22 As String '專用期止日
'Add By Cheng 2002/08/22
Dim m_strCust1 As String '申請人1
Dim m_strCust2 As String '申請人2
Dim m_strCust3 As String '申請人3
'add by nickc 2006/11/30
Dim m_strCust4 As String
Dim m_strCust5 As String
'Add By Sindy 2019/5/27
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_PrevForm As Form '前一畫面
'2019/5/27 END


Private Sub cmdCancel_Click()
   Unload Me
   frm010001.Show
End Sub

Private Sub cmdCaseProgress_Click()
   frm010012_03.SetData 0, m_TM01, True
   frm010012_03.SetData 1, m_TM02, False
   frm010012_03.SetData 2, m_TM03, False
   frm010012_03.SetData 3, m_TM04, False
   frm010012_03.SetData 4, m_CP09, False
   'Modified by Lydia 2020/04/21 改為Form型態
   'frm010012_03.SetParent "frm010012_02"
   frm010012_03.SetParent Me
   Me.Hide
   frm010012_03.Show
   frm010012_03.QueryData
End Sub

Private Sub cmdExit_Click()
   Unload Me
   Unload frm010001
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid = True Then
      '重新檢查欄位的正確性
      If ValidateInput() = False Then
         Exit Sub
      End If
      
      'Added by Lydia 2015/02/04 所有內部收文, 若有輸入本所期限或法定期限者, 檢查期限不可小於系統日
      'Modified by Lydia 2017/07/31 改為預設和檢查
      'If PUB_CheckCP0607(0, textCP06.Text, textCP07.Text) = False Then Exit Sub
      'Modified by Lyddia 2023/11/08 傳入必需欄位
      'If PUB_CheckCP0607(0, textCP06, textCP07) = False Then Exit Sub
      If PUB_CheckCP0607(0, textCP06, textCP07, "", textTM10, m_TM01, textCP10) = False Then Exit Sub
      
      OnUpdateField
        'Modify By Cheng 2002/11/06
'      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
      
      'Added by Lydia 2023/02/08 內部收文補收款，智權人員為SXX部門時，要發MAIL給杜協理及智權人員
      If (Left(m_TM01, 1) = "T" Or m_TM01 = "CFT" Or m_TM01 = "CFS" Or m_TM01 = "S") And textCP10 <> "" And InStr(textCP10_2, "補收款") > 0 And Left(GetST15(textCP13), 1) = "S" Then
          strExc(0) = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
          strExc(1) = "本所案號：" & strExc(0) & vbCrLf & _
                           "案件名稱：" & textTM05_1 & vbCrLf & _
                           "申請人1：" & textTM23 & " " & textTM23_2 & vbCrLf & _
                           "申請國家：" & textTM10_2 & vbCrLf & _
                           "補收款費用：" & Val(textCP16) & vbCrLf & _
                           "補收款備註：" & Trim(textCP64)
          strExc(2) = Pub_GetSpecMan("全所智權部主管")
          If InStr(strExc(2), textCP13) = 0 Then
              strExc(2) = strExc(2) & ";" & textCP13
          End If
          PUB_SendMail strUserNum, strExc(2), "", strExc(0) & "內部收文補收款通知!", strExc(1)
      End If
      'end 2023/02/08
      
      'Modify By Sindy 2019/5/27 信件內部收文執行完畢後,關閉視窗
      If m_strIR01 <> "" Then
         Unload frm010001
         Unload Me
      Else
      '2019/5/27 End
         ' 回到收文的畫面
         frm010001.SetData m_CP09, 0, True
         frm010001.SetData m_TM01, 1, False
         frm010001.SetData m_TM02, 2, False
         frm010001.SetData m_TM03, 3, False
         frm010001.SetData m_TM04, 4, False
         frm010001.Show
         ClearAll
         Unload Me
      End If
   End If
End Sub

Private Sub cmdPriority_Click()
   ' 修改優先權資料
   'Modify by Amy 2014/04/18 +, m_Priority(3)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   'Modify by Sindy 2019/1/23 + m_TM01 & m_TM02 & m_TM03 & m_TM04
   ModifyPriority m_Priority(1), m_Priority(2), m_Priority(3), , , m_TM01 & m_TM02 & m_TM03 & m_TM04, , , m_Priority(4), m_Priority(5), m_Priority(6)
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08_2.BackColor = &H8000000F
   textTM10_2.BackColor = &H8000000F
   textTM23_2.BackColor = &H8000000F
   textTM23_3.BackColor = &H8000000F
   textTM29_2.BackColor = &H8000000F
   textTM44_2.BackColor = &H8000000F
   textSP58_2.BackColor = &H8000000F
   textSP59_2.BackColor = &H8000000F
   textCP56_2.BackColor = &H8000000F
   'add by nickc 2006/12/01
   textTM80_2.BackColor = &H8000000F
   textTM81_2.BackColor = &H8000000F
   textCP89_2.BackColor = &H8000000F
   textCP90_2.BackColor = &H8000000F
   textCP91_2.BackColor = &H8000000F
   textCP92_2.BackColor = &H8000000F
   
   textCP10_2.BackColor = &H8000000F
   textCP13_2.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   
   SSTab1.Tab = 0
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/27
   m_strIR01 = frm010001.m_strIR01
   m_strIR02 = frm010001.m_strIR02
   m_strIR03 = frm010001.m_strIR03
   m_strIR04 = frm010001.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/27 END
End Sub

Private Sub InitialData()
   m_CPCount = 0
End Sub

Private Sub ClearAll()
   'ClearCPList
   ClearTMSPFieldList
   ClearCPFieldList
   'm_CP09 = Empty
   m_TM28 = Empty
   
   textTMKey = Empty
   textTM05 = Empty
   textTM05_1 = Empty
   textTM06 = Empty
   textTM07 = Empty
   textTM08 = Empty
   textTM08_2 = Empty
   textTM09 = Empty
   textTM10 = Empty
   textTM10_2 = Empty
   textTM23 = Empty
   textTM23_2 = Empty
   textTM24 = Empty
   textTM25 = Empty
   textTM26 = Empty
   textTM28 = Empty
   textTM29 = Empty
   textTM34 = Empty
   textTM44 = Empty
   textTM44_2 = Empty
   textTM45 = Empty
   textTM58 = Empty
   textSP58 = Empty
   textSP58_2 = Empty
   textSP59 = Empty
   textSP59_2 = Empty
   'add by nickc 2006/12/01
   textTM80 = Empty
   textTM80_2 = Empty
   textTM81 = Empty
   textTM81_2 = Empty
   
   textCP05 = Empty
   textCP06 = Empty
   textCP07 = Empty
   'textCP09_S = Empty
   textCP10 = Empty
   textCP10_2 = Empty
   textCP13 = Empty
   textCP13_2 = Empty
   textCP14 = Empty
   textCP14_2 = Empty
   textCP26 = Empty
   textCP43 = Empty
   textCP64 = Empty
   textCP01_S = Empty
   textCP02_S = Empty
   textCP03_S = Empty
   textCP04_S = Empty
   
   m_strCountry = Empty
End Sub

Private Sub AddCPToList(ByVal strCP09 As String)
   Dim bFind As Boolean
   Dim nIndex As Integer
   bFind = False
   For nIndex = 0 To m_CPKeyCount - 1
      If m_CPKeyList(nIndex) = strCP09 Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_CPKeyList(m_CPKeyCount + 1)
      m_CPKeyList(m_CPKeyCount) = strCP09
      m_CPKeyCount = m_CPKeyCount + 1
   End If
End Sub

Private Sub ClearCPList()
   If m_CPKeyCount > 0 Then
      Erase m_CPKeyList
   End If
   m_CPKeyCount = 0
End Sub

' 清除商標基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiOldData = strFieldData
         m_TMSPList(nPos).fiNewData = strFieldData
         m_TMSPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TMSPList(m_TMSPCount + 1)
      m_TMSPList(m_TMSPCount).fiName = strFieldName
      m_TMSPList(m_TMSPCount).fiOldData = strFieldData
      m_TMSPList(m_TMSPCount).fiNewData = strFieldData
      m_TMSPList(m_TMSPCount).fiType = nFieldType
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 清除案件進度檔檔案欄位串列
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

Public Sub SetData(ByVal strData As String, ByVal nType As Integer, ByVal bClear As Boolean)
   If bClear Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP10 = Empty
      m_CP56 = Empty
      '92.03.27 nick
      m_CP09 = Empty
      'add by nickc 2006/12/01
      m_CP89 = Empty
      m_CP90 = Empty
      m_CP91 = Empty
      m_CP92 = Empty
   End If
   
   Select Case nType
      Case 0: m_TM01 = strData
      Case 1: m_TM02 = strData
      Case 2: m_TM03 = strData & String(1 - Len(strData), "0")
      Case 3: m_TM04 = strData & String(2 - Len(strData), "0")
      Case 4:
              m_CP10 = strData
                '911113 nick 當案件性質是501 時 ，移轉申請人才可出現
                If m_CP10 = "501" Then
                   Label36.Visible = True
                   textCP56.Enabled = True
                   textCP56.Visible = True
                   textCP56.TabStop = True
                   textCP56_2.Visible = True
                   'add by nickc 2007/01/10
                   Label44.Visible = True
                   textCP89.Enabled = True
                   textCP89.Visible = True
                   textCP89.TabStop = True
                   textCP89_2.Visible = True
                   Label45.Visible = True
                   textCP90.Enabled = True
                   textCP90.Visible = True
                   textCP90.TabStop = True
                   textCP90_2.Visible = True
                   Label46.Visible = True
                   textCP91.Enabled = True
                   textCP91.Visible = True
                   textCP91.TabStop = True
                   textCP91_2.Visible = True
                   Label47.Visible = True
                   textCP92.Enabled = True
                   textCP92.Visible = True
                   textCP92.TabStop = True
                   textCP92_2.Visible = True
                Else
                   Label36.Visible = False
                   textCP56.Enabled = False
                   textCP56.Visible = False
                   textCP56.TabStop = False
                   textCP56_2.Visible = False
                   'add by nickc 2007/01/10
                   Label44.Visible = False
                   textCP89.Enabled = False
                   textCP89.Visible = False
                   textCP89.TabStop = False
                   textCP89_2.Visible = False
                   Label45.Visible = False
                   textCP90.Enabled = False
                   textCP90.Visible = False
                   textCP90.TabStop = False
                   textCP90_2.Visible = False
                   Label46.Visible = False
                   textCP91.Enabled = False
                   textCP91.Visible = False
                   textCP91.TabStop = False
                   textCP91_2.Visible = False
                   Label47.Visible = False
                   textCP92.Enabled = False
                   textCP92.Visible = False
                   textCP92.TabStop = False
                   textCP92_2.Visible = False
                End If
              
      Case 5:
         If Not IsEmptyText(strData) Then
            m_CP56 = strData & String(9 - Len(strData), "0")
         End If
      Case 6:
         m_CP43 = strData
         textCP43 = m_CP43
         'Added by Lydia 2017/10/16
         If m_CP43 <> "" Then
            textCP43_Validate False
         End If
         'end 2017/10/16
         
      Case 7:
         m_CP09 = strData
      'add by nickc 2006/12/01
      Case 8
         If Not IsEmptyText(strData) Then
            m_CP89 = strData & String(9 - Len(strData), "0")
         End If
      Case 9
         If Not IsEmptyText(strData) Then
            m_CP90 = strData & String(9 - Len(strData), "0")
         End If
      Case 10
         If Not IsEmptyText(strData) Then
            m_CP91 = strData & String(9 - Len(strData), "0")
         End If
      Case 11
         If Not IsEmptyText(strData) Then
            m_CP92 = strData & String(9 - Len(strData), "0")
         End If
   End Select
End Sub

' 取得商標基本檔的欄位內容
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset

   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
'      ' 案件中文名稱
'      If IsNull(rsTmp.Fields("TM05")) = False Then
'         textTM05 = rsTmp.Fields("TM05")
'      End If
'      SetTMSPFieldOldData "TM05", textTM05, 0
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05_1 = rsTmp.Fields("TM05")
      End If
      SetTMSPFieldOldData "TM05", textTM05_1, 0
'      ' 案件英文名稱
'      If IsNull(rsTmp.Fields("TM06")) = False Then
'         textTM06 = rsTmp.Fields("TM06")
'      End If
'      SetTMSPFieldOldData "TM06", textTM06, 0
'      ' 案件日文名稱
'      If IsNull(rsTmp.Fields("TM07")) = False Then
'         textTM07 = rsTmp.Fields("TM07")
'      End If
'      SetTMSPFieldOldData "TM07", textTM07, 0
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         textTM08 = rsTmp.Fields("TM08")
         textTM08_Validate False
      End If
      SetTMSPFieldOldData "TM08", textTM08, 0
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      SetTMSPFieldOldData "TM09", textTM09, 0
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         textTM10 = rsTmp.Fields("TM10")
         textTM10_Validate False
         m_TM10 = rsTmp.Fields("TM10")
      End If
      SetTMSPFieldOldData "TM10", m_TM10, 0
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = rsTmp.Fields("TM23")
         'textTM23_Validate False
         textTM23_2 = GetCustomerName(textTM23, 0)
         'Added by Lydia 2017/10/16 申請人國籍
         strExc(0) = GetCustomerNation(m_TM23)
         If strExc(0) <> "" Then
            textTM23_3 = GetNationName(Mid(strExc(0), 1, 3), 0)
         End If
         'end 2017/10/16
      End If
      SetTMSPFieldOldData "TM23", textTM23, 0
      'add by nickc 2006/11/30
      ' 申請人2
      If IsNull(rsTmp.Fields("TM78")) = False Then
         textSP58 = rsTmp.Fields("TM78")
         textSP58_2 = GetCustomerName(textSP58, 0)
      End If
      SetTMSPFieldOldData "TM78", textSP58, 0
      ' 申請人3
      If IsNull(rsTmp.Fields("TM79")) = False Then
         textSP59 = rsTmp.Fields("TM79")
         textSP59_2 = GetCustomerName(textSP59, 0)
      End If
      SetTMSPFieldOldData "TM79", textSP59, 0
      ' 申請人4
      If IsNull(rsTmp.Fields("TM80")) = False Then
         textTM80 = rsTmp.Fields("TM80")
         textTM80_2 = GetCustomerName(textTM80, 0)
      End If
      SetTMSPFieldOldData "TM80", textTM80, 0
      ' 申請人5
      If IsNull(rsTmp.Fields("TM81")) = False Then
         textTM81 = rsTmp.Fields("TM81")
         textTM81_2 = GetCustomerName(textTM81, 0)
      End If
      SetTMSPFieldOldData "TM81", textTM81, 0
      'add by nickc 2007/01/10
      m_TM78 = textSP58
      m_TM79 = textSP59
      m_TM80 = textTM80
      m_TM81 = textTM81
      
      'Add By Cheng 2002/08/22
      m_strCust1 = "" & Me.textTM23.Text
      'edit by nickc 2006/11/30
      'm_strCust2 = ""
      'm_strCust3 = ""
      m_strCust2 = "" & Me.textSP58.Text
      m_strCust3 = "" & Me.textSP59.Text
      'add by nickc 2006/11/30
      m_strCust4 = "" & Me.textTM80.Text
      m_strCust5 = "" & Me.textTM81.Text
      ' 申請地址
      If IsNull(rsTmp.Fields("TM24")) = False Then
         textTM24 = rsTmp.Fields("TM24")
      End If
      SetTMSPFieldOldData "TM24", textTM24, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM25")) = False Then
         textTM25 = rsTmp.Fields("TM25")
      End If
      SetTMSPFieldOldData "TM25", textTM25, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("TM26")) = False Then
         textTM26 = rsTmp.Fields("TM26")
      End If
      SetTMSPFieldOldData "TM26", textTM26, 0
      'add by nickc 2006/12/01
      textTM82 = "" & rsTmp.Fields("TM82")
      SetTMSPFieldOldData "TM82", textTM82, 0
      textTM83 = "" & rsTmp.Fields("TM83")
      SetTMSPFieldOldData "TM83", textTM83, 0
      textTM84 = "" & rsTmp.Fields("TM84")
      SetTMSPFieldOldData "TM84", textTM84, 0
      textTM85 = "" & rsTmp.Fields("TM85")
      SetTMSPFieldOldData "TM85", textTM85, 0
      textTM86 = "" & rsTmp.Fields("TM86")
      SetTMSPFieldOldData "TM86", textTM86, 0
      textTM87 = "" & rsTmp.Fields("TM87")
      SetTMSPFieldOldData "TM87", textTM87, 0
      textTM88 = "" & rsTmp.Fields("TM88")
      SetTMSPFieldOldData "TM88", textTM88, 0
      textTM89 = "" & rsTmp.Fields("TM89")
      SetTMSPFieldOldData "TM89", textTM89, 0
      textTM90 = "" & rsTmp.Fields("TM90")
      SetTMSPFieldOldData "TM90", textTM90, 0
      textTM91 = "" & rsTmp.Fields("TM91")
      SetTMSPFieldOldData "TM91", textTM91, 0
      textTM92 = "" & rsTmp.Fields("TM92")
      SetTMSPFieldOldData "TM92", textTM92, 0
      textTM93 = "" & rsTmp.Fields("TM93")
      SetTMSPFieldOldData "TM93", textTM93, 0
      textTM32 = "" & rsTmp.Fields("TM32")
      SetTMSPFieldOldData "TM32", textTM32, 0
      
      
      ' 卷宗性質
      If IsNull(rsTmp.Fields("TM28")) = False Then
         m_TM28 = rsTmp.Fields("TM28")
         textTM28 = rsTmp.Fields("TM28")
      End If
      SetTMSPFieldOldData "TM28", textTM28, 0
      ' 是否閉卷
      If IsNull(rsTmp.Fields("TM29")) = False Then
         m_TM29 = rsTmp.Fields("TM29")
      End If
      ' 分所案號
      If IsNull(rsTmp.Fields("TM34")) = False Then
         textTM34 = rsTmp.Fields("TM34")
      End If
      SetTMSPFieldOldData "TM34", textTM34, 0
      ' 客戶案件案號
      If IsNull(rsTmp.Fields("TM35")) = False Then
         textTM35 = rsTmp.Fields("TM35")
      End If
      SetTMSPFieldOldData "TM35", textTM35, 0
      ' FC代理人
      If IsNull(rsTmp.Fields("TM44")) = False Then
         textTM44 = rsTmp.Fields("TM44")
         textTM44_Validate False
      End If
      SetTMSPFieldOldData "TM44", textTM44, 0
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then
         textTM45 = rsTmp.Fields("TM45")
      End If
      SetTMSPFieldOldData "TM45", textTM45, 0
      ' 案件備註
      If IsNull(rsTmp.Fields("TM58")) = False Then
         textTM58 = rsTmp.Fields("TM58")
      End If
      SetTMSPFieldOldData "TM58", textTM58, 0
      'Add By Cheng 2002/06/12
      '取得專用期止日
      m_TM22 = "" & rsTmp.Fields("TM22")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得服務業務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textTM05 = rsTmp.Fields("SP05")
      End If
      SetTMSPFieldOldData "SP05", textTM05, 0
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textTM06 = rsTmp.Fields("SP06")
      End If
      SetTMSPFieldOldData "SP06", textTM06, 0
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textTM07 = rsTmp.Fields("SP07")
      End If
      SetTMSPFieldOldData "SP07", textTM07, 0
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textTM23 = rsTmp.Fields("SP08")
         'textTM23_Validate False
         textTM23_2 = GetCustomerName(textTM23, 0)
      End If
      SetTMSPFieldOldData "SP08", textTM23, 0
      'Add By Cheng 2002/08/22
      m_strCust1 = "" & Me.textTM23.Text
'edit by nickc 2007/01/10
'      ' 第二申請人及第三申請人
'      If m_TM01 = "CFC" Then
         If IsNull(rsTmp.Fields("SP58")) = False Then
            textSP58 = rsTmp.Fields("SP58")
            textSP58_Validate False
         End If
         If IsNull(rsTmp.Fields("SP59")) = False Then
            textSP59 = rsTmp.Fields("SP59")
            textSP59_Validate False
         End If
         'Add By Cheng 2002/08/22
         m_strCust2 = "" & Me.textSP58.Text
         m_strCust3 = "" & Me.textSP59.Text
'      End If
      'add by nickc 2006/11/30
      m_strCust4 = ""
      m_strCust5 = ""
      If IsNull(rsTmp.Fields("SP65")) = False Then
         textTM80 = rsTmp.Fields("SP65")
         textTM80_Validate False
      End If
      If IsNull(rsTmp.Fields("SP66")) = False Then
         textTM81 = rsTmp.Fields("SP66")
         textTM81_Validate False
      End If
      m_strCust4 = "" & Me.textTM80.Text
      m_strCust5 = "" & Me.textTM81.Text
      m_TM78 = textSP58
      m_TM79 = textSP59
      m_TM80 = textTM80
      m_TM81 = textTM81
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
         textTM10 = rsTmp.Fields("SP09")
         textTM10_Validate False
      End If
      SetTMSPFieldOldData "SP09", m_TM10, 0
      ' 是否閉卷
      If IsNull(rsTmp.Fields("SP15")) = False Then
         m_TM29 = rsTmp.Fields("SP15")
      End If
      ' FC代理人
      If IsNull(rsTmp.Fields("SP26")) = False Then
         textTM44 = rsTmp.Fields("SP26")
         textTM44_Validate False
      End If
      SetTMSPFieldOldData "SP26", textTM44, 0
      ' 彼所案號
      If IsNull(rsTmp.Fields("SP27")) = False Then
         textTM45 = rsTmp.Fields("SP27")
      End If
      SetTMSPFieldOldData "SP27", textTM45, 0
      ' 案件備註
      If IsNull(rsTmp.Fields("SP18")) = False Then
         textTM58 = rsTmp.Fields("SP18")
      End If
      SetTMSPFieldOldData "SP18", textTM58, 0
      'add by nickc 2006/12/04
      If IsNull(rsTmp.Fields("SP73")) = False Then
         textTM09 = rsTmp.Fields("SP73")
      End If
      SetTMSPFieldOldData "SP73", textTM09, 0
      If IsNull(rsTmp.Fields("SP74")) = False Then
         textTM32 = rsTmp.Fields("SP74")
      End If
      SetTMSPFieldOldData "SP74", textTM32, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgressWithNewCP()
   Dim strSql As String
   Dim strTemp As String
   Dim strSubSQL As String
   Dim rsSubTmp As ADODB.Recordset
   
   SetCPFieldOldData "CP01", Empty, 0
   SetCPFieldOldData "CP02", Empty, 0
   SetCPFieldOldData "CP03", Empty, 0
   SetCPFieldOldData "CP04", Empty, 0
   SetCPFieldOldData "CP05", Empty, 1
   SetCPFieldOldData "CP06", Empty, 1
   SetCPFieldOldData "CP07", Empty, 1
   SetCPFieldOldData "CP09", Empty, 0
   SetCPFieldOldData "CP10", Empty, 0
   ' 業務區
   SetCPFieldOldData "CP12", Empty, 0
   ' 智權人員
   SetCPFieldOldData "CP13", Empty, 0
   ' 承辦人員
   SetCPFieldOldData "CP14", Empty, 0
   ' 費用
   SetCPFieldOldData "CP16", Empty, 1
   ' 規費
   SetCPFieldOldData "CP17", Empty, 1
   ' 點數
   SetCPFieldOldData "CP18", Empty, 1
   ' 是否算案件數
   SetCPFieldOldData "CP26", Empty, 0
   ' 對造案件中文名稱
   SetCPFieldOldData "CP37", Empty, 0
   ' 對造案件英文名稱
   SetCPFieldOldData "CP38", Empty, 0
   ' 對造案件日文名稱
   SetCPFieldOldData "CP39", Empty, 0
   ' 相關總收文號
   SetCPFieldOldData "CP43", Empty, 0
   ' 承辦期限
   SetCPFieldOldData "CP48", Empty, 1
   SetCPFieldOldData "CP64", Empty, 0
   
   '911108 nick
   ' 案件來源
   SetCPFieldOldData "CP11", Empty, 0
   ' 是否向客戶收款
   SetCPFieldOldData "CP20", Empty, 0
   SetCPFieldOldData "CP21", Empty, 0
   '911113 nick
   SetCPFieldOldData "CP32", Empty, 0
   'Added by Lydia 2019/08/21 電子送件
   SetCPFieldOldData "CP118", Empty, 0
   
      'add by nickc 2007/01/08
      m_CP55 = Empty
      m_CP93 = Empty
      m_CP94 = Empty
      m_CP95 = Empty
      m_CP96 = Empty
      SetCPFieldOldData "CP55", m_CP55, 0
      SetCPFieldOldData "CP93", m_CP93, 0
      SetCPFieldOldData "CP94", m_CP94, 0
      SetCPFieldOldData "CP95", m_CP95, 0
      SetCPFieldOldData "CP96", m_CP96, 0
      SetCPFieldOldData "CP56", Empty, 0
      SetCPFieldOldData "CP89", Empty, 0
      SetCPFieldOldData "CP90", Empty, 0
      SetCPFieldOldData "CP91", Empty, 0
      SetCPFieldOldData "CP92", Empty, 0
      
    'Modify By Cheng 2003/08/21
    '若系統類別為T, CFT, FCT, TF
    If m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "CFT" Or m_TM01 = "FCT" Then
        ' 卷宗性質不為1時, 案件中英日文名稱從案件進度檔中帶入
        If IsEmptyText(m_CP10) = False Then
           If m_TM28 <> "1" Then
              'textTM05 = Empty
              'textTM06 = Empty
              'textTM07 = Empty
              Set rsSubTmp = New ADODB.Recordset
              strSubSQL = "SELECT * FROM CaseProgress " & _
                          "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                "CP02 = '" & m_TM02 & "' AND " & _
                                "CP03 = '" & m_TM03 & "' AND " & _
                                "CP04 = '" & m_TM04 & "' AND " & _
                                "CP31 = 'Y' "
              rsSubTmp.CursorLocation = adUseClient
              rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
              If rsSubTmp.RecordCount > 0 Then
                 rsSubTmp.MoveFirst
'                 ' 對造案件中文名稱
'                 If IsNull(rsSubTmp.Fields("CP37")) = False Then
'                    If IsEmptyText(rsSubTmp.Fields("CP37")) = False Then
'                       textTM05 = rsSubTmp.Fields("CP37")
'                    End If
'                 End If
'                 SetCPFieldOldData "CP37", textTM05, 0
                 ' 對造案件名稱
                 If IsNull(rsSubTmp.Fields("CP37")) = False Then
                    If IsEmptyText(rsSubTmp.Fields("CP37")) = False Then
                       textTM05_1 = rsSubTmp.Fields("CP37")
                    End If
                 End If
                 SetCPFieldOldData "CP37", textTM05_1, 0
'                 ' 對造案件英文名稱
'                 If IsNull(rsSubTmp.Fields("CP38")) = False Then
'                    If IsEmptyText(rsSubTmp.Fields("CP38")) = False Then
'                       textTM06 = rsSubTmp.Fields("CP38")
'                    End If
'                 End If
'                 SetCPFieldOldData "CP38", textTM06, 0
'                 ' 對造案件日文名稱
'                 If IsNull(rsSubTmp.Fields("CP39")) = False Then
'                    If IsEmptyText(rsSubTmp.Fields("CP39")) = False Then
'                       textTM07 = rsSubTmp.Fields("CP39")
'                    End If
'                 End If
'                 SetCPFieldOldData "CP39", textTM07, 0
              End If
              rsSubTmp.Close
              Set rsSubTmp = Nothing
           End If
        End If
    End If
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   Dim strSubSQL As String
   Dim rsSubTmp As ADODB.Recordset
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         textCP10 = rsTmp.Fields("CP10")
         textCP10_Validate False
      End If
      SetCPFieldOldData "CP10", textCP10, 0
      ' 收文日
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP05")) = False Then
         strTemp = rsTmp.Fields("CP05")
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      SetCPFieldOldData "CP05", strTemp, 1
      ' 本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         textCP06 = TAIWANDATE(rsTmp.Fields("CP06"))
      End If
      SetCPFieldOldData "CP06", textCP06, 1
      'Add By Cheng 2002/06/12
      m_strCP06 = "" & rsTmp.Fields("CP06")
      
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         textCP07 = TAIWANDATE(rsTmp.Fields("CP07"))
      End If
      SetCPFieldOldData "CP07", textCP07, 1
      'Add By Cheng 2002/06/12
      m_strCP07 = "" & rsTmp.Fields("CP07")
      ' 業務區
      SetCPFieldOldData "CP12", rsTmp.Fields("CP12"), 0
      
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = rsTmp.Fields("CP13")
         textCP13_Validate False
      End If
      SetCPFieldOldData "CP13", textCP13, 0
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         textCP14_Validate False
      End If
      SetCPFieldOldData "CP14", textCP14, 0
      ' 費用
      If IsNull(rsTmp.Fields("CP16")) = False Then
         textCP16 = rsTmp.Fields("CP16")
      End If
      SetCPFieldOldData "CP16", textCP16, 1
      ' 規費
      If IsNull(rsTmp.Fields("CP17")) = False Then
         textCP17 = rsTmp.Fields("CP17")
      End If
      SetCPFieldOldData "CP17", textCP17, 1
      ' 點數
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      SetCPFieldOldData "CP18", textCP18, 1
        'Add By Cheng 2004/03/16
      ' 是否向客戶收款
      If IsNull(rsTmp.Fields("CP20")) = False Then
         textCP20 = rsTmp.Fields("CP20")
      End If
      SetCPFieldOldData "CP20", textCP20, 0
        'End
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         textCP43 = rsTmp.Fields("CP43")
      End If
      SetCPFieldOldData "CP43", textCP43, 0
      ' 是否算案件數
      If IsNull(rsTmp.Fields("CP26")) = False Then
         textCP26 = rsTmp.Fields("CP26")
      End If
      SetCPFieldOldData "CP26", textCP26, 0
      ' 對造案件中文名稱
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP37")) = False Then
         strTemp = rsTmp.Fields("CP37")
      End If
      SetCPFieldOldData "CP37", strTemp, 0
      ' 對造案件英文名稱
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP38")) = False Then
         strTemp = rsTmp.Fields("CP38")
      End If
      SetCPFieldOldData "CP38", strTemp, 0
      ' 對造案件日文名稱
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP39")) = False Then
         strTemp = rsTmp.Fields("CP39")
      End If
      SetCPFieldOldData "CP39", strTemp, 0
      ' 承辦期限
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP48")) = False Then
         textCP48 = TAIWANDATE(rsTmp.Fields("CP48"))
         strTemp = rsTmp.Fields("CP48")
      End If
      SetCPFieldOldData "CP48", strTemp, 1
      
      '910626 Sieg 601
      '收據編號
      If IsNull(rsTmp.Fields("CP60")) = False Then
         m_CP60 = rsTmp.Fields("CP60")
      Else
         m_CP60 = ""
      End If
      
      'Add By Sindy 2019/7/11
      chkWebApp.Value = 0
      If IsNull(rsTmp.Fields("CP118")) = False Then
         If rsTmp.Fields("CP118") = "Y" Then
            chkWebApp.Value = 1
         End If
      End If
      SetCPFieldOldData "CP118", IIf(chkWebApp.Value = 1, "Y", ""), 0
      '2019/7/11 END
      
      ' 進度備註
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      'add by nickc 2007/01/10
      m_CP55 = CheckStr(rsTmp.Fields("CP55"))
      m_CP93 = CheckStr(rsTmp.Fields("CP93"))
      m_CP94 = CheckStr(rsTmp.Fields("CP94"))
      m_CP95 = CheckStr(rsTmp.Fields("CP95"))
      m_CP96 = CheckStr(rsTmp.Fields("CP96"))
      SetCPFieldOldData "CP55", m_CP55, 0
      SetCPFieldOldData "CP93", m_CP93, 0
      SetCPFieldOldData "CP94", m_CP94, 0
      SetCPFieldOldData "CP95", m_CP95, 0
      SetCPFieldOldData "CP96", m_CP96, 0
      m_CP56 = CheckStr(rsTmp.Fields("CP56"))
      m_CP89 = CheckStr(rsTmp.Fields("CP89"))
      m_CP90 = CheckStr(rsTmp.Fields("CP90"))
      m_CP91 = CheckStr(rsTmp.Fields("CP91"))
      m_CP92 = CheckStr(rsTmp.Fields("CP92"))
      textCP56 = m_CP56
      textCP56_Validate False
      textCP89 = m_CP89
      textCP89_Validate False
      textCP90 = m_CP90
      textCP90_Validate False
      textCP91 = m_CP91
      textCP91_Validate False
      textCP92 = m_CP92
      textCP92_Validate False
      SetCPFieldOldData "CP56", m_CP56, 0
      SetCPFieldOldData "CP89", m_CP89, 0
      SetCPFieldOldData "CP90", m_CP90, 0
      SetCPFieldOldData "CP91", m_CP91, 0
      SetCPFieldOldData "CP92", m_CP92, 0
      
        'Modify By Cheng 2003/08/21
        '若系統類別為T, CFT, FCT, TF
        If m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "CFT" Or m_TM01 = "FCT" Then
            ' 卷宗性質不為1時, 案件中英日文名稱從案件進度檔中帶入
            If IsEmptyText(m_CP10) = False Then
               If m_TM28 <> "1" Then
                  'textTM05 = Empty
                  'textTM06 = Empty
                  'textTM07 = Empty
                  Set rsSubTmp = New ADODB.Recordset
                  strSubSQL = "SELECT * FROM CaseProgress " & _
                              "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                    "CP02 = '" & m_TM02 & "' AND " & _
                                    "CP03 = '" & m_TM03 & "' AND " & _
                                    "CP04 = '" & m_TM04 & "' AND " & _
                                    "CP31 = 'Y' "
                  rsSubTmp.CursorLocation = adUseClient
                  rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsSubTmp.RecordCount > 0 Then
                     rsSubTmp.MoveFirst
'                     ' 對造案件中文名稱
'                     If IsNull(rsSubTmp.Fields("CP37")) = False Then
'                        If IsEmptyText(rsSubTmp.Fields("CP37")) = False Then
'                           textTM05 = rsSubTmp.Fields("CP37")
'                        End If
'                     End If
'                     SetCPFieldOldData "CP37", textTM05, 0
                     ' 對造案件名稱
                     If IsNull(rsSubTmp.Fields("CP37")) = False Then
                        If IsEmptyText(rsSubTmp.Fields("CP37")) = False Then
                           textTM05_1 = rsSubTmp.Fields("CP37")
                        End If
                     End If
                     SetCPFieldOldData "CP37", textTM05_1, 0
'                     ' 對造案件英文名稱
'                     If IsNull(rsSubTmp.Fields("CP38")) = False Then
'                        If IsEmptyText(rsSubTmp.Fields("CP38")) = False Then
'                           textTM06 = rsSubTmp.Fields("CP38")
'                        End If
'                     End If
'                     SetCPFieldOldData "CP38", textTM06, 0
'                     ' 對造案件日文名稱
'                     If IsNull(rsSubTmp.Fields("CP39")) = False Then
'                        If IsEmptyText(rsSubTmp.Fields("CP39")) = False Then
'                           textTM07 = rsSubTmp.Fields("CP39")
'                        End If
'                     End If
'                     SetCPFieldOldData "CP39", textTM07, 0
                  End If
                  rsSubTmp.Close
                  Set rsSubTmp = Nothing
               End If
            End If
        End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nickI As Integer
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
   
   ' 已閉卷
   m_TM29 = Empty
   textTM29_2 = Empty
   
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearTMSPFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   m_CP05 = TAIWANDATE(SystemDate())
   textCP05 = m_CP05
   textCP56 = m_CP56
   textCP56_Validate False
   'add by nickc 2006/12/01
   textCP89 = m_CP89
   textCP89_Validate False
   textCP90 = m_CP90
   textCP90_Validate False
   textCP91 = m_CP91
   textCP91_Validate False
   textCP92 = m_CP92
   textCP92_Validate False
   textTM80_Validate False
   textTM81_Validate False
   
   '2008/10/31 add by sonia
   If m_TM01 = "CFT" Or m_TM01 = "CFC" Then
      textCP13 = PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      textCP13_Validate False
   End If
   '2008/10/31 end
   
   textTM08_Validate False
   textTM10_Validate False
   textTM23_Validate False
   textTM44_Validate False
   textSP58_Validate False
   textSP59_Validate False
    'Add By Cheng 2003/11/10
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        Me.Label13.Visible = False
        Me.textTM05.Visible = False
        Me.textTM05.Enabled = False
        Me.Label12.Visible = False
        Me.textTM06.Visible = False
        Me.textTM06.Enabled = False
        Me.Label11.Visible = False
        Me.textTM07.Visible = False
        Me.textTM07.Enabled = False
    Case Else
        Me.Label41.Visible = False
        Me.textTM05_1.Visible = False
        Me.textTM05_1.Enabled = False
    End Select
    'End
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   
   ' 取得案件進度檔的欄位
   '92.03.27 nick 修正
   If frm010001.intModifyKind = 0 Then
        QueryCaseProgressWithNewCP
   Else
        QueryCaseProgress
   End If
      
   textCP10 = m_CP10
   textCP10_Validate False
   
   'Add By Sindy 2019/7/11 所有FCT案件性質都顯示
   If m_TM01 = "FCT" And textTM10 = "000" Then
      chkWebApp.Visible = True
      'Modify By Sindy 2022/7/18 FCT內部收文之「延期303」案請預設為電子送件
      If textCP10 = "303" Then
         chkWebApp.Value = 1
      End If
      '2022/7/18 END
   Else
      chkWebApp.Visible = False
   End If
   '2019/7/11 END
   
   ' 是否閉卷
   If m_TM29 = "Y" Then
      EnableTextBox textTM29, True
      textTM29_2 = "已閉卷"
   Else
      EnableTextBox textTM29, False
      textTM29_2 = Empty
   End If
   
   ' 計算承辦期限
   If IsEmptyText(textCP48) = True Then
      ReCaculateCP48
   End If
   'add by nickc 2008/01/04 加入回代時，承辦期限為本所收文日(當天不算)之第二個工作天
   If m_CP10 = "720" Then
      textCP48 = TAIWANDATE(CompWorkDay(3, DBDATE(m_CP05), 0))
   'Add By Sindy 2015/8/21
   ElseIf m_CP10 = "901" Then '催款，承辦期限為本所收文日(當天不算)+3個工作天
      textCP48 = TAIWANDATE(CompWorkDay(4, DBDATE(m_CP05), 0))
   '2015/8/21 END
   End If
   
'edit by nickc 2007/02/13
'   ' 系統類別為CFC時可輸入三個申請人
'   If m_TM01 = "CFC" Then
      EnableTextBox textSP58, True
      EnableTextBox textSP59, True
'   Else
'      EnableTextBox textSP58, False
'      EnableTextBox textSP59, False
'   End If
   
   ' 依讀取的是商標基本檔還是服務業務基本檔來更新控制項的狀態
'edit by nickc 2006/12/04 服務已經有類別和組群
'   Select Case m_TM01
'      Case "T", "TF", "CFT", "FCT":
         EnableTextBox textTM09, True
         'add by nickc 2006/12/04
         EnableTextBox textTM32, True
'      Case Else:
'         EnableTextBox textTM09, False
'   End Select
   
'Added by Lydia 2017/10/16 為了放大下一程序的高度,將申請人2~4隱藏
textSP58.Visible = False: textSP59.Visible = False
textSP58_2.Visible = False: textSP59_2.Visible = False
Label33.Visible = False: Label34.Visible = False
textTM80.Visible = False: textTM81.Visible = False
textTM80_2.Visible = False: textTM81_2.Visible = False
Label43.Visible = False: Label42.Visible = False
'end 2017/10/16

   ' 讀取優先權資料
   m_Pa(1) = m_TM01
   m_Pa(2) = m_TM02
   m_Pa(3) = m_TM03
   m_Pa(4) = m_TM04
   'edit by nickc 2007/02/02 不用 dll 了
   'objPublicData.ReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
   'Modify by Amy 2014/04/18 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   ClsPDReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)
   
   ' 更新本案期限的資料
   UpdateGrdList m_TM01, m_TM02, m_TM03, m_TM04
   '911018 nick 新增時要帶下一程序資料     本所期限，法定期限，收文號==>相關總收文號，備註==>進度備註    #只有一筆時，且本所案號和案件性質都要輸入且找的到
   If frm010001.intModifyKind = 0 Then
        If m_TM01 <> "" And m_TM02 <> "" And m_TM03 <> "" And m_TM04 <> "" And m_CP10 <> "" Then
            Dim nick911018rs As New ADODB.Recordset
            Dim nickstrsql As String
            Set nick911018rs = New ADODB.Recordset
            '911111 nick 邱小姐說要加入 np06 is null  np06<>'Y'(包含 null) 同意義
            'nickstrsql = "select * from nextprogress where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07=" & m_CP10 & " "
            '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
            'nickstrsql = "select * from nextprogress where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07=" & m_CP10 & " and (np06 <>'Y' or np06 is null) "
            nickstrsql = "select * from nextprogress where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07=" & m_CP10 & " and  np06 is null "
            nick911018rs.CursorLocation = adUseClient
            nick911018rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
            If nick911018rs.RecordCount = 1 Then
                textCP06 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np08").Value))
                textCP07 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np09").Value))
                textCP43 = CheckStr(nick911018rs.Fields("np01").Value)
                textCP64 = textCP64 & CheckStr(nick911018rs.Fields("np15").Value)
                '91.11.10 ADD BY SONIA
                textCP13 = CheckStr(nick911018rs.Fields("np10").Value)
                textCP13_Validate False
                '91.11.10 END
                '911030 nick 自動上勾
                
                For nickI = 1 To grdList.Rows - 1
                    'edit by nick 2004/09/08
                    'If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And grdList.TextMatrix(nickI, 2) = textCP06 And grdList.TextMatrix(nickI, 3) = textCP07 Then
                    If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And Val(grdList.TextMatrix(nickI, 2)) = Val(textCP06) And Val(grdList.TextMatrix(nickI, 3)) = Val(textCP07) And textCP10.Text <> "305" Then
                        grdList.TextMatrix(nickI, 0) = "V"
                    End If
                Next nickI
            Else
                '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
                If nick911018rs.RecordCount = 0 Then
                    nickstrsql = "select * from nextprogress where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07=" & m_CP10 & " and np06 <>'Y' "
                    Set nick911018rs = New ADODB.Recordset
                    nick911018rs.CursorLocation = adUseClient
                    nick911018rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
                    If nick911018rs.RecordCount = 1 Then
                        textCP06 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np08").Value))
                        textCP07 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np09").Value))
                        textCP43 = CheckStr(nick911018rs.Fields("np01").Value)
                        textCP64 = textCP64 & CheckStr(nick911018rs.Fields("np15").Value)
                        textCP13 = CheckStr(nick911018rs.Fields("np10").Value)
                        textCP13_Validate False
                        For nickI = 1 To grdList.Rows - 1
                            'edit by nick 2004/09/08
                            'If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And grdList.TextMatrix(nickI, 2) = textCP06 And grdList.TextMatrix(nickI, 3) = textCP07 Then
                            If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And Val(grdList.TextMatrix(nickI, 2)) = Val(textCP06) And Val(grdList.TextMatrix(nickI, 3)) = Val(textCP07) And textCP10.Text <> "305" Then
                                grdList.TextMatrix(nickI, 0) = "V"
                            End If
                        Next nickI
                    End If
                End If
            End If
        End If
   End If
   ' 設定輸入的位置
   SetInputEntry
   
'edit by nickc 2008/01/03 可以輸，不強制
'   ' 91.09.11 申請人不輸入
'   EnableTextBox textTM23, False
'   EnableTextBox textSP58, False
'   EnableTextBox textSP59, False
'   'add by nickc 2006/12/01
'   EnableTextBox textTM80, False
'   EnableTextBox textTM81, False
   
   '92.03.27 nick 當查詢時，將確定 disabled
   If frm010001.intModifyKind = 2 Then
        cmdOK.Enabled = False
   End If
   
   'Add By Sindy 2012/6/22 S-1334~1440 內部收文時不限案件性質,都預設該案號最後收文之A或B類的承辦人及智權人員
   'Modify By Sindy 2012/11/16 + "and cp10 not in('718') " & _ 排除取消收文
   '將游標設定在第三頁進度備註
   'modify by sonia 2016/5/19 因有新人加入故改用S案之案件名稱判斷 S-004536
   'If m_TM01 = "S" And (Val(m_TM02) >= 1334 And Val(m_TM02) <= 1440) Then
   If m_TM01 = "S" And textTM05 = "未成卷代理人往來函" Then
      strExc(0) = "Select cp05,cp09,cp13,cp14,cp01,cp02 " & _
                  "From caseprogress " & _
                  "Where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' " & _
                  "and cp09<'C' " & _
                  "and cp10 not in('718') " & _
                  "order by cp09 desc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          textCP14 = "" & RsTemp.Fields("cp14")
          textCP13 = "" & RsTemp.Fields("cp13")
          Call textCP14_Validate(False)
          Call textCP13_Validate(False)
      End If
      'Modify By Sindy 2012/7/5 阿蓮說只有720回覆告代理人時,游標才須要跳到進度備註欄
      If m_CP10 = "720" Then
         SSTab1.Tab = 2
         textCP64.SetFocus
      End If
      '2012/7/5 End
   End If
   '2012/6/22 End
   'Added by Lydia 2023/04/06 FCT,S案請控制內部收文739更換智權人員：直帶出承辦人員為輸入人員
   If frm010001.intModifyKind = 0 And (m_TM01 = "FCT" Or m_TM01 = "S") And textCP10 = "739" Then
      textCP14 = strUserNum
      textCP14_Validate False
      textCP13.SetFocus
      textCP13_GotFocus
   End If
   'end 2023/04/06
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   m_CP09 = Empty
   
   'Add By Sindy 2019/5/27
   If m_strIR01 <> "" Then
      If Not m_PrevForm Is Nothing Then
         Call m_PrevForm.GoNext
         Set m_PrevForm = Nothing
      End If
   End If
   '2019/5/27 END
   
   'Add By Cheng 2002/07/19
   Set frm010012_02 = Nothing
End Sub

Private Sub textCP01_S_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 查名本所案號第一欄
Private Sub textCP01_S_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If Not IsEmptyText(textCP01_S) Then
      If textCP01_S <> "S" Then
         MsgBox "查名本所案號的系統類別類輸入錯誤!!!", vbExclamation + vbOKOnly
         Cancel = True
         textCP01_S.SetFocus
         textCP01_S_GotFocus
      End If
   End If
End Sub

' 查名本所案號的第四欄
Private Sub textCP04_S_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If Not IsEmptyText(textCP01_S) Then
      If Not IsEmptyText(textCP02_S) Then
         If Not ExistServicePractice(textCP01_S, textCP02_S, textCP03_S, textCP04_S) Then
            strTit = "檢核資料"
            strMsg = "查名本所案號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP01_S.SetFocus
            textCP01_S_GotFocus
         End If
      Else
         strTit = "檢核資料"
         strMsg = "查名本所案號輸入不完整"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP01_S.SetFocus
         textCP01_S_GotFocus
      End If
   End If
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
      Else
         ReCaculateCP48
      End If
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "本所期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06_GotFocus
      'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/07
      End If
   End If
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "法定期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
      End If
   End If
End Sub

' 查名收文號
'Private Sub textCP09_S_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSQL As String
'
'   Cancel = False
'   If IsEmptyText(textCP09_S) = False Then
'      strSQL = "SELECT * FROM CaseProgress " & _
'               "WHERE CP09 = '" & textCP09_S & "' "
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount <= 0 Then
'         Cancel = True
'         strTit = "檢核資料"
'         strMsg = "查名收文號不存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP09_S_GotFocus
'      End If
'      rsTmp.Close
'   End If
'   Set rsTmp = Nothing
'End Sub

' 案件性質
Private Sub textCP10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse

   textCP10_2 = Empty
   Cancel = False
   If IsEmptyText(textCP10) = False Then
      If m_TM10 < "010" Then
         ' 取得國內的案件性質名稱
         textCP10_2 = GetCaseTypeName(m_TM01, textCP10, 0)
      Else
         textCP10_2 = GetCaseTypeName(m_TM01, textCP10, 1)
      End If
      If IsEmptyText(textCP10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP10_GotFocus
      End If
      
      '911113 nick 當案件性質是501 時 ，移轉申請人才可出現
      If textCP10 = "501" Then
         Label36.Visible = True
         textCP56.Enabled = True
         textCP56.Visible = True
         textCP56.TabStop = True
         textCP56_2.Visible = True
         'add by nickc 2006/12/01
         Label44.Visible = True
         textCP89.Enabled = True
         textCP89.Visible = True
         textCP89.TabStop = True
         textCP89_2.Visible = True
         Label45.Visible = True
         textCP90.Enabled = True
         textCP90.Visible = True
         textCP90.TabStop = True
         textCP90_2.Visible = True
         Label46.Visible = True
         textCP91.Enabled = True
         textCP91.Visible = True
         textCP91.TabStop = True
         textCP91_2.Visible = True
         Label47.Visible = True
         textCP92.Enabled = True
         textCP92.Visible = True
         textCP92.TabStop = True
         textCP92_2.Visible = True
      Else
         Label36.Visible = False
         textCP56.Enabled = False
         textCP56.Visible = False
         textCP56.TabStop = False
         textCP56_2.Visible = False
         'add by nickc 2006/12/01
         Label44.Visible = False
         textCP89.Enabled = False
         textCP89.Visible = False
         textCP89.TabStop = False
         textCP89_2.Visible = False
         Label45.Visible = False
         textCP90.Enabled = False
         textCP90.Visible = False
         textCP90.TabStop = False
         textCP90_2.Visible = False
         Label46.Visible = False
         textCP91.Enabled = False
         textCP91.Visible = False
         textCP91.TabStop = False
         textCP91_2.Visible = False
         Label47.Visible = False
         textCP92.Enabled = False
         textCP92.Visible = False
         textCP92.TabStop = False
         textCP92_2.Visible = False
      End If
   End If
End Sub

'Add By Sindy 2010/11/25
Private Sub textCP13_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 智權人員
Private Sub textCP13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   'Added by Lydia 2019/02/14
   Dim m_SalesST15 As String '畫面上智權人員的收文部門
   Dim m_Tuser As String '創新業務部預設收文人員
   
   Cancel = False
   textCP13_2 = Empty
   If IsEmptyText(textCP13) = False Then
      textCP13_2 = GetStaffName(textCP13)
      If IsEmptyText(textCP13_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "智權人員代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0 'Add By Sindy 2012/6/22
         textCP13_GotFocus
      'Added by Lydia 2019/02/14 創新業務部人員收文控管
      Else
         m_SalesST15 = GetST15(textCP13)
         'Added by Lydia 2020/04/08 檢查案件或智權人員是否為法務部
         If PUB_ChkSalesL(m_TM01, textCP13.Text) = False Then
             SSTab1.Tab = 0
             textCP13.SetFocus
             Call textCP13_GotFocus
             Cancel = True
             Exit Sub
         End If
         'end 2020/04/08
         If PUB_ChkIsT10T20("2", textCP13.Text, m_Tuser, strTit) = True Then
             SSTab1.Tab = 0
             textCP13.Text = m_Tuser
             textCP13_2.Text = strTit
             textCP13.SetFocus
             Call textCP13_GotFocus
             Cancel = True
             Exit Sub
         End If
      'end 2019/02/14
      End If
   End If
End Sub

'Add By Sindy 2010/11/25
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0 'Add By Sindy 2012/6/22
         textCP14_GotFocus
      'add by sonia 2023/5/23  外商人員收文案件FCT,S,還有FMT及CFT案,告代719回代720都要管制
      Else
         If Left(PUB_GetStaffST15(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), "1"), 2) = "F1" Then
            If (m_CP10 = "719" Or m_CP10 = "720") Then
               If PUB_GetStaffST15(textCP14, "1") <> "F10" And PUB_GetStaffST15(textCP14, "1") <> "F11" Then
                  Cancel = True
                  strTit = "資料檢核"
                  strMsg = "此類案件之告知代理人或回覆代理人的智權人員必須為外商承辦組人員！"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  SSTab1.Tab = 0 'Add By Sindy 2012/6/22
                  textCP13_GotFocus
               Else
                  textCP13.Text = textCP14
                  textCP13_2.Text = textCP14_2
               End If
            End If
         End If
      'end 2023/5/23
      End If
   '911111 nick 邱小姐說承辦人不可空白
   Else
      Cancel = True
      strTit = "資料檢核"
      strMsg = "承辦人不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0 'Add By Sindy 2012/6/22
      textCP14_GotFocus
   End If
   
'cancel by sonia 2023/5/23 移至上面
'   'Add By Sindy 2010/01/04 FCT案且為回覆代理人720或告知代理人719時,設定智權人員為輸入之承辦人
'   If (m_TM01 = "FCT" Or m_TM01 = "S") And (m_CP10 = "719" Or m_CP10 = "720") Then
'      textCP13.Text = textCP14
'      textCP13_2.Text = textCP14_2
'   End If
'   '2010/01/04 End
'end 2023/5/23
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
         strTit = "資料檢核"
         strMsg = "費用為數值資料"
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
         strTit = "資料檢核"
         strMsg = "規費為數值資料"
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
         strTit = "資料檢核"
         strMsg = "點數為數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP18_GotFocus
      End If
   End If
End Sub

Private Sub textCP20_GotFocus()
    TextInverse Me.textCP20
End Sub

Private Sub textCP20_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 78 Then
        KeyAscii = 0
    End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

' 相關總收文號
Private Sub textCP43_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCP64.Tag = "" 'Added by Lydia 2017/10/16
   If IsEmptyText(textCP43) = False Then
      If textCP43 = m_CP09 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號不可為本身之收文號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_TM01 & "' AND " & _
                     "CP02 = '" & m_TM02 & "' AND " & _
                     "CP03 = '" & m_TM03 & "' AND " & _
                     "CP04 = '" & m_TM04 & "' AND " & _
                     "CP09 = '" & textCP43 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         rsTmp.Close
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號資料不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2017/10/16 內部收文302更正，一定要輸入相關總收文，若點選為核准，則進度備註cp64='更改核准函' (原因參考:106063001 FCT延展、移轉、變更核准定稿暫不列印)
      ElseIf m_TM01 = "FCT" And textCP10 = "302" And rsTmp.Fields("CP10") = "1001" Then
         textCP64.Text = "更改核准函"
         textCP64.Tag = "更改核准函" '
      'end 2017/10/16
      End If
      rsTmp.Close
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 承辦期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "承辦期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
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
   Dim strCP56 As String
   Dim strTemp As String

   Cancel = False
   textCP56_2 = Empty
   If Not IsEmptyText(textCP56) Then
      strCP56 = textCP56
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If objPublicData.GetCustomer(strCP56, strTemp) Then
      'Modify By Sindy 2015/8/27 +m_TM01
      If GetCustomerAndState(strCP56, strTemp, , , , m_TM01) Then
         textCP56 = strCP56 & String(9 - Len(strCP56), "0")
         textCP56_2 = strTemp
      Else
         Cancel = True
         textCP56_GotFocus
      End If
   End If
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
      nResponse = MsgBox(strMsg, vbOKOnly, strTit) 'Added by Lydia 2017/10/16
      textCP64_GotFocus
   End If
   
   'Added by Lydia 2017/10/16 內部收文302更正，一定要輸入相關總收文，若點選為核准，則進度備註cp64='更改核准函' (原因參考:106063001 FCT延展、移轉、變更核准定稿暫不列印)
   If textCP64.Tag <> "" Then
      If textCP64.Text = "" Or (textCP64.Text <> "" And InStr(textCP64.Text, textCP64.Tag) = 0) Then
        Cancel = True
        strTit = "檢核資料"
        strMsg = "FCT內部收文更正，相關總收文若為核准，則進度備註為更改核准函!"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        textCP64_GotFocus
      End If
   End If
   'end 2017/10/16
   
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCP64.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub UpdateGrdList(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String)
   Dim nIndex As Integer
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   'Modify by Morgan 2009/12/25 下一程序要排除程序管制的案件性質
   '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
   strSql = "SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
            "WHERE NP02 = '" & strTM01 & "' AND " & _
                  "NP03 = '" & strTM02 & "' AND " & _
                  "NP04 = '" & strTM03 & "' AND " & _
                  "NP05 = '" & strTM04 & "' AND " & _
                  "(NP06 IS NULL OR NP06 <> 'Y') " & strNpSqlOfNoSalesDuty
   
   'Add by Morgan 2009/12/25 延期+AB類未發文未取消收文的程序
   If textCP10 = "303" Then
      textCP10.Enabled = False
      strSql = strSql & " UNION SELECT CP09,CP10,CP06,CP07,CP13,CP57,CP58,CP08,CP40,CP64,0" & _
         " FROM CASEPROGRESS WHERE CP01 = '" & strTM01 & "' AND CP02 = '" & strTM02 & "'" & _
         " AND CP03 = '" & strTM03 & "' AND CP04 = '" & strTM04 & "'" & _
         " AND CP09<'C' and cp10<>'303' and cp07>0 AND CP27 IS NULL AND CP57 IS NULL"
   End If
   
   'Added by Lydia 2017/10/16 CFT,CFC,S,FCT內部收文時，下方之本案期限改由本所期限由大至小排序
   If m_TM01 = "CFT" Or m_TM01 = "CFC" Or m_TM01 = "S" Or m_TM01 = "FCT" Then
      strSql = strSql & " ORDER BY NP08 DESC,NP07 ASC"
   End If
   'end 2017/10/16
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         grdList.Rows = grdList.Rows + 1
         nIndex = grdList.Rows - 1
         
         ' 收文號
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(nIndex, 8) = rsTmp.Fields("NP01")
         End If
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            '911111 nick 案件性質要依國家判斷
            'grdList.TextMatrix(nIndex, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"))
            grdList.TextMatrix(nIndex, 1) = GetPrjState4(strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04, rsTmp.Fields("NP07"))
            
            grdList.TextMatrix(nIndex, 9) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(nIndex, 2) = ChangeWStringToTString(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(nIndex, 3) = ChangeWStringToTString(rsTmp.Fields("NP09"))
            End If
         End If
         ' 機關文號
         If IsNull(rsTmp.Fields("NP13")) = False Then
            grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("NP13")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(nIndex, 5) = rsTmp.Fields("NP14")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("NP15")
         End If
         ' 解除期限日期
         If IsNull(rsTmp.Fields("NP11")) = False Then
            grdList.TextMatrix(nIndex, 7) = rsTmp.Fields("NP11")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(nIndex, 10) = rsTmp.Fields("NP22")
         End If
         '911111 nick 智權人員
         If IsNull(rsTmp.Fields("NP10")) = False Then
            grdList.TextMatrix(nIndex, 11) = rsTmp.Fields("NP10")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/16
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/16
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Function IsCaseProgressExist(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   IsCaseProgressExist = False
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & strTM01 & "' AND " & _
                  "CP02 = '" & strTM02 & "' AND " & _
                  "CP03 = '" & strTM03 & "' AND " & _
                  "CP04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsCaseProgressExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function IsDataRecordExist(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   IsDataRecordExist = False
   Select Case strTM01
      Case "T", "TF", "FCT", "CFT":
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' AND " & _
                        "TM03 = '" & strTM03 & "' AND " & _
                        "TM04 = '" & strTM04 & "' "
      Case Else
         strSql = "SELECT * FROM ServicePractice " & _
                  "WHERE SP01 = '" & strTM01 & "' AND " & _
                        "SP02 = '" & strTM02 & "' AND " & _
                        "SP03 = '" & strTM03 & "' AND " & _
                        "SP04 = '" & strTM04 & "' "
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsDataRecordExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   '911111 nick
   'grdList.Cols = 11
   grdList.Cols = 12
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "下一程序"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "本所期限"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "法定期限"
   grdList.ColWidth(3) = 1000
   grdList.col = 4
   grdList.Text = "機關文號"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "相關人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "備註"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "解除期限日"
   grdList.ColWidth(7) = 1200
   grdList.col = 8
   grdList.Text = "收文號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "下一程序代號"
   grdList.ColWidth(9) = 0
   grdList.col = 10
   grdList.Text = "序號"
   grdList.ColWidth(10) = 0
   '911111 nick add
   grdList.col = 11
   grdList.Text = "序號"
   grdList.ColWidth(11) = 0
End Sub

Private Sub grdList_Click()
   If grdList.row > 0 Then
      grdList.col = 0
      If grdList.Text = "V" Then
         grdList.Text = Empty
      Else
             'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
             If Pub_CheckNpTheSameShow(m_TM01, textCP10, Trim("" & grdList.TextMatrix(grdList.row, 9))) = False Then
                 Exit Sub
             End If
             'end 2021/08/31
            'Modify by Morgan 2009/12/25 延期只更新期限不可點選
            'grdList.Text = "V"
            If textCP10 <> "303" Then
               grdList.Text = "V"
            End If
            '911018 nick 當有勾選第一筆時，將本所期限，法定期限，備註，相關總收文號更新
            '911111 nick 邱小姐說改成若本所期限沒值時，以勾的該筆代 本所期限，法定期限，備註，相關總收文號 到上方
            '            智權人員沒值時，以勾的該筆代智權人員
            'If grdList.Row = 1 Then
             If textCP06.Text = "" Then
                grdList.col = 2
                textCP06 = grdList.Text
                grdList.col = 3
                textCP07 = grdList.Text
                grdList.col = 8
                textCP43 = grdList.Text
                grdList.col = 6
                If textCP10 <> "303" Then textCP64 = textCP64 & grdList.Text 'modify by sonia 2017/4/19 收延期不帶備註
             End If
             If textCP13.Text = "" Then
                grdList.col = 11
                textCP13 = grdList.Text
                '911115 nick
                textCP13_2 = GetStaffName(textCP13)
             End If
            'End If
         
      End If
   End If
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   ' 案件性質必須為延期的才可以選取
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
             'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
             If Pub_CheckNpTheSameShow(m_TM01, textCP10, Trim("" & grdList.TextMatrix(grdList.row, 9))) = False Then
                 Exit Sub
             End If
             'end 2021/08/31
            'Modify by Morgan 2009/12/25 延期只更新期限不可點選
            'grdList.Text = "V"
            If textCP10 <> "303" Then
               grdList.Text = "V"
            End If
            'End 2009/12/25
            '911018 nick 當有勾選第一筆時，將本所期限，法定期限，備註，相關總收文號更新
            '911111 nick 邱小姐說改成若本所期限沒值時，以勾的該筆代 本所期限，法定期限，備註，相關總收文號 到上方
            '            智權人員沒值時，以勾的該筆代智權人員
            'If grdList.Row = 1 Then
             If textCP06.Text = "" Then
                grdList.col = 2
                textCP06 = grdList.Text
                grdList.col = 3
                textCP07 = grdList.Text
                grdList.col = 8
                textCP43 = grdList.Text
                grdList.col = 6
                If textCP10 <> "303" Then textCP64 = textCP64 & grdList.Text 'modify by sonia 2017/4/19 收延期不帶備註
             End If
             If textCP13.Text = "" Then
                grdList.col = 11
                textCP13 = grdList.Text
                '911115 nick
                textCP13_2 = GetStaffName(textCP13)
             End If
            'End If
            
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

'910722 Sieg
Private Function chkNewTMNo(strNo() As String, iChk As Integer) As Boolean
   chkNewTMNo = True
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetMaxNumber(strNo(1), strExc(0)) Then
   If ClsPDGetMaxNumber(strNo(1), strExc(0)) Then
      '2012/2/8 MODIFY BY SONIA 只判斷前二欄
      'If strNo(1) & strNo(2) & strNo(3) & strNo(4) > strNo(1) & String(6 - Len(strExc(0)), "0") & strExc(0) Then
      If strNo(1) & strNo(2) > strNo(1) & String(6 - Len(strExc(0)), "0") & strExc(0) Then
         MsgBox "新本所案號不可大於自動編號，請重新輸入 !", vbCritical
         iChk = 2
         chkNewTMNo = False
      Else
         If MsgBox("此本所案號不存在 ( " & strNo(1) & strNo(2) & strNo(3) & strNo(4) & " ) ，請確認 ?", vbQuestion + vbYesNo) = vbNo Then
            iChk = 1
            chkNewTMNo = False
         End If
      End If
   End If
End Function

' 更新欄位的內容
Private Sub OnUpdateField()
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   
   SetCPFieldNewData "CP01", m_TM01
   SetCPFieldNewData "CP02", m_TM02
   SetCPFieldNewData "CP03", m_TM03
   SetCPFieldNewData "CP04", m_TM04
   
   ' 收文日
   If IsEmptyText(textCP05) = False Then
      SetCPFieldNewData "CP05", DBDATE(textCP05)
   Else
      SetCPFieldNewData "CP05", Empty
   End If
   ' 本所期限
   If IsEmptyText(textCP06) = False Then
      SetCPFieldNewData "CP06", DBDATE(textCP06)
   Else
      SetCPFieldNewData "CP06", Empty
   End If
   ' 法定期限
   If IsEmptyText(textCP07) = False Then
      SetCPFieldNewData "CP07", DBDATE(textCP07)
   Else
      SetCPFieldNewData "CP07", Empty
   End If
   ' 收文號
   'Modify by Morgan 2004/2/18
   '新增才要重抓收文號
    If frm010001.intModifyKind = 0 Then
        m_CP09 = AutoNo("B", 6)
    End If
   SetCPFieldNewData "CP09", m_CP09
   ' 案件性質
   SetCPFieldNewData "CP10", textCP10
   
   '911108 nick
   ' 案件來源
   SetCPFieldNewData "CP11", "90"
   
   ' 業務區
   SetCPFieldNewData "CP12", GetSalesArea(textCP13)
   ' 智權人員
   SetCPFieldNewData "CP13", textCP13
   ' 承辦人員
   SetCPFieldNewData "CP14", textCP14
   ' 費用
   SetCPFieldNewData "CP16", textCP16
   ' 規費
   SetCPFieldNewData "CP17", textCP17
   
   'Add By Cheng 2002/06/12
   Select Case m_TM01
      ' 更新商標基本檔
      Case "T", "TF", "FCT", "CFT":
         '若案件性質為"延展"(102)
         If Me.textCP10.Text = "102" Then
            '若系統日小於等於法定期限
            If Val(ServerDate) <= Val(m_strCP07) Then
               '本所期限及法定期限不可修改
               SetCPFieldNewData "CP06", DBDATE(m_strCP06)
               SetCPFieldNewData "CP07", DBDATE(m_strCP07)
            '若系統日大於法定期限
            Else
               '法定期限為商標基本檔的"專用期止日"+案件國家收費表的"下次管制期限"
                'Modify By Cheng 2003/09/01
'               m_strCP07 = DBDATE(Format(DateSerial(Val(DBYEAR(m_TM22)), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)) + GetCF12(m_TM01, m_TM10, Me.textCP10.Text))))
               m_strCP07 = DBDATE(DateAdd("d", GetCF12(m_TM01, m_TM10, Me.textCP10.Text), ChangeWStringToWDateString(DBDATE(m_TM22))))
               SetCPFieldNewData "CP07", DBDATE(m_strCP07)
               '本所期限 = 法定期限 - 2天
                'Modify By Cheng 2003/09/01
'               m_strCP06 = DBDATE(Format(DateSerial(Val(DBYEAR(m_strCP07)), Val(DBMONTH(m_strCP07)), Val(DBDAY(m_strCP07)) - 2)))
               'Modify By Sindy 2014/10/6 台灣案之本所期限設定
               If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                  m_strCP06 = PUB_GetOurDeadline(DBDATE(m_strCP07))
               Else
               '2014/10/6 END
                  m_strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(m_strCP07))))
               End If
               m_strCP06 = PUB_GetWorkDay1(m_strCP06, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
               SetCPFieldNewData "CP06", DBDATE(m_strCP06)
            End If
            '規費
            'edit by nickc 2006/12/26 修正原先錯誤
            'SetCPFieldNewData "CP07", (Val(GetCF08(m_TM01, m_TM10, Me.textCP10.Text)) * 2)
            SetCPFieldNewData "CP17", (Val(GetCF08(m_TM01, m_TM10, Me.textCP10.Text)) * 2)
         End If
   End Select
   
   '911113 nick
   ' 是否向客戶收款
   SetCPFieldNewData "CP26", "N"
   
   ' 點數
   SetCPFieldNewData "CP18", textCP18
   
    'Add By Cheng 2004/03/16
   ' 是否向客戶收款
   SetCPFieldNewData "CP20", textCP20
   
   '2011/7/22 MODIFY BY SONIA
   '' 是否開電腦收據
   'SetCPFieldNewData "CP32", "N"
   SetCPFieldNewData "CP32", textCP20
   '2011/7/22 END
   
    'End
   ' 相關總收文號
   SetCPFieldNewData "CP43", textCP43
   ' 是否算案件數
   SetCPFieldNewData "CP26", textCP26
   ' 承辦期限
    'Modify By Cheng 2002/10/31
'   SetCPFieldNewData "CP48", textCP48
   SetCPFieldNewData "CP48", DBDATE(textCP48)
   ' 移轉申請人
   '911113 nick
   If textCP10 = "501" Then
      If frm010001.intModifyKind = 0 Then 'Add By Sindy 2009/10/19
        '若移轉人與原申請人不同時
        If ChangeCustomerL(m_CP55) <> ChangeCustomerL(m_TM23) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP55", ChangeCustomerL(m_TM23)
        End If
        '若移轉人與原申請人不同時
        If ChangeCustomerL(m_CP93) <> ChangeCustomerL(m_TM78) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP93", ChangeCustomerL(m_TM78)
        End If
        '若移轉人與原申請人不同時
        If ChangeCustomerL(m_CP94) <> ChangeCustomerL(m_TM79) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP94", ChangeCustomerL(m_TM79)
        End If
        '若移轉人與原申請人不同時
        If ChangeCustomerL(m_CP95) <> ChangeCustomerL(m_TM80) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP95", ChangeCustomerL(m_TM80)
        End If
        '若移轉人與原申請人不同時
        If ChangeCustomerL(m_CP96) <> ChangeCustomerL(m_TM81) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP96", ChangeCustomerL(m_TM81)
        End If
      End If
   End If
   
   If Not IsEmptyText(textCP56) Then
      SetCPFieldNewData "CP56", textCP56 & String(9 - Len(textCP56), "0")
   Else
      SetCPFieldNewData "CP56", Empty
   End If
   'add by nickc 2006/12/01
   If Not IsEmptyText(textCP89) Then
      SetCPFieldNewData "CP89", textCP89 & String(9 - Len(textCP89), "0")
   Else
      SetCPFieldNewData "CP89", Empty
   End If
   If Not IsEmptyText(textCP90) Then
      SetCPFieldNewData "CP90", textCP90 & String(9 - Len(textCP90), "0")
   Else
      SetCPFieldNewData "CP90", Empty
   End If
   If Not IsEmptyText(textCP91) Then
      SetCPFieldNewData "CP91", textCP91 & String(9 - Len(textCP91), "0")
   Else
      SetCPFieldNewData "CP91", Empty
   End If
   If Not IsEmptyText(textCP92) Then
      SetCPFieldNewData "CP92", textCP92 & String(9 - Len(textCP92), "0")
   Else
      SetCPFieldNewData "CP92", Empty
   End If
   
   'Add By Sindy 2019/7/11
   ' 是否電子送件
   If chkWebApp.Visible = True Then
      If chkWebApp.Value = 1 Then
         SetCPFieldNewData "CP118", "Y"
      Else
         SetCPFieldNewData "CP118", Empty
      End If
   Else
      SetCPFieldNewData "CP118", Empty
   End If
   '2019/7/11 END
   
   ' 進度備註
   SetCPFieldNewData "CP64", textCP64
   ' 卷宗性質為非申請時, 更新案件進度檔的對造案件名稱
   If textTM28 <> "1" Then
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            ' 對造案件名稱
            SetCPFieldNewData "CP37", textTM05_1
        Case Else
            ' 對造案件名稱(中)
            SetCPFieldNewData "CP37", textTM05
            ' 對造案件名稱(英)
            SetCPFieldNewData "CP38", textTM06
            ' 對造案件名稱(日)
            SetCPFieldNewData "CP39", textTM07
        End Select
   End If
   
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "T", "TF", "CFT", "FCT":
         ' 卷宗性質為非申請時, 不更新基本檔
         If textTM28 = "1" Then
'            ' 案件中文名稱
'            SetTMSPFieldNewData "TM05", textTM05
            ' 案件名稱
            SetTMSPFieldNewData "TM05", textTM05_1
'            ' 案件英文名稱
'            SetTMSPFieldNewData "TM06", textTM06
'            ' 案件日文名稱
'            SetTMSPFieldNewData "TM07", textTM07
         End If
         ' 商標種類
         SetTMSPFieldNewData "TM08", textTM08
         ' 商品類別
         SetTMSPFieldNewData "TM09", textTM09
         ' 申請國家
         SetTMSPFieldNewData "TM10", textTM10
         ' 申請人
         If Not IsEmptyText(textTM23) Then
            SetTMSPFieldNewData "TM23", textTM23 & String(9 - Len(textTM23), "0")
         Else
            SetTMSPFieldNewData "TM23", Empty
         End If
         ' 申請地址(中)
         SetTMSPFieldNewData "TM24", textTM24
         ' 申請地址(英)
         SetTMSPFieldNewData "TM25", textTM25
         ' 申請地址(日)
         SetTMSPFieldNewData "TM26", textTM26
         'add by nickc 2006/12/01
         SetTMSPFieldNewData "TM78", textSP58
         SetTMSPFieldNewData "TM79", textSP59
         SetTMSPFieldNewData "TM80", textTM80
         SetTMSPFieldNewData "TM81", textTM81
         SetTMSPFieldNewData "TM82", textTM82
         SetTMSPFieldNewData "TM83", textTM83
         SetTMSPFieldNewData "TM84", textTM84
         SetTMSPFieldNewData "TM85", textTM85
         SetTMSPFieldNewData "TM86", textTM86
         SetTMSPFieldNewData "TM87", textTM87
         SetTMSPFieldNewData "TM88", textTM88
         SetTMSPFieldNewData "TM89", textTM89
         SetTMSPFieldNewData "TM90", textTM90
         SetTMSPFieldNewData "TM91", textTM91
         SetTMSPFieldNewData "TM92", textTM92
         SetTMSPFieldNewData "TM93", textTM93
         SetTMSPFieldNewData "TM32", textTM32
         
         
         
         ' 卷宗性質
         SetTMSPFieldNewData "TM28", textTM28
         ' 分所案號
         SetTMSPFieldNewData "TM34", textTM34
         ' 客戶案件案號
         SetTMSPFieldNewData "TM35", textTM35
         ' FC代理人
         If IsEmptyText(textTM44) = False Then
            SetTMSPFieldNewData "TM44", textTM44 & String(9 - Len(textTM44), "0")
         Else
            SetTMSPFieldNewData "TM44", textTM44
         End If
         ' 彼所案號
         SetTMSPFieldNewData "TM45", textTM45
         ' 案件備註
         SetTMSPFieldNewData "TM58", textTM58
      Case Else:
         ' 卷宗性質為非申請時, 不更新基本檔
         If textTM28 = "1" Then
            ' 案件中文名稱
            SetTMSPFieldNewData "SP05", textTM05
            ' 案件英文名稱
            SetTMSPFieldNewData "SP06", textTM06
            ' 案件日文名稱
            SetTMSPFieldNewData "SP07", textTM07
         End If
         ' 申請人
         SetTMSPFieldNewData "SP08", textTM23
'edit by nickc 2007/02/13
'         If m_TM01 = "CFC" Then
            ' 申請人2
            If Not IsEmptyText(textSP58) Then
               SetTMSPFieldNewData "SP58", textSP58 & String(9 - Len(textSP58), "0")
            Else
               SetTMSPFieldNewData "SP58", Empty
            End If
            ' 申請人3
            If Not IsEmptyText(textSP59) Then
               SetTMSPFieldNewData "SP59", textSP59 & String(9 - Len(textSP59), "0")
            Else
               SetTMSPFieldNewData "SP59", Empty
            End If
'         End If
'edit by nickc 2007/02/13
            ' 申請人4
            If Not IsEmptyText(textTM80) Then
               SetTMSPFieldNewData "SP65", textTM80 & String(9 - Len(textTM80), "0")
            Else
               SetTMSPFieldNewData "SP65", Empty
            End If
            ' 申請人5
            If Not IsEmptyText(textTM81) Then
               SetTMSPFieldNewData "SP66", textTM81 & String(9 - Len(textTM81), "0")
            Else
               SetTMSPFieldNewData "SP66", Empty
            End If
         ' 申請國家
         SetTMSPFieldNewData "SP09", m_TM10
         ' FC代理人
         If IsEmptyText(textTM44) = False Then
            SetTMSPFieldNewData "SP26", textTM44 & String(9 - Len(textTM44), "0")
         Else
            SetTMSPFieldNewData "SP26", textTM44
         End If
         ' 彼所案號
         SetTMSPFieldNewData "SP27", textTM45
         ' 案件備註
         SetTMSPFieldNewData "SP18", textTM58
         'add by nickc 2006/12/04
         SetTMSPFieldNewData "SP73", textTM09
         SetTMSPFieldNewData "SP74", textTM32
   End Select
End Sub

Private Function NextRecord() As Boolean
   Dim nIndex As Integer
   NextRecord = False
   
   For nIndex = 0 To m_CPKeyCount - 1
      If m_CP09 = m_CPKeyList(nIndex) Then
         If nIndex < m_CPKeyCount - 1 Then
            m_CP09 = m_CPKeyList(nIndex + 1)
            NextRecord = True
            Exit For
         End If
      End If
   Next nIndex
End Function

' 更新商標基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateTradeMark()
Private Function OnUpdateTradeMark() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
                      
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnUpdateTradeMark = True
   
   'Modify By Sindy 2009/10/23
'   '910702 Sieg 先檢查是否有修改申請人1，參照 501
'   Dim strTmp1(1 To 3) As String
'   If textTM23 <> "" Then
'      'edit by nickc 2007/02/02 不用 dll 了
'      'If objPublicData.GetCustomerNameAndAddress(textTM23, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'      If ClsPDGetCustomerNameAndAddress(textTM23, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'         '修改申請人時
'         If InStr(ChangeCustomerL(m_TM23), ChangeCustomerL(textTM23)) = 0 Then
'            If m_CP60 <> "" Then
'               strExc(1) = m_TM01
'               strExc(2) = m_TM02
'               strExc(3) = m_TM03
'               strExc(4) = m_TM04
'               strExc(5) = m_CP60
'               strExc(6) = textTM23
'               strExc(7) = strExc(0)
'               'edit by nickc 2007/02/05 不用 dll 了
'               'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
'               If Not ClsLawUpdAcc0k0(strExc(), True) Then
'                  textTM23.SetFocus
'                  Exit Function
'               End If
'            End If
'            SetTMSPFieldNewData "TM24", strTmp1(1)
'            SetTMSPFieldNewData "TM25", strTmp1(2)
'            SetTMSPFieldNewData "TM26", strTmp1(3)
'            SetTMSPFieldNewData "TM47", ""
'            SetTMSPFieldNewData "TM48", ""
'            SetTMSPFieldNewData "TM49", ""
'            SetTMSPFieldNewData "TM50", ""
'            SetTMSPFieldNewData "TM51", ""
'            SetTMSPFieldNewData "TM52", ""
'         End If
'      End If
'   Else
'      SetTMSPFieldNewData "TM24", ""
'      SetTMSPFieldNewData "TM25", ""
'      SetTMSPFieldNewData "TM26", ""
'      SetTMSPFieldNewData "TM47", ""
'      SetTMSPFieldNewData "TM48", ""
'      SetTMSPFieldNewData "TM49", ""
'      SetTMSPFieldNewData "TM50", ""
'      SetTMSPFieldNewData "TM51", ""
'      SetTMSPFieldNewData "TM52", ""
'   End If
'   'add by nickc 2006/12/01
'    If textSP58 <> "" Then
'      'edit by nickc 2007/02/02 不用 dll 了
'      'If objPublicData.GetCustomerNameAndAddress(textSP58, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'      If ClsPDGetCustomerNameAndAddress(textSP58, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'         '修改申請人時
'         If InStr(ChangeCustomerL(m_strCust2), ChangeCustomerL(textSP58)) = 0 Then
'            If m_CP60 <> "" Then
'               strExc(1) = m_TM01
'               strExc(2) = m_TM02
'               strExc(3) = m_TM03
'               strExc(4) = m_TM04
'               strExc(5) = m_CP60
'               strExc(6) = textSP58
'               strExc(7) = strExc(0)
'               'edit by nickc 2007/02/05 不用 dll 了
'               'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
'               If Not ClsLawUpdAcc0k0(strExc(), True) Then
'                  textSP58.SetFocus
'                  Exit Function
'               End If
'            End If
'            SetTMSPFieldNewData "TM82", strTmp1(1)
'            SetTMSPFieldNewData "TM86", strTmp1(2)
'            SetTMSPFieldNewData "TM90", strTmp1(3)
'            SetTMSPFieldNewData "TM94", ""
'            SetTMSPFieldNewData "TM95", ""
'            SetTMSPFieldNewData "TM96", ""
'            SetTMSPFieldNewData "TM97", ""
'            SetTMSPFieldNewData "TM98", ""
'            SetTMSPFieldNewData "TM99", ""
'         End If
'      End If
'   Else
'      SetTMSPFieldNewData "TM82", ""
'      SetTMSPFieldNewData "TM86", ""
'      SetTMSPFieldNewData "TM90", ""
'      SetTMSPFieldNewData "TM94", ""
'      SetTMSPFieldNewData "TM95", ""
'      SetTMSPFieldNewData "TM96", ""
'      SetTMSPFieldNewData "TM97", ""
'      SetTMSPFieldNewData "TM98", ""
'      SetTMSPFieldNewData "TM99", ""
'   End If
'    If textSP59 <> "" Then
'      'edit by nickc 2007/02/02 不用 dll 了
'      'If objPublicData.GetCustomerNameAndAddress(textSP59, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'      If ClsPDGetCustomerNameAndAddress(textSP59, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'         '修改申請人時
'         If InStr(ChangeCustomerL(m_strCust3), ChangeCustomerL(textSP59)) = 0 Then
'            If m_CP60 <> "" Then
'               strExc(1) = m_TM01
'               strExc(2) = m_TM02
'               strExc(3) = m_TM03
'               strExc(4) = m_TM04
'               strExc(5) = m_CP60
'               strExc(6) = textSP59
'               strExc(7) = strExc(0)
'               'edit by nickc 2007/02/05 不用 dll 了
'               'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
'               If Not ClsLawUpdAcc0k0(strExc(), True) Then
'                  textSP59.SetFocus
'                  Exit Function
'               End If
'            End If
'            SetTMSPFieldNewData "TM83", strTmp1(1)
'            SetTMSPFieldNewData "TM87", strTmp1(2)
'            SetTMSPFieldNewData "TM91", strTmp1(3)
'            SetTMSPFieldNewData "TM100", ""
'            SetTMSPFieldNewData "TM101", ""
'            SetTMSPFieldNewData "TM102", ""
'            SetTMSPFieldNewData "TM103", ""
'            SetTMSPFieldNewData "TM104", ""
'            SetTMSPFieldNewData "TM105", ""
'         End If
'      End If
'   Else
'      SetTMSPFieldNewData "TM83", ""
'      SetTMSPFieldNewData "TM87", ""
'      SetTMSPFieldNewData "TM91", ""
'      SetTMSPFieldNewData "TM100", ""
'      SetTMSPFieldNewData "TM101", ""
'      SetTMSPFieldNewData "TM102", ""
'      SetTMSPFieldNewData "TM103", ""
'      SetTMSPFieldNewData "TM104", ""
'      SetTMSPFieldNewData "TM105", ""
'   End If
'    If textTM80 <> "" Then
'      'edit by nickc 2007/02/02 不用 dll 了
'      'If objPublicData.GetCustomerNameAndAddress(textTM80, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'      If ClsPDGetCustomerNameAndAddress(textTM80, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'         '修改申請人時
'         If InStr(ChangeCustomerL(m_strCust4), ChangeCustomerL(textTM80)) = 0 Then
'            If m_CP60 <> "" Then
'               strExc(1) = m_TM01
'               strExc(2) = m_TM02
'               strExc(3) = m_TM03
'               strExc(4) = m_TM04
'               strExc(5) = m_CP60
'               strExc(6) = textTM80
'               strExc(7) = strExc(0)
'               'edit by nickc 2007/02/05 不用 dll 了
'               'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
'               If Not ClsLawUpdAcc0k0(strExc(), True) Then
'                  textTM80.SetFocus
'                  Exit Function
'               End If
'            End If
'            SetTMSPFieldNewData "TM84", strTmp1(1)
'            SetTMSPFieldNewData "TM88", strTmp1(2)
'            SetTMSPFieldNewData "TM92", strTmp1(3)
'            SetTMSPFieldNewData "TM106", ""
'            SetTMSPFieldNewData "TM107", ""
'            SetTMSPFieldNewData "TM108", ""
'            SetTMSPFieldNewData "TM109", ""
'            SetTMSPFieldNewData "TM110", ""
'            SetTMSPFieldNewData "TM111", ""
'         End If
'      End If
'   Else
'      SetTMSPFieldNewData "TM84", ""
'      SetTMSPFieldNewData "TM88", ""
'      SetTMSPFieldNewData "TM92", ""
'      SetTMSPFieldNewData "TM106", ""
'      SetTMSPFieldNewData "TM107", ""
'      SetTMSPFieldNewData "TM108", ""
'      SetTMSPFieldNewData "TM109", ""
'      SetTMSPFieldNewData "TM110", ""
'      SetTMSPFieldNewData "TM111", ""
'   End If
'    If textTM81 <> "" Then
'      'edit by nickc 2007/02/02 不用 dll 了
'      'If objPublicData.GetCustomerNameAndAddress(textTM81, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'      If ClsPDGetCustomerNameAndAddress(textTM81, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
'         '修改申請人時
'         If InStr(ChangeCustomerL(m_strCust5), ChangeCustomerL(textTM81)) = 0 Then
'            If m_CP60 <> "" Then
'               strExc(1) = m_TM01
'               strExc(2) = m_TM02
'               strExc(3) = m_TM03
'               strExc(4) = m_TM04
'               strExc(5) = m_CP60
'               strExc(6) = textTM81
'               strExc(7) = strExc(0)
'               'edit by nickc 2007/02/05 不用 dll 了
'               'If Not objLawDll.UpdAcc0k0(strExc(), True) Then
'               If Not ClsLawUpdAcc0k0(strExc(), True) Then
'                  textTM81.SetFocus
'                  Exit Function
'               End If
'            End If
'            SetTMSPFieldNewData "TM85", strTmp1(1)
'            SetTMSPFieldNewData "TM89", strTmp1(2)
'            SetTMSPFieldNewData "TM93", strTmp1(3)
'            SetTMSPFieldNewData "TM112", ""
'            SetTMSPFieldNewData "TM113", ""
'            SetTMSPFieldNewData "TM114", ""
'            SetTMSPFieldNewData "TM115", ""
'            SetTMSPFieldNewData "TM116", ""
'            SetTMSPFieldNewData "TM117", ""
'         End If
'      End If
'   Else
'      SetTMSPFieldNewData "TM85", ""
'      SetTMSPFieldNewData "TM89", ""
'      SetTMSPFieldNewData "TM93", ""
'      SetTMSPFieldNewData "TM112", ""
'      SetTMSPFieldNewData "TM113", ""
'      SetTMSPFieldNewData "TM114", ""
'      SetTMSPFieldNewData "TM115", ""
'      SetTMSPFieldNewData "TM116", ""
'      SetTMSPFieldNewData "TM117", ""
'   End If
   '2009/10/23 End
   
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & "NULL"
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateTradeMark = False
End Function

' 更新服務業務基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateServicePractice()
Private Function OnUpdateServicePractice() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnUpdateServicePractice = True
      
   ' 更新案件進度檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & "NULL"
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateServicePractice = False
End Function

' 新增案件進度檔
'Modify By Cheng 2002/11/06
'Private Sub SaveNewCaseProgress()
Private Function SaveNewCaseProgress() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer

'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
SaveNewCaseProgress = True

   strSql = "INSERT INTO CaseProgress ("
   For nIndex = 0 To m_CPCount - 1
      If Not IsEmptyText(m_CPList(nIndex).fiNewData) Then
         If nIndex <> 0 Then strSql = strSql & ","
         strSql = strSql & m_CPList(nIndex).fiName
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   For nIndex = 0 To m_CPCount - 1
      If Not IsEmptyText(m_CPList(nIndex).fiNewData) Then
         If nIndex <> 0 Then strSql = strSql & ","
         If m_CPList(nIndex).fiType = 0 Then
            '911028 nick 加 chgsql
            'strSQL = strSQL & "'" & m_CPList(nIndex).fiNewData & "'"
            strSql = strSql & "'" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            strSql = strSql & m_CPList(nIndex).fiNewData
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   
   cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    SaveNewCaseProgress = False
End Function

' 更新案件進度檔
Private Sub OnUpdateCaseProgress()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & "NULL"
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
               strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
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
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

' 檢查服務業務基本檔是否存在該筆本所案號
Private Function ExistServicePractice(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim bExist As Boolean
   
   bExist = False
   strSP03 = strSP03 & String(1 - Len(strSP03), "0")
   strSP04 = strSP04 & String(2 - Len(strSP04), "0")
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & strSP01 & "' AND " & _
                  "SP02 = '" & strSP02 & "' AND " & _
                  "SP03 = '" & strSP03 & "' AND " & _
                  "SP04 = '" & strSP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      bExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Modify By Cheng 2002/11/06
'Private Function OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strCF13 As String
   Dim strCF14 As String
   Dim strDay As String
   Dim strDate As String
   Dim strNP07 As String
   Dim strNP22 As String
   Dim strTemp As String
   Dim nIndex As Integer
   Dim nSubIndex As Integer
   Dim strCountry As String
   Dim strProduct As String
   Dim objCopyTM As ClsCopyTM
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔
   'OnUpdateCaseProgress
    'Modify By Cheng 2002/11/06
'   SaveNewCaseProgress
    'Modify By Cheng 2003/03/28
    '若為新增
    If frm010001.intModifyKind = 0 Then
        If SaveNewCaseProgress = False Then GoTo ErrorHandler
    '若為修改
    ElseIf frm010001.intModifyKind = 1 Then
        OnUpdateCaseProgress
    End If
   'Added by Lydia 2023/03/27 FCT,S案請控制內部收文739更換智權人員存檔時，同時上發文日=系統日。
   If (m_TM01 = "FCT" Or m_TM01 = "S") And textCP10 = "739" Then
       strSql = "Update CaseProgress Set CP27=" & strSrvDate(1) & " Where CP09='" & m_CP09 & "' and cp27 is null"
       cnnConnection.Execute strSql
       'Added by Lydia 2023/04/06 同時更換下一程序資料維護中「是否續辦」欄為空白之「智權人員」
       strSql = "update nextprogress set np10='" & textCP13 & "', np15=sqldatet(to_char(sysdate,'yyyymmdd'))||'更換智權人員,原為'||np10||getstaffnamelist(np10) " & _
                   "where np02='" & m_TM01 & "' and np03='" & m_TM02 & "'  and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np06 is null "
       cnnConnection.Execute strSql
       'end 2023/04/06
   End If
   'end 203/03/27
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Select Case m_TM01
      ' 更新商標基本檔
      Case "T", "TF", "CFT", "FCT":
        'Modify By Cheng 2002/11/06
'         OnUpdateTradeMark
         If OnUpdateTradeMark = False Then GoTo ErrorHandler
      ' 更新服務業務基本檔
      Case Else:
        'Modify By Cheng 2002/11/06
'         OnUpdateServicePractice
         If OnUpdateServicePractice = False Then GoTo ErrorHandler
   End Select

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 儲存優先權資料
    'Modify By Cheng 2002/11/06
'   objPublicData.SavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.SavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)) = False Then GoTo ErrorHandler
   'Modify by Amy 2014/04/18 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 +, m_Priority(6)
   If ClsPDSavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)) = False Then GoTo ErrorHandler
   
   ' 機關文號
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         If Not IsEmptyText(grdList.TextMatrix(nIndex, 4)) Then
            strSql = "UPDATE CASEPROGRESS SET CP08 = '" & grdList.TextMatrix(nIndex, 4) & "' " & _
                     "WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql
         End If
         Exit For
      End If
   Next nIndex

   ' 對造案件名稱
   If textTM28 <> "1" Then
      '911028 nick 加chgsql
      'strSQL = "UPDATE CASEPROGRESS SET CP37 = '" & textTM05 & "', " & _
                                       "CP38 = '" & textTM06 & "', " & _
                                       "CP39 = '" & textTM07 & "' " & _
                "WHERE CP09 = '" & m_CP09 & "' "
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            strSql = "UPDATE CASEPROGRESS SET CP37 = '" & ChgSQL(textTM05_1) & "' " & _
                      "WHERE CP09 = '" & m_CP09 & "' "
        Case Else
            strSql = "UPDATE CASEPROGRESS SET CP37 = '" & ChgSQL(textTM05) & "', " & _
                                             "CP38 = '" & ChgSQL(textTM06) & "', " & _
                                             "CP39 = '" & ChgSQL(textTM07) & "' " & _
                      "WHERE CP09 = '" & m_CP09 & "' "
        End Select
      cnnConnection.Execute strSql
   End If
   
   ' 更新基本檔是否閉卷, 閉卷日期, 閉卷原因
   If textTM29 = "Y" Then
      Select Case m_TM01
      ' 更新商標基本檔
         Case "T", "TF", "CFT", "FCT":
            strSql = "UPDATE TRADEMARK SET TM29=NULL, TM30=NULL,TM31=NULL " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "TM04 = '" & m_TM04 & "' "
         Case Else:
            strSql = "UPDATE SERVICEPRACTICE SET SP15=NULL, SP16=NULL,SP17=NULL " & _
                     "WHERE SP01 = '" & m_TM01 & "' AND " & _
                           "SP02 = '" & m_TM02 & "' AND " & _
                           "SP03 = '" & m_TM03 & "' AND " & _
                           "SP04 = '" & m_TM04 & "' "
      End Select
      cnnConnection.Execute strSql
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '91.11.10 cancel by sonia
   ' 若有修改案件性質時,
   'If textCP10 <> m_CP10 Then
   '   strCF13 = "0"
   '   strCF14 = "0"
   '   strSQL = "SELECT * FROM CaseFee " & _
   '            "WHERE CF01 = '" & m_TM01 & "' AND " & _
   '                  "CF02 = '" & textTM10 & "' AND " & _
   '                  "CF03 = '" & textCP10 & "' "
   '   rsTmp.CursorLocation = adUseClient
   '   rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
   '   If rsTmp.RecordCount > 0 Then
   '      rsTmp.MoveFirst
   '      If IsNull(rsTmp.Fields("CF13")) = False Then
   '         strCF13 = rsTmp.Fields("CF13")
   '      End If
   '      If IsNull(rsTmp.Fields("CF14")) = False Then
   '         strCF14 = rsTmp.Fields("CF14")
   '      End If
   '   End If
   '   rsTmp.Close
      ' 更新案件進度檔的標準價及底價欄位
   '   strSQL = "UPDATE CaseProgress SET CP33 = " & strCF13 & ", " & _
   '                                    "CP34 = " & strCF14 & " " & _
   '            "WHERE CP09 = '" & m_CP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '91.11.10 end
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 案件性質為查名, 補收款, 後金時更新發文日為系統日
   Select Case textCP10
      'edit by nickc 2008/01/04 加入告代直接上發文日
      'Case "01", "705", "909":
      Case "01", "705", "909", "719":
         strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(SystemDate()) & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         cnnConnection.Execute strSql
   End Select
   
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入查名收文號時, 更新此該查名收文號的案件進度資料的本所案號為本案的本所案號
   'If IsEmptyText(textCP09_S) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP01 = '" & m_TM01 & "', " & _
   '                                    "CP02 = '" & m_TM02 & "', " & _
   '                                    "CP03 = '" & m_TM03 & "', " & _
   '                                    "CP04 = '" & m_TM04 & "' " & _
   '            "WHERE CP09 = '" & textCP09_S & "' "
   '   cnnConnection.Execute strSQL
   'End If
   
   ' 若有輸入查名本所案號時
   If textCP01_S.Text = "S" And IsEmptyText(textCP02_S) = False Then
      Dim strCP01 As String
      Dim strCP02 As String
      Dim strCP03 As String
      Dim strCP04 As String
      strCP01 = textCP01_S
      strCP02 = textCP02_S
      strCP03 = textCP03_S & String(1 - Len(textCP03_S), "0")
      strCP04 = textCP04_S & String(2 - Len(textCP04_S), "0")
      strSql = "UPDATE CaseProgress SET CP01 = '" & m_TM01 & "', " & _
                                       "CP02 = '" & m_TM02 & "', " & _
                                       "CP03 = '" & m_TM03 & "', " & _
                                       "CP04 = '" & m_TM04 & "' " & _
               "WHERE CP01 = '" & strCP01 & "' AND " & _
                     "CP02 = '" & strCP02 & "' AND " & _
                     "CP03 = '" & strCP03 & "' AND " & _
                     "CP04 = '" & strCP04 & "' "
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若案件性質為救濟程序時或爭議程序更新基本檔的欄位
   Select Case Mid(textCP10, 1, 1)
      ' 救濟程序
      Case "4":
         Select Case m_TM01:
            'Modified by Lydia 2017/03/15 外商只用在FCT,CFT
            'Case "T", "TF", "FCT":
            Case "FCT", "CFT":
               strSql = "UPDATE TradeMark SET TM18 = 'Y' " & _
                        "WHERE TM01 = '" & m_TM01 & "' AND " & _
                              "TM02 = '" & m_TM02 & "' AND " & _
                              "TM03 = '" & m_TM03 & "' AND " & _
                              "TM04 = '" & m_TM04 & "' "
               cnnConnection.Execute strSql
            Case Else:
         End Select
      ' 爭議程序
      Case "6":
         Select Case m_TM01:
            'Modified by Lydia 2017/03/15 外商只用在FCT,CFT
            'Case "T", "TF", "FCT":
            Case "FCT", "CFT":
               strSql = "UPDATE TradeMark SET TM19 = 'Y' " & _
                        "WHERE TM01 = '" & m_TM01 & "' AND " & _
                              "TM02 = '" & m_TM02 & "' AND " & _
                              "TM03 = '" & m_TM03 & "' AND " & _
                              "TM04 = '" & m_TM04 & "' "
               cnnConnection.Execute strSql
            Case Else:
         End Select
   End Select

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strNP07 = grdList.TextMatrix(nIndex, 9)
         strNP22 = grdList.TextMatrix(nIndex, 10)
         'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(m_CP09) &
         strSql = "UPDATE NextProgress SET NP06 = 'Y',np24=" & CNULL(m_CP09) & _
                  " WHERE NP02 = '" & m_TM01 & "' AND " & _
                        "NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND " & _
                        "NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         Pub_SeekTbLog strSql 'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業，若畫面勾選下一程序期限且存檔有上續辦Y的都寫Log以便事後能追蹤
         cnnConnection.Execute strSql
      End If
   Next nIndex

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
      '911018 nick 當有相關總收文號時，要將總收文號該筆更新成續辦，因為只會有一筆時才會讀出來秀畫面，所以不用np22
   '91.11.10 MODIFY BY SONIA
   'If textCP43 <> "" Then
   '     strSQL = "update nextprogress set np06='Y' where np01='" & textCP43 & "' "
   '     cnnConnection.Execute strSQL
   'End If
   '91.11.10 END

   '2005/12/26 CANCEL BY SONIA 外商分案於92/3/13即刪除
   ' 若此案為母案且申請國家為歐盟, 美國或新加坡時
'   If m_TM03 = "0" And m_TM04 = "00" And (textTM10 = "239" Or textTM10 = "101" Or textTM10 = "014") Then
'      If textTM10 = "239" And textCP10 = "101" And IsEmptyText(m_strCountry) = False Then
'         If IsEmptyText(textTM09) = False Then
'            For nIndex = 1 To GetSubStringCount(textTM09)
'               strProduct = GetSubString(textTM09, nIndex)
'               For nSubIndex = 1 To GetSubStringCount(m_strCountry)
'                  strCountry = GetSubString(m_strCountry, nSubIndex)
'                  Set objCopyTM = New ClsCopyTM
'                  objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
'                  objCopyTM.SetDes m_TM01, m_TM02, CStr(Val(m_TM03 + nIndex)), Format(CStr(Val(m_TM04) + nSubIndex), "00")
'                  objCopyTM.SetExtraField "TM09", strProduct
'                  objCopyTM.SetExtraField "TM10", strCountry
'                  objCopyTM.CopyTradeMark
'                  Set objCopyTM = Nothing
'               Next nSubIndex
'            Next nIndex
'         Else
'            For nSubIndex = 1 To GetSubStringCount(m_strCountry)
'               strCountry = GetSubString(m_strCountry, nSubIndex)
'               Set objCopyTM = New ClsCopyTM
'               objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
'               objCopyTM.SetDes m_TM01, m_TM02, m_TM03, Format(CStr(Val(m_TM04) + nSubIndex), "00")
'               objCopyTM.SetExtraField "TM10", strCountry
'               objCopyTM.CopyTradeMark
'               Set objCopyTM = Nothing
'            Next nSubIndex
'         End If
'      Else
'         For nIndex = 1 To GetSubStringCount(textTM09)
'            strProduct = GetSubString(textTM09, nIndex)
'            Set objCopyTM = New ClsCopyTM
'            objCopyTM.SetSrc m_TM01, m_TM02, m_TM03, m_TM04
'            objCopyTM.SetDes m_TM01, m_TM02, CStr(Val(m_TM03 + nIndex)), m_TM04
'            objCopyTM.SetExtraField "TM09", strProduct
'            objCopyTM.CopyTradeMark
'            Set objCopyTM = Nothing
'         Next nIndex
'      End If
'   End If
'2005/12/26 END
   
   'Add by Sindy 2019/5/27
   Call PUB_TMFilePathToCPP(strTMCppFilePath, m_CP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm010001", IIf(Left(Pub_StrUserSt03, 1) = "F", m_CP09, "")
   End If
   '2019/5/27 END
   
   'Add By Cheng 2002/11/06
   cnnConnection.CommitTrans
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nVal As Currency
Dim ii As Integer
   
   CheckDataValid = False
   ' 案件性質不可為空白
   If IsEmptyText(textCP10) = True Then
      strTit = "檢核資料"
      strMsg = "案件性質不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP10.SetFocus
      GoTo EXITSUB
   End If
   ' 申請國家不可空白
   If IsEmptyText(textTM10) = True Then
      strTit = "檢核資料"
      strMsg = "申請國家不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM10.SetFocus
      GoTo EXITSUB
   End If
   ' 案件名稱不可同時為空白
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        If IsEmptyText(textTM05_1) = True Then
           strTit = "檢核資料"
           strMsg = "案件名稱不可空白"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM05_1.SetFocus
           GoTo EXITSUB
        End If
    Case Else
        If IsEmptyText(textTM05) = True And IsEmptyText(textTM06) = True And IsEmptyText(textTM07) = True Then
           strTit = "檢核資料"
           strMsg = "案件名稱(中)(英)(日)不可全為空白"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM05.SetFocus
           GoTo EXITSUB
        End If
    End Select
   ' 承辦期限不可超過本所期限
   If IsEmptyText(textCP06) = False And IsEmptyText(textCP48) = False Then
      If Val(textCP48) > Val(textCP06) Then
         strTit = "檢核資料"
         strMsg = "承辦期限不可超過本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 案件性質為延展或延期時本所期限及法定期限不可為空白
   If textCP10 = "102" Or textCP10 = "303" Then
      If IsEmptyText(textCP06) = True Then
         strTit = "檢核資料"
        'Modify By Cheng 2002/10/30
'         strMsg = "案件性質為延展, 本所期限不可為空白"
         strMsg = "案件性質為延展或延期時, 本所期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(textCP07) = True Then
         strTit = "檢核資料"
        'Modify By Cheng 2002/10/30
'         strMsg = "案件性質為延展, 法定期限不可為空白"
         strMsg = "案件性質為延展或延期時, 法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07.SetFocus
         GoTo EXITSUB
      End If
   End If
   If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
      If Val(textCP06) > Val(textCP07) Then
         strTit = "檢核資料"
         strMsg = "本所期限與法定期限範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 收文日
   If IsEmptyText(textCP05) = True Then
      strTit = "檢核資料"
      strMsg = "收文日不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP05.SetFocus
      GoTo EXITSUB
   End If
   ' 點數=(費用-規費) / 1000
   If IsEmptyText(textCP16) = False Then '有費用
      nVal = (Val(textCP16) - Val(textCP17)) / 1000
      If textCP18 <> CStr(nVal) Then
         strTit = "檢核資料"
         strMsg = "點數應為 " & CStr(nVal)
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP16.SetFocus
         GoTo EXITSUB
      End If
      'add by nickc 2008/01/09
      If textCP10 = "719" Or textCP10 = "720" Then
           textCP20 = ""
      End If
   Else '無費用
      nVal = 0
      If IsEmptyText(textCP18) = False Then
         If textCP18 <> "0" Then
            strTit = "檢核資料"
            strMsg = "點數應為空白或0"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP18.SetFocus
            GoTo EXITSUB
         End If
      End If
      'add by nickc 2008/01/09
      If textCP10 = "719" Or textCP10 = "720" Then
            textCP20 = "N"
      End If
   End If
   
   'Add By Sindy 2010/02/24
   If Val(textCP16) = 0 Then
      If m_TM01 = "FCT" Then
         '2010/6/11 MODIFY BY SONIA 阿蓮要求加612補充理由也預設不請款
         If textCP10 = "201" Or textCP10 = "302" Or textCP10 = "303" Or textCP10 = "612" Then
            textCP20 = "N"
         End If
      End If
   End If
   
   'Add By Sindy 2016/6/30 若有輸入費用時,是否向客戶收款欄不可為N
   If Val(textCP16) > 0 And textCP20 = "N" Then
      strTit = "檢核資料"
      strMsg = "是否向客戶收款欄不可為N"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP20.SetFocus
      GoTo EXITSUB
   End If
   '2016/6/30 END
   
   ' 申請人
'edit by nickc 2007/02/13
'   If m_TM01 = "CFC" Then
'add by nickc 2008/01/03 S 不用檢查
'2008/2/12 cancel by sonia 下面已有檢查申請人和代理人不可同時空白,此處不必檢查
'   If m_TM01 <> "S" Then
'      'If IsEmptyText(textTM23) = True And IsEmptyText(textSP58) = True And IsEmptyText(textSP59) = True Then
'      If IsEmptyText(textTM23) = True And IsEmptyText(textSP58) = True And IsEmptyText(textSP59) = True And IsEmptyText(textTM80) = True And IsEmptyText(textTM81) = True Then
'         strTit = "檢核資料"
'         strMsg = "申請人不可全為空白"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textTM23.SetFocus
'         GoTo EXITSUB
'      End If
'   End If
'2008/2/12 end
   
   'Add By Cheng 2002/07/11
   '若案件性質為"自請撤回"(306)或"自請撤銷"(307)時, 第二頁的"相關總收文號"欄不可空白
   If Me.textCP10.Text = "306" Or Me.textCP10.Text = "307" Then
      '相關總收文號不可為空白
      If IsEmptyText(Me.textCP43.Text) = True Then
         strTit = "檢核資料"
         strMsg = "相關總收文號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 1
         Me.textCP43.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 查名本所案號
   If Not IsEmptyText(textCP01_S) Then
      If Not IsEmptyText(textCP01_S) Then
         If Not ExistServicePractice(textCP01_S, textCP02_S, textCP03_S, textCP04_S) Then
            strTit = "檢核資料"
            strMsg = "查名本所案號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Me.SSTab1.Tab = 1
            Me.textCP01_S.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   ' 智權人員 ADD BY SONIA 91.11.3
   If IsEmptyText(textCP13) = True Then
      strTit = "檢核資料"
      strMsg = "智權人員不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP13.SetFocus
     
      GoTo EXITSUB
   End If
   '91.11.3 END
   
   '911113 nick 當申請國家是 '000' 時，申請人1 和代理人不可空白，因為申請人不能輸入，所以控制在代理人
   '                     不是                  不可空白，但是因為申請人不可輸入，所以使用者必須結束出去補申請人在進來
   If textTM10 = "000" Then
      If m_TM01 <> "S" Then   '2008/2/12 add by sonia  S案件不用檢查
        If IsEmptyText(textTM23) And IsEmptyText(textTM44) Then
            strTit = "檢核資料"
            strMsg = "申請人和代理人不可同時空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM44.SetFocus
            GoTo EXITSUB
        End If
      End If
      '2008/2/12 end
   Else
        If IsEmptyText(textTM23) Then
            strTit = "檢核資料"
            strMsg = "申請人不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
        End If
   End If
   
   '911113 nick  當案件性質是 501 時，移轉申請人不可空白
   If textCP10.Text = "501" Then
       If IsEmptyText(textCP56) Then
            strTit = "檢核資料"
            strMsg = "移轉申請人不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP56.SetFocus
            GoTo EXITSUB
       End If
   End If
    'Add By Cheng 2003/08/13
    '若案件性質為延期, 則不可點選本案期限
    If Me.textCP10.Text = "303" Then
        For ii = 1 To Me.grdList.Rows - 1
            If Me.grdList.TextMatrix(ii, 0) <> "" Then
                MsgBox "此案僅收文<延期>，不可點選下一程序期限資料，" & vbCrLf & "否則無法管制下一程序的期限!!!", vbExclamation + vbOKOnly
                GoTo EXITSUB
            End If
        Next ii
    End If
   '2006/10/31 ADD BY SONIA
   If Me.textCP10 = "310" And Me.textCP43 = "" Then
      MsgBox "暫緩審理案件, 請輸入相關總收文號!!! 可按 案件進度 按鈕 點選 !!", vbExclamation + vbOKOnly
      Me.textCP43.SetFocus
      GoTo EXITSUB
   End If
   '2006/10/31 END
   
   'Added by Lydia 2017/10/16 FCT內部收文302更正，一定要輸入相關總收文
   If m_TM01 = "FCT" And Me.textCP10 = "302" And Me.textCP43 = "" Then
      MsgBox "內部收文更正, 請輸入相關總收文號!!! 可按 案件進度 按鈕 點選 !!", vbExclamation + vbOKOnly
      Me.textCP43.SetFocus
      GoTo EXITSUB
   End If
   'end 2017/10/16
   
   CheckDataValid = True
EXITSUB:
End Function


'add by nickc 2006/11/30
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
   Dim strCP89 As String
   Dim strTemp As String

   Cancel = False
   textCP89_2 = Empty
   If Not IsEmptyText(textCP89) Then
      strCP89 = textCP89
      'Modify By Sindy 2015/8/27 +m_TM01
      If GetCustomerAndState(strCP89, strTemp, , , , m_TM01) Then
         textCP89 = strCP89 & String(9 - Len(strCP89), "0")
         textCP89_2 = strTemp
      Else
         Cancel = True
         textCP89_GotFocus
      End If
   End If
End Sub
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
   Dim strCP90 As String
   Dim strTemp As String

   Cancel = False
   textCP90_2 = Empty
   If Not IsEmptyText(textCP90) Then
      strCP90 = textCP90
      'Modify By Sindy 2015/8/27 +m_TM01
      If GetCustomerAndState(strCP90, strTemp, , , , m_TM01) Then
         textCP90 = strCP90 & String(9 - Len(strCP90), "0")
         textCP90_2 = strTemp
      Else
         Cancel = True
         textCP90_GotFocus
      End If
   End If
End Sub
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
   Dim strCP91 As String
   Dim strTemp As String

   Cancel = False
   textCP91_2 = Empty
   If Not IsEmptyText(textCP91) Then
      strCP91 = textCP91
      'Modify By Sindy 2015/8/27 +m_TM01
      If GetCustomerAndState(strCP91, strTemp, , , , m_TM01) Then
         textCP91 = strCP91 & String(9 - Len(strCP91), "0")
         textCP91_2 = strTemp
      Else
         Cancel = True
         textCP91_GotFocus
      End If
   End If
End Sub
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
   Dim strCP92 As String
   Dim strTemp As String

   Cancel = False
   textCP92_2 = Empty
   If Not IsEmptyText(textCP92) Then
      strCP92 = textCP92
      'Modify By Sindy 2015/8/27 +m_TM01
      If GetCustomerAndState(strCP92, strTemp, , , , m_TM01) Then
         textCP92 = strCP92 & String(9 - Len(strCP92), "0")
         textCP92_2 = strTemp
      Else
         Cancel = True
         textCP92_GotFocus
      End If
   End If
End Sub

Private Sub textSP58_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSP59_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM05_1_GotFocus()
    TextInverse Me.textTM05_1
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textTM05_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM05_1, textTM05_1.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件名稱內容太長"
      textTM05_1.SetFocus
      textTM05_1_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

' 案件中文名稱
Private Sub textTM05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM05, textTM05.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件中文名稱內容太長"
      '911111 nick
      textTM05.SetFocus
      
      textTM05_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 案件英文名稱
Private Sub textTM06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM06, textTM06.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件英文名稱內容太長"
      '911111 nick
      textTM06.SetFocus
      
      textTM06_GotFocus
   End If
End Sub

' 案件日文名稱
Private Sub textTM07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM07, textTM07.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件日文名稱內容太長"
      '911111 nick
      textTM07.SetFocus
      
      textTM07_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM07.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 商標種類
Private Sub textTM08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textTM08_2 = Empty
   If IsEmptyText(textTM08) = False Then
      'Modify By Sindy 2015/8/13
      'textTM08_2 = GetTradeMarkName(textTM08, 0)
      textTM08_2 = GetTradeMarkName(textTM08, IIf(textTM10 = "020", 1, 0))
      '2015/8/13 END
      If IsEmptyText(textTM08_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商標種類不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textTM08.SetFocus
         
         textTM08_GotFocus
      End If
   End If
End Sub

' 商品類別
Private Sub textTM09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textTM09) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textTM09)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM09, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商品類別<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textTM09.SetFocus
         
         textTM09_GotFocus
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM09, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM09, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品類別<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               '911111 nick
               textTM09.SetFocus
               
               textTM09_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
'add by nickc 2005/06/03
textTM09 = Replace(textTM09, " ", "")
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 申請國家
Private Sub textTM10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textTM10_2 = Empty
   If IsEmptyText(textTM10) = False Then
      '911111 nick 邱小姐說不能 001~009
      If textTM10 >= "001" And textTM10 <= "009" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請國家代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textTM10.SetFocus
         
         textTM10_GotFocus
         Exit Sub
      End If
   
      textTM10_2 = GetNationName(textTM10, 0)
      If IsEmptyText(textTM10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請國家代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textTM10.SetFocus
         
         textTM10_GotFocus
         GoTo EXITSUB
      End If
      '91.11.10 add by sonia
      If m_TM01 = "FCT" And textTM10 <> 台灣國家代號 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請國家代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textTM10.SetFocus
         
         textTM10_GotFocus
         GoTo EXITSUB
      End If
      '91.11.10 END
      
      ReCaculateCP48
   End If
EXITSUB:
End Sub

' 申請人
Private Sub textTM23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTemp As String
   
   Cancel = False
   textTM23_2 = Empty
   textTM23_3 = Empty
   If IsEmptyText(textTM23) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textTM23_2 = GetCustomerName(textTM23, 0)
      textTM23_2 = GetCustomerNameAndState(textTM23, 0, oState)
      If oState = False Then
            Cancel = True
            textTM23.SetFocus
            textTM23_GotFocus
            Exit Sub
      End If
      If IsEmptyText(textTM23_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM23 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textTM23.SetFocus
         textTM23_GotFocus
         Exit Sub
      Else
         '910701 Sieg 601
         If m_CP60 <> "" And InStr(ChangeCustomerL(m_TM23), ChangeCustomerL(textTM23)) = 0 Then
            strExc(1) = m_TM01
            strExc(2) = m_TM02
            strExc(3) = m_TM03
            strExc(4) = m_TM04
            strExc(5) = m_CP60
            strExc(6) = textTM23
            strExc(7) = textTM23_2
            'edit by nickc 2007/02/05 不用 dll 了
            'If Not objLawDll.UpdAcc0k0(strExc()) Then
            If Not ClsLawUpdAcc0k0(strExc()) Then
               textTM23_2 = ""
               textTM24 = ""
               textTM25 = ""
               textTM26 = ""
               Cancel = True
               '911111 nick
               textTM23.SetFocus
               textTM23_GotFocus
               Exit Sub
            End If
         End If
         
         strTemp = GetCustomerNation(textTM23)
         If IsEmptyText(strTemp) = False Then
            textTM23_3 = GetNationName(strTemp, 0)
         End If
         ' 91.01.22 modify by louis (更新申請人地址)
         'Modify By Sindy 2011/1/11
         'UpdateCustomerAddress
      End If
   End If
   
   'Add By Cheng 2002/08/22
   If Cancel = False Then
      If Me.textTM23.Text <> m_strCust1 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   
   'Modify By Sindy 2011/1/11
   If m_TM23 <> textTM23 Then
      Call UpdateCustomerAddress(1, textTM23)
   End If
End Sub

' 申請人
Private Sub textSP58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textSP58_2 = Empty
   If IsEmptyText(textSP58) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textSP58_2 = GetCustomerName(textSP58, 0)
      textSP58_2 = GetCustomerNameAndState(textSP58, 0, oState)
      If oState = False Then
        Cancel = True
        textSP58.SetFocus
        textSP58_GotFocus
        Exit Sub
      End If
      If textSP58_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textSP58 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textSP58.SetFocus
         textSP58_GotFocus
         Exit Sub
      End If
   End If
   
   'Add By Cheng 2002/08/22
   If Cancel = False Then
      If Me.textSP58.Text <> m_strCust2 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then
            Cancel = True
            '911111 nick
            textSP58.SetFocus
            textSP58_GotFocus
            Exit Sub
         End If
      End If
   End If
   
   'Modify By Sindy 2011/1/11
   If m_TM78 <> textSP58 Then
      Call UpdateCustomerAddress(2, textSP58)
   End If
End Sub

' 申請人
Private Sub textSP59_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textSP59_2 = Empty
   If IsEmptyText(textSP59) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textSP59_2 = GetCustomerName(textSP59, 0)
      textSP59_2 = GetCustomerNameAndState(textSP59, 0, oState)
      If oState = False Then
        Cancel = True
        textSP59.SetFocus
        textSP59_GotFocus
        Exit Sub
      End If
      If textSP59_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textSP59 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textSP59.SetFocus
         textSP59_GotFocus
         Exit Sub
      End If
   End If
   'Add By Cheng 2002/08/22
   If Cancel = False Then
      If Me.textSP59.Text <> m_strCust3 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then
            Cancel = True
            '911111 nick
            textSP59.SetFocus
            textSP59_GotFocus
            Exit Sub
         End If
      End If
   End If
   
   'Modify By Sindy 2011/1/11
   If m_TM79 <> textSP59 Then
      Call UpdateCustomerAddress(3, textSP59)
   End If
End Sub

' 申請地址(中)
Private Sub textTM24_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM24) = False Then
      If CheckLengthIsOK(textTM24, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址1(中)內容太長"
         '911111 nick
         textTM24.SetFocus
         
         textTM24_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM24.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 申請地址(英)
Private Sub textTM25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM25) = False Then
      If CheckLengthIsOK(textTM25, 154) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址1(英)內容太長"
         '911111 nick
         textTM25.SetFocus
         
         textTM25_GotFocus
      End If
   End If
End Sub

' 申請地址(日)
Private Sub textTM26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM26) = False Then
      If CheckLengthIsOK(textTM26, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址1(日)內容太長"
         '911111 nick
         textTM26.SetFocus
         
         textTM26_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM26.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 卷宗性質
Private Sub textTM28_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM28) = False Then
      If IsEmptyText(textCP10) = False Then
         Select Case textCP10
            ' 異議
            Case "601":
               If textTM28 <> "2" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  '911111 nick
                  textTM28.SetFocus
                  
                  textTM28_GotFocus
               End If
            ' 評定
            Case "603":
               If textTM28 <> "3" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  '911111 nick
                  textTM28.SetFocus
                  
                  textTM28_GotFocus
               End If
            ' 廢止
            Case "605":
               If textTM28 <> "4" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  '911111 nick
                  textTM28.SetFocus
                  
                  textTM28_GotFocus
               End If
            Case Else:
               '91.11.10 CANCEL BY SONIA
               'If textTM28 <> "1" Then
               '   Cancel = True
               '   strTit = "檢核資料"
               '   strMsg = "卷宗性質不正確"
               '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               '   textTM28_GotFocus
               'End If
               '91.11.10 END
         End Select
      End If
   End If
End Sub

Private Sub textTM29_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否閉卷
Private Sub textTM29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM29) = False Then
      Select Case textTM29
         Case "Y", " ":
         Case Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '911111 nick
            textTM29.SetFocus
            
            textTM29_GotFocus
      End Select
   End If
End Sub
'add by nickc 2006/11/30
Private Sub textTM32_GotFocus()
InverseTextBox textTM32
End Sub

Private Sub textTM32_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textTM32) = True Then
      GoTo EXITSUB
   End If
   
   'Modify By Sindy 2024/4/18 商品組群欄人員貼上資料後將全形或半形的「；」分號，轉為半形的逗號存入TM32。
   textTM32 = Replace(Replace(textTM32, ";", ","), "；", ",")
   '2024/4/18 END
   nCount = GetSubStringCount(textTM32)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商品組群<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM32.SetFocus
         textTM32_GotFocus
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM32, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品組群<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM32.SetFocus
               textTM32_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
textTM32 = Replace(textTM32, " ", "")
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub textTM34_GotFocus()
   InverseTextBox textTM34
End Sub

Private Sub textTM34_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM34, 50) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "分所號內容太長"
      '911111 nick
      textTM34.SetFocus
      
      textTM34_GotFocus
   End If
End Sub

'Add By Sindy 2012/6/22
Private Sub textTM44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' FC代理人
Private Sub textTM44_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM44_2 = Empty
   If IsEmptyText(textTM44) = False Then
      textTM44_2 = GetFAgentName(textTM44)
      If IsEmptyText(textTM44_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "FC代理人<" & textTM44 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textTM44.SetFocus
         
         textTM44_GotFocus
      End If
   End If
End Sub

' 案件備註
Private Sub textTM58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM58, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註內容太長"
      '911111 nick
      textTM58.SetFocus
      
      textTM58_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM58.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub ReCaculateCP48()
   Dim strDay As String
   Dim strDate As String
   Dim strTemp As String
   
   ' 檢查案件性質
   If IsEmptyText(textCP10) = True Then
      GoTo EXITSUB
   End If
   ' 檢查收文日
   If IsEmptyText(textCP05) = True Then
      GoTo EXITSUB
   End If
   ' 檢查申請國家
   If IsEmptyText(textTM10) = True Then
      GoTo EXITSUB
   End If
   
   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質搜尋案件收費表的工作天數
''edit by nickc 2007/10/11 改抓有時效的
''   strDay = Empty
''   strDay = GetWorkDays(m_TM01, textTM10, textCP10)
''   If IsEmptyText(strDay) = False And strDay <> Empty Then
''      strDate = DBDATE(textCP05)
''      ' 90.07.03 modify by louis (承辦期限以實際工作天數計算)
''      'strTemp = DBDATE(Format(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) + Val(strDay))))
''      strTemp = DBDATE(CompWorkDay(Val(strDay), DBDATE(strDate), 0))
   strTemp = Pub_GetHandleDay(m_TM01, textTM10, textCP10, DBDATE(textCP05))
   If Trim(strTemp) <> "" Then
      textCP48 = TAIWANDATE(strTemp)
   End If
   
EXITSUB:
End Sub

Private Sub SetInputEntry()
   textCP14.SetFocus
End Sub

Private Sub textTM10_GotFocus()
   InverseTextBox textTM10
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

'Private Sub textCP09_S_GotFocus()
'   InverseTextBox textCP09_S
'End Sub

Private Sub textCP10_GotFocus()
   InverseTextBox textCP10
End Sub

Private Sub textCP13_GotFocus()
   InverseTextBox textCP13
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP16_GotFocus()
   InverseTextBox textCP16
   'add by sonia 2016/11/28
   textCP18 = CStr((Val(textCP16) - Val(textCP17)) / 1000)
   If textCP16 = "" And textCP17 = "" Then textCP18 = ""
   'end 2016/11/28
End Sub

Private Sub textCP17_GotFocus()
   InverseTextBox textCP17
   'add by sonia 2016/11/28
   textCP18 = CStr((Val(textCP16) - Val(textCP17)) / 1000)
   If textCP16 = "" And textCP17 = "" Then textCP18 = ""
   'end 2016/11/28
End Sub

Private Sub textCP18_GotFocus()
   InverseTextBox textCP18
   'add by sonia 2016/11/28
   textCP18 = CStr((Val(textCP16) - Val(textCP17)) / 1000)
   If textCP16 = "" And textCP17 = "" Then textCP18 = ""
   'end 2016/11/28
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP64.IMEMode = 1
'   OpenIme 'Modify By Sindy 2012/7/13 Mark
End Sub

Private Sub textTM08_GotFocus()
   InverseTextBox textTM08
End Sub

Private Sub textTM28_GotFocus()
   InverseTextBox textTM28
End Sub

Private Sub textCP43_GotFocus()
   InverseTextBox textCP43
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP56_GotFocus()
   InverseTextBox textCP56
End Sub

Private Sub textTM23_GotFocus()
   InverseTextBox textTM23
End Sub

Private Sub textTM24_GotFocus()
   InverseTextBox textTM24
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM24.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM25_GotFocus()
   InverseTextBox textTM25
End Sub

Private Sub textTM26_GotFocus()
   InverseTextBox textTM26
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM26.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM29_GotFocus()
   InverseTextBox textTM29
End Sub

Private Sub textTM44_GotFocus()
   InverseTextBox textTM44
   CloseIme 'Add By Sindy 2012/6/22
End Sub

Private Sub textTM45_GotFocus()
   InverseTextBox textTM45
End Sub

Private Sub textSP58_GotFocus()
   InverseTextBox textSP58
End Sub

Private Sub textSP59_GotFocus()
   InverseTextBox textSP59
End Sub

Private Sub textTM09_GotFocus()
   InverseTextBox textTM09
End Sub

Private Sub textTM58_GotFocus()
   InverseTextBox textTM58
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM58.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM05_GotFocus()
   InverseTextBox textTM05
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM05.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM06_GotFocus()
   InverseTextBox textTM06
End Sub

Private Sub textTM07_GotFocus()
   InverseTextBox textTM07
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM07.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM35_GotFocus()
   InverseTextBox textTM35
End Sub

Private Sub textCP01_S_GotFocus()
   InverseTextBox textCP01_S
End Sub

Private Sub textCP02_S_GotFocus()
   InverseTextBox textCP02_S
End Sub

Private Sub textCP03_S_GotFocus()
   InverseTextBox textCP03_S
End Sub

Private Sub textCP04_S_GotFocus()
   InverseTextBox textCP04_S
End Sub

Private Sub UpdateCustomerAddress(intType As Integer, strIdNo As String)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strCU01 As String
   Dim strCU02 As String
   Dim strTM24 As String
   Dim strTM25 As String
   Dim strTM26 As String
   
   Select Case intType
      Case 1 '申請人1
         If IsEmptyText(textTM23) Then
            textTM24 = Empty
            textTM25 = Empty
            textTM26 = Empty
            Exit Sub
         End If
      Case 2 '申請人2
         If IsEmptyText(textSP58) Then
            textTM82 = Empty
            textTM86 = Empty
            textTM90 = Empty
            Exit Sub
         End If
      Case 3 '申請人3
         If IsEmptyText(textSP59) Then
            textTM83 = Empty
            textTM87 = Empty
            textTM91 = Empty
            Exit Sub
         End If
      Case 4 '申請人4
         If IsEmptyText(textTM80) Then
            textTM84 = Empty
            textTM88 = Empty
            textTM92 = Empty
            Exit Sub
         End If
      Case 5 '申請人5
         If IsEmptyText(textTM81) Then
            textTM85 = Empty
            textTM89 = Empty
            textTM93 = Empty
            Exit Sub
         End If
   End Select
   
   If Len(strIdNo) > 8 Then
      strCU01 = Mid(strIdNo, 1, 8)
      strCU02 = Mid(strIdNo, 9, 1)
   Else
      strCU01 = strIdNo & String(8 - Len(strIdNo), "0")
      strCU02 = "0"
   End If
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CUSTOMER " & _
            "WHERE CU01 = '" & strCU01 & "' AND " & _
                  "CU02 = '" & strCU02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      '中文地址
      If Not IsNull(rsTmp.Fields("CU23")) Then
         If Not IsEmptyText(rsTmp.Fields("CU23")) Then
            strTM24 = rsTmp.Fields("CU23")
         End If
      End If
      '英文地址
      If Not IsNull(rsTmp.Fields("CU24")) Then
         If Not IsEmptyText(rsTmp.Fields("CU24")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU24")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU25")) Then
         If Not IsEmptyText(rsTmp.Fields("CU25")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU25")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU26")) Then
         If Not IsEmptyText(rsTmp.Fields("CU26")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU26")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU27")) Then
         If Not IsEmptyText(rsTmp.Fields("CU27")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU27")
         End If
      End If
      If Not IsNull(rsTmp.Fields("CU28")) Then
         If Not IsEmptyText(rsTmp.Fields("CU28")) Then
            If Not IsEmptyText(strTM25) Then strTM25 = strTM25 & " "
            strTM25 = strTM25 & rsTmp.Fields("CU28")
         End If
      End If
      '日文地址
      If Not IsNull(rsTmp.Fields("CU29")) Then
         If Not IsEmptyText(rsTmp.Fields("CU29")) Then
            strTM26 = rsTmp.Fields("CU29")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   Select Case intType
      Case 1 '申請人1
         textTM24 = strTM24
         textTM25 = strTM25
         textTM26 = strTM26
      Case 2 '申請人2
         textTM82 = strTM24
         textTM86 = strTM25
         textTM90 = strTM26
      Case 3 '申請人3
         textTM83 = strTM24
         textTM87 = strTM25
         textTM91 = strTM26
      Case 4 '申請人4
         textTM84 = strTM24
         textTM88 = strTM25
         textTM92 = strTM26
      Case 5 '申請人5
         textTM85 = strTM24
         textTM89 = strTM25
         textTM93 = strTM26
   End Select
End Sub

' 確認使用者所輸入的都完全正確
Private Function ValidateInput() As Boolean
   Dim Cancel As Boolean

   ValidateInput = False
   
   If textCP05.Enabled = True Then
      Cancel = False
      textCP05_Validate Cancel
      If Cancel = True Then
         textCP05.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP06.Enabled = True Then
      Cancel = False
      textCP06_Validate Cancel
      If Cancel = True Then
         textCP06.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP07.Enabled = True Then
      Cancel = False
      textCP07_Validate Cancel
      If Cancel = True Then
         textCP07.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP10.Enabled = True Then
      Cancel = False
      textCP10_Validate Cancel
      If Cancel = True Then
         textCP10.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP13.Enabled = True Then
      Cancel = False
      textCP13_Validate Cancel
      If Cancel = True Then
         textCP13.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP14.Enabled = True Then
      Cancel = False
      textCP14_Validate Cancel
      If Cancel = True Then
         textCP14.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP26.Enabled = True Then
      Cancel = False
      textCP26_Validate Cancel
      If Cancel = True Then
         textCP26.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP43.Enabled = True Then
      Cancel = False
      textCP43_Validate Cancel
      If Cancel = True Then
         textCP43.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP56.Enabled = True Then
      Cancel = False
      textCP56_Validate Cancel
      If Cancel = True Then
         textCP56.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         textCP64.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textSP58.Enabled = True Then
      Cancel = False
      textSP58_Validate Cancel
      If Cancel = True Then
         textSP58.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textSP59.Enabled = True Then
      Cancel = False
      textSP59_Validate Cancel
      If Cancel = True Then
         textSP59.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM05.Enabled = True Then
      Cancel = False
      textTM05_Validate Cancel
      If Cancel = True Then
         textTM05.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   If textTM05_1.Enabled = True Then
      Cancel = False
      textTM05_1_Validate Cancel
      If Cancel = True Then
         textTM05_1.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM06.Enabled = True Then
      Cancel = False
      textTM06_Validate Cancel
      If Cancel = True Then
         textTM06.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM07.Enabled = True Then
      Cancel = False
      textTM07_Validate Cancel
      If Cancel = True Then
         textTM07.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM08.Enabled = True Then
      Cancel = False
      textTM08_Validate Cancel
      If Cancel = True Then
         textTM08.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM09.Enabled = True Then
      Cancel = False
      textTM09_Validate Cancel
      If Cancel = True Then
         textTM09.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM10.Enabled = True Then
      Cancel = False
      textTM10_Validate Cancel
      If Cancel = True Then
         textTM10.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM23.Enabled = True Then
      Cancel = False
      textTM23_Validate Cancel
      If Cancel = True Then
         textTM23.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM24.Enabled = True Then
      Cancel = False
      textTM24_Validate Cancel
      If Cancel = True Then
         textTM24.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM25.Enabled = True Then
      Cancel = False
      textTM25_Validate Cancel
      If Cancel = True Then
         textTM25.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM26.Enabled = True Then
      Cancel = False
      textTM26_Validate Cancel
      If Cancel = True Then
         textTM26.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM28.Enabled = True Then
      Cancel = False
      textTM28_Validate Cancel
      If Cancel = True Then
         textTM28.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM29.Enabled = True Then
      Cancel = False
      textTM29_Validate Cancel
      If Cancel = True Then
         textTM29.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM44.Enabled = True Then
      Cancel = False
      textTM44_Validate Cancel
      If Cancel = True Then
         textTM44.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textTM58.Enabled = True Then
      Cancel = False
      textTM58_Validate Cancel
      If Cancel = True Then
         textTM58.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP01_S.Enabled = True Then
      Cancel = False
      textCP01_S_Validate Cancel
      If Cancel = True Then
         textCP01_S.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   '911112 nick
   '***** start
   If textCP48.Enabled = True Then
      Cancel = False
      textCP48_Validate Cancel
      If Cancel = True Then
         textCP48.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP16.Enabled = True Then
      Cancel = False
      textCP16_Validate Cancel
      If Cancel = True Then
         textCP16.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP17.Enabled = True Then
      Cancel = False
      textCP17_Validate Cancel
      If Cancel = True Then
         textCP17.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP18.Enabled = True Then
      Cancel = False
      textCP18_Validate Cancel
      If Cancel = True Then
         textCP18.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   
   If textCP01_S.Enabled = True Then
      Cancel = False
      textCP01_S_Validate Cancel
      If Cancel = True Then
         textCP01_S.SetFocus 'Add By Sindy 2012/6/25
         Exit Function
      End If
   End If
   '***** end
   
   ValidateInput = True
End Function

'取得案件收費表的下次期限
Private Function GetCF12(strCF01 As String, strCF02 As String, strCF03 As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetCF12 = "0"
   strSql = "SELECT CF12 FROM CASEFEE " & _
            "WHERE CF01 = '" & strCF01 & "' AND " & _
                  "CF02 = '" & strCF02 & "' AND " & _
                  "CF03 = '" & strCF03 & "' AND " & _
                  "CF12 IS NOT NULL "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GetCF12 = rsTmp.Fields("CF12").Value
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'取得案件收費表的規費
Private Function GetCF08(strCF01 As String, strCF02 As String, strCF03 As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetCF08 = "0"
   strSql = "SELECT CF08 FROM CASEFEE " & _
            "WHERE CF01 = '" & strCF01 & "' AND " & _
                  "CF02 = '" & strCF02 & "' AND " & _
                  "CF03 = '" & strCF03 & "' AND " & _
                  "CF08 IS NOT NULL "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GetCF08 = rsTmp.Fields("CF08").Value
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'add by nickc 2006/12/01
Private Sub textTM80_GotFocus()
InverseTextBox textTM80
End Sub
Private Sub textTM80_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM80_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM80_2 = Empty
   If IsEmptyText(textTM80) = False Then
      Dim oState As Boolean
      textTM80_2 = GetCustomerNameAndState(textTM80, 0, oState)
      If oState = False Then
        Cancel = True
        textTM80.SetFocus
        textTM80_GotFocus
        Exit Sub
      End If
      If textTM80_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM80 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM80.SetFocus
         textTM80_GotFocus
         Exit Sub
      End If
   End If
   If Cancel = False Then
      If Me.textTM80.Text <> m_strCust2 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then
            Cancel = True
            textTM80.SetFocus
            textTM80_GotFocus
            Exit Sub
         End If
      End If
   End If
   
   'Modify By Sindy 2011/1/11
   If m_TM80 <> textTM80 Then
      Call UpdateCustomerAddress(4, textTM80)
   End If
End Sub
Private Sub textTM81_GotFocus()
InverseTextBox textTM81
End Sub
Private Sub textTM81_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM81_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM81_2 = Empty
   If IsEmptyText(textTM81) = False Then
      Dim oState As Boolean
      textTM81_2 = GetCustomerNameAndState(textTM81, 0, oState)
      If oState = False Then
        Cancel = True
        textTM81.SetFocus
        textTM81_GotFocus
        Exit Sub
      End If
      If textTM81_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM81 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM81.SetFocus
         textTM81_GotFocus
         Exit Sub
      End If
   End If
   If Cancel = False Then
      If Me.textTM81.Text <> m_strCust2 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then
            Cancel = True
            textTM81.SetFocus
            textTM81_GotFocus
            Exit Sub
         End If
      End If
   End If
   
   'Modify By Sindy 2011/1/11
   If m_TM81 <> textTM81 Then
      Call UpdateCustomerAddress(5, textTM81)
   End If
End Sub
Private Sub textTM82_GotFocus()
   InverseTextBox textTM82
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM82.IMEMode = 1
   OpenIme
End Sub
Private Sub textTM82_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM82) = False Then
      If CheckLengthIsOK(textTM82, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址2(中)內容太長"
         textTM82.SetFocus
         textTM82_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM82.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM83_GotFocus()
   InverseTextBox textTM83
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM83.IMEMode = 1
   OpenIme
End Sub
Private Sub textTM83_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM83) = False Then
      If CheckLengthIsOK(textTM83, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址3(中)內容太長"
         textTM83.SetFocus
         textTM83_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM83.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM84_GotFocus()
   InverseTextBox textTM84
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM84.IMEMode = 1
   OpenIme
End Sub
Private Sub textTM84_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM84) = False Then
      If CheckLengthIsOK(textTM84, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址4(中)內容太長"
         textTM84.SetFocus
         textTM84_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM84.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM85_GotFocus()
   InverseTextBox textTM85
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM85.IMEMode = 1
   OpenIme
End Sub
Private Sub textTM85_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM85) = False Then
      If CheckLengthIsOK(textTM85, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址5(中)內容太長"
         textTM85.SetFocus
         textTM85_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM85.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
Private Sub textTM86_GotFocus()
InverseTextBox textTM86
End Sub
Private Sub textTM86_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM86) = False Then
      If CheckLengthIsOK(textTM86, 154) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址2(英)內容太長"
         '911111 nick
         textTM86.SetFocus
         
         textTM86_GotFocus
      End If
   End If
End Sub
Private Sub textTM87_GotFocus()
InverseTextBox textTM87
End Sub

Private Sub textTM87_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM87) = False Then
      If CheckLengthIsOK(textTM87, 154) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址3(英)內容太長"
         '911111 nick
         textTM87.SetFocus
         
         textTM87_GotFocus
      End If
   End If
End Sub

Private Sub textTM88_GotFocus()
InverseTextBox textTM88
End Sub

Private Sub textTM88_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM88) = False Then
      If CheckLengthIsOK(textTM88, 154) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址4(英)內容太長"
         '911111 nick
         textTM88.SetFocus
         
         textTM88_GotFocus
      End If
   End If
End Sub

Private Sub textTM89_GotFocus()
InverseTextBox textTM89
End Sub

Private Sub textTM89_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM89) = False Then
      If CheckLengthIsOK(textTM89, 154) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址5(英)內容太長"
         '911111 nick
         textTM89.SetFocus
         
         textTM89_GotFocus
      End If
   End If
End Sub

Private Sub textTM90_GotFocus()
   InverseTextBox textTM90
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM90.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM90_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM90) = False Then
      If CheckLengthIsOK(textTM90, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址2(日)內容太長"
         '911111 nick
         textTM90.SetFocus
         
         textTM90_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM90.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM91_GotFocus()
   InverseTextBox textTM91
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM91.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM91_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM91) = False Then
      If CheckLengthIsOK(textTM91, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址3(日)內容太長"
         '911111 nick
         textTM91.SetFocus
         
         textTM91_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM91.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM92_GotFocus()
   InverseTextBox textTM92
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM92.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM92_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM92) = False Then
      If CheckLengthIsOK(textTM92, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址4(日)內容太長"
         '911111 nick
         textTM92.SetFocus
         
         textTM92_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM92.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM93_GotFocus()
   InverseTextBox textTM93
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM93.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM93_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM93) = False Then
      If CheckLengthIsOK(textTM93, 70) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請地址5(日)內容太長"
         '911111 nick
         textTM93.SetFocus
         
         textTM93_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTM93.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
