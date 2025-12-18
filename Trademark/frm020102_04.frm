VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_04 
   BorderStyle     =   1  '單線固定
   Caption         =   "變更事項"
   ClientHeight    =   5796
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8544
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   8544
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5424
      Locked          =   -1  'True
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   705
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5424
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   435
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1224
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   705
      Width           =   2532
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6840
      TabIndex        =   96
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7668
      TabIndex        =   97
      Top             =   10
      Width           =   800
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1224
      Locked          =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   435
      Width           =   2532
   End
   Begin TabDlg.SSTab tabCtrl 
      Height          =   4515
      Left            =   90
      TabIndex        =   24
      Top             =   1290
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   7980
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm020102_04.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label20"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label17"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label19"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label23"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label27"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label29"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label31"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textCE04_2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCE05_2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCE06_2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textCE07_2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCE08_2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCE17"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCE18"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCE19"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCE20"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCE21"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCE02"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCE04"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "checkCE09"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "checkCE03"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCE05"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCE06"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCE07"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCE08"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "checkCE22"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "checkCE52"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "checkCE54"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "checkCE56"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm020102_04.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(3)=   "Label9"
      Tab(1).Control(4)=   "Label10"
      Tab(1).Control(5)=   "Label11"
      Tab(1).Control(6)=   "Label33"
      Tab(1).Control(7)=   "Label37"
      Tab(1).Control(8)=   "Label39"
      Tab(1).Control(9)=   "Label41"
      Tab(1).Control(10)=   "Label43"
      Tab(1).Control(11)=   "Label45"
      Tab(1).Control(12)=   "Label46"
      Tab(1).Control(13)=   "Label47"
      Tab(1).Control(14)=   "Label48"
      Tab(1).Control(15)=   "Label49"
      Tab(1).Control(16)=   "Label50"
      Tab(1).Control(17)=   "Label51"
      Tab(1).Control(18)=   "Label52"
      Tab(1).Control(19)=   "Label53"
      Tab(1).Control(20)=   "Label54"
      Tab(1).Control(21)=   "Label55"
      Tab(1).Control(22)=   "Label56"
      Tab(1).Control(23)=   "Label57"
      Tab(1).Control(24)=   "Label58"
      Tab(1).Control(25)=   "Label59"
      Tab(1).Control(26)=   "Label60"
      Tab(1).Control(27)=   "Label61"
      Tab(1).Control(28)=   "Label62"
      Tab(1).Control(29)=   "Label63"
      Tab(1).Control(30)=   "textCE10"
      Tab(1).Control(31)=   "textCE12"
      Tab(1).Control(32)=   "textCE13"
      Tab(1).Control(33)=   "textCE15"
      Tab(1).Control(34)=   "textCE68"
      Tab(1).Control(35)=   "textCE70"
      Tab(1).Control(36)=   "textCE71"
      Tab(1).Control(37)=   "textCE73"
      Tab(1).Control(38)=   "textCE74"
      Tab(1).Control(39)=   "textCE76"
      Tab(1).Control(40)=   "textCE77"
      Tab(1).Control(41)=   "textCE79"
      Tab(1).Control(42)=   "textCE80"
      Tab(1).Control(43)=   "textCE82"
      Tab(1).Control(44)=   "textCE83"
      Tab(1).Control(45)=   "textCE85"
      Tab(1).Control(46)=   "textCE86"
      Tab(1).Control(47)=   "textCE88"
      Tab(1).Control(48)=   "textCE89"
      Tab(1).Control(49)=   "textCE91"
      Tab(1).Control(50)=   "checkCE16"
      Tab(1).Control(51)=   "textCE11"
      Tab(1).Control(52)=   "textCE14"
      Tab(1).Control(53)=   "textCE69"
      Tab(1).Control(54)=   "textCE72"
      Tab(1).Control(55)=   "textCE75"
      Tab(1).Control(56)=   "textCE78"
      Tab(1).Control(57)=   "textCE81"
      Tab(1).Control(58)=   "textCE84"
      Tab(1).Control(59)=   "textCE87"
      Tab(1).Control(60)=   "textCE90"
      Tab(1).ControlCount=   61
      TabCaption(2)   =   "第三頁"
      TabPicture(2)   =   "frm020102_04.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label24"
      Tab(2).Control(1)=   "Label25"
      Tab(2).Control(2)=   "Label26"
      Tab(2).Control(3)=   "Label64"
      Tab(2).Control(4)=   "Label65"
      Tab(2).Control(5)=   "Label66"
      Tab(2).Control(6)=   "Label67"
      Tab(2).Control(7)=   "Label68"
      Tab(2).Control(8)=   "Label69"
      Tab(2).Control(9)=   "Label70"
      Tab(2).Control(10)=   "Label71"
      Tab(2).Control(11)=   "Label72"
      Tab(2).Control(12)=   "Label73"
      Tab(2).Control(13)=   "Label74"
      Tab(2).Control(14)=   "Label75"
      Tab(2).Control(15)=   "textCE23"
      Tab(2).Control(16)=   "textCE25"
      Tab(2).Control(17)=   "textCE26"
      Tab(2).Control(18)=   "textCE28"
      Tab(2).Control(19)=   "textCE29"
      Tab(2).Control(20)=   "textCE31"
      Tab(2).Control(21)=   "textCE32"
      Tab(2).Control(22)=   "textCE34"
      Tab(2).Control(23)=   "textCE35"
      Tab(2).Control(24)=   "textCE37"
      Tab(2).Control(25)=   "checkCE38"
      Tab(2).Control(26)=   "textCE24"
      Tab(2).Control(27)=   "textCE27"
      Tab(2).Control(28)=   "textCE30"
      Tab(2).Control(29)=   "textCE33"
      Tab(2).Control(30)=   "textCE36"
      Tab(2).ControlCount=   31
      TabCaption(3)   =   "第四頁"
      TabPicture(3)   =   "frm020102_04.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label21"
      Tab(3).Control(1)=   "Label22"
      Tab(3).Control(2)=   "Label76"
      Tab(3).Control(3)=   "Label77"
      Tab(3).Control(4)=   "Label78"
      Tab(3).Control(5)=   "Label79"
      Tab(3).Control(6)=   "Label80"
      Tab(3).Control(7)=   "Label81"
      Tab(3).Control(8)=   "Label82"
      Tab(3).Control(9)=   "Label83"
      Tab(3).Control(10)=   "textCE63"
      Tab(3).Control(11)=   "textCE64"
      Tab(3).Control(12)=   "textCE92"
      Tab(3).Control(13)=   "textCE93"
      Tab(3).Control(14)=   "textCE94"
      Tab(3).Control(15)=   "textCE95"
      Tab(3).Control(16)=   "textCE96"
      Tab(3).Control(17)=   "textCE97"
      Tab(3).Control(18)=   "textCE98"
      Tab(3).Control(19)=   "textCE99"
      Tab(3).Control(20)=   "checkCE65"
      Tab(3).ControlCount=   21
      TabCaption(4)   =   "第五頁"
      TabPicture(4)   =   "frm020102_04.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label12"
      Tab(4).Control(1)=   "Label5"
      Tab(4).Control(2)=   "Label28"
      Tab(4).Control(3)=   "Label30"
      Tab(4).Control(4)=   "Label32"
      Tab(4).Control(5)=   "Label34"
      Tab(4).Control(6)=   "Label35"
      Tab(4).Control(7)=   "Label36"
      Tab(4).Control(8)=   "Label38"
      Tab(4).Control(9)=   "Label40"
      Tab(4).Control(10)=   "Label42"
      Tab(4).Control(11)=   "Label44"
      Tab(4).Control(12)=   "textCE41"
      Tab(4).Control(13)=   "textCE43"
      Tab(4).Control(14)=   "textCE41_1"
      Tab(4).Control(15)=   "textCE66"
      Tab(4).Control(16)=   "checkCE67"
      Tab(4).Control(17)=   "textCE39_2"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "checkCE62"
      Tab(4).Control(19)=   "checkCE50"
      Tab(4).Control(20)=   "checkCE48"
      Tab(4).Control(21)=   "checkCE46"
      Tab(4).Control(22)=   "checkCE44"
      Tab(4).Control(23)=   "checkCE60"
      Tab(4).Control(24)=   "checkCE40"
      Tab(4).Control(25)=   "checkCE58"
      Tab(4).Control(26)=   "textCE57"
      Tab(4).Control(27)=   "textCE39"
      Tab(4).Control(28)=   "textCE42"
      Tab(4).Control(29)=   "textCE45"
      Tab(4).Control(30)=   "textCE47"
      Tab(4).Control(31)=   "textCE49"
      Tab(4).Control(32)=   "textCE61"
      Tab(4).Control(33)=   "cmdGoods"
      Tab(4).ControlCount=   34
      Begin VB.CommandButton cmdGoods 
         Caption         =   "商品名稱"
         Height          =   315
         Left            =   -74670
         TabIndex        =   87
         Top             =   2280
         Width           =   1005
      End
      Begin VB.TextBox textCE61 
         Height          =   270
         Left            =   -73350
         MaxLength       =   2000
         TabIndex        =   93
         Top             =   3270
         Width           =   6135
      End
      Begin VB.TextBox textCE49 
         Height          =   264
         Left            =   -73350
         MaxLength       =   699
         TabIndex        =   91
         Top             =   2970
         Width           =   6135
      End
      Begin VB.TextBox textCE47 
         Height          =   264
         Left            =   -73350
         MaxLength       =   395
         TabIndex        =   89
         Top             =   2670
         Width           =   6135
      End
      Begin VB.TextBox textCE45 
         Height          =   264
         Left            =   -73350
         MaxLength       =   200
         TabIndex        =   86
         Top             =   2070
         Width           =   6135
      End
      Begin VB.TextBox textCE42 
         Height          =   264
         Left            =   -73350
         MaxLength       =   180
         TabIndex        =   82
         Top             =   1530
         Width           =   6135
      End
      Begin VB.TextBox textCE39 
         Height          =   264
         Left            =   -73350
         MaxLength       =   1
         TabIndex        =   78
         Top             =   690
         Width           =   1212
      End
      Begin VB.TextBox textCE57 
         Height          =   264
         Left            =   -73350
         MaxLength       =   20
         TabIndex        =   76
         Top             =   390
         Width           =   1212
      End
      Begin VB.CheckBox checkCE58 
         Height          =   180
         Left            =   -74910
         TabIndex        =   75
         Top             =   390
         Width           =   252
      End
      Begin VB.CheckBox checkCE40 
         Height          =   180
         Left            =   -74910
         TabIndex        =   77
         Top             =   690
         Width           =   252
      End
      Begin VB.CheckBox checkCE60 
         Height          =   180
         Left            =   -74910
         TabIndex        =   79
         Top             =   990
         Width           =   252
      End
      Begin VB.CheckBox checkCE44 
         Height          =   180
         Left            =   -74910
         TabIndex        =   80
         Top             =   1290
         Width           =   252
      End
      Begin VB.CheckBox checkCE46 
         Height          =   180
         Left            =   -74910
         TabIndex        =   85
         Top             =   2070
         Width           =   252
      End
      Begin VB.CheckBox checkCE48 
         Height          =   180
         Left            =   -74910
         TabIndex        =   88
         Top             =   2670
         Width           =   252
      End
      Begin VB.CheckBox checkCE50 
         Height          =   180
         Left            =   -74910
         TabIndex        =   90
         Top             =   2970
         Width           =   252
      End
      Begin VB.CheckBox checkCE62 
         Height          =   180
         Left            =   -74910
         TabIndex        =   92
         Top             =   3270
         Width           =   252
      End
      Begin VB.TextBox textCE39_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -72030
         Locked          =   -1  'True
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   690
         Width           =   4815
      End
      Begin VB.CheckBox checkCE67 
         Height          =   180
         Left            =   -74910
         TabIndex        =   94
         Top             =   3570
         Width           =   252
      End
      Begin VB.TextBox textCE66 
         Height          =   270
         Left            =   -73350
         MaxLength       =   25
         TabIndex        =   95
         Top             =   3570
         Width           =   6135
      End
      Begin VB.TextBox textCE36 
         Height          =   264
         Left            =   -73350
         MaxLength       =   154
         TabIndex        =   62
         Top             =   3915
         Width           =   6135
      End
      Begin VB.TextBox textCE33 
         Height          =   264
         Left            =   -73350
         MaxLength       =   154
         TabIndex        =   59
         Top             =   3045
         Width           =   6135
      End
      Begin VB.TextBox textCE30 
         Height          =   264
         Left            =   -73350
         MaxLength       =   154
         TabIndex        =   56
         Top             =   2205
         Width           =   6135
      End
      Begin VB.TextBox textCE27 
         Height          =   264
         Left            =   -73350
         MaxLength       =   154
         TabIndex        =   53
         Top             =   1365
         Width           =   6135
      End
      Begin VB.CheckBox checkCE65 
         Height          =   180
         Left            =   -74820
         TabIndex        =   64
         Top             =   450
         Width           =   252
      End
      Begin VB.TextBox textCE90 
         Height          =   285
         Left            =   -69660
         MaxLength       =   80
         TabIndex        =   46
         Top             =   3930
         Width           =   2925
      End
      Begin VB.TextBox textCE87 
         Height          =   285
         Left            =   -69660
         MaxLength       =   80
         TabIndex        =   43
         Top             =   3090
         Width           =   2925
      End
      Begin VB.TextBox textCE84 
         Height          =   285
         Left            =   -69660
         MaxLength       =   80
         TabIndex        =   40
         Top             =   2280
         Width           =   2925
      End
      Begin VB.TextBox textCE81 
         Height          =   285
         Left            =   -69660
         MaxLength       =   80
         TabIndex        =   37
         Top             =   1470
         Width           =   2925
      End
      Begin VB.TextBox textCE78 
         Height          =   285
         Left            =   -69660
         MaxLength       =   80
         TabIndex        =   34
         Top             =   600
         Width           =   2925
      End
      Begin VB.TextBox textCE75 
         Height          =   285
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   31
         Top             =   3930
         Width           =   2925
      End
      Begin VB.TextBox textCE72 
         Height          =   285
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   28
         Top             =   3090
         Width           =   2925
      End
      Begin VB.TextBox textCE69 
         Height          =   285
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   25
         Top             =   2280
         Width           =   2925
      End
      Begin VB.CheckBox checkCE56 
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   1980
         Width           =   252
      End
      Begin VB.CheckBox checkCE54 
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   2460
         Width           =   252
      End
      Begin VB.CheckBox checkCE52 
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   2220
         Width           =   252
      End
      Begin VB.CheckBox checkCE22 
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   2700
         Width           =   252
      End
      Begin VB.TextBox textCE14 
         Height          =   285
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   22
         Top             =   1470
         Width           =   2925
      End
      Begin VB.TextBox textCE11 
         Height          =   285
         Left            =   -73680
         MaxLength       =   80
         TabIndex        =   19
         Top             =   600
         Width           =   2925
      End
      Begin VB.CheckBox checkCE16 
         Height          =   180
         Left            =   -74940
         TabIndex        =   17
         Top             =   330
         Width           =   195
      End
      Begin VB.TextBox textCE08 
         Height          =   264
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   5
         Top             =   1440
         Width           =   1212
      End
      Begin VB.TextBox textCE07 
         Height          =   264
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   4
         Top             =   1170
         Width           =   1212
      End
      Begin VB.TextBox textCE06 
         Height          =   264
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   3
         Top             =   900
         Width           =   1212
      End
      Begin VB.TextBox textCE05 
         Height          =   264
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   2
         Top             =   630
         Width           =   1212
      End
      Begin VB.TextBox textCE24 
         Height          =   264
         Left            =   -73350
         MaxLength       =   154
         TabIndex        =   50
         Top             =   555
         Width           =   6135
      End
      Begin VB.CheckBox checkCE38 
         Height          =   180
         Left            =   -74910
         TabIndex        =   48
         Top             =   285
         Width           =   252
      End
      Begin VB.CheckBox checkCE03 
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   1710
         Width           =   252
      End
      Begin VB.CheckBox checkCE09 
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   252
      End
      Begin VB.TextBox textCE04 
         Height          =   264
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   1
         Top             =   360
         Width           =   1212
      End
      Begin VB.TextBox textCE02 
         Height          =   264
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   7
         Top             =   1710
         Width           =   1212
      End
      Begin MSForms.TextBox textCE41_1 
         Height          =   792
         Left            =   -73350
         TabIndex        =   84
         Top             =   1290
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "10821;1397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE99 
         Height          =   300
         Left            =   -73260
         TabIndex        =   74
         Top             =   3150
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE98 
         Height          =   300
         Left            =   -73260
         TabIndex        =   73
         Top             =   2850
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE97 
         Height          =   300
         Left            =   -73260
         TabIndex        =   72
         Top             =   2550
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE96 
         Height          =   300
         Left            =   -73260
         TabIndex        =   71
         Top             =   2250
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE95 
         Height          =   300
         Left            =   -73260
         TabIndex        =   70
         Top             =   1950
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE94 
         Height          =   300
         Left            =   -73260
         TabIndex        =   69
         Top             =   1650
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE93 
         Height          =   300
         Left            =   -73260
         TabIndex        =   68
         Top             =   1350
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE92 
         Height          =   300
         Left            =   -73260
         TabIndex        =   67
         Top             =   1050
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE43 
         Height          =   300
         Left            =   -73350
         TabIndex        =   83
         Top             =   1770
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   160
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41 
         Height          =   264
         Left            =   -73350
         TabIndex        =   81
         Top             =   1290
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   160
         Size            =   "10821;466"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE37 
         Height          =   300
         Left            =   -73350
         TabIndex        =   63
         Top             =   4185
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   70
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE35 
         Height          =   300
         Left            =   -73350
         TabIndex        =   61
         Top             =   3630
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE34 
         Height          =   300
         Left            =   -73350
         TabIndex        =   60
         Top             =   3315
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   70
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE32 
         Height          =   300
         Left            =   -73350
         TabIndex        =   58
         Top             =   2760
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE31 
         Height          =   270
         Left            =   -73350
         TabIndex        =   57
         Top             =   2475
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   70
         Size            =   "10821;466"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE29 
         Height          =   300
         Left            =   -73350
         TabIndex        =   55
         Top             =   1920
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE28 
         Height          =   270
         Left            =   -73350
         TabIndex        =   54
         Top             =   1635
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   70
         Size            =   "10821;466"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE26 
         Height          =   300
         Left            =   -73350
         TabIndex        =   52
         Top             =   1110
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE64 
         Height          =   300
         Left            =   -73260
         TabIndex        =   66
         Top             =   750
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE63 
         Height          =   300
         Left            =   -73260
         TabIndex        =   65
         Top             =   450
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE91 
         Height          =   300
         Left            =   -69660
         TabIndex        =   47
         Top             =   4200
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE89 
         Height          =   300
         Left            =   -69660
         TabIndex        =   45
         Top             =   3630
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE88 
         Height          =   300
         Left            =   -69660
         TabIndex        =   44
         Top             =   3360
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE86 
         Height          =   300
         Left            =   -69660
         TabIndex        =   42
         Top             =   2820
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE85 
         Height          =   300
         Left            =   -69660
         TabIndex        =   41
         Top             =   2550
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE83 
         Height          =   300
         Left            =   -69660
         TabIndex        =   39
         Top             =   2010
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE82 
         Height          =   300
         Left            =   -69660
         TabIndex        =   38
         Top             =   1740
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE80 
         Height          =   300
         Left            =   -69660
         TabIndex        =   36
         Top             =   1170
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE79 
         Height          =   300
         Left            =   -69660
         TabIndex        =   35
         Top             =   870
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE77 
         Height          =   300
         Left            =   -69660
         TabIndex        =   33
         Top             =   300
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE76 
         Height          =   300
         Left            =   -73680
         TabIndex        =   32
         Top             =   4200
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE74 
         Height          =   300
         Left            =   -73680
         TabIndex        =   30
         Top             =   3630
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE73 
         Height          =   300
         Left            =   -73680
         TabIndex        =   29
         Top             =   3360
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE71 
         Height          =   300
         Left            =   -73680
         TabIndex        =   27
         Top             =   2820
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE70 
         Height          =   300
         Left            =   -73680
         TabIndex        =   26
         Top             =   2550
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE68 
         Height          =   300
         Left            =   -73680
         TabIndex        =   105
         Top             =   2010
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE21 
         Height          =   300
         Left            =   1680
         TabIndex        =   16
         Top             =   3900
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE20 
         Height          =   300
         Left            =   1680
         TabIndex        =   15
         Top             =   3600
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE19 
         Height          =   300
         Left            =   1680
         TabIndex        =   14
         Top             =   3300
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE18 
         Height          =   300
         Left            =   1680
         TabIndex        =   13
         Top             =   3000
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE17 
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Top             =   2700
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE15 
         Height          =   300
         Left            =   -73680
         TabIndex        =   23
         Top             =   1740
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE13 
         Height          =   300
         Left            =   -73680
         TabIndex        =   21
         Top             =   1170
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE12 
         Height          =   300
         Left            =   -73680
         TabIndex        =   20
         Top             =   870
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   40
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE10 
         Height          =   300
         Left            =   -73680
         TabIndex        =   18
         Top             =   300
         Width           =   2925
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5159;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE25 
         Height          =   300
         Left            =   -73350
         TabIndex        =   51
         Top             =   825
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   70
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE23 
         Height          =   300
         Left            =   -73350
         TabIndex        =   49
         Top             =   270
         Width           =   6135
         VariousPropertyBits=   679493659
         MaxLength       =   80
         Size            =   "10821;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE08_2 
         Height          =   264
         Left            =   3000
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4815
         VariousPropertyBits=   679493663
         Size            =   "8493;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE07_2 
         Height          =   264
         Left            =   3000
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   1170
         Width           =   4815
         VariousPropertyBits=   679493663
         Size            =   "8493;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE06_2 
         Height          =   264
         Left            =   3000
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   900
         Width           =   4815
         VariousPropertyBits=   679493663
         Size            =   "8493;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE05_2 
         Height          =   264
         Left            =   3000
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   630
         Width           =   4815
         VariousPropertyBits=   679493663
         Size            =   "8493;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE04_2 
         Height          =   264
         Left            =   3000
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   360
         Width           =   4815
         VariousPropertyBits=   679493663
         Size            =   "8493;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         Caption         =   "代表人10中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   197
         Top             =   3180
         Width           =   1350
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "代表人9中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   196
         Top             =   2850
         Width           =   1260
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "代表人8中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   195
         Top             =   2550
         Width           =   1260
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "代表人7中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   194
         Top             =   2250
         Width           =   1260
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "代表人6中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   193
         Top             =   1950
         Width           =   1260
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "代表人5中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   192
         Top             =   1650
         Width           =   1260
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "代表人4中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   191
         Top             =   1350
         Width           =   1260
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "代表人3中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   190
         Top             =   1050
         Width           =   1260
      End
      Begin VB.Label Label44 
         Caption         =   "其它 :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   189
         Top             =   3270
         Width           =   615
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "商品組群 :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   188
         Top             =   2970
         Width           =   810
      End
      Begin VB.Label Label40 
         Caption         =   "商品類別 :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   187
         Top             =   2670
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "縮減商品 :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   186
         Top             =   2070
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "案件名稱(日) :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   185
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label Label35 
         Caption         =   "案件名稱(英) :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   184
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label Label34 
         Caption         =   "案件名稱(中) :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   183
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "圖樣 :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   182
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label30 
         Caption         =   "商標種類 :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   181
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "正商標號數 :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   180
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "密碼 :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   179
         Top             =   3570
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "案件名稱 :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   178
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(日) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   176
         Top             =   4185
         Width           =   1200
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(英) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   175
         Top             =   3915
         Width           =   1200
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "申請地址5(中) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   174
         Top             =   3645
         Width           =   1200
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(日) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   173
         Top             =   3315
         Width           =   1200
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(英) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   172
         Top             =   3045
         Width           =   1200
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "申請地址4(中) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   171
         Top             =   2775
         Width           =   1200
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(日) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   170
         Top             =   2475
         Width           =   1200
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(英) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   169
         Top             =   2205
         Width           =   1200
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "申請地址3(中) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   168
         Top             =   1935
         Width           =   1200
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(日) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   167
         Top             =   1635
         Width           =   1200
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(英) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   166
         Top             =   1365
         Width           =   1200
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "申請地址2(中) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   165
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "代表人2中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   164
         Top             =   750
         Width           =   1260
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "代表人1中譯文 :"
         Height          =   180
         Left            =   -74580
         TabIndex        =   163
         Top             =   450
         Width           =   1260
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(日) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   162
         Top             =   4230
         Width           =   1110
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(英) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   161
         Top             =   3960
         Width           =   1110
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(中) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   160
         Top             =   3660
         Width           =   1110
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(日) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   159
         Top             =   3390
         Width           =   1020
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(英) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   158
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(中) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   157
         Top             =   2850
         Width           =   1020
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(日) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   156
         Top             =   2580
         Width           =   1020
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(英) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   155
         Top             =   2310
         Width           =   1020
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(中) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   154
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(日) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   153
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(英) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   152
         Top             =   1500
         Width           =   1020
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(中) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   151
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(日) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   150
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(英) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   149
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(中) :"
         Height          =   180
         Left            =   -70710
         TabIndex        =   148
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(日) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   147
         Top             =   4230
         Width           =   1020
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(英) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   146
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(中) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   145
         Top             =   3660
         Width           =   1020
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(日) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   144
         Top             =   3390
         Width           =   1020
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(英) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   143
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(中) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   142
         Top             =   2850
         Width           =   1020
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(日) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   141
         Top             =   2580
         Width           =   1020
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(英) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   140
         Top             =   2310
         Width           =   1020
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(中) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   139
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "申請人5中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   138
         Top             =   3900
         Width           =   1260
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "申請人4中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   137
         Top             =   3600
         Width           =   1260
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "申請人3中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   136
         Top             =   3300
         Width           =   1260
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請人2中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   135
         Top             =   3000
         Width           =   1260
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請人1中譯文 :"
         Height          =   180
         Left            =   360
         TabIndex        =   134
         Top             =   2700
         Width           =   1260
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人印鑑 :"
         Height          =   180
         Left            =   360
         TabIndex        =   133
         Top             =   2220
         Width           =   990
      End
      Begin VB.Label Label15 
         Caption         =   "代表人印鑑 :"
         Height          =   255
         Left            =   360
         TabIndex        =   132
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代理人 :"
         Height          =   180
         Left            =   360
         TabIndex        =   131
         Top             =   1980
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(日) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   130
         Top             =   1770
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(英) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   129
         Top             =   1500
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(中) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   128
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(日) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   127
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(英) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   126
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(中) :"
         Height          =   180
         Left            =   -74730
         TabIndex        =   125
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label Label20 
         Caption         =   "申請人5 :"
         Height          =   255
         Left            =   360
         TabIndex        =   124
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "申請人4 :"
         Height          =   255
         Left            =   360
         TabIndex        =   122
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "申請人3 :"
         Height          =   255
         Left            =   360
         TabIndex        =   120
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "申請人2 :"
         Height          =   255
         Left            =   360
         TabIndex        =   118
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(日) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   116
         Top             =   825
         Width           =   1200
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(英) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   115
         Top             =   555
         Width           =   1200
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "申請地址1(中) :"
         Height          =   180
         Left            =   -74670
         TabIndex        =   114
         Top             =   285
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "申請人1 :"
         Height          =   252
         Left            =   360
         TabIndex        =   101
         Top             =   360
         Width           =   732
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請日 :"
         Height          =   180
         Left            =   360
         TabIndex        =   100
         Top             =   1710
         Width           =   630
      End
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5424
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   975
      Width           =   2532
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   264
      Left            =   1224
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   975
      Width           =   2532
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標種類 :"
      Height          =   180
      Index           =   4
      Left            =   4470
      TabIndex        =   113
      Top             =   747
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號 :"
      Height          =   180
      Index           =   9
      Left            =   4470
      TabIndex        =   111
      Top             =   477
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   4470
      TabIndex        =   110
      Top             =   1017
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "承辦人 :"
      Height          =   180
      Left            =   150
      TabIndex        =   107
      Top             =   1017
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   104
      Top             =   747
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   102
      Top             =   477
      Width           =   810
   End
End
Attribute VB_Name = "frm020102_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/23 Form2.0已修改 textCP13/textCP14/textCE04_2(申請人名).../textCE17(申請人中譯文).../textCE10(代表人中/日).../textCE23(申請地址中/日).../textCE63(代表人中譯文).../textCE41(案件名稱中/日)...
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CE01 As String
' 申請國家
Dim m_TM10 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_CEList() As FIELDITEM
Dim m_CECount As Integer

' 更新商標基本檔時所使用的變更事項檔欄位的暫存資料
' 申請日
Dim m_CE02 As String
' 申請人
Dim m_CE04 As String
'add by nickc 2006/12/25
Dim m_CE05 As String
Dim m_CE06 As String
Dim m_CE07 As String
Dim m_CE08 As String

' 商品種類代碼
Dim m_CE39 As String
' 前畫面
'Dim m_Parent As String
Dim m_Parent As Form '前一畫面 Modify By Sindy 2018/9/25
Dim m_frm020102_01 As Form 'Add By Sindy 2018/9/25

' 91.09.02 modify by louis
Private Type SRFIELDITEM
   siName As String
   siData As String
   siType As String
End Type
Dim m_SRList() As SRFIELDITEM
Dim m_SRListCount As Integer
' 案件性質
Dim m_CP10 As String
'911204 nick
Dim tmpOldCE04 As String
Dim tmpOldCE02 As String
Dim tmpOldCE10 As String
Dim tmpOldCE11 As String
Dim tmpOldCE12 As String
Dim tmpOldCE13 As String
Dim tmpOldCE14 As String
Dim tmpOldCE15 As String
Dim tmpOldCE23 As String
Dim tmpOldCE24 As String
Dim tmpOldCE25 As String
Dim tmpOldCE39 As String
Dim tmpOldCE41 As String
Dim tmpOldCE42 As String
Dim tmpOldCE43 As String
Dim tmpOldCE47 As String
Dim tmpOldCE49 As String
Dim tmpOldCE57 As String

'add by nickc 2006/12/25
Dim tmpOldCE05 As String
Dim tmpOldCE06 As String
Dim tmpOldCE07 As String
Dim tmpOldCE08 As String
Dim tmpOldCE26 As String
Dim tmpOldCE27 As String
Dim tmpOldCE28 As String
Dim tmpOldCE29 As String
Dim tmpOldCE30 As String
Dim tmpOldCE31 As String
Dim tmpOldCE32 As String
Dim tmpOldCE33 As String
Dim tmpOldCE34 As String
Dim tmpOldCE35 As String
Dim tmpOldCE36 As String
Dim tmpOldCE37 As String
Dim tmpOldCE68 As String
Dim tmpOldCE69 As String
Dim tmpOldCE70 As String
Dim tmpOldCE71 As String
Dim tmpOldCE72 As String
Dim tmpOldCE73 As String
Dim tmpOldCE74 As String
Dim tmpOldCE75 As String
Dim tmpOldCE76 As String
Dim tmpOldCE77 As String
Dim tmpOldCE78 As String
Dim tmpOldCE79 As String
Dim tmpOldCE80 As String
Dim tmpOldCE81 As String
Dim tmpOldCE82 As String
Dim tmpOldCE83 As String
Dim tmpOldCE84 As String
Dim tmpOldCE85 As String
Dim tmpOldCE86 As String
Dim tmpOldCE87 As String
Dim tmpOldCE88 As String
Dim tmpOldCE89 As String
Dim tmpOldCE90 As String
Dim tmpOldCE91 As String

'Add By Cheng 2003/04/10
Private Type TMFIELDITEM
   tiName As String
   tiData As String
   tiType As String
End Type
Dim m_TMList() As TMFIELDITEM
Dim m_TMListCount As Integer
'add by nickc 2007/04/02
Dim m_TM09 As String
Public ChkTG As Boolean
'Add By Sindy 2011/8/3
Dim m_TM23 As String
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
Dim m_CP31 As String 'Add By Sindy 2011/8/23 是否新案件


' 檢查該欄位是否存在
Private Function IsCEFieldExist(ByVal strField As String) As Boolean
   Dim nIndex As Integer
   IsCEFieldExist = False
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         IsCEFieldExist = True
         Exit For
      End If
   Next nIndex
End Function
' 設定欄位新值
Private Sub SetCEFieldData(ByVal strField As String, ByVal strNewData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   bFind = False
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         bFind = True
         m_CEList(nIndex).fiNewData = strNewData
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_CEList(m_CECount + 1)
      m_CEList(m_CECount).fiName = strField
      m_CEList(m_CECount).fiNewData = strNewData
      m_CEList(m_CECount).fiType = nType
      m_CECount = m_CECount + 1
   End If
End Sub
' 清除欄位串列
Private Sub ClearCEFields()
   Erase m_CEList
   m_CECount = 0
End Sub

' 更新欄位內容
Private Sub UpdateFieldNewData()
   SetCEFieldData "CE01", m_CE01, 0
   If checkCE03.Value = 1 Then
      SetCEFieldData "CE02", DBDATE(textCE02), 1
   End If
   If checkCE09.Value = 1 Then
      If Trim(textCE04) <> "" Then
        'Modify By Sindy 2013/1/22
        'SetCEFieldData "CE04", textCE04 & String(9 - Len(textCE04), "0"), 0
        textCE04 = textCE04 & String(9 - Len(textCE04), "0")
        SetCEFieldData "CE04", textCE04, 0
        '2013/1/22 End
      End If
      'add by nickc 2007/03/29
      If Trim(textCE05) <> "" Then
        'Modify By Sindy 2013/1/22
        'SetCEFieldData "CE05", textCE05 & String(9 - Len(textCE05), "0"), 0
        textCE05 = textCE05 & String(9 - Len(textCE05), "0")
        SetCEFieldData "CE05", textCE05, 0
        '2013/1/22 End
      End If
      If Trim(textCE06) <> "" Then
        'Modify By Sindy 2013/1/22
        'SetCEFieldData "CE06", textCE06 & String(9 - Len(textCE06), "0"), 0
        textCE06 = textCE06 & String(9 - Len(textCE06), "0")
        SetCEFieldData "CE06", textCE06, 0
        '2013/1/22 End
      End If
      If Trim(textCE07) <> "" Then
        'Modify By Sindy 2013/1/22
        'SetCEFieldData "CE07", textCE07 & String(9 - Len(textCE07), "0"), 0
        textCE07 = textCE07 & String(9 - Len(textCE07), "0")
        SetCEFieldData "CE07", textCE07, 0
        '2013/1/22 End
      End If
      If Trim(textCE08) <> "" Then
        'Modify By Sindy 2013/1/22
        'SetCEFieldData "CE08", textCE08 & String(9 - Len(textCE08), "0"), 0
        textCE08 = textCE08 & String(9 - Len(textCE08), "0")
        SetCEFieldData "CE08", textCE08, 0
        '2013/1/22 End
      End If
   End If
   If checkCE16.Value = 1 Then
      SetCEFieldData "CE10", textCE10, 0
      SetCEFieldData "CE11", textCE11, 0
      SetCEFieldData "CE12", textCE12, 0
      SetCEFieldData "CE13", textCE13, 0
      SetCEFieldData "CE14", textCE14, 0
      SetCEFieldData "CE15", textCE15, 0
      'add by nickc 2006/12/25
      SetCEFieldData "CE68", textCE68, 0
      SetCEFieldData "CE69", textCE69, 0
      SetCEFieldData "CE70", textCE70, 0
      SetCEFieldData "CE71", textCE71, 0
      SetCEFieldData "CE72", textCE72, 0
      SetCEFieldData "CE73", textCE73, 0
      SetCEFieldData "CE74", textCE74, 0
      SetCEFieldData "CE75", textCE75, 0
      SetCEFieldData "CE76", textCE76, 0
      SetCEFieldData "CE77", textCE77, 0
      SetCEFieldData "CE78", textCE78, 0
      SetCEFieldData "CE79", textCE79, 0
      SetCEFieldData "CE80", textCE80, 0
      SetCEFieldData "CE81", textCE81, 0
      SetCEFieldData "CE82", textCE82, 0
      SetCEFieldData "CE83", textCE83, 0
      SetCEFieldData "CE84", textCE84, 0
      SetCEFieldData "CE85", textCE85, 0
      SetCEFieldData "CE86", textCE86, 0
      SetCEFieldData "CE87", textCE87, 0
      SetCEFieldData "CE88", textCE88, 0
      SetCEFieldData "CE89", textCE89, 0
      SetCEFieldData "CE90", textCE90, 0
      SetCEFieldData "CE91", textCE91, 0
   End If
   If checkCE56.Value = 1 Then
      SetCEFieldData "CE55", "V", 0
   End If
   If checkCE54.Value = 1 Then
      SetCEFieldData "CE53", "V", 0
   End If
   If checkCE52.Value = 1 Then
      SetCEFieldData "CE51", "V", 0
   End If
   If checkCE22.Value = 1 Then
      SetCEFieldData "CE17", textCE17, 0
      'add by nickc 2006/12/25
      SetCEFieldData "CE18", textCE18, 0
      SetCEFieldData "CE19", textCE19, 0
      SetCEFieldData "CE20", textCE20, 0
      SetCEFieldData "CE21", textCE21, 0
   End If
   If checkCE65.Value = 1 Then
      SetCEFieldData "CE63", textCE63, 0
      SetCEFieldData "CE64", textCE64, 0
      'add by nickc 2006/12/25
      SetCEFieldData "CE92", textCE92, 0
      SetCEFieldData "CE93", textCE93, 0
      SetCEFieldData "CE94", textCE94, 0
      SetCEFieldData "CE95", textCE95, 0
      SetCEFieldData "CE96", textCE96, 0
      SetCEFieldData "CE97", textCE97, 0
      SetCEFieldData "CE98", textCE98, 0
      SetCEFieldData "CE99", textCE99, 0
   End If
   If checkCE38.Value = 1 Then
      SetCEFieldData "CE23", textCE23, 0
      SetCEFieldData "CE24", textCE24, 0
      SetCEFieldData "CE25", textCE25, 0
      'add by nickc 2006/12/25
      SetCEFieldData "CE26", textCE26, 0
      SetCEFieldData "CE27", textCE27, 0
      SetCEFieldData "CE28", textCE28, 0
      SetCEFieldData "CE29", textCE29, 0
      SetCEFieldData "CE30", textCE30, 0
      SetCEFieldData "CE31", textCE31, 0
      SetCEFieldData "CE32", textCE32, 0
      SetCEFieldData "CE33", textCE33, 0
      SetCEFieldData "CE34", textCE34, 0
      SetCEFieldData "CE35", textCE35, 0
      SetCEFieldData "CE36", textCE36, 0
      SetCEFieldData "CE37", textCE37, 0
   End If
   If checkCE58.Value = 1 Then
      SetCEFieldData "CE57", textCE57, 0
   End If
   If checkCE40.Value = 1 Then
      SetCEFieldData "CE39", textCE39, 0
   End If
   If checkCE60.Value = 1 Then
      SetCEFieldData "CE59", "V", 0
   End If
   If checkCE44.Value = 1 Then
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF", "TS"
            SetCEFieldData "CE41", Me.textCE41_1.Text, 0
        Case Else
            SetCEFieldData "CE41", textCE41, 0
            SetCEFieldData "CE42", textCE42, 0
            SetCEFieldData "CE43", textCE43, 0
        End Select
   End If
   If checkCE46.Value = 1 Then
      SetCEFieldData "CE45", textCE45, 0
   End If
   If checkCE48.Value = 1 Then
      SetCEFieldData "CE47", textCE47, 0
   End If
   If checkCE50.Value = 1 Then
      SetCEFieldData "CE49", textCE49, 0
   End If
   If checkCE62.Value = 1 Then
      SetCEFieldData "CE61", textCE61, 0
   End If
   If checkCE67.Value = 1 Then
      SetCEFieldData "CE66", textCE66, 0
   End If
End Sub

Private Sub checkCE38_Click()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
    'Add By Cheng 2003/04/17
    '若有勾選變更申請地址, 且未勾選變更申請人
    If Me.checkCE38.Value = vbChecked And Me.checkCE09.Value = vbUnchecked Then
        'edit by nickc 2006/12/25
        'StrSQLa = "Select TM23 From Trademark Where " & ChgTradeMark(Replace(textTMKey.Text, "-", ""))
        'StrSQLa = StrSQLa & " union Select SP08 From Servicepractice Where " & ChgService(Replace(textTMKey.Text, "-", ""))
        StrSQLa = "Select TM23,tm78,tm79,tm80,tm81 From Trademark Where " & ChgTradeMark(Replace(textTMKey.Text, "-", ""))
        StrSQLa = StrSQLa & " union Select SP08,sp58,sp59,sp65,sp66 From Servicepractice Where " & ChgService(Replace(textTMKey.Text, "-", ""))
        
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            '顯示申請人地址
            textCE23.Text = PUB_GetCustEachAdd("" & rsA.Fields(0).Value, "1")
            textCE24.Text = PUB_GetCustEachAdd("" & rsA.Fields(0).Value, "2")
            textCE25.Text = PUB_GetCustEachAdd("" & rsA.Fields(0).Value, "3")
            'add by nickc 2006/12/25
            textCE26.Text = PUB_GetCustEachAdd("" & rsA.Fields(1).Value, "1")
            textCE27.Text = PUB_GetCustEachAdd("" & rsA.Fields(1).Value, "2")
            textCE28.Text = PUB_GetCustEachAdd("" & rsA.Fields(1).Value, "3")
            textCE29.Text = PUB_GetCustEachAdd("" & rsA.Fields(2).Value, "1")
            textCE30.Text = PUB_GetCustEachAdd("" & rsA.Fields(2).Value, "2")
            textCE31.Text = PUB_GetCustEachAdd("" & rsA.Fields(2).Value, "3")
            textCE32.Text = PUB_GetCustEachAdd("" & rsA.Fields(3).Value, "1")
            textCE33.Text = PUB_GetCustEachAdd("" & rsA.Fields(3).Value, "2")
            textCE34.Text = PUB_GetCustEachAdd("" & rsA.Fields(3).Value, "3")
            textCE35.Text = PUB_GetCustEachAdd("" & rsA.Fields(4).Value, "1")
            textCE36.Text = PUB_GetCustEachAdd("" & rsA.Fields(4).Value, "2")
            textCE37.Text = PUB_GetCustEachAdd("" & rsA.Fields(4).Value, "3")
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    '若有勾選變更申請地址, 同時勾選變更申請人
    ElseIf Me.checkCE38.Value = vbChecked And Me.checkCE09.Value = vbChecked Then
        '顯示申請人地址
        textCE23.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE04.Text), "1")
        textCE24.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE04.Text), "2")
        textCE25.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE04.Text), "3")
        'add by nickc 2006/12/25
        textCE26.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE05.Text), "1")
        textCE27.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE05.Text), "2")
        textCE28.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE05.Text), "3")
        textCE29.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE06.Text), "1")
        textCE30.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE06.Text), "2")
        textCE31.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE06.Text), "3")
        textCE32.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE07.Text), "1")
        textCE33.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE07.Text), "2")
        textCE34.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE07.Text), "3")
        textCE35.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE08.Text), "1")
        textCE36.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE08.Text), "2")
        textCE37.Text = PUB_GetCustEachAdd(ChangeCustomerL(Me.textCE08.Text), "3")
    End If

End Sub

Private Sub cmdExit_Click()
   'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'   '列印接洽接案單
'   PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'   '刪除暫存資料
'   PUB_DeleteCaseCloseSheet strUserNum
   
   'Modify By Sindy 2018/9/25
'   Select Case m_Parent
'      Case "frm020102_05":
'         Unload frm020102_05
'      Case "frm020102_07":
'         Unload frm020102_07
'      Case "frm020102_08":
'         Unload frm020102_08
'      Case "frm020102_09":
'         Unload frm020102_09
'      Case "frm020102_10":
'         Unload frm020102_10
'   End Select
'   Unload frm020102_01
   If UCase(TypeName(m_Parent)) = UCase("frm020102_05") Or _
      UCase(TypeName(m_Parent)) = UCase("frm020102_07") Or _
      UCase(TypeName(m_Parent)) = UCase("frm020102_08") Or _
      UCase(TypeName(m_Parent)) = UCase("frm020102_09") Or _
      UCase(TypeName(m_Parent)) = UCase("frm020102_10") Then
      Unload m_Parent
   Else
      m_Parent.Show
   End If
   If UCase(TypeName(m_frm020102_01)) = UCase("frm020102_01") Then
      Unload m_frm020102_01
   End If
   '2018/9/25 END
   Unload Me
End Sub

Private Sub cmdGoods_Click()
frm03010303_04.Hide
Set frm03010303_04.UpForm = Me
frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
frm03010303_04.AllClass = m_TM09
frm03010303_04.cmdOK(2).Visible = True
Me.Hide
frm03010303_04.QueryData
frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
End Sub

Private Sub cmdok_Click()
   'Add By Cheng 2003/01/27
   '檢查輸入資料的有效性
   If CheckDataValidate = False Then Exit Sub
   UpdateFieldNewData
    'Modify By Cheng 2002/11/06
'   'OnSaveData
   If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
   Unload Me
   
   'Add By Sindy 2012/4/17 有變更事項資料時,必須按確定鍵
'   Select Case m_Parent
'      Case "frm020102_05"
'         frm020102_05.m_blnClkChgButton = True
'         frm020102_05.Show
'         frm020102_05.QueryData
'      Case "frm020102_07"
'         frm020102_07.m_blnClkChgButton = True
'         frm020102_07.Show
'         frm020102_07.QueryData
'      Case "frm020102_08"
'         frm020102_08.m_blnClkChgButton = True
'         frm020102_08.Show
'         frm020102_08.QueryData
'      Case "frm020102_09"
'         frm020102_09.m_blnClkChgButton = True
'         frm020102_09.Show
'         frm020102_09.QueryData
'      Case "frm020102_10"
'         frm020102_10.m_blnClkChgButton = True
'         frm020102_10.Show
'         frm020102_10.QueryData
'   End Select
   'Modify By Sindy 2018/9/25
   If UCase(TypeName(m_Parent)) = UCase("frm020102_05") Or _
      UCase(TypeName(m_Parent)) = UCase("frm020102_07") Or _
      UCase(TypeName(m_Parent)) = UCase("frm020102_08") Or _
      UCase(TypeName(m_Parent)) = UCase("frm020102_09") Or _
      UCase(TypeName(m_Parent)) = UCase("frm020102_10") Then
      m_Parent.m_blnClkChgButton = True
      m_Parent.Show
      m_Parent.QueryData
   Else
      m_Parent.Show
      If UCase(TypeName(m_Parent)) = UCase("frm090202_2") Then
         Call m_Parent.cmdMod_LostFocus
      End If
   End If
End Sub

Private Sub Form_Load()
   textTMKey.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCE04_2.BackColor = &H8000000F
   'add by nickc 2007/01/24
   textCE05_2.BackColor = &H8000000F
   textCE06_2.BackColor = &H8000000F
   textCE07_2.BackColor = &H8000000F
   textCE08_2.BackColor = &H8000000F
   
   textCE39_2.BackColor = &H8000000F
   
    'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    textCE10.MaxLength = Pub_MaxCEL10
    textCE11.MaxLength = Pub_MaxCEL11
    textCE13.MaxLength = Pub_MaxCEL10
    textCE14.MaxLength = Pub_MaxCEL11
    textCE68.MaxLength = Pub_MaxCEL10
    textCE69.MaxLength = Pub_MaxCEL11
    textCE71.MaxLength = Pub_MaxCEL10
    textCE72.MaxLength = Pub_MaxCEL11
    textCE74.MaxLength = Pub_MaxCEL10
    textCE75.MaxLength = Pub_MaxCEL11
    textCE77.MaxLength = Pub_MaxCEL10
    textCE78.MaxLength = Pub_MaxCEL11
    textCE80.MaxLength = Pub_MaxCEL10
    textCE81.MaxLength = Pub_MaxCEL11
    textCE83.MaxLength = Pub_MaxCEL10
    textCE84.MaxLength = Pub_MaxCEL11
    textCE86.MaxLength = Pub_MaxCEL10
    textCE87.MaxLength = Pub_MaxCEL11
    textCE89.MaxLength = Pub_MaxCEL10
    textCE90.MaxLength = Pub_MaxCEL11
    'end 2016/09/10
    
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    ClearCEFields
   'Add By Cheng 2002/07/18
   Set frm020102_04 = Nothing
End Sub

' 由客戶代碼取得客戶名稱
Private Function GetCustomer(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCustomer = Empty
   If IsEmptyText(strData) = False Then
      Set rsTmp = New ADODB.Recordset
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CU04")) = False Then
            GetCustomer = rsTmp.Fields("CU04")
         ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
            GetCustomer = rsTmp.Fields("CU05")
         ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
            GetCustomer = rsTmp.Fields("CU06")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 由客戶代號取得地址
' Input : strData ==> 客戶代號
'         nType ==> 種類
'                   0 : 表要取得的是中文地址
'                   1 : 表要取得的是英文地址
'                   2 : 表要取得的是日文地址
Private Function GetAddress(ByVal strData As String, ByVal nType As Integer) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetAddress = Empty
   If IsEmptyText(strData) = False Then
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Select Case nType
            Case 0:
               If IsNull(rsTmp.Fields("CU23")) = False Then
                  GetAddress = rsTmp.Fields("CU23")
               End If
            Case 1:
               If IsNull(rsTmp.Fields("CU24")) = False Then
                  GetAddress = rsTmp.Fields("CU24")
               End If
            Case 2:
               If IsNull(rsTmp.Fields("CU29")) = False Then
                  GetAddress = rsTmp.Fields("CU29")
               End If
         End Select
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CE01 = Empty
      ' 91.09.02 modify by louis
      m_CP10 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 收文號
      Case 4: m_CE01 = strData
      ' 91.09.02
      ' 案件性質
      Case 5: m_CP10 = strData
   End Select
End Sub

'Public Sub SetParent(ByVal strParent As String)
'   m_Parent = strParent
'End Sub
'Modify By Sindy 2018/9/25
Public Sub SetParent(ByVal fm As Form)
   Set m_Parent = fm
End Sub
'Add By Sindy 2018/9/25
Public Sub SetParent_MainForm(ByVal fm As Form)
   Set m_frm020102_01 = fm
End Sub

Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   
   '911204 nick 服務也要抓
   'strSQL = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
'edit by nickc 2007/04/02
'   strSQL = "SELECT TM08,TM45,TM10 FROM TradeMark " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "' "
'   strSQL = strSQL & " union SELECT '' as tm08,sp27 as tm45,sp09 as tm10 FROM servicepractice " & _
'            "WHERE sp01 = '" & m_TM01 & "' AND " & _
'                  "sp02 = '" & m_TM02 & "' AND " & _
'                  "sp03 = '" & m_TM03 & "' AND " & _
'                  "sp04 = '" & m_TM04 & "' "
   strSql = "SELECT TM08,TM45,TM10,tm09 FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
   strSql = strSql & " union SELECT '' as tm08,sp27 as tm45,sp09 as tm10,sp73 as tm09 FROM servicepractice " & _
            "WHERE sp01 = '" & m_TM01 & "' AND " & _
                  "sp02 = '" & m_TM02 & "' AND " & _
                  "sp03 = '" & m_TM03 & "' AND " & _
                  "sp04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         textTM08 = rsTmp.Fields("TM08")
      End If
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then
         textTM45 = rsTmp.Fields("TM45")
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      'add by nickc 2007/04/02
      m_TM09 = Empty
      If IsNull(rsTmp.Fields("TM09")) = False Then
         m_TM09 = rsTmp.Fields("TM09")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      'Add By Sindy 2011/8/23 是否新案件
      m_CP31 = ""
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
   End If
End Sub

Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
    'Add By Cheng 2003/11/10
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF", "TS"
        Me.Label34.Visible = False
        Me.Label35.Visible = False
        Me.Label36.Visible = False
        Me.textCE41.Visible = False
        Me.textCE41.Enabled = False
        Me.textCE42.Visible = False
        Me.textCE42.Enabled = False
        Me.textCE43.Visible = False
        Me.textCE43.Enabled = False
    Case Else
        Me.Label12.Visible = False
        Me.textCE41_1.Visible = False
        Me.textCE41_1.Enabled = False
    End Select
    'End
   'Add By Cheng 2002/07/17
   m_TM10 = ""
   QueryTradeMark
   QueryCaseProgress
   
   ' 清除暫存變數
   m_CE02 = Empty
   m_CE04 = Empty
   'Add By Cheng 2002/07/17
   m_CE39 = Empty
   
   '911204 nick
   tmpOldCE04 = Empty
   tmpOldCE02 = Empty
   tmpOldCE10 = Empty
   tmpOldCE11 = Empty
   tmpOldCE12 = Empty
   tmpOldCE13 = Empty
   tmpOldCE14 = Empty
   tmpOldCE15 = Empty
   tmpOldCE23 = Empty
   tmpOldCE24 = Empty
   tmpOldCE25 = Empty
   tmpOldCE39 = Empty
   tmpOldCE41 = Empty
   tmpOldCE42 = Empty
   tmpOldCE43 = Empty
   tmpOldCE47 = Empty
   tmpOldCE49 = Empty
   tmpOldCE57 = Empty
   
    'add by nickc 2006/12/25
    tmpOldCE05 = Empty
    tmpOldCE06 = Empty
    tmpOldCE07 = Empty
    tmpOldCE08 = Empty
    tmpOldCE26 = Empty
    tmpOldCE27 = Empty
    tmpOldCE28 = Empty
    tmpOldCE29 = Empty
    tmpOldCE30 = Empty
    tmpOldCE31 = Empty
    tmpOldCE32 = Empty
    tmpOldCE33 = Empty
    tmpOldCE34 = Empty
    tmpOldCE35 = Empty
    tmpOldCE36 = Empty
    tmpOldCE37 = Empty
    tmpOldCE68 = Empty
    tmpOldCE69 = Empty
    tmpOldCE70 = Empty
    tmpOldCE71 = Empty
    tmpOldCE72 = Empty
    tmpOldCE73 = Empty
    tmpOldCE74 = Empty
    tmpOldCE75 = Empty
    tmpOldCE76 = Empty
    tmpOldCE77 = Empty
    tmpOldCE78 = Empty
    tmpOldCE79 = Empty
    tmpOldCE80 = Empty
    tmpOldCE81 = Empty
    tmpOldCE82 = Empty
    tmpOldCE83 = Empty
    tmpOldCE84 = Empty
    tmpOldCE85 = Empty
    tmpOldCE86 = Empty
    tmpOldCE87 = Empty
    tmpOldCE88 = Empty
    tmpOldCE89 = Empty
    tmpOldCE90 = Empty
    tmpOldCE91 = Empty
   
   
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   'textCE01 = m_CE01
   
    'Add By Cheng 2003/09/03
    'Begin
    'edit by nickc 2006/12/25
    'StrSQLa = "Select TM11, TM23, TM47, TM48, TM49, TM50, TM51, TM52, TM08, TM05, TM06, TM07, TM09, TM32, TM27 From Trademark Where " & ChgTradeMark(textTMKey)
    'StrSQLa = StrSQLa & " Union Select SP10, SP08, SP42, '', '', '', '', '', '', SP05, SP06, SP07, '', '', '' From Servicepractice Where " & ChgService(textTMKey)
    StrSQLa = "Select TM11, TM23, TM47, TM48, TM49, TM50, TM51, TM52, TM08, TM05, TM06, TM07, TM09, TM32, TM27,tm78,tm79,tm80,tm81,tm94,tm95,tm96,tm97,tm98,tm99,tm100,tm101,tm102,tm103,tm104,tm105,tm106,tm107,tm108,tm109,tm110,tm111,tm112,tm113,tm114,tm115,tm116,tm117,TM24,TM25,TM26,TM82,TM86,TM90,TM83,TM87,TM91,TM84,TM88,TM92,TM85,TM89,TM93 From Trademark Where " & ChgTradeMark(textTMKey)
    StrSQLa = StrSQLa & " Union Select SP10, SP08, SP42, '', '', '', '', '', '', SP05, SP06, SP07, '', '', '',sp58,sp59,sp65,sp66,'','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','' From Servicepractice Where " & ChgService(textTMKey)
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
      ' 申請日
        tmpOldCE02 = "" & rsA.Fields(0).Value
      ' 申請人
        m_TM23 = "" & rsA.Fields(1).Value 'Add By Sindy 2011/8/3
        tmpOldCE04 = "" & rsA.Fields(1).Value
      '申請人地址
'        tmpOldCE23 = PUB_GetCustEachAdd(tmpOldCE04, "1")
'        tmpOldCE24 = PUB_GetCustEachAdd(tmpOldCE04, "2")
'        tmpOldCE25 = PUB_GetCustEachAdd(tmpOldCE04, "3")
        'Modify By Sindy 2011/2/1
        tmpOldCE23 = Trim("" & rsA.Fields(43).Value)
        tmpOldCE24 = Trim("" & rsA.Fields(44).Value)
        tmpOldCE25 = Trim("" & rsA.Fields(45).Value)
        '2011/2/1 End
      ' 代表人
        tmpOldCE10 = "" & rsA.Fields(2).Value
        tmpOldCE11 = "" & rsA.Fields(3).Value
        tmpOldCE12 = "" & rsA.Fields(4).Value
        tmpOldCE13 = "" & rsA.Fields(5).Value
        tmpOldCE14 = "" & rsA.Fields(6).Value
        tmpOldCE15 = "" & rsA.Fields(7).Value
      ' 專利商標種類代號
         tmpOldCE39 = "" & rsA.Fields(8).Value
      ' 案件名稱
        tmpOldCE41 = "" & rsA.Fields(9).Value
        tmpOldCE42 = "" & rsA.Fields(10).Value
        tmpOldCE43 = "" & rsA.Fields(11).Value
      ' 商品類別
        tmpOldCE47 = "" & rsA.Fields(12).Value
      ' 商品群組
        tmpOldCE49 = "" & rsA.Fields(13).Value
      ' 正商標號數
        tmpOldCE57 = "" & rsA.Fields(14).Value
        'add by nickc 2006/12/25
        tmpOldCE05 = "" & rsA.Fields(15).Value
        tmpOldCE06 = "" & rsA.Fields(16).Value
        tmpOldCE07 = "" & rsA.Fields(17).Value
        tmpOldCE08 = "" & rsA.Fields(18).Value
        m_TM78 = "" & rsA.Fields(15).Value 'Add By Sindy 2011/8/3
        m_TM79 = "" & rsA.Fields(16).Value 'Add By Sindy 2011/8/3
        m_TM80 = "" & rsA.Fields(17).Value 'Add By Sindy 2011/8/3
        m_TM81 = "" & rsA.Fields(18).Value 'Add By Sindy 2011/8/3
        
'        tmpOldCE26 = PUB_GetCustEachAdd(tmpOldCE05, "1")
'        tmpOldCE27 = PUB_GetCustEachAdd(tmpOldCE05, "2")
'        tmpOldCE28 = PUB_GetCustEachAdd(tmpOldCE05, "3")
'        tmpOldCE29 = PUB_GetCustEachAdd(tmpOldCE06, "1")
'        tmpOldCE30 = PUB_GetCustEachAdd(tmpOldCE06, "2")
'        tmpOldCE31 = PUB_GetCustEachAdd(tmpOldCE06, "3")
'        tmpOldCE32 = PUB_GetCustEachAdd(tmpOldCE07, "1")
'        tmpOldCE33 = PUB_GetCustEachAdd(tmpOldCE07, "2")
'        tmpOldCE34 = PUB_GetCustEachAdd(tmpOldCE07, "3")
'        tmpOldCE35 = PUB_GetCustEachAdd(tmpOldCE08, "1")
'        tmpOldCE36 = PUB_GetCustEachAdd(tmpOldCE08, "2")
'        tmpOldCE37 = PUB_GetCustEachAdd(tmpOldCE08, "3")
        'Modify By Sindy 2011/2/1
        tmpOldCE26 = Trim("" & rsA.Fields(46).Value)
        tmpOldCE27 = Trim("" & rsA.Fields(47).Value)
        tmpOldCE28 = Trim("" & rsA.Fields(48).Value)
        tmpOldCE29 = Trim("" & rsA.Fields(49).Value)
        tmpOldCE30 = Trim("" & rsA.Fields(50).Value)
        tmpOldCE31 = Trim("" & rsA.Fields(51).Value)
        tmpOldCE32 = Trim("" & rsA.Fields(52).Value)
        tmpOldCE33 = Trim("" & rsA.Fields(53).Value)
        tmpOldCE34 = Trim("" & rsA.Fields(54).Value)
        tmpOldCE35 = Trim("" & rsA.Fields(55).Value)
        tmpOldCE36 = Trim("" & rsA.Fields(56).Value)
        tmpOldCE37 = Trim("" & rsA.Fields(57).Value)
        '2011/2/1 End
        tmpOldCE68 = "" & rsA.Fields(19).Value
        tmpOldCE69 = "" & rsA.Fields(20).Value
        tmpOldCE70 = "" & rsA.Fields(21).Value
        tmpOldCE71 = "" & rsA.Fields(22).Value
        tmpOldCE72 = "" & rsA.Fields(23).Value
        tmpOldCE73 = "" & rsA.Fields(24).Value
        tmpOldCE74 = "" & rsA.Fields(25).Value
        tmpOldCE75 = "" & rsA.Fields(26).Value
        tmpOldCE76 = "" & rsA.Fields(27).Value
        tmpOldCE77 = "" & rsA.Fields(28).Value
        tmpOldCE78 = "" & rsA.Fields(29).Value
        tmpOldCE79 = "" & rsA.Fields(30).Value
        tmpOldCE80 = "" & rsA.Fields(31).Value
        tmpOldCE81 = "" & rsA.Fields(32).Value
        tmpOldCE82 = "" & rsA.Fields(33).Value
        tmpOldCE83 = "" & rsA.Fields(34).Value
        tmpOldCE84 = "" & rsA.Fields(35).Value
        tmpOldCE85 = "" & rsA.Fields(36).Value
        tmpOldCE86 = "" & rsA.Fields(37).Value
        tmpOldCE87 = "" & rsA.Fields(38).Value
        tmpOldCE88 = "" & rsA.Fields(39).Value
        tmpOldCE89 = "" & rsA.Fields(40).Value
        tmpOldCE90 = "" & rsA.Fields(41).Value
        tmpOldCE91 = "" & rsA.Fields(42).Value
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    'End
   
   'Modify By Sindy 2012/5/18 Mark暫存變數，因暫存變數是為比對基本檔和畫面上的欄位值是否相同,所以暫存變數不須再存取變更檔資料
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請日
      If IsNull(rsTmp.Fields("CE02")) = False Then
         checkCE03.Value = 1 'Add By Sindy 2012/3/5
         m_CE02 = rsTmp.Fields("CE02")
         '911204 nick
'         tmpOldCE02 = CheckStr(rsTmp.Fields("CE02"))
         textCE02 = ChangeWStringToTString(rsTmp.Fields("CE02"))
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("CE04")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/5
         m_CE04 = rsTmp.Fields("CE04")
         '911204 nick
'         tmpOldCE04 = CheckStr(rsTmp.Fields("CE04"))
         textCE04 = rsTmp.Fields("CE04")
         textCE04_2 = GetCustomer(rsTmp.Fields("CE04"))
         'Add By Cheng 2002/07/16
         '顯示申請人地址
         textCE23.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "1")
         textCE24.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "2")
         textCE25.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "3")
'         tmpOldCE23 = textCE23.Text
'         tmpOldCE24 = textCE24.Text
'         tmpOldCE25 = textCE25.Text
      End If
      'add by nickc 2006/12/25
      If IsNull(rsTmp.Fields("CE05")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/5
         m_CE05 = rsTmp.Fields("CE05")
'         tmpOldCE05 = CheckStr(rsTmp.Fields("CE05"))
         textCE05 = rsTmp.Fields("CE05")
         textCE05_2 = GetCustomer(rsTmp.Fields("CE05"))
         '顯示申請人地址
         textCE26.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "1")
         textCE27.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "2")
         textCE28.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "3")
'         tmpOldCE26 = textCE26.Text
'         tmpOldCE27 = textCE27.Text
'         tmpOldCE28 = textCE28.Text
      End If
      If IsNull(rsTmp.Fields("CE06")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/5
         m_CE06 = rsTmp.Fields("CE06")
'         tmpOldCE06 = CheckStr(rsTmp.Fields("CE06"))
         textCE06 = rsTmp.Fields("CE06")
         textCE06_2 = GetCustomer(rsTmp.Fields("CE06"))
         '顯示申請人地址
         textCE29.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "1")
         textCE30.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "2")
         textCE31.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "3")
'         tmpOldCE29 = textCE29.Text
'         tmpOldCE30 = textCE30.Text
'         tmpOldCE31 = textCE31.Text
      End If
      If IsNull(rsTmp.Fields("CE07")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/5
         m_CE07 = rsTmp.Fields("CE07")
'         tmpOldCE07 = CheckStr(rsTmp.Fields("CE07"))
         textCE07 = rsTmp.Fields("CE07")
         textCE07_2 = GetCustomer(rsTmp.Fields("CE07"))
         '顯示申請人地址
         textCE32.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "1")
         textCE33.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "2")
         textCE34.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "3")
'         tmpOldCE32 = textCE32.Text
'         tmpOldCE33 = textCE33.Text
'         tmpOldCE34 = textCE34.Text
      End If
      If IsNull(rsTmp.Fields("CE08")) = False Then
         checkCE09.Value = 1 'Add By Sindy 2012/3/5
         m_CE08 = rsTmp.Fields("CE08")
'         tmpOldCE08 = CheckStr(rsTmp.Fields("CE08"))
         textCE08 = rsTmp.Fields("CE08")
         textCE08_2 = GetCustomer(rsTmp.Fields("CE08"))
         '顯示申請人地址
         textCE35.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "1")
         textCE36.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "2")
         textCE37.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "3")
'         tmpOldCE35 = textCE35.Text
'         tmpOldCE36 = textCE36.Text
'         tmpOldCE37 = textCE37.Text
      End If
      
      'Add By Sindy 2012/3/5
      '申請人中譯文
      If IsNull(rsTmp.Fields("CE17")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE17 = rsTmp.Fields("CE17")
      End If
      If IsNull(rsTmp.Fields("CE18")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE18 = rsTmp.Fields("CE18")
      End If
      If IsNull(rsTmp.Fields("CE19")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE19 = rsTmp.Fields("CE19")
      End If
      If IsNull(rsTmp.Fields("CE20")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE20 = rsTmp.Fields("CE20")
      End If
      If IsNull(rsTmp.Fields("CE21")) = False Then
         checkCE22.Value = 1 'Add By Sindy 2012/3/5
         textCE21 = rsTmp.Fields("CE21")
      End If
      ' 代表人
      If IsNull(rsTmp.Fields("CE10")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE10 = rsTmp.Fields("CE10")
         '911204 nick
'         tmpOldCE10 = rsTmp.Fields("CE10")
      End If
      If IsNull(rsTmp.Fields("CE11")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE11 = rsTmp.Fields("CE11")
'         tmpOldCE11 = rsTmp.Fields("CE11")
      End If
      If IsNull(rsTmp.Fields("CE12")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE12 = rsTmp.Fields("CE12")
'         tmpOldCE12 = rsTmp.Fields("CE12")
      End If
      If IsNull(rsTmp.Fields("CE13")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE13 = rsTmp.Fields("CE13")
'         tmpOldCE13 = rsTmp.Fields("CE13")
      End If
      If IsNull(rsTmp.Fields("CE14")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE14 = rsTmp.Fields("CE14")
'         tmpOldCE14 = rsTmp.Fields("CE14")
      End If
      If IsNull(rsTmp.Fields("CE15")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE15 = rsTmp.Fields("CE15")
'         tmpOldCE15 = rsTmp.Fields("CE15")
      End If
      'add by nickc 2006/12/25
      If IsNull(rsTmp.Fields("CE68")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE68 = CheckStr(rsTmp.Fields("CE68"))
'         tmpOldCE68 = CheckStr(rsTmp.Fields("CE68"))
      End If
      If IsNull(rsTmp.Fields("CE69")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE69 = CheckStr(rsTmp.Fields("CE69"))
'         tmpOldCE69 = CheckStr(rsTmp.Fields("CE69"))
      End If
      If IsNull(rsTmp.Fields("CE70")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE70 = CheckStr(rsTmp.Fields("CE70"))
'         tmpOldCE70 = CheckStr(rsTmp.Fields("CE70"))
      End If
      If IsNull(rsTmp.Fields("CE71")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE71 = CheckStr(rsTmp.Fields("CE71"))
'         tmpOldCE71 = CheckStr(rsTmp.Fields("CE71"))
      End If
      If IsNull(rsTmp.Fields("CE72")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE72 = CheckStr(rsTmp.Fields("CE72"))
'         tmpOldCE72 = CheckStr(rsTmp.Fields("CE72"))
      End If
      If IsNull(rsTmp.Fields("CE73")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE73 = CheckStr(rsTmp.Fields("CE73"))
'         tmpOldCE73 = CheckStr(rsTmp.Fields("CE73"))
      End If
      If IsNull(rsTmp.Fields("CE74")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE74 = CheckStr(rsTmp.Fields("CE74"))
'         tmpOldCE74 = CheckStr(rsTmp.Fields("CE74"))
      End If
      If IsNull(rsTmp.Fields("CE75")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE75 = CheckStr(rsTmp.Fields("CE75"))
'         tmpOldCE75 = CheckStr(rsTmp.Fields("CE75"))
      End If
      If IsNull(rsTmp.Fields("CE76")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE76 = CheckStr(rsTmp.Fields("CE76"))
'         tmpOldCE76 = CheckStr(rsTmp.Fields("CE76"))
      End If
      If IsNull(rsTmp.Fields("CE77")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE77 = CheckStr(rsTmp.Fields("CE77"))
'         tmpOldCE77 = CheckStr(rsTmp.Fields("CE77"))
      End If
      If IsNull(rsTmp.Fields("CE78")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE78 = CheckStr(rsTmp.Fields("CE78"))
'         tmpOldCE78 = CheckStr(rsTmp.Fields("CE78"))
      End If
      If IsNull(rsTmp.Fields("CE79")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE79 = CheckStr(rsTmp.Fields("CE79"))
'         tmpOldCE79 = CheckStr(rsTmp.Fields("CE79"))
      End If
      If IsNull(rsTmp.Fields("CE80")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE80 = CheckStr(rsTmp.Fields("CE80"))
'         tmpOldCE80 = CheckStr(rsTmp.Fields("CE80"))
      End If
      If IsNull(rsTmp.Fields("CE81")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE81 = CheckStr(rsTmp.Fields("CE81"))
'         tmpOldCE81 = CheckStr(rsTmp.Fields("CE81"))
      End If
      If IsNull(rsTmp.Fields("CE82")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE82 = CheckStr(rsTmp.Fields("CE82"))
'         tmpOldCE82 = CheckStr(rsTmp.Fields("CE82"))
      End If
      If IsNull(rsTmp.Fields("CE83")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE83 = CheckStr(rsTmp.Fields("CE83"))
'         tmpOldCE83 = CheckStr(rsTmp.Fields("CE83"))
      End If
      If IsNull(rsTmp.Fields("CE84")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE84 = CheckStr(rsTmp.Fields("CE84"))
'         tmpOldCE84 = CheckStr(rsTmp.Fields("CE84"))
      End If
      If IsNull(rsTmp.Fields("CE85")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE85 = CheckStr(rsTmp.Fields("CE85"))
'         tmpOldCE85 = CheckStr(rsTmp.Fields("CE85"))
      End If
      If IsNull(rsTmp.Fields("CE86")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE86 = CheckStr(rsTmp.Fields("CE86"))
'         tmpOldCE86 = CheckStr(rsTmp.Fields("CE86"))
      End If
      If IsNull(rsTmp.Fields("CE87")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE87 = CheckStr(rsTmp.Fields("CE87"))
'         tmpOldCE87 = CheckStr(rsTmp.Fields("CE87"))
      End If
      If IsNull(rsTmp.Fields("CE88")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE88 = CheckStr(rsTmp.Fields("CE88"))
'         tmpOldCE88 = CheckStr(rsTmp.Fields("CE88"))
      End If
      If IsNull(rsTmp.Fields("CE89")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE89 = CheckStr(rsTmp.Fields("CE89"))
'         tmpOldCE89 = CheckStr(rsTmp.Fields("CE89"))
      End If
      If IsNull(rsTmp.Fields("CE90")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE90 = CheckStr(rsTmp.Fields("CE90"))
'         tmpOldCE90 = CheckStr(rsTmp.Fields("CE90"))
      End If
      If IsNull(rsTmp.Fields("CE91")) = False Then
         checkCE16.Value = 1 'Add By Sindy 2012/3/5
         textCE91 = CheckStr(rsTmp.Fields("CE91"))
'         tmpOldCE91 = CheckStr(rsTmp.Fields("CE91"))
      End If
      
      ' 申請地址
      If IsNull(rsTmp.Fields("CE23")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE23 = rsTmp.Fields("CE23")
'         tmpOldCE23 = textCE23.Text
      End If
      If IsNull(rsTmp.Fields("CE24")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE24 = rsTmp.Fields("CE24")
'         tmpOldCE24 = textCE24.Text
      End If
      If IsNull(rsTmp.Fields("CE25")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE25 = rsTmp.Fields("CE25")
'         tmpOldCE25 = textCE25.Text
      End If
      'add by nickc 2006/12/25
      If IsNull(rsTmp.Fields("CE26")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE26 = rsTmp.Fields("CE26")
'         tmpOldCE26 = textCE26.Text
      End If
      If IsNull(rsTmp.Fields("CE27")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE27 = rsTmp.Fields("CE27")
'         tmpOldCE27 = textCE27.Text
      End If
      If IsNull(rsTmp.Fields("CE28")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE28 = rsTmp.Fields("CE28")
'         tmpOldCE28 = textCE28.Text
      End If
      If IsNull(rsTmp.Fields("CE29")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE29 = rsTmp.Fields("CE29")
'         tmpOldCE29 = textCE29.Text
      End If
      If IsNull(rsTmp.Fields("CE30")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE30 = rsTmp.Fields("CE30")
'         tmpOldCE30 = textCE30.Text
      End If
      If IsNull(rsTmp.Fields("CE31")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE31 = rsTmp.Fields("CE31")
'         tmpOldCE31 = textCE31.Text
      End If
      If IsNull(rsTmp.Fields("CE32")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE32 = rsTmp.Fields("CE32")
'         tmpOldCE32 = textCE32.Text
      End If
      If IsNull(rsTmp.Fields("CE33")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE33 = rsTmp.Fields("CE33")
'         tmpOldCE33 = textCE33.Text
      End If
      If IsNull(rsTmp.Fields("CE34")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE34 = rsTmp.Fields("CE34")
'         tmpOldCE34 = textCE34.Text
      End If
      If IsNull(rsTmp.Fields("CE35")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE35 = rsTmp.Fields("CE35")
'         tmpOldCE35 = textCE35.Text
      End If
      If IsNull(rsTmp.Fields("CE36")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE36 = rsTmp.Fields("CE36")
'         tmpOldCE36 = textCE36.Text
      End If
      If IsNull(rsTmp.Fields("CE37")) = False Then
         checkCE38.Value = 1 'Add By Sindy 2012/3/5
         textCE37 = rsTmp.Fields("CE37")
'         tmpOldCE37 = textCE37.Text
      End If
      
      ' 專利商標種類代號
      If IsNull(rsTmp.Fields("CE39")) = False Then
         checkCE40.Value = 1 'Add By Sindy 2012/3/5
         m_CE39 = rsTmp.Fields("CE39")
         textCE39 = rsTmp.Fields("CE39")
'         tmpOldCE39 = textCE39.Text
         If IsEmptyText(textCE39) = False Then: textCE39_Validate (False)
      End If
      Select Case m_TM01
      Case "T", "FCT", "CFT", "TF", "TS"
          ' 案件名稱
          If IsNull(rsTmp.Fields("CE41")) = False Then
             checkCE44.Value = 1 'Add By Sindy 2012/3/5
             Me.textCE41_1.Text = rsTmp.Fields("CE41")
             '911204 nick
'             tmpOldCE41 = rsTmp.Fields("CE41")
          End If
      Case Else
          ' 案件名稱
          If IsNull(rsTmp.Fields("CE41")) = False Then
             checkCE44.Value = 1 'Add By Sindy 2012/3/5
             textCE41 = rsTmp.Fields("CE41")
             '911204 nick
'             tmpOldCE41 = rsTmp.Fields("CE41")
          End If
          If IsNull(rsTmp.Fields("CE42")) = False Then
             checkCE44.Value = 1 'Add By Sindy 2012/3/5
             textCE42 = rsTmp.Fields("CE42")
             '911204 nick
'             tmpOldCE42 = rsTmp.Fields("CE42")
          End If
          If IsNull(rsTmp.Fields("CE43")) = False Then
             checkCE44.Value = 1 'Add By Sindy 2012/3/5
             textCE43 = rsTmp.Fields("CE43")
             '911204 nick
'             tmpOldCE43 = rsTmp.Fields("CE43")
          End If
      End Select
      ' 縮減商品
      If IsNull(rsTmp.Fields("CE45")) = False Then
         checkCE46.Value = 1 'Add By Sindy 2012/3/5
         textCE45 = rsTmp.Fields("CE45")
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("CE47")) = False Then
         checkCE48.Value = 1 'Add By Sindy 2012/3/5
         textCE47 = rsTmp.Fields("CE47")
'         tmpOldCE47 = textCE47.Text
      End If
      ' 商品群組
      If IsNull(rsTmp.Fields("CE49")) = False Then
         checkCE50.Value = 1 'Add By Sindy 2012/3/5
         textCE49 = rsTmp.Fields("CE49")
'         tmpOldCE49 = textCE49.Text
      End If
      ' 申請人印鑑
      If IsNull(rsTmp.Fields("CE51")) = False Then
         checkCE52.Value = 1 'Add By Sindy 2012/3/5
         'textCE51 = rsTmp.Fields("CE51")
      End If
      ' 代表人印鑑
      If IsNull(rsTmp.Fields("CE53")) = False Then
         checkCE54.Value = 1 'Add By Sindy 2012/3/5
         'textCE53 = rsTmp.Fields("CE53")
      End If
      '代理人
      If IsNull(rsTmp.Fields("CE55")) = False Then
         checkCE56.Value = 1 'Add By Sindy 2012/3/5
         'textCE55 = rsTmp.Fields("CE55")
      End If
      ' 正商標號數
      If IsNull(rsTmp.Fields("CE57")) = False Then
         checkCE58.Value = 1 'Add By Sindy 2012/3/5
         textCE57 = rsTmp.Fields("CE57")
'         tmpOldCE57 = textCE57.Text
      End If
      ' 圖樣
      If IsNull(rsTmp.Fields("CE59")) = False Then
         checkCE60.Value = 1 'Add By Sindy 2012/3/5
         'textCE59 = rsTmp.Fields("CE59")
      End If
      ' 其它
      If IsNull(rsTmp.Fields("CE61")) = False Then
         checkCE62.Value = 1 'Add By Sindy 2012/3/5
         textCE61 = rsTmp.Fields("CE61")
      End If
      ' 代表人譯文
      If IsNull(rsTmp.Fields("CE63")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE63 = rsTmp.Fields("CE63")
      End If
      If IsNull(rsTmp.Fields("CE64")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE64 = rsTmp.Fields("CE64")
      End If
      'add by nickc 2006/12/25
      If IsNull(rsTmp.Fields("CE92")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE92 = CheckStr(rsTmp.Fields("CE92"))
      End If
      If IsNull(rsTmp.Fields("CE93")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE93 = CheckStr(rsTmp.Fields("CE93"))
      End If
      If IsNull(rsTmp.Fields("CE94")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE94 = CheckStr(rsTmp.Fields("CE94"))
      End If
      If IsNull(rsTmp.Fields("CE95")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE95 = CheckStr(rsTmp.Fields("CE95"))
      End If
      If IsNull(rsTmp.Fields("CE96")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE96 = CheckStr(rsTmp.Fields("CE96"))
      End If
      If IsNull(rsTmp.Fields("CE97")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE97 = CheckStr(rsTmp.Fields("CE97"))
      End If
      If IsNull(rsTmp.Fields("CE98")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE98 = CheckStr(rsTmp.Fields("CE98"))
      End If
      If IsNull(rsTmp.Fields("CE99")) = False Then
         checkCE65.Value = 1 'Add By Sindy 2012/3/5
         textCE99 = CheckStr(rsTmp.Fields("CE99"))
      End If
      
      ' 密碼
      If IsNull(rsTmp.Fields("CE66")) = False Then
         checkCE67.Value = 1 'Add By Sindy 2012/3/5
         textCE66 = rsTmp.Fields("CE66")
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

'Modify By Cheng 2002/11/06
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim rsTmp As New ADODB.Recordset
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   ' 先刪除掉已存在的資料
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.Close
      strSql = "DELETE FROM ChangeEvent " & _
               "WHERE CE01 = '" & m_CE01 & "' "
      cnnConnection.Execute strSql
   Else
      rsTmp.Close
   End If
            
   ' 新增一筆資料到變更事項檔
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO ChangeEvent ("
   For nIndex = 0 To m_CECount - 1
      strTmp = m_CEList(nIndex).fiName
      If IsEmptyText(strTmp) = False And IsEmptyText(m_CEList(nIndex).fiNewData) = False Then
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
   For nIndex = 0 To m_CECount - 1
      strTmp = Empty
      If m_CEList(nIndex).fiType = 0 Then
        'Modify By Cheng 2003/01/21
        '避免單引號發生的存檔錯誤
'         If IsEmptyText(m_CEList(nIndex).fiNewData) = False Then: strTmp = "'" & m_CEList(nIndex).fiNewData & "'"
         If IsEmptyText(m_CEList(nIndex).fiNewData) = False Then: strTmp = "'" & ChgSQL(m_CEList(nIndex).fiNewData) & "'"
      Else
         strTmp = m_CEList(nIndex).fiNewData
      End If
      If IsEmptyText(m_CEList(nIndex).fiName) = False And IsEmptyText(m_CEList(nIndex).fiNewData) = False Then
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
      cnnConnection.Execute strSql
   End If
    'Modify By Cheng 2003/04/10
    '基本檔都要更新
'   ' 91.09.02 modify by louis
'   ' 系統類別為著作權, 案件性質為301的需即時更新到基本檔
'   If m_TM01 = "TC" And m_CP10 = "301" Then
'        'Modify By Cheng 2002/11/06
''      OnSaveServicePractice
    Select Case m_TM01
    Case "CFT", "FCT", "T", "TF"
        If OnSaveTrademark = False Then GoTo ErrorHandler
    Case Else
        If OnSaveServicePractice = False Then GoTo ErrorHandler
    End Select
'   End If
   Set rsTmp = Nothing
'Add By Cheng 2002/11/06
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

'add by nickc 2006/12/19
Private Sub textCE02_GotFocus()
InverseTextBox textCE02
End Sub

' 申請日
Private Sub textCE02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textCE02) = False Then
      If CheckIsTaiwanDate(textCE02, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的申請日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

'add by nickc 2006/12/19
Private Sub textCE04_GotFocus()
   InverseTextBox textCE04
   CloseIme
End Sub
'add by nickc 2007/01/24
Private Sub textCE04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCE04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE04_2 = Empty
   If IsEmptyText(textCE04) = False Then
        'Add By Cheng 2003/04/14
        '補滿9碼
        Me.textCE04.Text = Left(Me.textCE04.Text & "000000000", 9)
       'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
       Dim oState As Boolean
       oState = True
      'textCE04_2 = GetCustomerName(textCE04)
      textCE04_2 = GetCustomerNameAndState(textCE04, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
     End If
      If textCE04_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE04 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      'Add By Cheng 2002/07/16
      '顯示申請人地址
      textCE23.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "1")
      textCE24.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "2")
      textCE25.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "3")
   End If
End Sub
'add by nickc 2006/12/19
Private Sub textCE05_GotFocus()
InverseTextBox textCE05
End Sub
'add by nickc 2007/01/24
Private Sub textCE05_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'add by nickc 2006/12/19
Private Sub textCE05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE05_2 = Empty
   If IsEmptyText(textCE05) = False Then

        Me.textCE05.Text = Left(Me.textCE05.Text & "000000000", 9)
       Dim oState As Boolean
       oState = True
      textCE05_2 = GetCustomerNameAndState(textCE05, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
     End If
      If textCE05_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE05 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      '顯示申請人地址
      textCE26.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "1")
      textCE27.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "2")
      textCE28.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "3")
   End If
End Sub
'add by nickc 2006/12/19
Private Sub textCE06_GotFocus()
InverseTextBox textCE06
End Sub
'add by nickc 2007/01/24
Private Sub textCE06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCE06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE06_2 = Empty
   If IsEmptyText(textCE06) = False Then

        Me.textCE06.Text = Left(Me.textCE06.Text & "000000000", 9)
       Dim oState As Boolean
       oState = True
      textCE06_2 = GetCustomerNameAndState(textCE06, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
     End If
      If textCE06_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE06 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      '顯示申請人地址
      textCE29.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "1")
      textCE30.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "2")
      textCE31.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "3")
   End If
End Sub

'add by nickc 2006/12/19
Private Sub textCE07_GotFocus()
InverseTextBox textCE07
End Sub
'add by nickc 2007/01/24
Private Sub textCE07_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCE07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE07_2 = Empty
   If IsEmptyText(textCE07) = False Then

        Me.textCE07.Text = Left(Me.textCE07.Text & "000000000", 9)
       Dim oState As Boolean
       oState = True
      textCE07_2 = GetCustomerNameAndState(textCE07, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
     End If
      If textCE07_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE07 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      '顯示申請人地址
      textCE32.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "1")
      textCE33.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "2")
      textCE34.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "3")
   End If
End Sub

'add by nickc 2006/12/19
Private Sub textCE08_GotFocus()
InverseTextBox textCE08
End Sub
'add by nickc 2007/01/24
Private Sub textCE08_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCE08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE08_2 = Empty
   If IsEmptyText(textCE08) = False Then

        Me.textCE08.Text = Left(Me.textCE08.Text & "000000000", 9)
       Dim oState As Boolean
       oState = True
      textCE08_2 = GetCustomerNameAndState(textCE08, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
     End If
      If textCE08_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE08 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      '顯示申請人地址
      textCE35.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "1")
      textCE36.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "2")
      textCE37.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "3")
   End If
End Sub
'add by nickc 2006/12/19
Private Sub textCE10_GotFocus()
   InverseTextBox textCE10
   OpenIme
End Sub
Private Sub textCE10_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE10) Then Exit Sub
If CheckLengthIsOK(textCE10.Text, textCE10.MaxLength) = False Then
    MsgBox "代表人1(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE10.SetFocus
    textCE10_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE11_GotFocus()
   InverseTextBox textCE11
   CloseIme
End Sub
Private Sub textCE11_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE11) Then Exit Sub
If CheckLengthIsOK(textCE11.Text, textCE11.MaxLength) = False Then
    MsgBox "代表人1(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE11.SetFocus
    textCE11_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE12_GotFocus()
   InverseTextBox textCE12
   OpenIme
End Sub
Private Sub textCE12_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE12) Then Exit Sub
If CheckLengthIsOK(textCE12.Text, textCE12.MaxLength) = False Then
    MsgBox "代表人1(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE12.SetFocus
    textCE12_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE13_GotFocus()
   InverseTextBox textCE13
   OpenIme
End Sub
Private Sub textCE13_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE13) Then Exit Sub
If CheckLengthIsOK(textCE13.Text, textCE13.MaxLength) = False Then
    MsgBox "代表人2(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE13.SetFocus
    textCE13_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE14_GotFocus()
   InverseTextBox textCE14
   CloseIme
End Sub
Private Sub textCE14_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE14) Then Exit Sub
If CheckLengthIsOK(textCE14.Text, textCE14.MaxLength) = False Then
    MsgBox "代表人2(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE14.SetFocus
    textCE14_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE15_GotFocus()
   InverseTextBox textCE15
   OpenIme
End Sub
Private Sub textCE15_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE15) Then Exit Sub
If CheckLengthIsOK(textCE15.Text, textCE15.MaxLength) = False Then
    MsgBox "代表人2(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE15.SetFocus
    textCE15_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
'add by nickc 2006/12/19
Private Sub textCE17_GotFocus()
   InverseTextBox textCE17
   OpenIme
End Sub
Private Sub textCE17_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE17) Then Exit Sub
If CheckLengthIsOK(textCE17.Text, textCE17.MaxLength) = False Then
    MsgBox "申請人1中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 0
    Me.textCE17.SetFocus
    textCE17_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE18_GotFocus()
   InverseTextBox textCE18
   OpenIme
End Sub
Private Sub textCE18_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE18) Then Exit Sub
If CheckLengthIsOK(textCE18.Text, textCE18.MaxLength) = False Then
    MsgBox "申請人2中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 0
    Me.textCE18.SetFocus
    textCE18_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE19_GotFocus()
   InverseTextBox textCE19
   OpenIme
End Sub
Private Sub textCE19_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE19) Then Exit Sub
If CheckLengthIsOK(textCE19.Text, textCE19.MaxLength) = False Then
    MsgBox "申請人3中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 0
    Me.textCE19.SetFocus
    textCE19_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE20_GotFocus()
   InverseTextBox textCE20
   OpenIme
End Sub
Private Sub textCE20_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE20) Then Exit Sub
If CheckLengthIsOK(textCE20.Text, textCE20.MaxLength) = False Then
    MsgBox "申請人4中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 0
    Me.textCE20.SetFocus
    textCE20_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE21_GotFocus()
   InverseTextBox textCE21
   OpenIme
End Sub
Private Sub textCE21_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE21) Then Exit Sub
If CheckLengthIsOK(textCE21.Text, textCE21.MaxLength) = False Then
    MsgBox "申請人5中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 0
    Me.textCE21.SetFocus
    textCE21_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE23_GotFocus()
   InverseTextBox textCE23
   OpenIme
End Sub
Private Sub textCE23_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE23) Then Exit Sub
If CheckLengthIsOK(textCE23.Text, textCE23.MaxLength) = False Then
    MsgBox "申請地址1(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE23.SetFocus
    textCE23_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE24_GotFocus()
   InverseTextBox textCE24
   CloseIme
End Sub
Private Sub textCE24_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE24) Then Exit Sub
If CheckLengthIsOK(textCE24.Text, textCE24.MaxLength) = False Then
    MsgBox "申請地址1(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE24.SetFocus
    textCE24_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE25_GotFocus()
   InverseTextBox textCE25
   OpenIme
End Sub
Private Sub textCE25_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE25) Then Exit Sub
If CheckLengthIsOK(textCE25.Text, textCE25.MaxLength) = False Then
    MsgBox "申請地址1(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE25.SetFocus
    textCE25_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE26_GotFocus()
   InverseTextBox textCE26
   OpenIme
End Sub
Private Sub textCE26_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE26) Then Exit Sub
If CheckLengthIsOK(textCE26.Text, textCE26.MaxLength) = False Then
    MsgBox "申請地址2(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE26.SetFocus
    textCE26_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE27_GotFocus()
   InverseTextBox textCE27
   CloseIme
End Sub
Private Sub textCE27_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE27) Then Exit Sub
If CheckLengthIsOK(textCE27.Text, textCE27.MaxLength) = False Then
    MsgBox "申請地址2(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE27.SetFocus
    textCE27_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE28_GotFocus()
   InverseTextBox textCE28
   OpenIme
End Sub
Private Sub textCE28_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE28) Then Exit Sub
If CheckLengthIsOK(textCE28.Text, textCE28.MaxLength) = False Then
    MsgBox "申請地址2(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE28.SetFocus
    textCE28_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE29_GotFocus()
   InverseTextBox textCE29
   OpenIme
End Sub
Private Sub textCE29_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE29) Then Exit Sub
If CheckLengthIsOK(textCE29.Text, textCE29.MaxLength) = False Then
    MsgBox "申請地址3(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE29.SetFocus
    textCE29_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE30_GotFocus()
   InverseTextBox textCE30
   CloseIme
End Sub
Private Sub textCE30_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE30) Then Exit Sub
If CheckLengthIsOK(textCE30.Text, textCE30.MaxLength) = False Then
    MsgBox "申請地址3(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE30.SetFocus
    textCE30_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE31_GotFocus()
   InverseTextBox textCE31
   OpenIme
End Sub
Private Sub textCE31_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE31) Then Exit Sub
If CheckLengthIsOK(textCE31.Text, textCE31.MaxLength) = False Then
    MsgBox "申請地址3(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE31.SetFocus
    textCE31_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE32_GotFocus()
   InverseTextBox textCE32
   OpenIme
End Sub
Private Sub textCE32_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE32) Then Exit Sub
If CheckLengthIsOK(textCE32.Text, textCE32.MaxLength) = False Then
    MsgBox "申請地址4(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE32.SetFocus
    textCE32_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE33_GotFocus()
   InverseTextBox textCE33
   CloseIme
End Sub
Private Sub textCE33_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE33) Then Exit Sub
If CheckLengthIsOK(textCE33.Text, textCE33.MaxLength) = False Then
    MsgBox "申請地址4(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE33.SetFocus
    textCE33_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE34_GotFocus()
   InverseTextBox textCE34
   OpenIme
End Sub
Private Sub textCE34_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE34) Then Exit Sub
If CheckLengthIsOK(textCE34.Text, textCE34.MaxLength) = False Then
    MsgBox "申請地址4(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE34.SetFocus
    textCE34_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE35_GotFocus()
   InverseTextBox textCE35
   OpenIme
End Sub
Private Sub textCE35_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE35) Then Exit Sub
If CheckLengthIsOK(textCE35.Text, textCE35.MaxLength) = False Then
    MsgBox "申請地址5(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE35.SetFocus
    textCE35_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE36_GotFocus()
   InverseTextBox textCE36
   CloseIme
End Sub
Private Sub textCE36_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE36) Then Exit Sub
If CheckLengthIsOK(textCE36.Text, textCE36.MaxLength) = False Then
    MsgBox "申請地址5(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE36.SetFocus
    textCE36_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE37_GotFocus()
   InverseTextBox textCE37
   OpenIme
End Sub
Private Sub textCE37_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE37) Then Exit Sub
If CheckLengthIsOK(textCE37.Text, textCE37.MaxLength) = False Then
    MsgBox "申請地址5(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 2
    Me.textCE37.SetFocus
    textCE37_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE39_2_GotFocus()
InverseTextBox textCE39_2
End Sub
Private Sub textCE39_GotFocus()
InverseTextBox textCE39
End Sub

' 商標種類
Private Sub textCE39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textCE39_2 = Empty
   Cancel = False
   If IsEmptyText(textCE39) = False Then
      textCE39_2 = GetTradeMarkName(textCE39, 0)
      If IsEmptyText(textCE39_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商標種類不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Exit Sub
      End If
      If CheckLengthIsOK(textCE39.Text, textCE39.MaxLength) = False Then
         MsgBox "商標種類超過長度!!!", vbExclamation + vbOKOnly
         Me.tabCtrl.Tab = 4
         Me.textCE39.SetFocus
         textCE39_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

' 設定欄位新值
Private Sub SetSRFieldData(ByVal strField As String, ByVal strNewData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   ' 搜尋是否存在該欄位
   bFind = False
   For nIndex = 0 To m_SRListCount - 1
      If m_SRList(nIndex).siName = strField Then
         bFind = True
         m_SRList(nIndex).siData = strNewData
         Exit For
      End If
   Next nIndex
   ' 不存在則新增該欄位
   If bFind = False Then
      ReDim Preserve m_SRList(m_SRListCount + 1)
      m_SRList(m_SRListCount).siName = strField
      m_SRList(m_SRListCount).siData = strNewData
      m_SRList(m_SRListCount).siType = nType
      m_SRListCount = m_SRListCount + 1
   End If
End Sub

' 91.09.02 modify by louis
'Modify By Cheng 2002/11/06
'Private Sub OnSaveServicePractice()
Private Function OnSaveServicePractice() As Boolean
   Dim strSql As String
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strTmp As String
   Dim nIndex As Integer
   '911204 nick
   Dim tmpCp64 As String
   Dim rsnick911204 As New ADODB.Recordset
   
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnSaveServicePractice = True

   '911204 nick
   tmpCp64 = ""
    'Modify By Cheng 2003/04/10
    '取消限制
'   ' 只有系統類別為TC及案件性質為變更301才更新基本檔
'   If m_TM01 <> "TC" Or m_CP10 <> "301" Then
'      Exit Function
'   End If
   
   '911204 nick
   tmpCp64 = " select cp64 from caseprogress where cP09= '" & m_CE01 & "'"
   Set rsnick911204 = New ADODB.Recordset
   rsnick911204.CursorLocation = adUseClient
   rsnick911204.Open tmpCp64, cnnConnection, adOpenStatic, adLockReadOnly
   tmpCp64 = ""
   If rsnick911204.RecordCount > 0 Then
        tmpCp64 = CheckStr(rsnick911204.Fields(0).Value) & " "
   End If
   
   ' 申請人
'   If checkCE09.Value = True Then
   If checkCE09.Value = vbChecked Then
      '911204 nick
      If tmpOldCE04 <> textCE04 Then
            SetSRFieldData "SP08", textCE04, 0
            '911204 nick
'            tmpCp64 = tmpCp64 & "原申請人:" & tmpOldCE04 & " "
            If tmpOldCE04 <> "" Then tmpCp64 = tmpCp64 & "原申請人1:" & tmpOldCE04 & " "
      End If
      'add by nickc 2006/12/25
      If tmpOldCE05 <> textCE05 Then
            SetSRFieldData "SP58", textCE05, 0
            If tmpOldCE05 <> "" Then tmpCp64 = tmpCp64 & "原申請人2:" & tmpOldCE05 & " "
      End If
      If tmpOldCE06 <> textCE06 Then
            SetSRFieldData "SP59", textCE06, 0
            If tmpOldCE06 <> "" Then tmpCp64 = tmpCp64 & "原申請人3:" & tmpOldCE06 & " "
      End If
      If tmpOldCE07 <> textCE07 Then
            SetSRFieldData "SP65", textCE07, 0
            If tmpOldCE07 <> "" Then tmpCp64 = tmpCp64 & "原申請人4:" & tmpOldCE07 & " "
      End If
      If tmpOldCE08 <> textCE08 Then
            SetSRFieldData "SP66", textCE08, 0
            If tmpOldCE08 <> "" Then tmpCp64 = tmpCp64 & "原申請人5:" & tmpOldCE08 & " "
      End If
   
   End If
   ' 申請日
'   If checkCE03.Value = True Then
   If checkCE03.Value = vbChecked Then
      '911204 nick
      If tmpOldCE02 <> DBDATE(textCE02) Then
            SetSRFieldData "SP10", DBDATE(textCE02), 1
            '911204 nick
'            tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
            If tmpOldCE02 <> "" Then tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
      End If
   End If
   '911204 nick 新增代表人 只判斷是否有變更
   If checkCE16.Value = 1 Then
      If textCE10 <> tmpOldCE10 Then
            'edit by nickc 2007/09/11
            'SetSRFieldData "SP42", textCE10, 0
            SetSRFieldData "SP42", ChgSQL(textCE10), 0
'            tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE10 & " "
            If tmpOldCE10 <> "" Then tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE10 & " "
      End If
   End If
   ' 案件名稱
'   If checkCE44.Value = True Then
   If checkCE44.Value = vbChecked Then
        Select Case m_TM01
        Case "TS"
            If tmpOldCE41 <> textCE41_1 Then
                'edit by nickc 2007/09/11
                'SetSRFieldData "SP05", textCE41_1, 0
                SetSRFieldData "SP05", ChgSQL(textCE41_1), 0
                '911204 nick
    '            tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
                If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
            End If
        Case Else
            '911204 nick
            If tmpOldCE41 <> textCE41 Then
                'edit by nickc 2007/09/11
                'SetSRFieldData "SP05", textCE41, 0
                SetSRFieldData "SP05", textCE41, 0
                '911204 nick
    '            tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
                If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
            End If
        End Select
        '911204 nick
        If tmpOldCE42 <> textCE42 Then
            'edit by nickc 2007/09/11
            'SetSRFieldData "SP06", textCE42, 0
            SetSRFieldData "SP06", ChgSQL(textCE42), 0
            '911204 nick
'            tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
            If tmpOldCE42 <> "" Then tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
        End If
        '911204 nick
        If tmpOldCE43 <> textCE43 Then
            'edit by nickc 2007/09/11
            'SetSRFieldData "SP07", textCE43, 0
            SetSRFieldData "SP07", ChgSQL(textCE43), 0
            '911204 nick
'            tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
            If tmpOldCE43 <> "" Then tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
        End If
   End If
   
   ' 更新服務業務基本檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_SRListCount - 1
      strTmp = Empty
      If m_SRList(nIndex).siType = 0 Then
         strTmp = m_SRList(nIndex).siName & " = '" & m_SRList(nIndex).siData & "'"
      Else
         If m_SRList(nIndex).siData = Empty Then
            strTmp = m_SRList(nIndex).siName & " = " & 0
         Else
            strTmp = m_SRList(nIndex).siName & " = " & m_SRList(nIndex).siData
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
   ' 組成SQL語法
   strSql = strSql & " " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "'"
   ' 執行SQL指令
   If bDifference = True Then
      cnnConnection.Execute strSql
      '911226 nick 更新回原本收文號的備註
      'edit by nickc 2005/04/12 單引號判斷
      'StrSql = "update caseprogress set cp64='" & tmpCp64 & "' where cp09='" & m_CE01 & "' "
      strSql = "update caseprogress set cp64='" & ChgSQL(tmpCp64) & "' where cp09='" & m_CE01 & "' "
      cnnConnection.Execute strSql
   End If
   
   ' 清除所佔用的記憶體
   If m_SRListCount > 0 Then
      Erase m_SRList
      m_SRListCount = 0
   End If
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnSaveServicePractice = False
End Function

'Add By Cheng 2003/01/27
Private Function CheckDataValidate() As Boolean
Dim Cancel As Boolean
Dim bUpdate As Boolean 'Add By Sindy 2012/3/7
    
   CheckDataValidate = False
   
   'Add by Amy 2021/12/23檢查畫面的 TextBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
      Exit Function
   End If

   'Add By Sindy 2012/3/7
   bUpdate = False
   If checkCE03.Value = vbChecked Then bUpdate = True
   If checkCE09.Value = vbChecked Then bUpdate = True
   If checkCE22.Value = vbChecked Then bUpdate = True
   If checkCE16.Value = vbChecked Then bUpdate = True
   If checkCE38.Value = vbChecked Then bUpdate = True
   If checkCE40.Value = vbChecked Then bUpdate = True
   If checkCE44.Value = vbChecked Then bUpdate = True
   If checkCE46.Value = vbChecked Then bUpdate = True
   If checkCE48.Value = vbChecked Then bUpdate = True
   If checkCE50.Value = vbChecked Then bUpdate = True
   If checkCE52.Value = vbChecked Then bUpdate = True
   If checkCE54.Value = vbChecked Then bUpdate = True
   If checkCE56.Value = vbChecked Then bUpdate = True
   If checkCE58.Value = vbChecked Then bUpdate = True
   If checkCE60.Value = vbChecked Then bUpdate = True
   If checkCE62.Value = vbChecked Then bUpdate = True
   If checkCE65.Value = vbChecked Then bUpdate = True
   If checkCE67.Value = vbChecked Then bUpdate = True
   If bUpdate = False Then
      'Add By Sindy 2023/9/28
      'MsgBox "請勾選變更項目 !", vbCritical, "檢核資料"
      'Exit Function
      If MsgBox("無勾選任何變更項目，確定要繼續嗎？", vbYesNo) = vbNo Then
         Exit Function
      End If
      '2023/9/28 END
   End If
   '2012/3/7 End
    
    Cancel = False
    '若有勾申請人
    If Me.checkCE09.Value = vbChecked Then
        'add by nickc 2006/12/19
        If textCE04 = "" And textCE05 = "" And textCE06 = "" And textCE07 = "" And textCE08 = "" Then
           MsgBox "有勾選申請人時，申請人不可空白 !", vbCritical
           tabCtrl.Tab = 0
           textCE04.SetFocus
           Exit Function
        End If
        'Add By Sindy 2011/8/3
        If m_CP31 <> "Y" Then 'Add By Sindy 2011/8/23 新案時不檢查
            If ChangeCustomerL(textCE04) = m_TM23 And ChangeCustomerL(textCE05) = m_TM78 And ChangeCustomerL(textCE06) = m_TM79 And ChangeCustomerL(textCE07) = m_TM80 And ChangeCustomerL(textCE08) = m_TM81 Then
               'Modify By Sindy 2018/11/20 變更事項已開放給承辦人勾選,程序操作時不宜直接鎖死訊息,改用詢問方式
               'MsgBox "新申請人編號與目前相同 !", vbCritical
               If MsgBox("新申請人編號與目前相同，確定資料正確嗎？", vbYesNo) = vbNo Then
               '2018/11/20 END
                  tabCtrl.Tab = 0
                  textCE04.SetFocus
                  Exit Function
               End If
            End If
        End If
        '2011/8/3 End
        textCE04_Validate Cancel
        If Cancel = True Then Exit Function
        'add by nickc 2006/12/19
        textCE05_Validate Cancel
        If Cancel = True Then Exit Function
        textCE06_Validate Cancel
        If Cancel = True Then Exit Function
        textCE07_Validate Cancel
        If Cancel = True Then Exit Function
        textCE08_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    'add by nickc 2006/12/19
    If checkCE03.Value = vbChecked Then
        If textCE02 = "" Then
           MsgBox "有勾選申請日時，申請日不可空白 !", vbCritical
           tabCtrl.Tab = 0
           textCE02.SetFocus
           Exit Function
        End If
        textCE02_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE22.Value = vbChecked Then
        If textCE17 = "" And textCE18 = "" And textCE19 = "" And textCE20 = "" And textCE21 = "" Then
           MsgBox "有勾選申請人中譯文時，申請人中譯文不可空白 !", vbCritical
           tabCtrl.Tab = 0
           textCE17.SetFocus
           Exit Function
        End If
        textCE17_Validate Cancel
        If Cancel = True Then Exit Function
        textCE18_Validate Cancel
        If Cancel = True Then Exit Function
        textCE19_Validate Cancel
        If Cancel = True Then Exit Function
        textCE20_Validate Cancel
        If Cancel = True Then Exit Function
        textCE21_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE16.Value = vbChecked Then
        If textCE10 = "" And textCE11 = "" And textCE12 = "" And textCE13 = "" And textCE14 = "" And textCE15 = "" And textCE68 = "" And textCE69 = "" And textCE70 = "" And textCE71 = "" And textCE72 = "" And textCE73 = "" _
           And textCE74 = "" And textCE75 = "" And textCE76 = "" And textCE77 = "" And textCE78 = "" And textCE79 = "" And textCE80 = "" And textCE81 = "" And textCE82 = "" And textCE83 = "" And textCE84 = "" And textCE85 = "" _
           And textCE86 = "" And textCE87 = "" And textCE88 = "" And textCE89 = "" And textCE90 = "" And textCE91 = "" Then
           MsgBox "有勾選代表人時，代表人不可空白 !", vbCritical
           tabCtrl.Tab = 1
           textCE10.SetFocus
           Exit Function
        End If
        textCE10_Validate Cancel
        If Cancel = True Then Exit Function
        textCE11_Validate Cancel
        If Cancel = True Then Exit Function
        textCE12_Validate Cancel
        If Cancel = True Then Exit Function
        textCE13_Validate Cancel
        If Cancel = True Then Exit Function
        textCE14_Validate Cancel
        If Cancel = True Then Exit Function
        textCE15_Validate Cancel
        If Cancel = True Then Exit Function
        textCE68_Validate Cancel
        If Cancel = True Then Exit Function
        textCE69_Validate Cancel
        If Cancel = True Then Exit Function
        textCE70_Validate Cancel
        If Cancel = True Then Exit Function
        textCE71_Validate Cancel
        If Cancel = True Then Exit Function
        textCE72_Validate Cancel
        If Cancel = True Then Exit Function
        textCE73_Validate Cancel
        If Cancel = True Then Exit Function
        textCE74_Validate Cancel
        If Cancel = True Then Exit Function
        textCE75_Validate Cancel
        If Cancel = True Then Exit Function
        textCE76_Validate Cancel
        If Cancel = True Then Exit Function
        textCE77_Validate Cancel
        If Cancel = True Then Exit Function
        textCE78_Validate Cancel
        If Cancel = True Then Exit Function
        textCE79_Validate Cancel
        If Cancel = True Then Exit Function
        textCE80_Validate Cancel
        If Cancel = True Then Exit Function
        textCE81_Validate Cancel
        If Cancel = True Then Exit Function
        textCE82_Validate Cancel
        If Cancel = True Then Exit Function
        textCE83_Validate Cancel
        If Cancel = True Then Exit Function
        textCE84_Validate Cancel
        If Cancel = True Then Exit Function
        textCE85_Validate Cancel
        If Cancel = True Then Exit Function
        textCE86_Validate Cancel
        If Cancel = True Then Exit Function
        textCE87_Validate Cancel
        If Cancel = True Then Exit Function
        textCE88_Validate Cancel
        If Cancel = True Then Exit Function
        textCE89_Validate Cancel
        If Cancel = True Then Exit Function
        textCE90_Validate Cancel
        If Cancel = True Then Exit Function
        textCE91_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE38.Value = vbChecked Then
        If textCE23 = "" And textCE24 = "" And textCE25 = "" And textCE26 = "" And textCE27 = "" And textCE28 = "" And textCE29 = "" And textCE30 = "" And textCE31 = "" And textCE32 = "" And textCE33 = "" And textCE34 = "" _
           And textCE35 = "" And textCE36 = "" And textCE37 = "" Then
           MsgBox "有勾選申請地址時，申請地址不可空白 !", vbCritical
           tabCtrl.Tab = 2
           textCE23.SetFocus
           Exit Function
        End If
        textCE23_Validate Cancel
        If Cancel = True Then Exit Function
        textCE24_Validate Cancel
        If Cancel = True Then Exit Function
        textCE25_Validate Cancel
        If Cancel = True Then Exit Function
        textCE26_Validate Cancel
        If Cancel = True Then Exit Function
        textCE27_Validate Cancel
        If Cancel = True Then Exit Function
        textCE28_Validate Cancel
        If Cancel = True Then Exit Function
        textCE29_Validate Cancel
        If Cancel = True Then Exit Function
        textCE30_Validate Cancel
        If Cancel = True Then Exit Function
        textCE31_Validate Cancel
        If Cancel = True Then Exit Function
        textCE32_Validate Cancel
        If Cancel = True Then Exit Function
        textCE33_Validate Cancel
        If Cancel = True Then Exit Function
        textCE34_Validate Cancel
        If Cancel = True Then Exit Function
        textCE35_Validate Cancel
        If Cancel = True Then Exit Function
        textCE36_Validate Cancel
        If Cancel = True Then Exit Function
        textCE37_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE65.Value = vbChecked Then
        If textCE63 = "" And textCE64 = "" And textCE92 = "" And textCE93 = "" And textCE94 = "" And textCE95 = "" And textCE96 = "" And textCE97 = "" And textCE98 = "" And textCE99 = "" Then
           MsgBox "有勾選申請地址時，申請地址不可空白 !", vbCritical
           tabCtrl.Tab = 3
           textCE63.SetFocus
           Exit Function
        End If
        textCE63_Validate Cancel
        If Cancel = True Then Exit Function
        textCE64_Validate Cancel
        If Cancel = True Then Exit Function
        textCE92_Validate Cancel
        If Cancel = True Then Exit Function
        textCE93_Validate Cancel
        If Cancel = True Then Exit Function
        textCE94_Validate Cancel
        If Cancel = True Then Exit Function
        textCE95_Validate Cancel
        If Cancel = True Then Exit Function
        textCE96_Validate Cancel
        If Cancel = True Then Exit Function
        textCE97_Validate Cancel
        If Cancel = True Then Exit Function
        textCE98_Validate Cancel
        If Cancel = True Then Exit Function
        textCE99_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE58.Value = vbChecked Then
        If textCE57 = "" Then
           MsgBox "有勾選正商標號數時，正商標號數不可空白 !", vbCritical
           tabCtrl.Tab = 4
           textCE57.SetFocus
           Exit Function
        End If
        textCE57_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE40.Value = vbChecked Then
        If textCE39 = "" Then
           MsgBox "有勾選商標種類時，商標種類不可空白 !", vbCritical
           tabCtrl.Tab = 4
           textCE39.SetFocus
           Exit Function
        End If
        textCE39_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE46.Value = vbChecked Then
        If textCE45 = "" Then
           MsgBox "有勾選縮減商品時，縮減商品不可空白 !", vbCritical
           tabCtrl.Tab = 4
           textCE45.SetFocus
           Exit Function
        End If
        textCE45_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE48.Value = vbChecked Then
        If textCE47 = "" Then
           MsgBox "有勾選商品類別時，商品類別不可空白 !", vbCritical
           tabCtrl.Tab = 4
           textCE47.SetFocus
           Exit Function
        End If
        textCE47_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE50.Value = vbChecked Then
        If textCE49 = "" Then
           MsgBox "有勾選商品組群時，商品組群不可空白 !", vbCritical
           tabCtrl.Tab = 4
           textCE49.SetFocus
           Exit Function
        End If
        textCE49_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE62.Value = vbChecked Then
        If textCE61 = "" Then
           MsgBox "有勾選其他時，其他不可空白 !", vbCritical
           tabCtrl.Tab = 4
           textCE61.SetFocus
           Exit Function
        End If
        textCE61_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    If checkCE67.Value = vbChecked Then
        If textCE66 = "" Then
           MsgBox "有勾選密碼時，密碼不可空白 !", vbCritical
           tabCtrl.Tab = 4
           textCE66.SetFocus
           Exit Function
        End If
        textCE66_Validate Cancel
        If Cancel = True Then Exit Function
    End If
    
    If Me.checkCE44.Value = vbChecked Then
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF", "TS"
            If Me.textCE41_1.Text = "" Then
                MsgBox "請輸入案件名稱!!!", vbExclamation + vbOKOnly
                Me.tabCtrl.Tab = 4 'edit by nickc 2006/12/19 1
                Me.textCE41_1.SetFocus
                Exit Function
            End If
            If CheckLengthIsOK(Me.textCE41_1.Text, textCE41_1.MaxLength) = False Then
                MsgBox "案件名稱超過長度!!!", vbExclamation + vbOKOnly
                Me.tabCtrl.Tab = 4 'edit by nickc 2006/12/19 1
                Me.textCE41_1.SetFocus
                textCE41_1_GotFocus
                Exit Function
            End If
        End Select
    End If
    CheckDataValidate = True
End Function

'Add By Cheng 2003/04/10
Private Function OnSaveTrademark() As Boolean
Dim strSql As String
Dim bFirst As Boolean
Dim bDifference As Boolean
Dim strTmp As String
Dim nIndex As Integer
Dim tmpCp64 As String
Dim rsnick911204 As New ADODB.Recordset
   
On Error GoTo ErrorHandler
    
    OnSaveTrademark = True
    tmpCp64 = ""
    tmpCp64 = " select cp64 from caseprogress where cP09= '" & m_CE01 & "'"
    Set rsnick911204 = New ADODB.Recordset
    rsnick911204.CursorLocation = adUseClient
    rsnick911204.Open tmpCp64, cnnConnection, adOpenStatic, adLockReadOnly
    tmpCp64 = ""
    If rsnick911204.RecordCount > 0 Then
         tmpCp64 = CheckStr(rsnick911204.Fields(0).Value) & " "
    End If
    ' 申請人
'    If checkCE09.Value = True Then
    If checkCE09.Value = vbChecked Then
       If tmpOldCE04 <> textCE04 Then
             SetTMFieldData "TM23", textCE04, 0
'             tmpCp64 = tmpCp64 & "原申請人:" & tmpOldCE04 & " "
             If tmpOldCE04 <> "" Then tmpCp64 = tmpCp64 & "原申請人1:" & tmpOldCE04 & " "
            '申請地址
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM24", PUB_GetCustEachAdd(textCE04, 1), 0
             'SetTMFieldData "TM25", PUB_GetCustEachAdd(textCE04, 2), 0
             'SetTMFieldData "TM26", PUB_GetCustEachAdd(textCE04, 3), 0
             SetTMFieldData "TM24", ChgSQL(PUB_GetCustEachAdd(textCE04, 1)), 0
             SetTMFieldData "TM25", ChgSQL(PUB_GetCustEachAdd(textCE04, 2)), 0
             SetTMFieldData "TM26", ChgSQL(PUB_GetCustEachAdd(textCE04, 3)), 0
       End If
       'add by nickc 2006/12/25
       If tmpOldCE05 <> textCE05 Then
             SetTMFieldData "TM78", textCE05, 0
             If tmpOldCE05 <> "" Then tmpCp64 = tmpCp64 & "原申請人2:" & tmpOldCE05 & " "
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM82", PUB_GetCustEachAdd(textCE05, 1), 0
             'SetTMFieldData "TM86", PUB_GetCustEachAdd(textCE05, 2), 0
             'SetTMFieldData "TM90", PUB_GetCustEachAdd(textCE05, 3), 0
             SetTMFieldData "TM82", ChgSQL(PUB_GetCustEachAdd(textCE05, 1)), 0
             SetTMFieldData "TM86", ChgSQL(PUB_GetCustEachAdd(textCE05, 2)), 0
             SetTMFieldData "TM90", ChgSQL(PUB_GetCustEachAdd(textCE05, 3)), 0
       End If
       If tmpOldCE06 <> textCE06 Then
             SetTMFieldData "TM79", textCE06, 0
             If tmpOldCE06 <> "" Then tmpCp64 = tmpCp64 & "原申請人3:" & tmpOldCE06 & " "
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM83", PUB_GetCustEachAdd(textCE06, 1), 0
             'SetTMFieldData "TM87", PUB_GetCustEachAdd(textCE06, 2), 0
             'SetTMFieldData "TM91", PUB_GetCustEachAdd(textCE06, 3), 0
             SetTMFieldData "TM83", ChgSQL(PUB_GetCustEachAdd(textCE06, 1)), 0
             SetTMFieldData "TM87", ChgSQL(PUB_GetCustEachAdd(textCE06, 2)), 0
             SetTMFieldData "TM91", ChgSQL(PUB_GetCustEachAdd(textCE06, 3)), 0
       End If
       If tmpOldCE07 <> textCE07 Then
             SetTMFieldData "TM80", textCE07, 0
             If tmpOldCE07 <> "" Then tmpCp64 = tmpCp64 & "原申請人4:" & tmpOldCE07 & " "
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM84", PUB_GetCustEachAdd(textCE07, 1), 0
             'SetTMFieldData "TM88", PUB_GetCustEachAdd(textCE07, 2), 0
             'SetTMFieldData "TM92", PUB_GetCustEachAdd(textCE07, 3), 0
             SetTMFieldData "TM84", ChgSQL(PUB_GetCustEachAdd(textCE07, 1)), 0
             SetTMFieldData "TM88", ChgSQL(PUB_GetCustEachAdd(textCE07, 2)), 0
             SetTMFieldData "TM92", ChgSQL(PUB_GetCustEachAdd(textCE07, 3)), 0
       End If
       If tmpOldCE08 <> textCE08 Then
             SetTMFieldData "TM81", textCE08, 0
             If tmpOldCE08 <> "" Then tmpCp64 = tmpCp64 & "原申請人5:" & tmpOldCE08 & " "
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM85", PUB_GetCustEachAdd(textCE08, 1), 0
             'SetTMFieldData "TM89", PUB_GetCustEachAdd(textCE08, 2), 0
             'SetTMFieldData "TM93", PUB_GetCustEachAdd(textCE08, 3), 0
             SetTMFieldData "TM85", ChgSQL(PUB_GetCustEachAdd(textCE08, 1)), 0
             SetTMFieldData "TM89", ChgSQL(PUB_GetCustEachAdd(textCE08, 2)), 0
             SetTMFieldData "TM93", ChgSQL(PUB_GetCustEachAdd(textCE08, 3)), 0
       End If
    End If
    ' 申請日
'    If checkCE03.Value = True Then
    If checkCE03.Value = vbChecked Then
       If tmpOldCE02 <> DBDATE(textCE02) Then
             SetTMFieldData "TM11", DBDATE(textCE02), 1
'             tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
             If tmpOldCE02 <> "" Then tmpCp64 = tmpCp64 & "原申請日:" & tmpOldCE02 & " "
       End If
    End If
    '911204 nick 新增代表人 只判斷是否有變更
    If checkCE16.Value = 1 Then
       If textCE10 <> tmpOldCE10 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM47", textCE10, 0
             SetTMFieldData "TM47", ChgSQL(textCE10), 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE10 & " "
             If tmpOldCE10 <> "" Then tmpCp64 = tmpCp64 & "原代表人1(中):" & tmpOldCE10 & " "
       End If
       If textCE11 <> tmpOldCE11 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM48", textCE11, 0
             SetTMFieldData "TM48", ChgSQL(textCE11), 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE11 & " "
             If tmpOldCE11 <> "" Then tmpCp64 = tmpCp64 & "原代表人1(英):" & tmpOldCE11 & " "
       End If
       If textCE12 <> tmpOldCE12 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM49", textCE12, 0
             SetTMFieldData "TM49", ChgSQL(textCE12), 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE12 & " "
             If tmpOldCE12 <> "" Then tmpCp64 = tmpCp64 & "原代表人1(日):" & tmpOldCE12 & " "
       End If
       If textCE13 <> tmpOldCE13 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM50", textCE13, 0
             SetTMFieldData "TM50", ChgSQL(textCE13), 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE13 & " "
             If tmpOldCE13 <> "" Then tmpCp64 = tmpCp64 & "原代表人2(中):" & tmpOldCE13 & " "
       End If
       If textCE14 <> tmpOldCE14 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM51", textCE14, 0
             SetTMFieldData "TM51", ChgSQL(textCE14), 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE14 & " "
             If tmpOldCE14 <> "" Then tmpCp64 = tmpCp64 & "原代表人2(英):" & tmpOldCE14 & " "
       End If
       If textCE15 <> tmpOldCE15 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM52", textCE15, 0
             SetTMFieldData "TM52", ChgSQL(textCE15), 0
'             tmpCp64 = tmpCp64 & "原代表人:" & tmpOldCE15 & " "
             If tmpOldCE15 <> "" Then tmpCp64 = tmpCp64 & "原代表人2(日):" & tmpOldCE15 & " "
       End If
       'add by nickc 2006/12/25
       If textCE68 <> tmpOldCE68 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM94", textCE68, 0
             If tmpOldCE68 <> "" Then tmpCp64 = tmpCp64 & "原代表人3(中):" & tmpOldCE68 & " "
             SetTMFieldData "TM94", ChgSQL(textCE68), 0
       End If
       If textCE69 <> tmpOldCE69 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM95", textCE69, 0
             If tmpOldCE69 <> "" Then tmpCp64 = tmpCp64 & "原代表人3(英):" & tmpOldCE69 & " "
             SetTMFieldData "TM95", ChgSQL(textCE69), 0
       End If
       If textCE70 <> tmpOldCE70 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM96", textCE70, 0
             If tmpOldCE70 <> "" Then tmpCp64 = tmpCp64 & "原代表人3(日):" & tmpOldCE70 & " "
             SetTMFieldData "TM96", ChgSQL(textCE70), 0
       End If
       If textCE71 <> tmpOldCE71 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM97", textCE71, 0
             If tmpOldCE71 <> "" Then tmpCp64 = tmpCp64 & "原代表人4(中):" & tmpOldCE71 & " "
             SetTMFieldData "TM97", ChgSQL(textCE71), 0
       End If
       If textCE72 <> tmpOldCE72 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM98", textCE72, 0
             If tmpOldCE72 <> "" Then tmpCp64 = tmpCp64 & "原代表人4(英):" & tmpOldCE72 & " "
             SetTMFieldData "TM98", ChgSQL(textCE72), 0
       End If
       If textCE73 <> tmpOldCE73 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM99", textCE73, 0
             If tmpOldCE73 <> "" Then tmpCp64 = tmpCp64 & "原代表人4(日):" & tmpOldCE73 & " "
             SetTMFieldData "TM99", ChgSQL(textCE73), 0
       End If
       If textCE68 <> tmpOldCE68 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM94", textCE68, 0
             If tmpOldCE68 <> "" Then tmpCp64 = tmpCp64 & "原代表人5(中):" & tmpOldCE74 & " "
             SetTMFieldData "TM94", ChgSQL(textCE68), 0
       End If
       If textCE69 <> tmpOldCE69 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM95", textCE69, 0
             If tmpOldCE69 <> "" Then tmpCp64 = tmpCp64 & "原代表人5(英):" & tmpOldCE75 & " "
             SetTMFieldData "TM95", ChgSQL(textCE69), 0
       End If
       If textCE70 <> tmpOldCE70 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM96", textCE70, 0
             If tmpOldCE70 <> "" Then tmpCp64 = tmpCp64 & "原代表人5(日):" & tmpOldCE76 & " "
             SetTMFieldData "TM96", ChgSQL(textCE70), 0
       End If
       If textCE71 <> tmpOldCE71 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM97", textCE71, 0
             If tmpOldCE71 <> "" Then tmpCp64 = tmpCp64 & "原代表人6(中):" & tmpOldCE77 & " "
             SetTMFieldData "TM97", ChgSQL(textCE71), 0
       End If
       If textCE72 <> tmpOldCE72 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM98", textCE72, 0
             If tmpOldCE72 <> "" Then tmpCp64 = tmpCp64 & "原代表人6(英):" & tmpOldCE78 & " "
             SetTMFieldData "TM98", ChgSQL(textCE72), 0
       End If
       If textCE73 <> tmpOldCE73 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM99", textCE73, 0
             If tmpOldCE73 <> "" Then tmpCp64 = tmpCp64 & "原代表人6(日):" & tmpOldCE79 & " "
             SetTMFieldData "TM99", ChgSQL(textCE73), 0
       End If
       If textCE68 <> tmpOldCE68 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM94", textCE68, 0
             If tmpOldCE68 <> "" Then tmpCp64 = tmpCp64 & "原代表人7(中):" & tmpOldCE80 & " "
             SetTMFieldData "TM94", ChgSQL(textCE68), 0
       End If
       If textCE69 <> tmpOldCE69 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM95", textCE69, 0
             If tmpOldCE69 <> "" Then tmpCp64 = tmpCp64 & "原代表人7(英):" & tmpOldCE81 & " "
             SetTMFieldData "TM95", ChgSQL(textCE69), 0
       End If
       If textCE70 <> tmpOldCE70 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM96", textCE70, 0
             If tmpOldCE70 <> "" Then tmpCp64 = tmpCp64 & "原代表人7(日):" & tmpOldCE82 & " "
             SetTMFieldData "TM96", ChgSQL(textCE70), 0
       End If
       If textCE71 <> tmpOldCE71 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM97", textCE71, 0
             If tmpOldCE71 <> "" Then tmpCp64 = tmpCp64 & "原代表人8(中):" & tmpOldCE83 & " "
             SetTMFieldData "TM97", ChgSQL(textCE71), 0
       End If
       If textCE72 <> tmpOldCE72 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM98", textCE72, 0
             If tmpOldCE72 <> "" Then tmpCp64 = tmpCp64 & "原代表人8(英):" & tmpOldCE84 & " "
             SetTMFieldData "TM98", ChgSQL(textCE72), 0
       End If
       If textCE73 <> tmpOldCE73 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM99", textCE73, 0
             If tmpOldCE73 <> "" Then tmpCp64 = tmpCp64 & "原代表人8(日):" & tmpOldCE85 & " "
             SetTMFieldData "TM99", ChgSQL(textCE73), 0
       End If
       If textCE68 <> tmpOldCE68 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM94", textCE68, 0
             If tmpOldCE68 <> "" Then tmpCp64 = tmpCp64 & "原代表人9(中):" & tmpOldCE86 & " "
             SetTMFieldData "TM94", ChgSQL(textCE68), 0
       End If
       If textCE69 <> tmpOldCE69 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM95", textCE69, 0
             If tmpOldCE69 <> "" Then tmpCp64 = tmpCp64 & "原代表人9(英):" & tmpOldCE87 & " "
             SetTMFieldData "TM95", ChgSQL(textCE69), 0
       End If
       If textCE70 <> tmpOldCE70 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM96", textCE70, 0
             If tmpOldCE70 <> "" Then tmpCp64 = tmpCp64 & "原代表人9(日):" & tmpOldCE88 & " "
             SetTMFieldData "TM96", ChgSQL(textCE70), 0
       End If
       If textCE71 <> tmpOldCE71 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM97", textCE71, 0
             If tmpOldCE71 <> "" Then tmpCp64 = tmpCp64 & "原代表人10(中):" & tmpOldCE89 & " "
             SetTMFieldData "TM97", ChgSQL(textCE71), 0
       End If
       If textCE72 <> tmpOldCE72 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM98", textCE72, 0
             If tmpOldCE72 <> "" Then tmpCp64 = tmpCp64 & "原代表人10(英):" & tmpOldCE90 & " "
             SetTMFieldData "TM98", ChgSQL(textCE72), 0
       End If
       If textCE73 <> tmpOldCE73 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM99", textCE73, 0
             If tmpOldCE73 <> "" Then tmpCp64 = tmpCp64 & "原代表人10(日):" & tmpOldCE91 & " "
             SetTMFieldData "TM99", ChgSQL(textCE73), 0
       End If
       
    End If
    ' 申請地址
'    If checkCE38.Value = True Then
    If checkCE38.Value = vbChecked Then
         If tmpOldCE23 <> textCE23 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM24", textCE23, 0
             SetTMFieldData "TM24", ChgSQL(textCE23), 0
'             tmpCp64 = tmpCp64 & "原申請中文地址:" & tmpOldCE23 & " "
             If tmpOldCE23 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址1:" & tmpOldCE23 & " "
         End If
         If tmpOldCE24 <> textCE24 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM25", textCE24, 0
             SetTMFieldData "TM25", ChgSQL(textCE24), 0
'             tmpCp64 = tmpCp64 & "原申請英文地址:" & tmpOldCE24 & " "
             If tmpOldCE24 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址1:" & tmpOldCE24 & " "
         End If
         If tmpOldCE25 <> textCE25 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM26", textCE25, 0
             SetTMFieldData "TM26", ChgSQL(textCE25), 0
'             tmpCp64 = tmpCp64 & "原申請日文地址:" & tmpOldCE25 & " "
             If tmpOldCE25 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址1:" & tmpOldCE25 & " "
         End If
         'add by nickc 2006/12/25
         If tmpOldCE26 <> textCE26 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM82", textCE26, 0
             If tmpOldCE26 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址2:" & tmpOldCE26 & " "
             SetTMFieldData "TM82", ChgSQL(textCE26), 0
         End If
         If tmpOldCE27 <> textCE27 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM86", textCE27, 0
             If tmpOldCE27 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址2:" & tmpOldCE27 & " "
             SetTMFieldData "TM86", ChgSQL(textCE27), 0
         End If
         If tmpOldCE28 <> textCE28 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM90", textCE28, 0
             If tmpOldCE28 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址2:" & tmpOldCE28 & " "
             SetTMFieldData "TM90", ChgSQL(textCE28), 0
         End If
         If tmpOldCE29 <> textCE29 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM83", textCE29, 0
             If tmpOldCE29 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址3:" & tmpOldCE29 & " "
             SetTMFieldData "TM83", ChgSQL(textCE29), 0
         End If
         If tmpOldCE30 <> textCE30 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM87", textCE30, 0
             If tmpOldCE30 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址3:" & tmpOldCE30 & " "
             SetTMFieldData "TM87", ChgSQL(textCE30), 0
         End If
         If tmpOldCE31 <> textCE31 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM91", textCE31, 0
             If tmpOldCE31 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址3:" & tmpOldCE31 & " "
             SetTMFieldData "TM91", ChgSQL(textCE31), 0
         End If
         If tmpOldCE32 <> textCE32 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM84", textCE32, 0
             If tmpOldCE32 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址4:" & tmpOldCE32 & " "
             SetTMFieldData "TM84", ChgSQL(textCE32), 0
         End If
         If tmpOldCE33 <> textCE33 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM88", textCE33, 0
             If tmpOldCE33 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址4:" & tmpOldCE33 & " "
             SetTMFieldData "TM88", ChgSQL(textCE33), 0
         End If
         If tmpOldCE34 <> textCE34 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM92", textCE34, 0
             If tmpOldCE34 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址4:" & tmpOldCE34 & " "
             SetTMFieldData "TM92", ChgSQL(textCE34), 0
         End If
         If tmpOldCE35 <> textCE35 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM85", textCE35, 0
             If tmpOldCE35 <> "" Then tmpCp64 = tmpCp64 & "原申請中文地址5:" & tmpOldCE35 & " "
             SetTMFieldData "TM85", ChgSQL(textCE35), 0
         End If
         If tmpOldCE36 <> textCE36 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM89", textCE36, 0
             If tmpOldCE36 <> "" Then tmpCp64 = tmpCp64 & "原申請英文地址5:" & tmpOldCE36 & " "
             SetTMFieldData "TM89", ChgSQL(textCE36), 0
         End If
         If tmpOldCE37 <> textCE37 Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM93", textCE37, 0
             If tmpOldCE37 <> "" Then tmpCp64 = tmpCp64 & "原申請日文地址5:" & tmpOldCE37 & " "
             SetTMFieldData "TM93", ChgSQL(textCE37), 0
         End If
    End If
    '正商標號數
'    If checkCE58.Value = True Then
    If checkCE58.Value = vbChecked Then
         If tmpOldCE57 <> textCE57 Then
             SetTMFieldData "TM27", textCE57, 0
'             tmpCp64 = tmpCp64 & "原正商標號數:" & tmpOldCE57 & " "
             If tmpOldCE57 <> "" Then tmpCp64 = tmpCp64 & "原正商標號數:" & tmpOldCE57 & " "
         End If
    End If
    '商標種類
'    If checkCE40.Value = True Then
    If checkCE40.Value = vbChecked Then
         If tmpOldCE39 <> textCE39 Then
             SetTMFieldData "TM08", textCE39, 0
'             tmpCp64 = tmpCp64 & "原商標種類:" & tmpOldCE39 & " "
             If tmpOldCE39 <> "" Then tmpCp64 = tmpCp64 & "原商標種類:" & tmpOldCE39 & " "
            'Add By Cheng 2003/09/09
            '聯合商標變更為正商標, 清除基本檔的正商標號數
            If (tmpOldCE39 = "2" And Me.textCE39.Text = "1") Or (tmpOldCE39 = "5" And Me.textCE39.Text = "4") Then
                SetTMFieldData "TM27", "", 0
            End If
         End If
    End If
    ' 案件名稱
'    If checkCE44.Value = True Then
    If checkCE44.Value = vbChecked Then
'         '911204 nick
'         If tmpOldCE41 <> textCE41 Then
'             SetTMFieldData "TM05", textCE41, 0
'             '911204 nick
''             tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
'             If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件中文名稱:" & tmpOldCE41 & " "
'         End If
         If tmpOldCE41 <> Me.textCE41_1.Text Then
             'edit by nickc 2007/09/11
             'SetTMFieldData "TM05", Me.textCE41_1.Text, 0
             SetTMFieldData "TM05", ChgSQL(Me.textCE41_1.Text), 0
             If tmpOldCE41 <> "" Then tmpCp64 = tmpCp64 & "原案件名稱:" & tmpOldCE41 & " "
         End If
'         '911204 nick
'         If tmpOldCE42 <> textCE42 Then
'             SetTMFieldData "TM06", textCE42, 0
'             '911204 nick
''             tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
'             If tmpOldCE42 <> "" Then tmpCp64 = tmpCp64 & "原案件英文名稱:" & tmpOldCE42 & " "
'         End If
'         '911204 nick
'         If tmpOldCE43 <> textCE43 Then
'             SetTMFieldData "TM07", textCE43, 0
'             '911204 nick
''             tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
'             If tmpOldCE43 <> "" Then tmpCp64 = tmpCp64 & "原案件日文名稱:" & tmpOldCE43 & " "
'         End If
    End If
    '商品類別
'    If checkCE48.Value = True Then
    If checkCE48.Value = vbChecked Then
         If tmpOldCE47 <> textCE47 Then
             SetTMFieldData "TM09", textCE47, 0
'             tmpCp64 = tmpCp64 & "原商標類別:" & tmpOldCE47 & " "
             If tmpOldCE47 <> "" Then tmpCp64 = tmpCp64 & "原商標類別:" & tmpOldCE47 & " "
         End If
    End If
    '商品群組
'    If checkCE50.Value = True Then
    If checkCE50.Value = vbChecked Then
         If tmpOldCE49 <> textCE49 Then
             SetTMFieldData "TM32", textCE49, 0
'             tmpCp64 = tmpCp64 & "原商標群組:" & tmpOldCE49 & " "
             If tmpOldCE49 <> "" Then tmpCp64 = tmpCp64 & "原商標群組:" & tmpOldCE49 & " "
         End If
    End If
    ' 更新商標基本檔
    strSql = "UPDATE Trademark SET "
    bFirst = True
    bDifference = False
    For nIndex = 0 To m_TMListCount - 1
       strTmp = Empty
       If m_TMList(nIndex).tiType = 0 Then
          strTmp = m_TMList(nIndex).tiName & " = '" & m_TMList(nIndex).tiData & "'"
       Else
          If m_TMList(nIndex).tiData = Empty Then
             strTmp = m_TMList(nIndex).tiName & " = " & 0
          Else
             strTmp = m_TMList(nIndex).tiName & " = " & m_TMList(nIndex).tiData
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
    ' 組成SQL語法
    strSql = strSql & " " & _
                   "WHERE TM01 = '" & m_TM01 & "' AND " & _
                         "TM02 = '" & m_TM02 & "' AND " & _
                         "TM03 = '" & m_TM03 & "' AND " & _
                         "TM04 = '" & m_TM04 & "'"
    ' 執行SQL指令
    If bDifference = True Then
       cnnConnection.Execute strSql
       '911226 nick 更新回原本收文號的備註
       'edit by nickc 2005/04/12 單引號判斷
       'StrSql = "update caseprogress set cp64='" & tmpCp64 & "' where cp09='" & m_CE01 & "' "
       strSql = "update caseprogress set cp64='" & ChgSQL(tmpCp64) & "' where cp09='" & m_CE01 & "' "
       cnnConnection.Execute strSql
    End If
    
    ' 清除所佔用的記憶體
    If m_TMListCount > 0 Then
       Erase m_TMList
       m_TMListCount = 0
    End If
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnSaveTrademark = False
End Function

' 設定欄位新值
Private Sub SetTMFieldData(ByVal strField As String, ByVal strNewData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   ' 搜尋是否存在該欄位
   bFind = False
   For nIndex = 0 To m_TMListCount - 1
      If m_TMList(nIndex).tiName = strField Then
         bFind = True
         m_TMList(nIndex).tiData = strNewData
         Exit For
      End If
   Next nIndex
   ' 不存在則新增該欄位
   If bFind = False Then
      ReDim Preserve m_TMList(m_TMListCount + 1)
      m_TMList(m_TMListCount).tiName = strField
      m_TMList(m_TMListCount).tiData = strNewData
      m_TMList(m_TMListCount).tiType = nType
      m_TMListCount = m_TMListCount + 1
   End If
End Sub

Private Sub textCE41_1_GotFocus()
    TextInverse Me.textCE41_1
End Sub

'add by nickc 2006/12/19
Private Sub textCE41_GotFocus()
InverseTextBox textCE41
End Sub
Private Sub textCE41_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE41) Then Exit Sub
If CheckLengthIsOK(textCE41.Text, textCE41.MaxLength) = False Then
    MsgBox "案件名稱(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 4
    Me.textCE41.SetFocus
    textCE41_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE42_GotFocus()
InverseTextBox textCE42
End Sub
Private Sub textCE42_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE42) Then Exit Sub
If CheckLengthIsOK(textCE42.Text, textCE42.MaxLength) = False Then
    MsgBox "案件名稱(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 4
    Me.textCE42.SetFocus
    textCE42_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE43_GotFocus()
InverseTextBox textCE43
End Sub
Private Sub textCE43_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE43) Then Exit Sub
If CheckLengthIsOK(textCE43.Text, textCE43.MaxLength) = False Then
    MsgBox "案件名稱(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 4
    Me.textCE43.SetFocus
    textCE43_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub textCE45_GotFocus()
InverseTextBox textCE45
End Sub
Private Sub textCE45_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE45) Then Exit Sub
If CheckLengthIsOK(textCE45.Text, textCE45.MaxLength) = False Then
    MsgBox "縮減商品超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 4
    Me.textCE45.SetFocus
    textCE45_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE47_GotFocus()
InverseTextBox textCE47
End Sub
Private Sub textCE47_Validate(Cancel As Boolean)
'add by nickc 2005/06/03
textCE47 = Replace(textCE47, " ", "")
Cancel = False
If IsEmpty(textCE47) Then Exit Sub
If CheckLengthIsOK(textCE47.Text, textCE47.MaxLength) = False Then
    MsgBox "商品類別超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 4
    Me.textCE47.SetFocus
    textCE47_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
'add by nickc 2006/12/19
Private Sub textCE49_GotFocus()
InverseTextBox textCE49
End Sub
Private Sub textCE49_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE49) Then Exit Sub
If CheckLengthIsOK(textCE49.Text, textCE49.MaxLength) = False Then
    MsgBox "商品組群超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 4
    Me.textCE49.SetFocus
    textCE49_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE57_GotFocus()
   InverseTextBox textCE57
   CloseIme
End Sub
Private Sub textCE57_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE57) Then Exit Sub
If CheckLengthIsOK(textCE57.Text, textCE57.MaxLength) = False Then
    MsgBox "正商標號數超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 4
    Me.textCE57.SetFocus
    textCE57_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE61_GotFocus()
InverseTextBox textCE61
End Sub
Private Sub textCE61_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE61) Then Exit Sub
If CheckLengthIsOK(textCE61.Text, textCE61.MaxLength) = False Then
    MsgBox "其他超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 4
    Me.textCE61.SetFocus
    textCE61_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE63_GotFocus()
   InverseTextBox textCE63
   OpenIme
End Sub
Private Sub textCE63_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE63) Then Exit Sub
If CheckLengthIsOK(textCE63.Text, textCE63.MaxLength) = False Then
    MsgBox "代表人1中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE63.SetFocus
    textCE63_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE64_GotFocus()
   InverseTextBox textCE64
   OpenIme
End Sub
Private Sub textCE64_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE64) Then Exit Sub
If CheckLengthIsOK(textCE64.Text, textCE64.MaxLength) = False Then
    MsgBox "代表人2中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE64.SetFocus
    textCE64_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE66_GotFocus()
InverseTextBox textCE66
End Sub
Private Sub textCE66_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE66) Then Exit Sub
If CheckLengthIsOK(textCE66.Text, textCE66.MaxLength) = False Then
    MsgBox "密碼超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 4
    Me.textCE66.SetFocus
    textCE66_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE68_GotFocus()
   InverseTextBox textCE68
   OpenIme
End Sub
Private Sub textCE68_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE68) Then Exit Sub
If CheckLengthIsOK(textCE68.Text, textCE68.MaxLength) = False Then
    MsgBox "代表人3(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE68.SetFocus
    textCE68_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE69_GotFocus()
   InverseTextBox textCE69
   CloseIme
End Sub
Private Sub textCE69_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE69) Then Exit Sub
If CheckLengthIsOK(textCE69.Text, textCE69.MaxLength) = False Then
    MsgBox "代表人3(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE69.SetFocus
    textCE69_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE70_GotFocus()
   InverseTextBox textCE70
   OpenIme
End Sub
Private Sub textCE70_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE70) Then Exit Sub
If CheckLengthIsOK(textCE70.Text, textCE70.MaxLength) = False Then
    MsgBox "代表人3(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE70.SetFocus
    textCE70_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE71_GotFocus()
   InverseTextBox textCE71
   OpenIme
End Sub
Private Sub textCE71_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE71) Then Exit Sub
If CheckLengthIsOK(textCE71.Text, textCE70.MaxLength) = False Then
    MsgBox "代表人4(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE71.SetFocus
    textCE71_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE72_GotFocus()
   InverseTextBox textCE72
   CloseIme
End Sub
Private Sub textCE72_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE72) Then Exit Sub
If CheckLengthIsOK(textCE72.Text, textCE72.MaxLength) = False Then
    MsgBox "代表人4(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE72.SetFocus
    textCE72_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE73_GotFocus()
   InverseTextBox textCE73
   OpenIme
End Sub
Private Sub textCE73_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE73) Then Exit Sub
If CheckLengthIsOK(textCE73.Text, textCE73.MaxLength) = False Then
    MsgBox "代表人4(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE73.SetFocus
    textCE73_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE74_GotFocus()
   InverseTextBox textCE74
   OpenIme
End Sub
Private Sub textCE74_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE74) Then Exit Sub
If CheckLengthIsOK(textCE74.Text, textCE74.MaxLength) = False Then
    MsgBox "代表人5(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE74.SetFocus
    textCE74_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE75_GotFocus()
   InverseTextBox textCE75
   CloseIme
End Sub
Private Sub textCE75_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE75) Then Exit Sub
If CheckLengthIsOK(textCE75.Text, textCE75.MaxLength) = False Then
    MsgBox "代表人5(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE75.SetFocus
    textCE75_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE76_GotFocus()
   InverseTextBox textCE76
   OpenIme
End Sub
Private Sub textCE76_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE76) Then Exit Sub
If CheckLengthIsOK(textCE76.Text, textCE76.MaxLength) = False Then
    MsgBox "代表人5(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE76.SetFocus
    textCE76_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE77_GotFocus()
   InverseTextBox textCE77
   OpenIme
End Sub
Private Sub textCE77_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE77) Then Exit Sub
If CheckLengthIsOK(textCE77.Text, textCE77.MaxLength) = False Then
    MsgBox "代表人6(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE77.SetFocus
    textCE77_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE78_GotFocus()
   InverseTextBox textCE78
End Sub
Private Sub textCE78_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE78) Then Exit Sub
If CheckLengthIsOK(textCE78.Text, textCE78.MaxLength) = False Then
    MsgBox "代表人6(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE78.SetFocus
    textCE78_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE79_GotFocus()
   InverseTextBox textCE79
   OpenIme
End Sub
Private Sub textCE79_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE79) Then Exit Sub
If CheckLengthIsOK(textCE79.Text, textCE79.MaxLength) = False Then
    MsgBox "代表人6(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE79.SetFocus
    textCE79_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE80_GotFocus()
   InverseTextBox textCE80
   OpenIme
End Sub
Private Sub textCE80_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE80) Then Exit Sub
If CheckLengthIsOK(textCE80.Text, textCE80.MaxLength) = False Then
    MsgBox "代表人7(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE80.SetFocus
    textCE80_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE81_GotFocus()
   InverseTextBox textCE81
   CloseIme
End Sub
Private Sub textCE81_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE81) Then Exit Sub
If CheckLengthIsOK(textCE81.Text, textCE81.MaxLength) = False Then
    MsgBox "代表人7(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE81.SetFocus
    textCE81_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE82_GotFocus()
   InverseTextBox textCE82
   OpenIme
End Sub
Private Sub textCE82_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE82) Then Exit Sub
If CheckLengthIsOK(textCE82.Text, textCE82.MaxLength) = False Then
    MsgBox "代表人7(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE82.SetFocus
    textCE82_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE83_GotFocus()
   InverseTextBox textCE83
   OpenIme
End Sub
Private Sub textCE83_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE83) Then Exit Sub
If CheckLengthIsOK(textCE83.Text, textCE83.MaxLength) = False Then
    MsgBox "代表人8(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE83.SetFocus
    textCE83_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE84_GotFocus()
   InverseTextBox textCE84
   CloseIme
End Sub
Private Sub textCE84_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE84) Then Exit Sub
If CheckLengthIsOK(textCE84.Text, textCE84.MaxLength) = False Then
    MsgBox "代表人8(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE84.SetFocus
    textCE84_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE85_GotFocus()
   InverseTextBox textCE85
   OpenIme
End Sub
Private Sub textCE85_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE85) Then Exit Sub
If CheckLengthIsOK(textCE85.Text, textCE85.MaxLength) = False Then
    MsgBox "代表人8(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE85.SetFocus
    textCE85_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE86_GotFocus()
   InverseTextBox textCE86
   OpenIme
End Sub
Private Sub textCE86_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE86) Then Exit Sub
If CheckLengthIsOK(textCE86.Text, textCE86.MaxLength) = False Then
    MsgBox "代表人9(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE86.SetFocus
    textCE86_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE87_GotFocus()
   InverseTextBox textCE87
   CloseIme
End Sub
Private Sub textCE87_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE87) Then Exit Sub
If CheckLengthIsOK(textCE87.Text, textCE87.MaxLength) = False Then
    MsgBox "代表人9(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE87.SetFocus
    textCE87_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE88_GotFocus()
   InverseTextBox textCE88
   OpenIme
End Sub
Private Sub textCE88_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE88) Then Exit Sub
If CheckLengthIsOK(textCE88.Text, textCE88.MaxLength) = False Then
    MsgBox "代表人9(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE88.SetFocus
    textCE88_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE89_GotFocus()
   InverseTextBox textCE89
   OpenIme
End Sub
Private Sub textCE89_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE89) Then Exit Sub
If CheckLengthIsOK(textCE89.Text, textCE89.MaxLength) = False Then
    MsgBox "代表人10(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE89.SetFocus
    textCE89_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE90_GotFocus()
   InverseTextBox textCE90
   CloseIme
End Sub
Private Sub textCE90_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE90) Then Exit Sub
If CheckLengthIsOK(textCE90.Text, textCE90.MaxLength) = False Then
    MsgBox "代表人10(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE90.SetFocus
    textCE90_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE91_GotFocus()
   InverseTextBox textCE91
   OpenIme
End Sub
Private Sub textCE91_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE91) Then Exit Sub
If CheckLengthIsOK(textCE91.Text, textCE91.MaxLength) = False Then
    MsgBox "代表人10(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 1
    Me.textCE91.SetFocus
    textCE91_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE92_GotFocus()
   InverseTextBox textCE92
   OpenIme
End Sub
Private Sub textCE92_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE92) Then Exit Sub
If CheckLengthIsOK(textCE92.Text, textCE92.MaxLength) = False Then
    MsgBox "代表人3中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE92.SetFocus
    textCE92_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE93_GotFocus()
   InverseTextBox textCE93
   OpenIme
End Sub
Private Sub textCE93_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE93) Then Exit Sub
If CheckLengthIsOK(textCE93.Text, textCE93.MaxLength) = False Then
    MsgBox "代表人4中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE93.SetFocus
    textCE93_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE94_GotFocus()
   InverseTextBox textCE94
   OpenIme
End Sub
Private Sub textCE94_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE94) Then Exit Sub
If CheckLengthIsOK(textCE94.Text, textCE94.MaxLength) = False Then
    MsgBox "代表人5中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE94.SetFocus
    textCE94_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE95_GotFocus()
   InverseTextBox textCE95
   OpenIme
End Sub
Private Sub textCE95_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE95) Then Exit Sub
If CheckLengthIsOK(textCE95.Text, textCE95.MaxLength) = False Then
    MsgBox "代表人6中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE95.SetFocus
    textCE95_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE96_GotFocus()
   InverseTextBox textCE96
   OpenIme
End Sub
Private Sub textCE96_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE96) Then Exit Sub
If CheckLengthIsOK(textCE96.Text, textCE96.MaxLength) = False Then
    MsgBox "代表人7中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE96.SetFocus
    textCE96_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE97_GotFocus()
   InverseTextBox textCE97
   OpenIme
End Sub
Private Sub textCE97_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE97) Then Exit Sub
If CheckLengthIsOK(textCE97.Text, textCE97.MaxLength) = False Then
    MsgBox "代表人8中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE97.SetFocus
    textCE97_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE98_GotFocus()
   InverseTextBox textCE98
   OpenIme
End Sub
Private Sub textCE98_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE98) Then Exit Sub
If CheckLengthIsOK(textCE98.Text, textCE98.MaxLength) = False Then
    MsgBox "代表人9中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE98.SetFocus
    textCE98_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCE99_GotFocus()
   InverseTextBox textCE99
   OpenIme
End Sub
Private Sub textCE99_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE99) Then Exit Sub
If CheckLengthIsOK(textCE99.Text, textCE99.MaxLength) = False Then
    MsgBox "代表人10中譯文超過長度!!!", vbExclamation + vbOKOnly
    Me.tabCtrl.Tab = 3
    Me.textCE99.SetFocus
    textCE99_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub textCP10_GotFocus()
InverseTextBox textCP10
End Sub
Private Sub textCP13_GotFocus()
InverseTextBox textCP13
End Sub
Private Sub textCP14_GotFocus()
InverseTextBox textCP14
End Sub
Private Sub textTM08_GotFocus()
InverseTextBox textTM08
End Sub
Private Sub textTM45_GotFocus()
InverseTextBox textTM45
End Sub
Private Sub textTMKey_GotFocus()
InverseTextBox textTMKey
End Sub
