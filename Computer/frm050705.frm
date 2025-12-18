VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050705 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外代理人資料維護"
   ClientHeight    =   6120
   ClientLeft      =   110
   ClientTop       =   940
   ClientWidth     =   9140
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9140
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   120
      TabIndex        =   126
      Top             =   990
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   9013
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   7
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm050705.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Labeld1(0)"
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(2)=   "Label18"
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(4)=   "Label13"
      Tab(0).Control(5)=   "Label2(0)"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label27"
      Tab(0).Control(8)=   "Label29"
      Tab(0).Control(9)=   "Label30"
      Tab(0).Control(10)=   "Label2(1)"
      Tab(0).Control(11)=   "Label58"
      Tab(0).Control(12)=   "Label59"
      Tab(0).Control(13)=   "Label60"
      Tab(0).Control(14)=   "Label61"
      Tab(0).Control(15)=   "Label41(32)"
      Tab(0).Control(16)=   "Label41(2)"
      Tab(0).Control(17)=   "Label41(3)"
      Tab(0).Control(18)=   "Label41(4)"
      Tab(0).Control(19)=   "Label41(5)"
      Tab(0).Control(20)=   "Label41(6)"
      Tab(0).Control(21)=   "Label41(10)"
      Tab(0).Control(22)=   "Label41(13)"
      Tab(0).Control(23)=   "Label41(14)"
      Tab(0).Control(24)=   "Label41(15)"
      Tab(0).Control(25)=   "Label1(24)"
      Tab(0).Control(26)=   "Label1(53)"
      Tab(0).Control(27)=   "textFA23"
      Tab(0).Control(28)=   "textFA06"
      Tab(0).Control(29)=   "textFA04"
      Tab(0).Control(30)=   "textFA17"
      Tab(0).Control(31)=   "textFA19"
      Tab(0).Control(32)=   "textFA20"
      Tab(0).Control(33)=   "textFA21"
      Tab(0).Control(34)=   "textFA22"
      Tab(0).Control(35)=   "textFA18"
      Tab(0).Control(36)=   "textFA05"
      Tab(0).Control(37)=   "textFA63"
      Tab(0).Control(38)=   "textFA64"
      Tab(0).Control(39)=   "textFA65"
      Tab(0).Control(40)=   "textFA70"
      Tab(0).Control(41)=   "Label1(2)"
      Tab(0).Control(42)=   "Label1(4)"
      Tab(0).Control(43)=   "Label1(5)"
      Tab(0).Control(44)=   "txtXYS03"
      Tab(0).Control(45)=   "LblSourceN"
      Tab(0).Control(46)=   "textFA11"
      Tab(0).Control(47)=   "textFA10"
      Tab(0).Control(48)=   "textFA03"
      Tab(0).Control(49)=   "textFA55"
      Tab(0).Control(50)=   "textFA10_2"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textFA55_2"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "TextFA76"
      Tab(0).Control(53)=   "textFA77"
      Tab(0).Control(54)=   "textFA100"
      Tab(0).Control(55)=   "Combo1"
      Tab(0).Control(56)=   "cmdTW(0)"
      Tab(0).Control(57)=   "cboSource"
      Tab(0).Control(58)=   "txtXYS02"
      Tab(0).Control(59)=   "cmdIntroduce"
      Tab(0).ControlCount=   60
      TabCaption(1)   =   "聯絡資料"
      TabPicture(1)   =   "frm050705.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFA(105)"
      Tab(1).Control(1)=   "textFA14"
      Tab(1).Control(2)=   "textFA82"
      Tab(1).Control(3)=   "textFA81"
      Tab(1).Control(4)=   "textFA80"
      Tab(1).Control(5)=   "textFA79"
      Tab(1).Control(6)=   "textFA12"
      Tab(1).Control(7)=   "textFA15"
      Tab(1).Control(8)=   "textFA13"
      Tab(1).Control(9)=   "textFA16"
      Tab(1).Control(10)=   "textFA36"
      Tab(1).Control(11)=   "textFA34"
      Tab(1).Control(12)=   "textFA32"
      Tab(1).Control(13)=   "textFA35"
      Tab(1).Control(14)=   "textFA33"
      Tab(1).Control(15)=   "textFA57"
      Tab(1).Control(16)=   "textFA53"
      Tab(1).Control(17)=   "textFA08"
      Tab(1).Control(18)=   "textFA58"
      Tab(1).Control(19)=   "textFA56"
      Tab(1).Control(20)=   "textFA78"
      Tab(1).Control(21)=   "textFA54"
      Tab(1).Control(22)=   "textFA52"
      Tab(1).Control(23)=   "textFA09"
      Tab(1).Control(24)=   "textFA07"
      Tab(1).Control(25)=   "Label71"
      Tab(1).Control(26)=   "Label24(2)"
      Tab(1).Control(27)=   "Label24(1)"
      Tab(1).Control(28)=   "Label24(0)"
      Tab(1).Control(29)=   "Label66"
      Tab(1).Control(30)=   "Label65"
      Tab(1).Control(31)=   "Label64"
      Tab(1).Control(32)=   "Label63"
      Tab(1).Control(33)=   "Label41(12)"
      Tab(1).Control(34)=   "Label41(11)"
      Tab(1).Control(35)=   "Label41(9)"
      Tab(1).Control(36)=   "Label41(8)"
      Tab(1).Control(37)=   "Label41(7)"
      Tab(1).Control(38)=   "Label62"
      Tab(1).Control(39)=   "Label25"
      Tab(1).Control(40)=   "Label12"
      Tab(1).Control(41)=   "Label3"
      Tab(1).Control(42)=   "Label36"
      Tab(1).Control(43)=   "Label37"
      Tab(1).Control(44)=   "Label38"
      Tab(1).Control(45)=   "Label39"
      Tab(1).Control(46)=   "Label40"
      Tab(1).Control(47)=   "Label41(0)"
      Tab(1).Control(48)=   "Label4"
      Tab(1).ControlCount=   49
      TabCaption(2)   =   "專利"
      TabPicture(2)   =   "frm050705.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Labeld1(3)"
      Tab(2).Control(1)=   "Labeld1(4)"
      Tab(2).Control(2)=   "Labeld1(5)"
      Tab(2).Control(3)=   "Label31"
      Tab(2).Control(4)=   "Label22"
      Tab(2).Control(5)=   "Label21"
      Tab(2).Control(6)=   "Label20"
      Tab(2).Control(7)=   "Label19"
      Tab(2).Control(8)=   "Label15"
      Tab(2).Control(9)=   "Label14"
      Tab(2).Control(10)=   "Label11(0)"
      Tab(2).Control(11)=   "Label10"
      Tab(2).Control(12)=   "Label9"
      Tab(2).Control(13)=   "Label8(0)"
      Tab(2).Control(14)=   "Labeld1(10)"
      Tab(2).Control(15)=   "Label1(156)"
      Tab(2).Control(16)=   "lblFA(95)"
      Tab(2).Control(17)=   "lblFA(96)"
      Tab(2).Control(18)=   "Label23"
      Tab(2).Control(19)=   "Label34(0)"
      Tab(2).Control(20)=   "Label35(0)"
      Tab(2).Control(21)=   "Label34(3)"
      Tab(2).Control(22)=   "Label34(2)"
      Tab(2).Control(23)=   "Label68"
      Tab(2).Control(24)=   "Label67(0)"
      Tab(2).Control(25)=   "Label67(1)"
      Tab(2).Control(26)=   "Label44"
      Tab(2).Control(27)=   "Label49"
      Tab(2).Control(28)=   "Label6"
      Tab(2).Control(29)=   "Label55"
      Tab(2).Control(30)=   "Label72"
      Tab(2).Control(31)=   "textFA45"
      Tab(2).Control(32)=   "textFA30_2"
      Tab(2).Control(33)=   "textFA61_2"
      Tab(2).Control(34)=   "textFA62_2"
      Tab(2).Control(35)=   "textFA72_2"
      Tab(2).Control(36)=   "textFA71_2"
      Tab(2).Control(37)=   "Label74"
      Tab(2).Control(38)=   "Label76"
      Tab(2).Control(39)=   "textFA39"
      Tab(2).Control(40)=   "textFA28"
      Tab(2).Control(41)=   "textFA27"
      Tab(2).Control(42)=   "textFA26"
      Tab(2).Control(43)=   "textFA25"
      Tab(2).Control(44)=   "textFA44"
      Tab(2).Control(45)=   "textFA42"
      Tab(2).Control(46)=   "textFA41"
      Tab(2).Control(47)=   "textFA40"
      Tab(2).Control(48)=   "textFA85"
      Tab(2).Control(49)=   "txtFA(96)"
      Tab(2).Control(50)=   "textFA30"
      Tab(2).Control(51)=   "textFA61"
      Tab(2).Control(52)=   "textFA62"
      Tab(2).Control(53)=   "textFA71"
      Tab(2).Control(54)=   "textFA72"
      Tab(2).Control(55)=   "txtFA(86)"
      Tab(2).Control(56)=   "textFA89"
      Tab(2).Control(57)=   "textFA87"
      Tab(2).Control(58)=   "txtFA(98)"
      Tab(2).Control(59)=   "Combo2(0)"
      Tab(2).Control(60)=   "Combo3(0)"
      Tab(2).Control(61)=   "txtFA(104)"
      Tab(2).Control(62)=   "txtFA(124)"
      Tab(2).Control(63)=   "txtFA(95)"
      Tab(2).Control(64)=   "txtFA(128)"
      Tab(2).Control(65)=   "txtFA(136)"
      Tab(2).ControlCount=   66
      TabCaption(3)   =   "商標"
      TabPicture(3)   =   "frm050705.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "textFA138"
      Tab(3).Control(1)=   "textFA137"
      Tab(3).Control(2)=   "textFA139"
      Tab(3).Control(3)=   "TextFA93"
      Tab(3).Control(4)=   "txtFA(129)"
      Tab(3).Control(5)=   "Combo5"
      Tab(3).Control(6)=   "textFA120"
      Tab(3).Control(7)=   "Combo4"
      Tab(3).Control(8)=   "textFA117"
      Tab(3).Control(9)=   "Combo3(1)"
      Tab(3).Control(10)=   "Combo2(1)"
      Tab(3).Control(11)=   "textFA107"
      Tab(3).Control(12)=   "textFA112"
      Tab(3).Control(13)=   "textFA111"
      Tab(3).Control(14)=   "textFA106"
      Tab(3).Control(15)=   "textFA109"
      Tab(3).Control(16)=   "textFA73"
      Tab(3).Control(17)=   "textFA74"
      Tab(3).Control(18)=   "textFA75"
      Tab(3).Control(19)=   "txtFA(99)"
      Tab(3).Control(20)=   "txtFA(91)"
      Tab(3).Control(21)=   "textFA88"
      Tab(3).Control(22)=   "textFA90"
      Tab(3).Control(23)=   "textFA66"
      Tab(3).Control(24)=   "textFA67"
      Tab(3).Control(25)=   "textFA68"
      Tab(3).Control(26)=   "Label79"
      Tab(3).Control(27)=   "Label78"
      Tab(3).Control(28)=   "Label77"
      Tab(3).Control(29)=   "Label54"
      Tab(3).Control(30)=   "Label57"
      Tab(3).Control(31)=   "Label42"
      Tab(3).Control(32)=   "Label69"
      Tab(3).Control(33)=   "Label8(2)"
      Tab(3).Control(34)=   "Label35(1)"
      Tab(3).Control(35)=   "Label34(1)"
      Tab(3).Control(36)=   "Label75"
      Tab(3).Control(37)=   "textFA107_2"
      Tab(3).Control(38)=   "textFA111_2"
      Tab(3).Control(39)=   "textFA112_2"
      Tab(3).Control(40)=   "textFA67_2"
      Tab(3).Control(41)=   "textFA66_2"
      Tab(3).Control(42)=   "textFA110"
      Tab(3).Control(43)=   "Label73"
      Tab(3).Control(44)=   "LblFA120"
      Tab(3).Control(45)=   "Label53"
      Tab(3).Control(46)=   "Label51"
      Tab(3).Control(47)=   "Label43"
      Tab(3).Control(48)=   "Label1(28)"
      Tab(3).Control(49)=   "Label50"
      Tab(3).Control(50)=   "Label48"
      Tab(3).Control(51)=   "Label34(5)"
      Tab(3).Control(52)=   "Label34(4)"
      Tab(3).Control(53)=   "Label47"
      Tab(3).Control(54)=   "Label46"
      Tab(3).Control(55)=   "Label45"
      Tab(3).Control(56)=   "Label11(1)"
      Tab(3).Control(57)=   "Label56"
      Tab(3).Control(58)=   "Label67(3)"
      Tab(3).Control(59)=   "Label67(2)"
      Tab(3).Control(60)=   "Label8(1)"
      Tab(3).ControlCount=   61
      TabCaption(4)   =   "其他"
      TabPicture(4)   =   "frm050705.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label1(3)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label67(6)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label70"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label1(23)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label17"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label84"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label28"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label26"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label32"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Label33"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Label52"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Label67(4)"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Label67(5)"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Label67(7)"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "Label67(8)"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "Label1(1)"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "lstDeveloper"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "textFA92"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "textFA37"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "textFA38_2"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "textFA60"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "textFA59_2"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "txtFA(102)"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "textFA97"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "textFA24"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "textFA31"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "textFA38"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "textFA59"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "cboStatus"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "txtFA(83)"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "txtFA(101)"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).Control(31)=   "txtFA(121)"
      Tab(4).Control(31).Enabled=   0   'False
      Tab(4).Control(32)=   "txtFA(122)"
      Tab(4).Control(32).Enabled=   0   'False
      Tab(4).Control(33)=   "txtFA(123)"
      Tab(4).Control(33).Enabled=   0   'False
      Tab(4).Control(34)=   "Frame1K"
      Tab(4).Control(34).Enabled=   0   'False
      Tab(4).ControlCount=   35
      TabCaption(5)   =   "參考備註"
      TabPicture(5)   =   "frm050705.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "textFA29"
      Tab(5).Control(1)=   "cmdIns"
      Tab(5).ControlCount=   2
      Begin VB.TextBox textFA138 
         Height          =   270
         Left            =   -67320
         MaxLength       =   2
         TabIndex        =   89
         Top             =   1230
         Width           =   375
      End
      Begin VB.TextBox textFA137 
         Height          =   270
         Left            =   -69030
         MaxLength       =   2
         TabIndex        =   88
         Top             =   1230
         Width           =   375
      End
      Begin VB.TextBox textFA139 
         Height          =   270
         Left            =   -70020
         MaxLength       =   7
         TabIndex        =   91
         Top             =   1530
         Width           =   1095
      End
      Begin VB.TextBox TextFA93 
         Height          =   270
         Left            =   -73140
         MaxLength       =   1
         TabIndex        =   83
         Top             =   930
         Width           =   372
      End
      Begin VB.TextBox txtFA 
         Height          =   270
         Index           =   136
         Left            =   -67272
         MaxLength       =   1
         TabIndex        =   70
         Top             =   2550
         Width           =   345
      End
      Begin VB.Frame Frame1K 
         Height          =   280
         Left            =   3720
         TabIndex        =   278
         Top             =   2040
         Width           =   4930
         Begin VB.CheckBox Chk1K 
            Caption         =   "帳單另寄"
            Height          =   180
            Index           =   0
            Left            =   1740
            TabIndex        =   116
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "上傳平台"
            Height          =   180
            Index           =   1
            Left            =   2790
            TabIndex        =   117
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "月帳單"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   118
            Top             =   60
            Width           =   1030
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "請款單寄送類型："
            Height          =   180
            Index           =   26
            Left            =   150
            TabIndex        =   279
            Top             =   60
            Width           =   1440
         End
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   129
         Left            =   -72930
         MaxLength       =   1
         TabIndex        =   106
         Top             =   4770
         Width           =   330
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   128
         Left            =   -68820
         MaxLength       =   1
         TabIndex        =   62
         Top             =   1740
         Width           =   330
      End
      Begin VB.CommandButton cmdIntroduce 
         Caption         =   "被介紹者"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71220
         Style           =   1  '圖片外觀
         TabIndex        =   275
         Top             =   4410
         Width           =   1000
      End
      Begin VB.TextBox txtXYS02 
         Height          =   315
         Left            =   -73860
         MaxLength       =   8
         TabIndex        =   25
         Top             =   4740
         Width           =   1000
      End
      Begin VB.ComboBox cboSource 
         Height          =   260
         ItemData        =   "frm050705.frx":00A8
         Left            =   -74010
         List            =   "frm050705.frx":00AA
         Style           =   2  '單純下拉式
         TabIndex        =   24
         Top             =   4430
         Width           =   2750
      End
      Begin VB.ComboBox Combo5 
         Height          =   260
         ItemData        =   "frm050705.frx":00AC
         Left            =   -69630
         List            =   "frm050705.frx":00BC
         Style           =   2  '單純下拉式
         TabIndex        =   95
         Top             =   2130
         Width           =   2580
      End
      Begin VB.TextBox txtFA 
         Height          =   270
         Index           =   95
         Left            =   -73290
         MaxLength       =   2
         TabIndex        =   65
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtFA 
         Height          =   270
         Index           =   124
         Left            =   -73290
         MaxLength       =   1
         TabIndex        =   63
         Top             =   2040
         Width           =   372
      End
      Begin VB.TextBox txtFA 
         Height          =   315
         Index           =   105
         Left            =   -69555
         TabIndex        =   44
         Top             =   3150
         Width           =   3180
      End
      Begin VB.TextBox txtFA 
         Height          =   270
         Index           =   123
         Left            =   3435
         MaxLength       =   1
         TabIndex        =   125
         Top             =   4040
         Width           =   372
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   122
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   124
         Top             =   3312
         Width           =   330
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   121
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   123
         Top             =   3000
         Width           =   330
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   101
         Left            =   2196
         MaxLength       =   1
         TabIndex        =   122
         Top             =   3636
         Width           =   330
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   83
         Left            =   2196
         MaxLength       =   1
         TabIndex        =   121
         Top             =   3312
         Width           =   330
      End
      Begin VB.TextBox txtFA 
         Height          =   270
         Index           =   104
         Left            =   -67035
         MaxLength       =   1
         TabIndex        =   60
         Top             =   1500
         Width           =   345
      End
      Begin VB.TextBox textFA120 
         Height          =   270
         Left            =   -69120
         MaxLength       =   10
         TabIndex        =   93
         Top             =   1830
         Width           =   1000
      End
      Begin VB.ComboBox Combo4 
         Height          =   260
         ItemData        =   "frm050705.frx":00F8
         Left            =   -73650
         List            =   "frm050705.frx":0105
         TabIndex        =   105
         Text            =   "Combo4"
         Top             =   4470
         Width           =   7470
      End
      Begin VB.CommandButton cmdIns 
         Caption         =   "各項指示"
         Height          =   300
         Left            =   -74880
         TabIndex        =   132
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdTW 
         Caption         =   "臺灣地址格式"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.5
            Charset         =   136
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -71640
         TabIndex        =   258
         Top             =   300
         Width           =   1160
      End
      Begin VB.ComboBox cboStatus 
         Height          =   260
         ItemData        =   "frm050705.frx":013C
         Left            =   1590
         List            =   "frm050705.frx":0146
         TabIndex        =   115
         Text            =   "cboStatus"
         Top             =   2050
         Width           =   2055
      End
      Begin VB.TextBox textFA117 
         Height          =   270
         Left            =   -68760
         MaxLength       =   1
         TabIndex        =   104
         Top             =   4230
         Width           =   330
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   1
         ItemData        =   "frm050705.frx":0166
         Left            =   -70470
         List            =   "frm050705.frx":0179
         Style           =   2  '單純下拉式
         TabIndex        =   97
         Top             =   2430
         Width           =   1470
      End
      Begin VB.ComboBox Combo2 
         Height          =   260
         Index           =   1
         ItemData        =   "frm050705.frx":01AD
         Left            =   -73650
         List            =   "frm050705.frx":01AF
         Style           =   2  '單純下拉式
         TabIndex        =   96
         Top             =   2430
         Width           =   990
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   0
         ItemData        =   "frm050705.frx":01B1
         Left            =   -70470
         List            =   "frm050705.frx":01C4
         Style           =   2  '單純下拉式
         TabIndex        =   53
         Top             =   270
         Width           =   1470
      End
      Begin VB.ComboBox Combo2 
         Height          =   260
         Index           =   0
         ItemData        =   "frm050705.frx":01F8
         Left            =   -73680
         List            =   "frm050705.frx":01FA
         Style           =   2  '單純下拉式
         TabIndex        =   52
         Top             =   270
         Width           =   990
      End
      Begin VB.ComboBox Combo1 
         Height          =   260
         ItemData        =   "frm050705.frx":01FC
         Left            =   -69330
         List            =   "frm050705.frx":01FE
         Style           =   2  '單純下拉式
         TabIndex        =   23
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox textFA100 
         Height          =   315
         Left            =   -73155
         MaxLength       =   1
         TabIndex        =   22
         Top             =   4080
         Width           =   330
      End
      Begin VB.TextBox textFA107 
         Height          =   270
         Left            =   -73260
         MaxLength       =   8
         TabIndex        =   100
         Top             =   3360
         Width           =   972
      End
      Begin VB.TextBox textFA112 
         Height          =   270
         Left            =   -73260
         MaxLength       =   8
         TabIndex        =   102
         Top             =   3930
         Width           =   972
      End
      Begin VB.TextBox textFA111 
         Height          =   270
         Left            =   -72930
         MaxLength       =   8
         TabIndex        =   101
         Top             =   3630
         Width           =   972
      End
      Begin VB.TextBox textFA106 
         Height          =   270
         Left            =   -73080
         MaxLength       =   30
         TabIndex        =   103
         Top             =   4200
         Width           =   2532
      End
      Begin VB.TextBox textFA109 
         Height          =   270
         Left            =   -67050
         MaxLength       =   1
         TabIndex        =   98
         Top             =   2430
         Width           =   330
      End
      Begin VB.TextBox textFA59 
         Height          =   270
         Left            =   1590
         MaxLength       =   8
         TabIndex        =   112
         Top             =   1470
         Width           =   972
      End
      Begin VB.TextBox textFA73 
         Height          =   270
         Left            =   -73650
         MaxLength       =   2
         TabIndex        =   86
         Top             =   1230
         Width           =   375
      End
      Begin VB.TextBox textFA74 
         Height          =   270
         Left            =   -71100
         MaxLength       =   2
         TabIndex        =   87
         Top             =   1230
         Width           =   375
      End
      Begin VB.TextBox textFA75 
         Height          =   270
         Left            =   -73080
         MaxLength       =   7
         TabIndex        =   90
         Top             =   1530
         Width           =   1095
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   99
         Left            =   -72900
         MaxLength       =   1
         TabIndex        =   94
         Top             =   2130
         Width           =   330
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   91
         Left            =   -72900
         MaxLength       =   1
         TabIndex        =   92
         Top             =   1830
         Width           =   330
      End
      Begin VB.TextBox textFA88 
         Height          =   285
         Left            =   -70080
         MaxLength       =   1
         TabIndex        =   84
         Top             =   930
         Width           =   330
      End
      Begin VB.TextBox textFA90 
         Height          =   285
         Left            =   -68010
         MaxLength       =   1
         TabIndex        =   85
         Top             =   930
         Width           =   330
      End
      Begin VB.TextBox textFA38 
         Height          =   270
         Left            =   1590
         MaxLength       =   8
         TabIndex        =   110
         Top             =   870
         Width           =   975
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   98
         Left            =   -69864
         MaxLength       =   1
         TabIndex        =   79
         Top             =   4455
         Width           =   330
      End
      Begin VB.TextBox textFA87 
         Height          =   285
         Left            =   -73290
         MaxLength       =   1
         TabIndex        =   76
         Top             =   4185
         Width           =   330
      End
      Begin VB.TextBox textFA89 
         Height          =   270
         Left            =   -73290
         MaxLength       =   1
         TabIndex        =   78
         Top             =   4455
         Width           =   330
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   86
         Left            =   -69864
         MaxLength       =   1
         TabIndex        =   77
         Top             =   4185
         Width           =   330
      End
      Begin VB.TextBox textFA72 
         Height          =   270
         Left            =   -73290
         MaxLength       =   8
         TabIndex        =   75
         Top             =   3900
         Width           =   972
      End
      Begin VB.TextBox textFA71 
         Height          =   270
         Left            =   -72960
         MaxLength       =   8
         TabIndex        =   74
         Top             =   3630
         Width           =   972
      End
      Begin VB.TextBox textFA62 
         Height          =   270
         Left            =   -73290
         MaxLength       =   8
         TabIndex        =   72
         Top             =   3090
         Width           =   972
      End
      Begin VB.TextBox textFA61 
         Height          =   270
         Left            =   -73290
         MaxLength       =   8
         TabIndex        =   73
         Top             =   3360
         Width           =   972
      End
      Begin VB.TextBox textFA30 
         Height          =   270
         Left            =   -73290
         MaxLength       =   8
         TabIndex        =   71
         Top             =   2820
         Width           =   972
      End
      Begin VB.TextBox txtFA 
         Height          =   270
         Index           =   96
         Left            =   -70632
         MaxLength       =   2
         TabIndex        =   66
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox textFA85 
         Height          =   270
         Left            =   -67035
         MaxLength       =   1
         TabIndex        =   58
         Top             =   1200
         Width           =   345
      End
      Begin VB.TextBox textFA40 
         Height          =   270
         Left            =   -70560
         MaxLength       =   1
         TabIndex        =   57
         Top             =   1200
         Width           =   372
      End
      Begin VB.TextBox textFA41 
         Height          =   270
         Left            =   -73290
         MaxLength       =   1
         TabIndex        =   59
         Top             =   1500
         Width           =   372
      End
      Begin VB.TextBox textFA42 
         Height          =   270
         Left            =   -73290
         MaxLength       =   1
         TabIndex        =   61
         Top             =   1770
         Width           =   372
      End
      Begin VB.TextBox textFA44 
         Height          =   270
         Left            =   -67050
         MaxLength       =   1
         TabIndex        =   54
         Top             =   306
         Width           =   330
      End
      Begin VB.TextBox textFA25 
         Height          =   270
         Left            =   -73290
         MaxLength       =   2
         TabIndex        =   68
         Top             =   2550
         Width           =   375
      End
      Begin VB.TextBox textFA26 
         Height          =   270
         Left            =   -67080
         MaxLength       =   2
         TabIndex        =   67
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox textFA27 
         Height          =   270
         Left            =   -70128
         MaxLength       =   7
         TabIndex        =   69
         Top             =   2550
         Width           =   1095
      End
      Begin VB.TextBox textFA28 
         Height          =   270
         Left            =   -68820
         MaxLength       =   30
         TabIndex        =   64
         Top             =   2010
         Width           =   2532
      End
      Begin VB.TextBox textFA39 
         Height          =   270
         Left            =   -73800
         MaxLength       =   1
         TabIndex        =   56
         Top             =   1200
         Width           =   372
      End
      Begin VB.TextBox textFA31 
         Height          =   270
         Left            =   1590
         MaxLength       =   1
         TabIndex        =   109
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox textFA24 
         Height          =   270
         Left            =   1590
         MaxLength       =   1
         TabIndex        =   107
         Top             =   330
         Width           =   372
      End
      Begin VB.TextBox textFA97 
         Height          =   270
         Left            =   5820
         MaxLength       =   1
         TabIndex        =   108
         Top             =   300
         Width           =   330
      End
      Begin VB.TextBox txtFA 
         Height          =   285
         Index           =   102
         Left            =   2196
         MaxLength       =   1
         TabIndex        =   120
         Top             =   3000
         Width           =   330
      End
      Begin VB.TextBox textFA14 
         Height          =   315
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   29
         Top             =   660
         Width           =   3492
      End
      Begin VB.TextBox textFA82 
         Height          =   315
         Left            =   -73620
         MaxLength       =   50
         TabIndex        =   43
         Top             =   3150
         Width           =   3180
      End
      Begin VB.TextBox textFA81 
         Height          =   315
         Left            =   -69555
         MaxLength       =   50
         TabIndex        =   42
         Top             =   2850
         Width           =   3180
      End
      Begin VB.TextBox textFA80 
         Height          =   315
         Left            =   -73620
         MaxLength       =   50
         TabIndex        =   41
         Top             =   2850
         Width           =   3180
      End
      Begin VB.TextBox textFA79 
         Height          =   315
         Left            =   -69555
         TabIndex        =   40
         Top             =   2550
         Width           =   3180
      End
      Begin VB.TextBox textFA77 
         Height          =   315
         Left            =   -67380
         MaxLength       =   1
         TabIndex        =   19
         Top             =   3435
         Width           =   375
      End
      Begin VB.TextBox TextFA76 
         Height          =   315
         Left            =   -69765
         MaxLength       =   1
         TabIndex        =   21
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox textFA12 
         Height          =   315
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   27
         Top             =   330
         Width           =   3492
      End
      Begin VB.TextBox textFA15 
         Height          =   315
         Left            =   -69840
         MaxLength       =   20
         TabIndex        =   30
         Top             =   660
         Width           =   3492
      End
      Begin VB.TextBox textFA13 
         Height          =   315
         Left            =   -69840
         MaxLength       =   20
         TabIndex        =   28
         Top             =   330
         Width           =   3492
      End
      Begin VB.TextBox textFA16 
         Height          =   315
         Left            =   -73620
         MaxLength       =   50
         TabIndex        =   39
         Top             =   2550
         Width           =   3180
      End
      Begin VB.TextBox textFA66 
         Height          =   270
         Left            =   -73650
         MaxLength       =   8
         TabIndex        =   81
         Top             =   600
         Width           =   972
      End
      Begin VB.TextBox textFA67 
         Height          =   270
         Left            =   -73650
         MaxLength       =   8
         TabIndex        =   80
         Top             =   300
         Width           =   972
      End
      Begin VB.TextBox textFA68 
         Height          =   270
         Left            =   -73500
         MaxLength       =   1
         TabIndex        =   82
         Top             =   780
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textFA55_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   -72810
         Locked          =   -1  'True
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   3780
         Width           =   2325
      End
      Begin VB.TextBox textFA10_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   225
         Left            =   -72810
         Locked          =   -1  'True
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2325
      End
      Begin VB.TextBox textFA55 
         Height          =   315
         Left            =   -73440
         MaxLength       =   3
         TabIndex        =   20
         Top             =   3750
         Width           =   612
      End
      Begin VB.TextBox textFA03 
         Height          =   315
         Left            =   -73440
         MaxLength       =   8
         TabIndex        =   2
         Top             =   330
         Width           =   1092
      End
      Begin VB.TextBox textFA10 
         Height          =   315
         Left            =   -73440
         MaxLength       =   4
         TabIndex        =   17
         Top             =   3450
         Width           =   612
      End
      Begin VB.TextBox textFA11 
         Height          =   315
         Left            =   -69360
         MaxLength       =   7
         TabIndex        =   18
         Top             =   3450
         Width           =   975
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "延展折扣 :          %"
         Height          =   180
         Left            =   -68220
         TabIndex        =   283
         Top             =   1280
         Width           =   1460
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "繳註冊費折扣 :          %"
         Height          =   180
         Left            =   -70290
         TabIndex        =   282
         Top             =   1280
         Width           =   1820
      End
      Begin VB.Label Label77 
         Caption         =   "商標全部折扣終止日："
         Height          =   260
         Left            =   -71820
         TabIndex        =   281
         Top             =   1580
         Width           =   1970
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "商標全部折扣 :          %"
         Height          =   180
         Left            =   -74880
         TabIndex        =   236
         Top             =   1280
         Width           =   1820
      End
      Begin VB.Label Label57 
         Caption         =   "商標全部折扣起始日："
         Height          =   260
         Left            =   -74880
         TabIndex        =   234
         Top             =   1580
         Width           =   1970
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "商標 Email 同時寄紙本：       （Y：是）"
         Height          =   195
         Left            =   -74880
         TabIndex        =   233
         Top             =   2175
         Width           =   3120
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "商標以 EMail 通知：               （Y：是   D：僅D/N）"
         Height          =   180
         Left            =   -74880
         TabIndex        =   232
         Top             =   1875
         Width           =   4005
      End
      Begin VB.Label Label8 
         Caption         =   "FCT註冊費自動代繳：        (Y:自動代繳)"
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   205
         Top             =   980
         Width           =   3170
      End
      Begin VB.Label Label35 
         Caption         =   "延展請款對象："
         Height          =   260
         Index           =   1
         Left            =   -74880
         TabIndex        =   177
         Top             =   300
         Width           =   1340
      End
      Begin VB.Label Label34 
         Caption         =   "延展代理人："
         Height          =   260
         Index           =   1
         Left            =   -74880
         TabIndex        =   176
         Top             =   600
         Width           =   1340
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "專利不得請雜費 :            (Y:是)"
         Height          =   180
         Left            =   -68784
         TabIndex        =   280
         Top             =   2556
         Width           =   2400
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "台灣案商標註冊證形式 :          (1:電子 2:紙本)"
         Height          =   180
         Left            =   -74880
         TabIndex        =   277
         Top             =   4800
         Width           =   3495
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "台灣案專利證書形式 :          (1:電子 2:紙本)"
         Height          =   180
         Left            =   -70590
         TabIndex        =   276
         Top             =   1770
         Width           =   3315
      End
      Begin MSForms.Label LblSourceN 
         Height          =   285
         Left            =   -72690
         TabIndex        =   274
         Top             =   4740
         Width           =   2700
         VariousPropertyBits=   27
         Caption         =   "LblSourceN"
         Size            =   "4762;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtXYS03 
         Height          =   615
         Left            =   -69330
         TabIndex        =   26
         Top             =   4440
         Width           =   3105
         VariousPropertyBits=   -1466941413
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "5468;1085"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "其他     說明："
         Height          =   495
         Index           =   5
         Left            =   -69840
         TabIndex        =   273
         Top             =   4425
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹者編號："
         Height          =   180
         Index           =   4
         Left            =   -74910
         TabIndex        =   272
         Top             =   4740
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "來所原因："
         Height          =   180
         Index           =   2
         Left            =   -74910
         TabIndex        =   271
         Top             =   4425
         Width           =   900
      End
      Begin MSForms.TextBox textFA36 
         Height          =   315
         Left            =   -73275
         TabIndex        =   35
         Top             =   1590
         Width           =   3330
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5874;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA34 
         Height          =   315
         Left            =   -73275
         TabIndex        =   33
         Top             =   1275
         Width           =   3330
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5874;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA32 
         Height          =   315
         Left            =   -73275
         TabIndex        =   31
         Top             =   975
         Width           =   3330
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5874;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA35 
         Height          =   315
         Left            =   -69675
         TabIndex        =   34
         Top             =   1275
         Width           =   3330
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5874;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA33 
         Height          =   315
         Left            =   -69675
         TabIndex        =   32
         Top             =   975
         Width           =   3330
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5874;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA57 
         Height          =   315
         Left            =   -69555
         TabIndex        =   37
         Top             =   1905
         Width           =   3180
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "5609;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA53 
         Height          =   315
         Left            =   -70320
         TabIndex        =   49
         Top             =   4095
         Width           =   3975
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "7011;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA08 
         Height          =   315
         Left            =   -70320
         TabIndex        =   46
         Top             =   3465
         Width           =   3975
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "7011;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA70 
         Height          =   315
         Left            =   -69705
         TabIndex        =   15
         Top             =   2820
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA65 
         Height          =   315
         Left            =   -69705
         TabIndex        =   7
         Top             =   1230
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA64 
         Height          =   315
         Left            =   -73305
         TabIndex        =   6
         Top             =   1230
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA63 
         Height          =   315
         Left            =   -69705
         TabIndex        =   5
         Top             =   930
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA05 
         Height          =   315
         Left            =   -73305
         TabIndex        =   4
         Top             =   930
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA18 
         Height          =   315
         Left            =   -73305
         TabIndex        =   10
         Top             =   2160
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA22 
         Height          =   315
         Left            =   -73305
         TabIndex        =   14
         Top             =   2820
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA21 
         Height          =   315
         Left            =   -69705
         TabIndex        =   13
         Top             =   2490
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA20 
         Height          =   315
         Left            =   -73305
         TabIndex        =   12
         Top             =   2490
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA19 
         Height          =   315
         Left            =   -69705
         TabIndex        =   11
         Top             =   2160
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "5927;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA107_2 
         Height          =   285
         Left            =   -72270
         TabIndex        =   250
         TabStop         =   0   'False
         Top             =   3360
         Width           =   6045
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "10663;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA111_2 
         Height          =   285
         Left            =   -71940
         TabIndex        =   248
         TabStop         =   0   'False
         Top             =   3630
         Width           =   5715
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "10081;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA112_2 
         Height          =   285
         Left            =   -72270
         TabIndex        =   247
         TabStop         =   0   'False
         Top             =   3930
         Width           =   6045
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "10663;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA71_2 
         Height          =   285
         Left            =   -71940
         TabIndex        =   226
         TabStop         =   0   'False
         Top             =   3630
         Width           =   5685
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "8555;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA72_2 
         Height          =   285
         Left            =   -72270
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   3900
         Width           =   6012
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "8555;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA62_2 
         Height          =   285
         Left            =   -72270
         TabIndex        =   224
         TabStop         =   0   'False
         Top             =   3090
         Width           =   6012
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "8555;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA61_2 
         Height          =   285
         Left            =   -72270
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   3360
         Width           =   6012
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "8555;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA30_2 
         Height          =   285
         Left            =   -72270
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   2820
         Width           =   6012
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "8555;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA67_2 
         Height          =   290
         Left            =   -72660
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   300
         Width           =   6260
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "11033;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA66_2 
         Height          =   290
         Left            =   -72660
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   600
         Width           =   6260
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "11033;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA58 
         Height          =   315
         Left            =   -73110
         TabIndex        =   38
         Top             =   2220
         Width           =   6735
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "11880;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA56 
         Height          =   315
         Left            =   -73110
         TabIndex        =   36
         Top             =   1905
         Width           =   1740
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "3069;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA78 
         Height          =   315
         Left            =   -73440
         TabIndex        =   51
         Top             =   4710
         Width           =   7095
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12515;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA54 
         Height          =   315
         Left            =   -73440
         TabIndex        =   50
         Top             =   4395
         Width           =   7095
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12515;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA52 
         Height          =   315
         Left            =   -73440
         TabIndex        =   48
         Top             =   4095
         Width           =   1245
         VariousPropertyBits=   671105051
         Size            =   "2196;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA09 
         Height          =   315
         Left            =   -73440
         TabIndex        =   47
         Top             =   3780
         Width           =   7095
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12515;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA07 
         Height          =   315
         Left            =   -73440
         TabIndex        =   45
         Top             =   3465
         Width           =   1245
         VariousPropertyBits=   671105051
         Size            =   "2196;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA17 
         Height          =   315
         Left            =   -73440
         TabIndex        =   9
         Top             =   1860
         Width           =   7095
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "12515;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA04 
         Height          =   315
         Left            =   -73440
         TabIndex        =   3
         Top             =   630
         Width           =   7095
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12515;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA06 
         Height          =   315
         Left            =   -73440
         TabIndex        =   8
         Top             =   1560
         Width           =   7095
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "12515;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA23 
         Height          =   315
         Left            =   -73440
         TabIndex        =   16
         Top             =   3150
         Width           =   7095
         VariousPropertyBits=   671105051
         MaxLength       =   70
         Size            =   "12515;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA29 
         Height          =   4005
         Left            =   -74910
         TabIndex        =   134
         Top             =   720
         Width           =   8700
         VariousPropertyBits=   -1466941413
         MaxLength       =   4000
         ScrollBars      =   2
         Size            =   "15346;7064"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA110 
         Height          =   645
         Left            =   -73650
         TabIndex        =   99
         Top             =   2700
         Width           =   7455
         VariousPropertyBits=   -1466941413
         MaxLength       =   180
         ScrollBars      =   2
         Size            =   "13150;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA59_2 
         Height          =   285
         Left            =   2580
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   1470
         Width           =   6012
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "8555;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA60 
         Height          =   285
         Left            =   1590
         TabIndex        =   113
         Top             =   1770
         Width           =   2805
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "4948;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA38_2 
         Height          =   285
         Left            =   2580
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   870
         Width           =   6012
         VariousPropertyBits=   679493659
         BackColor       =   16777215
         Size            =   "8555;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA37 
         Height          =   285
         Left            =   1590
         TabIndex        =   111
         Top             =   1170
         Width           =   2805
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "4948;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA45 
         Height          =   615
         Left            =   -73680
         TabIndex        =   55
         Top             =   570
         Width           =   7455
         VariousPropertyBits=   -1466941413
         MaxLength       =   180
         ScrollBars      =   2
         Size            =   "13150;1085"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textFA92 
         Height          =   615
         Left            =   1005
         TabIndex        =   119
         Top             =   2340
         Width           =   7725
         VariousPropertyBits=   -1466941413
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "13626;1085"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstDeveloper 
         Height          =   315
         Left            =   5820
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1215
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2302;556"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單之本所帳戶："
         Height          =   180
         Left            =   -71625
         TabIndex        =   270
         Top             =   2175
         Width           =   1980
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "FCP實審自動代繳 :               (Y:自動代繳)"
         Height          =   180
         Left            =   -74910
         TabIndex        =   269
         Top             =   2040
         Width           =   3150
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "(財務CF)："
         Height          =   180
         Left            =   -70440
         TabIndex        =   268
         Top             =   3150
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否同意歐盟通用資料保護規範(GDPR)：        （W:待回覆 Y:同意 N:不同意）"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   267
         Top             =   4080
         Width           =   6060
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發促銷信：　　（Y：一定要寄，N：一定不要寄）"
         Height          =   180
         Index           =   8
         Left            =   3960
         TabIndex        =   266
         Top             =   3354
         Width           =   4560
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發竹曆：　　    （Y：一定要寄，N：一定不要寄）"
         Height          =   180
         Index           =   7
         Left            =   3960
         TabIndex        =   265
         Top             =   3045
         Width           =   4560
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "財務處是否寄發催款單：          (1： 每月寄對帳單　2. 客戶要求不寄對帳單　3. 其他)"
         Height          =   180
         Index           =   5
         Left            =   150
         TabIndex        =   264
         Top             =   3690
         Width           =   6645
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "財務處是否寄發FC收據：         (N：不寄)"
         Height          =   180
         Index           =   4
         Left            =   156
         TabIndex        =   263
         Top             =   3354
         Width           =   3216
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否電子送件 :                 (Y:是)"
         Height          =   180
         Left            =   -68910
         TabIndex        =   262
         Top             =   1530
         Width           =   2700
      End
      Begin VB.Label LblFA120 
         Caption         =   "LblFA120"
         Height          =   255
         Left            =   -68040
         TabIndex        =   261
         Top             =   1875
         Width           =   1305
      End
      Begin VB.Label Label53 
         Caption         =   "管控智權人員："
         Height          =   255
         Left            =   -70440
         TabIndex        =   260
         Top             =   1875
         Width           =   1300
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "陸代定稿加註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   259
         Top             =   4530
         Width           =   1260
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "(僅提醒)"
         Height          =   180
         Left            =   -74490
         TabIndex        =   257
         Top             =   2950
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "(僅提醒)"
         Height          =   180
         Left            =   -74500
         TabIndex        =   256
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "不催延展：        (Y:不催)"
         Height          =   180
         Index           =   28
         Left            =   -69630
         TabIndex        =   255
         Top             =   4260
         Width           =   1905
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單列印幣別格式："
         Height          =   180
         Left            =   -72630
         TabIndex        =   254
         Top             =   2460
         Width           =   2160
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "專利請款單列印幣別格式："
         Height          =   180
         Left            =   -72630
         TabIndex        =   253
         Top             =   345
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "帳單幣別："
         Height          =   240
         Index           =   53
         Left            =   -70200
         TabIndex        =   252
         Top             =   4140
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發專利雙週報：       (N:不寄)"
         Height          =   180
         Index           =   24
         Left            =   -74910
         TabIndex        =   251
         Top             =   4140
         Width           =   2760
      End
      Begin VB.Label Label48 
         Caption         =   "商標固定請款對象："
         Height          =   255
         Left            =   -74880
         TabIndex        =   249
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label34 
         Caption         =   "商標D/N固定列印對象："
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   246
         Top             =   3630
         Width           =   1905
      End
      Begin VB.Label Label34 
         Caption         =   "延展D/N列印對象："
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   245
         Top             =   3930
         Width           =   1575
      End
      Begin VB.Label Label47 
         Caption         =   "代理人商標財務編號："
         Height          =   255
         Left            =   -74880
         TabIndex        =   244
         Top             =   4260
         Width           =   1815
      End
      Begin VB.Label Label46 
         Caption         =   "商標D/N備註："
         Height          =   255
         Left            =   -74880
         TabIndex        =   243
         Top             =   2730
         Width           =   1215
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N是否印申請人：        (Y:印)"
         Height          =   180
         Left            =   -69000
         TabIndex        =   242
         Top             =   2460
         Width           =   2780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "商標請款幣別"
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   241
         Top             =   2460
         Width           =   1080
      End
      Begin VB.Label Label52 
         Caption         =   "代理人狀態:"
         Height          =   200
         Left            =   150
         TabIndex        =   240
         Top             =   2070
         Width           =   1220
      End
      Begin VB.Label Label33 
         Caption         =   "實體副本聯絡人:"
         Height          =   255
         Left            =   150
         TabIndex        =   239
         Top             =   1770
         Width           =   1695
      End
      Begin VB.Label Label32 
         Caption         =   "實體副本收受人:"
         Height          =   255
         Left            =   150
         TabIndex        =   238
         Top             =   1470
         Width           =   1695
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "商標申請/翻譯折扣 :          %"
         Height          =   180
         Left            =   -72790
         TabIndex        =   235
         Top             =   1280
         Width           =   2230
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單份數："
         Height          =   200
         Index           =   3
         Left            =   -69510
         TabIndex        =   231
         Top             =   980
         Width           =   1460
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "商標定稿份數："
         Height          =   200
         Index           =   2
         Left            =   -71400
         TabIndex        =   230
         Top             =   980
         Width           =   1280
      End
      Begin VB.Label Label26 
         Caption         =   "副本收受人:"
         Height          =   255
         Left            =   150
         TabIndex        =   229
         Top             =   870
         Width           =   1455
      End
      Begin VB.Label Label28 
         Caption         =   "副本聯絡人:"
         Height          =   255
         Left            =   150
         TabIndex        =   228
         Top             =   1170
         Width           =   1455
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "專利 Email 同時寄紙本：       （Y：是）"
         Height          =   192
         Left            =   -71820
         TabIndex        =   221
         Top             =   4452
         Width           =   3120
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "專利請款單份數："
         Height          =   195
         Index           =   1
         Left            =   -74910
         TabIndex        =   220
         Top             =   4455
         Width           =   1455
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "專利定稿份數："
         Height          =   195
         Index           =   0
         Left            =   -74910
         TabIndex        =   219
         Top             =   4185
         Width           =   1275
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "專利以 EMail 通知：               （Y：是   D：僅D/N）"
         Height          =   180
         Left            =   -71820
         TabIndex        =   218
         Top             =   4188
         Width           =   4008
      End
      Begin VB.Label Label34 
         Caption         =   "專利D/N固定列印對象："
         Height          =   255
         Index           =   2
         Left            =   -74910
         TabIndex        =   217
         Top             =   3630
         Width           =   1905
      End
      Begin VB.Label Label34 
         Caption         =   "年費D/N列印對象："
         Height          =   255
         Index           =   3
         Left            =   -74910
         TabIndex        =   216
         Top             =   3900
         Width           =   1935
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文 :                           (1:中 2:英 3:日)"
         Height          =   180
         Left            =   150
         TabIndex        =   215
         Top             =   600
         Width           =   3180
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "是否寄台一雜誌 :               (N : 不寄)"
         Height          =   180
         Left            =   150
         TabIndex        =   214
         Top             =   330
         Width           =   2760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報:              （N:不寄）"
         Height          =   180
         Index           =   23
         Left            =   4500
         TabIndex        =   213
         Top             =   330
         Width           =   2640
      End
      Begin VB.Label Label70 
         Caption         =   "帳單備註:"
         Height          =   255
         Left            =   150
         TabIndex        =   204
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "年費請款對象："
         Height          =   255
         Index           =   0
         Left            =   -74910
         TabIndex        =   212
         Top             =   3090
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label34 
         Caption         =   "年費代理人："
         Height          =   255
         Index           =   0
         Left            =   -74910
         TabIndex        =   211
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "專利固定請款對象："
         Height          =   255
         Left            =   -74910
         TabIndex        =   210
         Top             =   2820
         Width           =   1695
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "是否用LEDES電子帳單：          (Y：是)"
         Height          =   180
         Index           =   6
         Left            =   150
         TabIndex        =   209
         Top             =   3045
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發人員:"
         Height          =   180
         Index           =   3
         Left            =   4860
         TabIndex        =   208
         Top             =   1770
         Width           =   765
      End
      Begin VB.Label lblFA 
         Caption         =   "專利年費折扣 :          %"
         Height          =   180
         Index           =   96
         Left            =   -71832
         TabIndex        =   207
         Top             =   2328
         Width           =   2292
      End
      Begin VB.Label lblFA 
         Caption         =   "專利領證折扣 :                    %"
         Height          =   180
         Index           =   95
         Left            =   -74904
         TabIndex        =   206
         Top             =   2280
         Width           =   2508
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人日文名稱:"
         Height          =   180
         Index           =   2
         Left            =   -74865
         TabIndex        =   157
         Top             =   2220
         Width           =   1665
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人英文名稱:"
         Height          =   180
         Index           =   1
         Left            =   -71280
         TabIndex        =   156
         Top             =   1935
         Width           =   1665
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人中文名稱 :"
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   155
         Top             =   1935
         Width           =   1710
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "(其他3)："
         Height          =   180
         Left            =   -74370
         TabIndex        =   203
         Top             =   3150
         Width           =   750
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "(其他2)："
         Height          =   180
         Left            =   -70320
         TabIndex        =   202
         Top             =   2850
         Width           =   750
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "(其他1)："
         Height          =   180
         Left            =   -74370
         TabIndex        =   201
         Top             =   2850
         Width           =   750
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(財務)："
         Height          =   180
         Left            =   -70230
         TabIndex        =   200
         Top             =   2550
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否核對已准專利 :         (N:否)"
         Height          =   180
         Index           =   156
         Left            =   -68910
         TabIndex        =   199
         Top             =   1230
         Width           =   2700
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   15
         Left            =   -69825
         TabIndex        =   198
         Top             =   1260
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   14
         Left            =   -73425
         TabIndex        =   197
         Top             =   1260
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   13
         Left            =   -69825
         TabIndex        =   196
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   10
         Left            =   -73425
         TabIndex        =   195
         Top             =   990
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   12
         Left            =   -69825
         TabIndex        =   194
         Top             =   1005
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   11
         Left            =   -69825
         TabIndex        =   193
         Top             =   1275
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   9
         Left            =   -73425
         TabIndex        =   192
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   8
         Left            =   -73425
         TabIndex        =   191
         Top             =   1275
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   7
         Left            =   -73425
         TabIndex        =   190
         Top             =   1005
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   6
         Left            =   -69825
         TabIndex        =   189
         Top             =   2850
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   5
         Left            =   -69825
         TabIndex        =   188
         Top             =   2520
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   4
         Left            =   -69825
         TabIndex        =   187
         Top             =   2190
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   3
         Left            =   -73425
         TabIndex        =   186
         Top             =   2850
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   2
         Left            =   -73425
         TabIndex        =   185
         Top             =   2520
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   32
         Left            =   -73425
         TabIndex        =   183
         Top             =   2190
         Width           =   90
      End
      Begin VB.Label Label62 
         Caption         =   "聯絡人部門(日)："
         Height          =   255
         Left            =   -74880
         TabIndex        =   182
         Top             =   4710
         Width           =   1455
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Left            =   -66885
         TabIndex        =   181
         Top             =   3510
         Width           =   465
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "呆帳記錄："
         Height          =   180
         Left            =   -68310
         TabIndex        =   180
         Top             =   3495
         Width           =   900
      End
      Begin VB.Label Label59 
         Caption         =   "(A:律師事務所 B:公司直接委辦 C:其他)"
         Height          =   225
         Left            =   -69345
         TabIndex        =   179
         Top             =   3780
         Width           =   3090
      End
      Begin VB.Label Label58 
         Caption         =   "性質："
         Height          =   225
         Left            =   -70320
         TabIndex        =   178
         Top             =   3780
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "延展單筆不跑：          (Y:單筆不跑)"
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   175
         Top             =   810
         Visible         =   0   'False
         Width           =   2770
      End
      Begin VB.Label Labeld1 
         Height          =   255
         Index           =   10
         Left            =   -72000
         TabIndex        =   169
         Top             =   3780
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費通知函單筆不跑 :          (Y:單筆不跑)"
         Height          =   180
         Index           =   0
         Left            =   -72630
         TabIndex        =   168
         Top             =   1230
         Width           =   3465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費自動代繳 :               (Y:自動代繳 / N:寄證書後年費不續辦)"
         Height          =   180
         Left            =   -74910
         TabIndex        =   167
         Top             =   1500
         Width           =   5070
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "FCP領證自動代繳 :               (Y:自動代繳)"
         Height          =   180
         Left            =   -74910
         TabIndex        =   166
         Top             =   1770
         Width           =   3150
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "專利請款幣別"
         Height          =   180
         Index           =   0
         Left            =   -74910
         TabIndex        =   165
         Top             =   345
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N是否印申請人：          (Y:印)"
         Height          =   180
         Left            =   -68970
         TabIndex        =   164
         Top             =   345
         Width           =   2820
      End
      Begin VB.Label Label15 
         Caption         =   "專利D/N備註："
         Height          =   255
         Left            =   -74910
         TabIndex        =   163
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "專利全部折扣 :                    %"
         Height          =   180
         Left            =   -74904
         TabIndex        =   162
         Top             =   2556
         Width           =   2280
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "專利申請/翻譯折扣 :             %"
         Height          =   180
         Left            =   -68790
         TabIndex        =   161
         Top             =   2325
         Width           =   2295
      End
      Begin VB.Label Label21 
         Caption         =   "專利全部折扣起始日:"
         Height          =   252
         Left            =   -71832
         TabIndex        =   160
         Top             =   2556
         Width           =   1668
      End
      Begin VB.Label Label22 
         Caption         =   "代理人專利財務編號 :"
         Height          =   255
         Left            =   -70590
         TabIndex        =   159
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "收款後辦案 :           (Y:先收)"
         Height          =   180
         Left            =   -74910
         TabIndex        =   158
         Top             =   1230
         Width           =   2130
      End
      Begin VB.Label Label2 
         Caption         =   "地址國籍 :"
         Height          =   225
         Index           =   1
         Left            =   -74910
         TabIndex        =   153
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label Labeld1 
         Height          =   255
         Index           =   5
         Left            =   -71640
         TabIndex        =   152
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Labeld1 
         Height          =   255
         Index           =   4
         Left            =   -72840
         TabIndex        =   151
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Labeld1 
         Height          =   255
         Index           =   3
         Left            =   -73920
         TabIndex        =   150
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "代理人名稱(日)："
         Height          =   255
         Left            =   -74910
         TabIndex        =   148
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label29 
         Caption         =   "代理人名稱(英)："
         Height          =   255
         Left            =   -74910
         TabIndex        =   147
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "代理人名稱(中)："
         Height          =   255
         Left            =   -74910
         TabIndex        =   146
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "客戶編號："
         Height          =   255
         Left            =   -74910
         TabIndex        =   145
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "代理人國籍："
         Height          =   225
         Index           =   0
         Left            =   -74910
         TabIndex        =   144
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "代理人地址(日)："
         Height          =   225
         Left            =   -74910
         TabIndex        =   143
         Top             =   3150
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "代理人地址(英)："
         Height          =   255
         Left            =   -74910
         TabIndex        =   142
         Top             =   2190
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "代理人地址(中)："
         Height          =   255
         Left            =   -74910
         TabIndex        =   141
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "開發日期："
         Height          =   255
         Left            =   -70320
         TabIndex        =   140
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "POB："
         Height          =   255
         Left            =   -74880
         TabIndex        =   139
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail(代表)："
         Height          =   180
         Left            =   -74865
         TabIndex        =   138
         Top             =   2550
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "TEL："
         Height          =   255
         Left            =   -74880
         TabIndex        =   137
         Top             =   330
         Width           =   615
      End
      Begin VB.Label Label36 
         Caption         =   "聯絡人１(中)："
         Height          =   255
         Left            =   -74880
         TabIndex        =   136
         Top             =   3465
         Width           =   1455
      End
      Begin VB.Label Label37 
         Caption         =   "聯絡人１(英)："
         Height          =   255
         Left            =   -71625
         TabIndex        =   135
         Top             =   3510
         Width           =   1275
      End
      Begin VB.Label Label38 
         Caption         =   "聯絡人１(日)："
         Height          =   255
         Left            =   -74880
         TabIndex        =   133
         Top             =   3780
         Width           =   1455
      End
      Begin VB.Label Label39 
         Caption         =   "聯絡人２(中)："
         Height          =   255
         Left            =   -74880
         TabIndex        =   131
         Top             =   4095
         Width           =   1455
      End
      Begin VB.Label Label40 
         Caption         =   "聯絡人２(英)："
         Height          =   255
         Left            =   -71625
         TabIndex        =   130
         Top             =   4155
         Width           =   1455
      End
      Begin VB.Label Label41 
         Caption         =   "聯絡人２(日)："
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   129
         Top             =   4395
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "FAX："
         Height          =   255
         Left            =   -74880
         TabIndex        =   128
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Labeld1 
         Height          =   255
         Index           =   0
         Left            =   -72840
         TabIndex        =   127
         Top             =   3600
         Width           =   1335
      End
   End
   Begin VB.TextBox textFA01 
      Height          =   264
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Top             =   675
      Width           =   1092
   End
   Begin VB.TextBox textFA02 
      Height          =   264
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   1
      Top             =   675
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8505
      Top             =   90
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":0200
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":051C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":0838
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":0A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":0D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":104C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":1368
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":1684
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":19A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":1CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050705.frx":1FD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   149
      Top             =   0
      Width           =   9140
      _ExtentX        =   16122
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
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
      Left            =   3000
      TabIndex        =   170
      TabStop         =   0   'False
      Top             =   660
      Width           =   6075
      VariousPropertyBits=   16415
      BackColor       =   16777215
      Size            =   "5741;503"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   180
      Index           =   1
      Left            =   0
      TabIndex        =   184
      Top             =   0
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "代理人編號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   154
      Top             =   675
      Width           =   1455
   End
End
Attribute VB_Name = "frm050705"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Memo By Sindy 2021/12/07 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'Memo by Lydia 2020/11/16 原本取得字串長度的模組StrLength改成GetTextLength
Option Explicit

'Modify By Cheng 2003/09/23
'Const MAX_FIELD = 70
'Const MAX_FIELD = 72
'Const MAX_FIELD = 74
'edit by nickc 2005/12/02
'Const MAX_FIELD = 75
'Modify by Morgan 2006/10/18
'Const MAX_FIELD = 77
'edit by nickc 2006/10/24
'Const MAX_FIELD = 78

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
'edit by nickc 2006/10/24
'Dim m_FieldList(MAX_FIELD) As FIELDITEM
Dim m_FieldList() As FIELDITEM

' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer

' 辦識其為外商還是內商的程式
' 0 表內商
' 1 表外商
Dim m_SysKind As Integer

' 第一筆資料的本所案號
Dim m_FirstKEY(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_Txt As Object 'Add by Morgan 2008/11/13
Dim i As Integer 'Add By Sindy 2012/6/5
Dim bolShow100135 As Boolean 'Add by Amy 2016/06/30 避免地址格式有誤按確定鈕出現錯誤
'Added by Lydia 2018/10/24
Dim m_PrevForm As Form '前一畫面
Dim m_PrevNo As String '傳入代理人編號
Dim bCancel As Boolean 'Add by Amy 2022/11/25


'Added by Lydia 2018/10/24 傳入前一畫面
Public Sub SetParent(ByVal pFM As Form, ByVal pNo As String)
    Set m_PrevForm = pFM
    m_PrevNo = ChangeCustomerL(pNo)
End Sub
'end 2018/10/24

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT FA01,FA02 FROM FAGENT " & _
            "WHERE FA01 = (SELECT MIN(FA01) FROM FAGENT) AND " & _
                  "FA02 = (SELECT MIN(FA02) FROM FAGENT " & _
                           "WHERE FA01 = (SELECT MIN(FA01) FROM FAGENT)) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("FA01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("FA01")
      If IsNull(rsTmp.Fields("FA02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("FA02")
   End If
   rsTmp.Close

   strSql = "SELECT FA01,FA02 FROM FAGENT " & _
            "WHERE FA01 = (SELECT MAX(FA01) FROM FAGENT) AND " & _
                  "FA02 = (SELECT MAX(FA02) FROM FAGENT " & _
                           "WHERE FA01 = (SELECT MAX(FA01) FROM FAGENT)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("FA01")) = False Then: m_LastKEY(0) = rsTmp.Fields("FA01")
      If IsNull(rsTmp.Fields("FA02")) = False Then: m_LastKEY(1) = rsTmp.Fields("FA02")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

'Add by Amy 2022/11/25 來所原因(原名稱:代理人來源)
Private Sub cboSource_Click()
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    If cboSource = MsgText(601) Then Exit Sub
   
    'Modify by Amy 2022/12/28 bug修改時為11.其他,改成其餘選項再改回11.其他欄位會被鎖住
    'If m_FieldList(126).fiOldData <> Left(cboSource, 2) Then
        'Modify by 2024/11/29 欄位是否鎖住改抓共用函數,避免有未改的
        'txtXYS02.Locked = True: txtXYS03.Locked = True
        txtXYS02.Text = "": LblSourceN.Caption = ""
        'txtXYS03.Text = ""'Mark by Amy 2023/07/20 原有值不清空-秀玲
         Call Pub_SetCboComeSource(9, Me.Name, cboSource, , txtXYS02, txtXYS03)
    'End If
End Sub

Private Sub cboStatus_GotFocus()
   If Pub_StrUserSt03 = "M51" Then OpenIme
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
   'Add by Amy 2015/08/24 只限M51可以自行輸入,其他人只能下拉
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If Pub_StrUserSt03 <> "M51" Then KeyAscii = 0
End Sub

Private Sub cboStatus_LostFocus()
   CloseIme
End Sub

Private Sub cboStatus_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(cboStatus) = False Then
      If GetTextLength(cboStatus) > 12 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人狀態內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         cboStatus_GotFocus
      End If
   End If
End Sub

'Added by Lydia 2016/11/24 各項指示
Private Sub cmdIns_Click()
   If Me.textFA01.Text = "" Then
      MsgBox "請輸入代理人編號", vbInformation
      Exit Sub
   End If
   
   'Added by Lydia 2017/08/03
   If m_EditMode <> 0 And m_EditMode <> 4 Then
      MsgBox IIf(m_EditMode = 1, "新增中", "修改中") & "不可執行！", vbInformation
      Exit Sub
   End If
   'end 2017/08/03
   
   'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
   If PUB_CheckFormExist("frm12040159") Then
       MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
       Exit Sub
   End If
   'end 2020/05/05
   
   frm12040159.SetParent "E", Trim(Me.textFA01.Text & Me.textFA02.Text), Me
   frm12040159.Show
End Sub

'被介紹者鈕
Private Sub cmdIntroduce_Click()
    Dim stName As String
    
    If cmdIntroduce.BackColor = &HFFFF80 Then
        '英->中->日
        If textFA05 = MsgText(601) Then
            If textFA04 = MsgText(601) Then
                stName = stName & " " & textFA06
            Else
                stName = stName & " " & textFA04
            End If
        Else
            stName = textFA05
            If textFA63 <> MsgText(601) Then stName = stName & " " & textFA63
            If textFA64 <> MsgText(601) Then stName = stName & " " & textFA64
            If textFA65 <> MsgText(601) Then stName = stName & " " & textFA65
        End If
        frm050705_1.txtNo = textFA01
        frm050705_1.lbl1(0) = textFA10
        frm050705_1.lbl1(1) = textFA10_2
        frm050705_1.lbl1(3) = stName

        frm050705_1.SetParent Me
        frm050705_1.QueryData
        frm050705_1.Show
        Me.Hide
    End If
End Sub

'Add by Amy 2016/06/30
Private Sub cmdTW_Click(Index As Integer)
   frm100135.Show vbModal
End Sub

'Add By Sindy 2013/1/17
Private Sub Combo2_Click(Index As Integer)
   Call GetCurrType(Index)
End Sub

'Add By Sindy 2013/1/17
Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2013/1/17
Private Sub Combo2_Validate(Index As Integer, Cancel As Boolean)
   If Combo2(Index) = MsgText(601) Then
      Combo2(Index).Tag = Combo2(Index).Text
      Combo3(Index).ListIndex = 0
      Combo3(Index).Enabled = False
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo2(Index), Label11(Index)) = False Then
      Cancel = True
      Combo2(Index).SetFocus
   End If
   If Combo2(Index) <> "USD" Then
      If ExistCheck("DebitNoteRate", "DNR01", Combo2(Index), Label11(Index) & "匯率") = False Then
         Cancel = True
         Combo2(Index).SetFocus
         Exit Sub
      End If
   End If
   Call GetCurrType(Index)
End Sub

'Add By Sindy 2013/1/17
Private Sub GetCurrType(Index As Integer)
Dim intType As Integer
   
   If Combo2(Index) = MsgText(601) Then
      Combo2(Index).Tag = Combo2(Index).Text
      Combo3(Index).ListIndex = 0
      Combo3(Index).Enabled = False
      Exit Sub
   End If
   '若更改請款幣別
   If Me.Combo2(Index).Text <> Me.Combo2(Index).Tag Then
      Me.Combo2(Index).Tag = Me.Combo2(Index).Text
      '請款幣別變更要重新預設列印幣別
      '台幣
      If Me.Combo2(Index).Text = "NTD" Then
         intType = 1 '純台幣
      '人民幣
      ElseIf Me.Combo2(Index).Text = "RMB" Then
         intType = 4 '外幣+美金合計
      '其他幣別
      Else
         intType = 2 '台幣+外幣合計
      End If
      Combo3(Index).ListIndex = intType
      '若為台幣時則格式欄位鎖住不可修改
      If Me.Combo2(Index).Text = "NTD" Then
         Combo3(Index).Enabled = False
      Else
         Combo3(Index).Enabled = True
      End If
   End If
End Sub

Private Sub Form_Initialize()
   'add by nickc 2006/10/24
   ReDim m_FieldList(TF_FA) As FIELDITEM
End Sub

' Load Form
Private Sub Form_Load()
   SSTab1.Tab = 0
   
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm050705", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm050705", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm050705", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm050705", strFind, False)
   
   lstDeveloper.Height = 600
   lstDeveloper.Width = 1300
   
   textFA10_2.BackColor = &H8000000F
   textFA30_2.BackColor = &H8000000F
   textFA38_2.BackColor = &H8000000F
   textFA55_2.BackColor = &H8000000F
   textFA59_2.BackColor = &H8000000F
   textFA61_2.BackColor = &H8000000F
   textFA62_2.BackColor = &H8000000F
   textFA66_2.BackColor = &H8000000F
   textFA67_2.BackColor = &H8000000F
   textFA71_2.BackColor = &H8000000F
   textFA72_2.BackColor = &H8000000F
   'Add By Sindy 2011/3/4
   textFA107_2.BackColor = &H8000000F
   textFA111_2.BackColor = &H8000000F
   textFA112_2.BackColor = &H8000000F
   '2011/3/4 End
   textCUID.BackColor = &H8000000F
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   InitialField
   
   'Add By Sindy 2012/6/5
   Combo1.Clear
   Combo1.AddItem ""
   strExc(0) = "SELECT A1Y01||'-'||A1Y02 FROM ACC1Y0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Do While Not RsTemp.EOF
      Combo1.AddItem RsTemp.Fields(0)
      RsTemp.MoveNext
   Loop
   '2012/6/5 End
   'Add By Sindy 2013/1/17
   '抓有輸入過匯率的請款幣別
   For i = 0 To 1
      Combo2(i).Clear
      Combo2(i).AddItem ""
      Combo2(i).AddItem "USD"
      If RsTemp.State <> adStateClosed Then RsTemp.Close
      RsTemp.CursorLocation = adUseClient
      RsTemp.Open "select distinct DNR01 from DebitNoteRate order by DNR01 asc", adoTaie, adOpenStatic, adLockReadOnly
      Do While RsTemp.EOF = False
         Combo2(i).AddItem RsTemp.Fields("DNR01").Value
         RsTemp.MoveNext
      Loop
      RsTemp.Close
   Next i
   '2013/1/17 End
   
   'Add by Amy 2022/11/25 來所原因(原名稱:代理人來源)
   'Modify by 2024/11/29 改抓共用函數,避免有未改到
   cboSource.ListIndex = -1
   LblSourceN = ""
   Call Pub_SetCboComeSource(0, Me.Name, cboSource)
   'end 2024/11/29
   
   RefreshRange
   If m_PrevNo = "" Then 'Added by Lydia 2018/10/24 判斷是否有傳入代理人編號
        ShowFirstRecord
   'Added by Lydia 2018/10/24 有傳入代理人編號
   Else
        ShowCurrRecord Mid(m_PrevNo, 1, 8), Mid(m_PrevNo, 9, 1)
   End If
   'end 2018/10/24
   
   UpdateToolbarState
   SetCtrlReadOnly True
   
   'Add by Morgan 2006/5/29
   '考慮共榮(X22558000)的案件需要於客戶檔設定年費代理人，為避免邏輯過於複雜故取消代理人檔的年費代理人(FA61)&年費請款對象(FA62)
   Label35(0).Visible = False: textFA62.Visible = False: textFA62_2.Visible = False
   Label34(0).Visible = False: textFA61.Visible = False: textFA61_2.Visible = False
   'end 2006/5/29
   
   'Add by Amy 2015/08/24 +代理人狀態下拉選單
   cboStatus.Clear
   cboStatus.AddItem ""
   cboStatus.AddItem "刪址"
   cboStatus.AddItem "倒閉"
   cboStatus.AddItem "遷移不明"
   cboStatus.AddItem "解散"
   cboStatus.AddItem "廢止"
   cboStatus.AddItem "撤銷"
   cboStatus.AddItem "停業"
   cboStatus.AddItem "往生"
   cboStatus.AddItem "其他"
   cboStatus.AddItem "業務自行處理"
   'end 2015/07/24
   cboStatus.AddItem "國內同業" 'Add by Amy 2021/11/26
   'add by sonia 2025/4/24
   If Pub_StrUserSt03 = "M51" Then
      cboStatus.AddItem "不再使用"
      cboStatus.AddItem "不得代理"
      cboStatus.AddItem "不得代理專利"
      cboStatus.AddItem "不得代理商標"
      cboStatus.AddItem "宣告破產"
   End If
   'end 2025/4/24
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
      textFA29.Top = 360
      textFA29.Height = 4335
   End If
   
   'Add by Amy 2024/01/22 國外潛在客戶維護轉號存檔切至此畫面-陳金蓮
   If m_PrevForm Is Nothing = False Then
      If UCase(m_PrevForm.Name) = "FRM140402" Then
         OnAction vbKeyF3
      End If
   End If
   
   Frame1K.BorderStyle = 0 'Add By Sindy 2025/1/7
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To TF_FA   'edit by nickc 2006/10/24  MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "FA" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
        'Modify By Cheng 2003/11/17
'         Case 11, 25, 26, 27, 47, 48, 50, 51:
         Case 11, 25, 26, 27, 47, 48, 50, 51, 73, 74, 75:
        'End
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To TF_FA - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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

'2011/10/11 cancel by sonia
'' 取得新的代理人編號
'Private Function GetNewAgentNo() As String
'   Dim strTemp As String
'   Dim strSql As String
'   Dim rsTmp As New ADODB.Recordset
'   Dim bExist As Boolean
'
'   strTemp = "0"
'   bExist = False
'   strSql = "SELECT * FROM AutoNumber " & _
'            "WHERE AU01 = '" & "Y" & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      bExist = True
'      If IsNull(rsTmp.Fields("AU03")) = False Then
'         strTemp = rsTmp.Fields("AU03")
'      End If
'   End If
'   rsTmp.Close
'
'   strTemp = CStr(Val(strTemp))
'
'   GetNewAgentNo = "Y" & String(5 - Len(strTemp), "0") & strTemp & "00"
'
'   strTemp = CStr(Val(strTemp) + 1)
'
'   If bExist = True Then
'      strSql = "UPDATE AUTONUMBER SET AU03 = " & strTemp & " " & _
'               "WHERE AU01 = '" & "Y" & "' "
'   Else
'      strSql = "INSERT INTO AUTONUMBER (AU01,AU02,AU03) VALUES ('Y'," & DBDATE(SystemDate()) & "," & strTemp & ") "
'   End If
'   cnnConnection.Execute strSql
'
'   Set rsTmp = Nothing
'End Function
'2011/10/11 end

' 更新欄位的內容
Private Sub UpdateFieldNewData()
Dim strTmp  As String
   
   '若新增資料
   If m_EditMode = 1 Then
        '若未輸入代理人編號
      If IsEmptyText(textFA01) = True Then
         '2011/10/11 modify by sonia 因GetNewAgentNo是抓出au03用再+1更新回去,與國外潛在客戶轉入新代理人的編號(先+1更新回去同時用新值)重覆
         'textFA01 = GetNewAgentNo
         If ClsPDGetAutoNumber("Y", strTmp, True, False) Then
            strTmp = "Y" + Right(strTmp, 5)
            textFA01 = strTmp
         Else
            ShowMsg "讀取自動編號檔錯誤，請洽系統管理者 !"
            Exit Sub
         End If
         '2011/10/11 end
      End If
      If IsEmptyText(textFA02) = True Then
         textFA02 = "0"
      End If
   End If
   
   If IsEmptyText(textFA01) = False Then
      SetFieldNewData "FA01", textFA01 & String(8 - Len(textFA01), "0")
   Else
      SetFieldNewData "FA01", textFA01
   End If
   SetFieldNewData "FA02", textFA02
   ' 客戶編號
   If IsEmptyText(textFA03) = False Then
      SetFieldNewData "FA03", textFA03 & String(8 - Len(textFA03), "0")
   Else
      SetFieldNewData "FA03", textFA03
   End If
   SetFieldNewData "FA04", textFA04
   SetFieldNewData "FA05", textFA05
   SetFieldNewData "FA06", textFA06
   SetFieldNewData "FA07", textFA07
   SetFieldNewData "FA08", textFA08
   SetFieldNewData "FA09", textFA09
   SetFieldNewData "FA10", textFA10
   ' 開發日期
   If IsEmptyText(textFA11) = False Then
      SetFieldNewData "FA11", DBDATE(textFA11)
   Else
      SetFieldNewData "FA11", textFA11
   End If
   SetFieldNewData "FA12", textFA12
   SetFieldNewData "FA13", textFA13
   SetFieldNewData "FA14", textFA14
   SetFieldNewData "FA15", textFA15
   SetFieldNewData "FA16", textFA16
   SetFieldNewData "FA17", textFA17
   SetFieldNewData "FA18", textFA18
   SetFieldNewData "FA19", textFA19
   SetFieldNewData "FA20", textFA20
   SetFieldNewData "FA21", textFA21
   SetFieldNewData "FA22", textFA22
   SetFieldNewData "FA70", textFA70
   SetFieldNewData "FA23", textFA23
   SetFieldNewData "FA24", textFA24
   SetFieldNewData "FA25", textFA25
   SetFieldNewData "FA26", textFA26
   ' 全部折扣起始日
   If IsEmptyText(textFA27) = False Then
      SetFieldNewData "FA27", DBDATE(textFA27)
   Else
      SetFieldNewData "FA27", textFA27
   End If
   SetFieldNewData "FA28", textFA28
   SetFieldNewData "FA29", textFA29
   If IsEmptyText(textFA30) = False Then
      SetFieldNewData "FA30", textFA30 & String(9 - Len(textFA30), "0")
   Else
      SetFieldNewData "FA30", textFA30
   End If
   SetFieldNewData "FA31", textFA31
   SetFieldNewData "FA32", textFA32
   SetFieldNewData "FA33", textFA33
   SetFieldNewData "FA34", textFA34
   SetFieldNewData "FA35", textFA35
   SetFieldNewData "FA36", textFA36
   SetFieldNewData "FA37", textFA37
   If IsEmptyText(textFA38) = False Then
      SetFieldNewData "FA38", textFA38 & String(9 - Len(textFA38), "0")
   Else
      SetFieldNewData "FA38", textFA38
   End If
   SetFieldNewData "FA39", textFA39
   SetFieldNewData "FA40", textFA40
   SetFieldNewData "FA41", textFA41
   SetFieldNewData "FA42", textFA42
   'Modify By Sindy 2013/1/17
'   SetFieldNewData "FA43", textFA43
   SetFieldNewData "FA43", Combo2(0).Text
   '2013/1/17 End
   SetFieldNewData "FA44", textFA44
   SetFieldNewData "FA45", textFA45
   SetFieldNewData "FA52", textFA52
   SetFieldNewData "FA53", textFA53
   SetFieldNewData "FA54", textFA54
   SetFieldNewData "FA55", textFA55
   SetFieldNewData "FA56", textFA56
   SetFieldNewData "FA57", textFA57
   SetFieldNewData "FA58", textFA58
   If IsEmptyText(textFA59) = False Then
      SetFieldNewData "FA59", textFA59 & String(9 - Len(textFA59), "0")
   Else
      SetFieldNewData "FA59", textFA59
   End If
   SetFieldNewData "FA60", textFA60
   If IsEmptyText(textFA61) = False Then
      SetFieldNewData "FA61", textFA61 & String(9 - Len(textFA61), "0")
   Else
      SetFieldNewData "FA61", textFA61
   End If
   If IsEmptyText(textFA62) = False Then
      SetFieldNewData "FA62", textFA62 & String(9 - Len(textFA62), "0")
   Else
      SetFieldNewData "FA62", textFA62
   End If
   SetFieldNewData "FA63", textFA63
   SetFieldNewData "FA64", textFA64
   SetFieldNewData "FA65", textFA65
   If IsEmptyText(textFA66) = False Then
      SetFieldNewData "FA66", textFA66 & String(9 - Len(textFA66), "0")
   Else
      SetFieldNewData "FA66", textFA66
   End If
   If IsEmptyText(textFA67) = False Then
      SetFieldNewData "FA67", textFA67 & String(9 - Len(textFA67), "0")
   Else
      SetFieldNewData "FA67", textFA67
   End If
   SetFieldNewData "FA68", textFA68
   'Modify by Amy 2015/08/24 改成下拉選單
   'SetFieldNewData "FA69", textFA69
   SetFieldNewData "FA69", LTrim(cboStatus.Text)
   
    'Add By Cheng 2003/09/23
    'Begin
   SetFieldNewData "FA71", textFA71
   If IsEmptyText(textFA71) = False Then
      SetFieldNewData "FA71", textFA71 & String(9 - Len(textFA71), "0")
   Else
      SetFieldNewData "FA71", textFA71
   End If
   SetFieldNewData "FA72", textFA72
   If IsEmptyText(textFA72) = False Then
      SetFieldNewData "FA72", textFA72 & String(9 - Len(textFA72), "0")
   Else
      SetFieldNewData "FA72", textFA72
   End If
    'End
    'Add By Cheng 2003/11/17
   SetFieldNewData "FA73", textFA73
   SetFieldNewData "FA74", textFA74
   ' 全部折扣起始日
   If IsEmptyText(textFA75) = False Then
      SetFieldNewData "FA75", DBDATE(textFA75)
   Else
      SetFieldNewData "FA75", textFA75
   End If
    'End
   
   'Add By Sindy 2025/3/10
   SetFieldNewData "FA137", textFA137
   SetFieldNewData "FA138", textFA138
   ' 全部折扣終止日
   If IsEmptyText(textFA139) = False Then
      SetFieldNewData "FA139", DBDATE(textFA139)
   Else
      SetFieldNewData "FA139", textFA139
   End If
   '2025/3/10 END
   
   'add by nickc 2005/12/02
   SetFieldNewData "FA76", TextFA76
   SetFieldNewData "FA77", textFA77
   SetFieldNewData "FA78", textFA78 'Add by Morgan 2006/10/18
   SetFieldNewData "FA85", textFA85 'Add by Morgan 2007/10/26
'   SetFieldNewData "FA79", textFA79 'Add by Morgan 2008/1/16 'Modify By Sindy 2018/3/16 Mark
   SetFieldNewData "FA80", textFA80 'Add by Morgan 2008/1/16
   SetFieldNewData "FA81", textFA81 'Add by Morgan 2008/1/16
   SetFieldNewData "FA82", textFA82 'Add by Morgan 2008/1/16
   SetFieldNewData "FA87", textFA87 'Add by Morgan 2008/3/13
   SetFieldNewData "FA88", textFA88 'Add by Morgan 2008/3/13
   SetFieldNewData "FA89", textFA89 'Add by Morgan 2008/3/13
   SetFieldNewData "FA90", textFA90 'Add by Morgan 2008/3/13
   SetFieldNewData "FA92", textFA92 'Add by Morgan 2008/6/3
   
   SetFieldNewData "FA93", TextFA93 'add by Toni 2008/10/21
   SetFieldNewData "FA97", textFA97 '2008/12/9 add by sonia
   
   SetFieldNewData "FA100", textFA100 'Add By Sindy 2011/3/10
   
   'Add By Sindy 2011/3/4
   SetFieldNewData "FA106", textFA106
   SetFieldNewData "FA107", textFA107
   If IsEmptyText(textFA107) = False Then
      SetFieldNewData "FA107", textFA107 & String(9 - Len(textFA107), "0")
   Else
      SetFieldNewData "FA107", textFA107
   End If
   'Modify By Sindy 2013/1/17
'   SetFieldNewData "FA108", textFA108
   SetFieldNewData "FA108", Combo2(1).Text
   '2013/1/17 End
   SetFieldNewData "FA109", textFA109
   SetFieldNewData "FA110", textFA110
   SetFieldNewData "FA111", textFA111
   If IsEmptyText(textFA111) = False Then
      SetFieldNewData "FA111", textFA111 & String(9 - Len(textFA111), "0")
   Else
      SetFieldNewData "FA111", textFA111
   End If
   SetFieldNewData "FA112", textFA112
   If IsEmptyText(textFA112) = False Then
      SetFieldNewData "FA112", textFA112 & String(9 - Len(textFA112), "0")
   Else
      SetFieldNewData "FA112", textFA112
   End If
   '2011/3/4 End
   
   SetFieldNewData "FA113", Left(Trim(Combo1.Text), 3) 'Add By Sindy 2012/6/5
   SetFieldNewData "FA115", IIf(Combo3(0).Text <> "", Combo3(0).ListIndex, "") 'Add By Sindy 2013/1/17 專利
   SetFieldNewData "FA116", IIf(Combo3(1).Text <> "", Combo3(1).ListIndex, "") 'Add By Sindy 2013/1/17 商標
   SetFieldNewData "FA117", textFA117 'Add By Sindy 2013/8/15
   
   'Add by Morgan 2008/11/13 改用陣列以免控制項超過且以後再新增欄位也不必改
   For Each m_Txt In txtFA
      SetFieldNewData "FA" & m_Txt.Index, m_Txt
   Next
   'end 2008/11/13
   
   SetFieldNewData "FA119", Combo4.Text 'Add By Sindy 2016/12/5
   SetFieldNewData "FA120", textFA120 'Add by Amy 2017/01/05
   SetFieldNewData "FA126", Left(Combo5.Text, 1) 'Add By Sindy 2021/3/3
   SetFieldNewData "FA127", IIf(cboSource.Text <> "", Left(cboSource.Text, 2), "") 'Add by Amy 2022/11/25
   
   'Add By Sindy 2025/1/7
   strExc(10) = ""
   For Each m_Txt In Chk1K
      If m_Txt.Value = 1 Then
         strExc(10) = strExc(10) & "," & m_Txt.Index + 1
      End If
   Next
   If strExc(10) <> "" Then strExc(10) = Mid(strExc(10), 2)
   SetFieldNewData "FA135", strExc(10)
   '2025/1/7 END
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To TF_FA - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Private Sub QueryDB()
   RefreshRange
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   
   textFA01 = Empty
   textFA02 = Empty
   textFA03 = Empty
   textFA04 = Empty
   textFA05 = Empty
   textFA06 = Empty
   textFA07 = Empty
   textFA08 = Empty
   textFA09 = Empty
   textFA10 = Empty
   textFA10_2 = Empty
   textFA11 = Empty
   If m_EditMode = 1 Then textFA11 = strSrvDate(2)
   textFA12 = Empty
   textFA13 = Empty
   textFA14 = Empty
   textFA15 = Empty
   textFA16 = Empty
   textFA17 = Empty
   textFA18 = Empty
   textFA19 = Empty
   textFA20 = Empty
   textFA21 = Empty
   textFA22 = Empty
   textFA70 = Empty
   textFA23 = Empty
   textFA24 = Empty
   textFA25 = Empty
   textFA26 = Empty
   textFA27 = Empty
   textFA28 = Empty
   textFA29 = Empty
   textFA30 = Empty
   textFA30_2 = Empty
   textFA31 = Empty
   textFA32 = Empty
   textFA33 = Empty
   textFA34 = Empty
   textFA35 = Empty
   textFA36 = Empty
   textFA37 = Empty
   textFA38 = Empty
   textFA38_2 = Empty
   textFA39 = Empty
   textFA40 = Empty
   textFA41 = Empty
   textFA42 = Empty
'   textFA43 = Empty
   textFA44 = Empty
   textFA45 = Empty
   textFA52 = Empty
   textFA53 = Empty
   textFA54 = Empty
   textFA55 = Empty
   textFA55_2 = Empty
   textFA56 = Empty
   textFA57 = Empty
   textFA58 = Empty
   textFA59 = Empty
   textFA59_2 = Empty
   textFA60 = Empty
   textFA61 = Empty
   textFA61_2 = Empty
   textFA62 = Empty
   textFA62_2 = Empty
   textFA63 = Empty
   textFA64 = Empty
   textFA65 = Empty
   textFA66 = Empty
   textFA66_2 = Empty
   textFA67 = Empty
   textFA67_2 = Empty
   textFA68 = Empty
   'Modify by Amy 2015/08/24 改為下拉選單 原:textFA69
   cboStatus = Empty
   textFA71 = Empty
   textFA72 = Empty
    'Add By Cheng 2003/11/17
   textFA73 = Empty
   textFA74 = Empty
   textFA75 = Empty
   'Add By Sindy 2025/3/10
   textFA137 = Empty
   textFA138 = Empty
   textFA139 = Empty
   '2025/3/10 END
   'add by nickc 2005/12/02
   TextFA76 = Empty
   textFA77 = Empty
    'End
   textFA78 = Empty 'Add by Morgan 2006/10/18
   textFA85 = Empty 'Add by Morgan 2007/10/26
   textFA79 = Empty 'Add by Morgan 2008/1/16
   textFA80 = Empty 'Add by Morgan 2008/1/16
   textFA81 = Empty 'Add by Morgan 2008/1/16
   textFA82 = Empty 'Add by Morgan 2008/1/16
   textFA87 = Empty 'Add by Morgan 2008/3/13
   textFA88 = Empty 'Add by Morgan 2008/3/13
   textFA89 = Empty 'Add by Morgan 2008/3/13
   textFA90 = Empty 'Add by Morgan 2008/3/13
   textFA92 = Empty 'Add by Morgan 2008/6/3
   
   TextFA93 = Empty 'Add by Toni 2008/10/21
   textFA97 = Empty '2008/12/9 add by sonia
   
   textFA100 = Empty 'Add By Sindy 2011/3/10
   
   'Add By Sindy 2011/3/4
   textFA106 = Empty
   textFA107 = Empty
   textFA107_2 = Empty
'   textFA108 = Empty
   textFA109 = Empty
   textFA110 = Empty
   textFA111 = Empty
   textFA111_2 = Empty
   textFA112 = Empty
   textFA112_2 = Empty
   '2011/3/4 End
   textFA117 = Empty 'Add By Sindy 2013/8/15
   'Add by Amy 2017/01/05
   lblFA120 = Empty '控管智權人員名稱
   textFA120 = Empty '控管智權人員編號
   textFA120.Tag = Empty 'Add by Amy 2017/03/10
   textFA41.Tag = Empty 'Added by Lydia 2019/11/27
   
   'Add By Sindy 2012/6/5
   Combo1.ListIndex = 0
   '2012/6/5 End
   'Add By Sindy 2013/1/17
   For i = 0 To 1
      Combo2(i).ListIndex = 0
      Combo3(i).ListIndex = 0
   Next i
   '2013/1/17 End
   
   'Add by Morgan 2008/11/13 改用陣列以免控制項超過且以後再新增欄位也不必改
   For Each m_Txt In txtFA
      m_Txt = Empty
   Next
   lstDeveloper.Clear
   'end 2008/11/13
   
   For nIndex = 0 To TF_FA - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   'add by nickc 2006/03/17
   textCUID = ""
   Combo4.Text = "" 'Add By Sindy 2016/12/5
   Combo5.ListIndex = -1 'Add By Sindy 2021/3/3
   'Add by Amy 2022/11/25 +代理人來源
   cboSource.ListIndex = -1
   txtXYS02 = ""
   txtXYS03 = ""
   LblSourceN.Caption = "" 'X or Y編號名稱
   'end 2022/11/25
   'Add by Amy 2022/12/28
   txtXYS02.Tag = ""
   txtXYS03.Tag = ""
   cmdIntroduce.BackColor = &H8000000F 'Add by 2024/11/29
   
   'Add By Sindy 2025/1/7
   For Each m_Txt In Chk1K
      m_Txt = Empty
   Next
   '2025/1/7 END
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   Dim intSourceState As Integer 'Add by 2024/11/29
   
   textFA01.Locked = bEnable
   textFA02.Locked = bEnable
   textFA03.Locked = bEnable
   textFA04.Locked = bEnable
   textFA05.Locked = bEnable
   textFA06.Locked = bEnable
   textFA07.Locked = bEnable
   textFA08.Locked = bEnable
   textFA09.Locked = bEnable
   textFA10.Locked = bEnable
   textFA11.Locked = bEnable
   textFA12.Locked = bEnable
   textFA13.Locked = bEnable
   textFA14.Locked = bEnable
   textFA15.Locked = bEnable
   textFA16.Locked = bEnable
   textFA17.Locked = bEnable
   textFA18.Locked = bEnable
   textFA19.Locked = bEnable
   textFA20.Locked = bEnable
   textFA21.Locked = bEnable
   textFA22.Locked = bEnable
   textFA70.Locked = bEnable
   textFA23.Locked = bEnable
   textFA24.Locked = bEnable
   textFA25.Locked = bEnable
   textFA26.Locked = bEnable
   textFA27.Locked = bEnable
   textFA28.Locked = bEnable
   textFA29.Locked = bEnable
   textFA30.Locked = bEnable
   textFA31.Locked = bEnable
   textFA32.Locked = bEnable
   textFA33.Locked = bEnable
   textFA34.Locked = bEnable
   textFA35.Locked = bEnable
   textFA36.Locked = bEnable
   textFA37.Locked = bEnable
   textFA38.Locked = bEnable
   textFA39.Locked = bEnable
   textFA40.Locked = bEnable
   textFA41.Locked = bEnable
   textFA42.Locked = bEnable
'   textFA43.Locked = bEnable
   textFA44.Locked = bEnable
   textFA45.Locked = bEnable
   textFA52.Locked = bEnable
   textFA53.Locked = bEnable
   textFA54.Locked = bEnable
   textFA55.Locked = bEnable
   textFA56.Locked = bEnable
   textFA57.Locked = bEnable
   textFA58.Locked = bEnable
   textFA59.Locked = bEnable
   textFA60.Locked = bEnable
   textFA61.Locked = bEnable
   textFA62.Locked = bEnable
   textFA63.Locked = bEnable
   textFA64.Locked = bEnable
   textFA65.Locked = bEnable
   textFA66.Locked = bEnable
   textFA67.Locked = bEnable
   textFA68.Locked = bEnable
   'Modify by Amy 2015/08/24 改為下拉選單 原:textFA69
   cboStatus.Locked = bEnable
   'Add by Amy 2025/4/24 代理人狀態,並非操作者權限的下拉選項內容時,鎖住代理人狀態欄,不可修改
   If Pub_StrUserSt03 <> "M51" And (cboStatus = "不再使用" Or cboStatus = "不得代理" _
        Or cboStatus = "不得代理專利" Or cboStatus = "不得代理商標" Or cboStatus = "宣告破產") Then
      cboStatus.Locked = True
   End If
   'end 2025/4/24
   textFA71.Locked = bEnable
   textFA72.Locked = bEnable
    'Add By Cheng 2003/11/17
   textFA73.Locked = bEnable
   textFA74.Locked = bEnable
   textFA75.Locked = bEnable
    'End
   'Add By Sindy 2025/3/10
   textFA137.Locked = bEnable
   textFA138.Locked = bEnable
   textFA139.Locked = bEnable
   '2025/3/10 END
   'add by nickc 2005/12/02
   TextFA76.Locked = bEnable
   textFA77.Locked = bEnable
   textFA78.Locked = bEnable 'Add by Morgan 2006/10/18
   textFA85.Locked = bEnable 'Add by Morgan 2007/10/26
   textFA79.Locked = True 'Add by Morgan 2008/1/16
   textFA79.Enabled = bEnable 'Added by Morgan 2018/1/16
   textFA80.Locked = bEnable 'Add by Morgan 2008/1/16
   textFA81.Locked = bEnable 'Add by Morgan 2008/1/16
   textFA82.Locked = bEnable 'Add by Morgan 2008/1/16
   textFA87.Locked = bEnable 'Add by Morgan 2008/3/13
   textFA88.Locked = bEnable 'Add by Morgan 2008/3/13
   textFA89.Locked = bEnable 'Add by Morgan 2008/3/13
   textFA90.Locked = bEnable 'Add by Morgan 2008/3/13
   textFA92.Locked = bEnable 'Add by Morgan 2008/6/3
   
   TextFA93.Locked = bEnable 'Add by Toni 2008/10/21
   textFA97.Locked = bEnable '2008/12/9 add by sonia
   
   textFA100.Locked = bEnable 'Add By Sindy 2011/3/10
   
   'Add By Sindy 2011/3/4
   textFA106.Locked = bEnable
   textFA107.Locked = bEnable
'   textFA108.Locked = bEnable
   textFA109.Locked = bEnable
   textFA110.Locked = bEnable
   textFA111.Locked = bEnable
   textFA112.Locked = bEnable
   '2011/3/4 End
   textFA117.Locked = bEnable 'Add By Sindy 2013/8/15
   textFA120.Locked = bEnable 'Add by Amy 2017/01/05
   
   'Add By Sindy 2012/6/5
   Combo1.Locked = bEnable
   '2012/6/5 End
   'Add By Sindy 2013/1/17
   For i = 0 To 1
      Combo2(i).Locked = bEnable
      Combo3(i).Locked = bEnable
   Next i
   '2013/1/17 End
   
   'Add by Morgan 2008/11/13 改用陣列以免控制項超過且以後再新增欄位也不必改
   For Each m_Txt In txtFA
      m_Txt.Locked = bEnable
   Next
   'end 2008/11/13
   
   'Added by Morgan 2018/1/16
   txtFA(83).Locked = True
   txtFA(83).Enabled = bEnable
   txtFA(101).Locked = True
   txtFA(101).Enabled = bEnable
   'end 2018/1/16
   
   'Added by Lydia 2018/07/20 財務信箱(CF)
   txtFA(105).Locked = True
   txtFA(105).Enabled = bEnable
   'end 2018/07/20
   
   Combo4.Locked = bEnable 'Add By Sindy 2016/12/5
   Combo5.Locked = bEnable 'Add By Sindy 2021/3/3
   'Add by Amy 2022/11/25
   cboSource.Locked = True
   'Modify by Amy 2022/12/28 +if m_EditMode = 1,bug新增時未檢查必輸
   'Modify by 2024/11/29 改抓共用函數,避免有未改到
   If m_EditMode = 1 Then
      intSourceState = 6
   '[非]更名資料才可改
   ElseIf m_EditMode = 2 And textFA02 = "0" Then
      intSourceState = 7
   Else
      intSourceState = 8
   End If
   Call Pub_SetCboComeSource(intSourceState, Me.Name, cboSource, , txtXYS02, txtXYS03)
   'end 2024/11/29
   'end 2022/11/25
   
   Frame1K.Enabled = Not bEnable 'Add By Sindy 2025/1/7
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textFA01.Locked = bEnable
   textFA02.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strTp(3) As String 'Add by Amy 2022/11/25
Dim arrID 'Add By Sindy 2025/1/7
   
   strSql = "SELECT * FROM FAGENT " & _
            "WHERE FA01 = '" & m_CurrKEY(0) & "' AND " & _
                  "FA02 = '" & m_CurrKEY(1) & "' "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("FA01")) = False Then: textFA01 = rsTmp.Fields("FA01")
      If IsNull(rsTmp.Fields("FA02")) = False Then: textFA02 = rsTmp.Fields("FA02")
      If IsNull(rsTmp.Fields("FA03")) = False Then: textFA03 = rsTmp.Fields("FA03")
      If IsNull(rsTmp.Fields("FA04")) = False Then: textFA04 = rsTmp.Fields("FA04")
      If IsNull(rsTmp.Fields("FA05")) = False Then: textFA05 = rsTmp.Fields("FA05")
      If IsNull(rsTmp.Fields("FA06")) = False Then: textFA06 = rsTmp.Fields("FA06")
      If IsNull(rsTmp.Fields("FA07")) = False Then: textFA07 = rsTmp.Fields("FA07")
      If IsNull(rsTmp.Fields("FA08")) = False Then: textFA08 = rsTmp.Fields("FA08")
      If IsNull(rsTmp.Fields("FA09")) = False Then: textFA09 = rsTmp.Fields("FA09")
      If IsNull(rsTmp.Fields("FA10")) = False Then: textFA10 = rsTmp.Fields("FA10")
      ' 開發日期
      If IsNull(rsTmp.Fields("FA11")) = False Then
         If rsTmp.Fields("FA11") <> "0" Then
            textFA11 = TAIWANDATE(rsTmp.Fields("FA11"))
         End If
      End If
      If IsNull(rsTmp.Fields("FA12")) = False Then: textFA12 = rsTmp.Fields("FA12")
      If IsNull(rsTmp.Fields("FA13")) = False Then: textFA13 = rsTmp.Fields("FA13")
      If IsNull(rsTmp.Fields("FA14")) = False Then: textFA14 = rsTmp.Fields("FA14")
      If IsNull(rsTmp.Fields("FA15")) = False Then: textFA15 = rsTmp.Fields("FA15")
      If IsNull(rsTmp.Fields("FA16")) = False Then: textFA16 = rsTmp.Fields("FA16")
      If IsNull(rsTmp.Fields("FA17")) = False Then: textFA17 = rsTmp.Fields("FA17")
      If IsNull(rsTmp.Fields("FA18")) = False Then: textFA18 = rsTmp.Fields("FA18")
      If IsNull(rsTmp.Fields("FA19")) = False Then: textFA19 = rsTmp.Fields("FA19")
      If IsNull(rsTmp.Fields("FA20")) = False Then: textFA20 = rsTmp.Fields("FA20")
      If IsNull(rsTmp.Fields("FA21")) = False Then: textFA21 = rsTmp.Fields("FA21")
      If IsNull(rsTmp.Fields("FA22")) = False Then: textFA22 = rsTmp.Fields("FA22")
      If IsNull(rsTmp.Fields("FA70")) = False Then: textFA70 = rsTmp.Fields("FA70")
      If IsNull(rsTmp.Fields("FA23")) = False Then: textFA23 = rsTmp.Fields("FA23")
      If IsNull(rsTmp.Fields("FA24")) = False Then: textFA24 = rsTmp.Fields("FA24")
      If IsNull(rsTmp.Fields("FA25")) = False Then: textFA25 = rsTmp.Fields("FA25")
      If IsNull(rsTmp.Fields("FA26")) = False Then: textFA26 = rsTmp.Fields("FA26")
      ' 全部折扣起始日
      If IsNull(rsTmp.Fields("FA27")) = False Then
         If rsTmp.Fields("FA27") <> "0" Then
            textFA27 = TAIWANDATE(rsTmp.Fields("FA27"))
         End If
      End If
      If IsNull(rsTmp.Fields("FA28")) = False Then: textFA28 = rsTmp.Fields("FA28")
      If IsNull(rsTmp.Fields("FA29")) = False Then: textFA29 = rsTmp.Fields("FA29")
      If IsNull(rsTmp.Fields("FA30")) = False Then: textFA30 = rsTmp.Fields("FA30")
      If IsNull(rsTmp.Fields("FA31")) = False Then: textFA31 = rsTmp.Fields("FA31")
      If IsNull(rsTmp.Fields("FA32")) = False Then: textFA32 = rsTmp.Fields("FA32")
      If IsNull(rsTmp.Fields("FA33")) = False Then: textFA33 = rsTmp.Fields("FA33")
      If IsNull(rsTmp.Fields("FA34")) = False Then: textFA34 = rsTmp.Fields("FA34")
      If IsNull(rsTmp.Fields("FA35")) = False Then: textFA35 = rsTmp.Fields("FA35")
      If IsNull(rsTmp.Fields("FA36")) = False Then: textFA36 = rsTmp.Fields("FA36")
      If IsNull(rsTmp.Fields("FA37")) = False Then: textFA37 = rsTmp.Fields("FA37")
      If IsNull(rsTmp.Fields("FA38")) = False Then: textFA38 = rsTmp.Fields("FA38")
      If IsNull(rsTmp.Fields("FA39")) = False Then: textFA39 = rsTmp.Fields("FA39")
      If IsNull(rsTmp.Fields("FA40")) = False Then: textFA40 = rsTmp.Fields("FA40")
      If IsNull(rsTmp.Fields("FA41")) = False Then: textFA41 = rsTmp.Fields("FA41")
      textFA41.Tag = textFA41.Text 'Added by Lydia 2019/11/27
      If IsNull(rsTmp.Fields("FA42")) = False Then: textFA42 = rsTmp.Fields("FA42")
      'Modify By Sindy 2013/1/17
'      If IsNull(rsTmp.Fields("FA43")) = False Then: textFA43 = rsTmp.Fields("FA43")
      If IsNull(rsTmp.Fields("FA43")) = False Then
         For i = 0 To Combo2(0).ListCount - 1
            Combo2(0).ListIndex = i
            If InStr(Combo2(0).Text, rsTmp.Fields("FA43")) > 0 Then
               Exit For
            End If
         Next
      Else
         Combo2(0).ListIndex = 0
      End If
      '2013/1/17 End
      If IsNull(rsTmp.Fields("FA44")) = False Then: textFA44 = rsTmp.Fields("FA44")
      If IsNull(rsTmp.Fields("FA45")) = False Then: textFA45 = rsTmp.Fields("FA45")
      If IsNull(rsTmp.Fields("FA52")) = False Then: textFA52 = rsTmp.Fields("FA52")
      If IsNull(rsTmp.Fields("FA53")) = False Then: textFA53 = rsTmp.Fields("FA53")
      If IsNull(rsTmp.Fields("FA54")) = False Then: textFA54 = rsTmp.Fields("FA54")
      If IsNull(rsTmp.Fields("FA55")) = False Then: textFA55 = rsTmp.Fields("FA55")
      If IsNull(rsTmp.Fields("FA56")) = False Then: textFA56 = rsTmp.Fields("FA56")
      If IsNull(rsTmp.Fields("FA57")) = False Then: textFA57 = rsTmp.Fields("FA57")
      If IsNull(rsTmp.Fields("FA58")) = False Then: textFA58 = rsTmp.Fields("FA58")
      If IsNull(rsTmp.Fields("FA59")) = False Then: textFA59 = rsTmp.Fields("FA59")
      If IsNull(rsTmp.Fields("FA60")) = False Then: textFA60 = rsTmp.Fields("FA60")
      If IsNull(rsTmp.Fields("FA61")) = False Then: textFA61 = rsTmp.Fields("FA61")
      If IsNull(rsTmp.Fields("FA62")) = False Then: textFA62 = rsTmp.Fields("FA62")
      If IsNull(rsTmp.Fields("FA63")) = False Then: textFA63 = rsTmp.Fields("FA63")
      If IsNull(rsTmp.Fields("FA64")) = False Then: textFA64 = rsTmp.Fields("FA64")
      If IsNull(rsTmp.Fields("FA65")) = False Then: textFA65 = rsTmp.Fields("FA65")
      If IsNull(rsTmp.Fields("FA66")) = False Then: textFA66 = rsTmp.Fields("FA66")
      If IsNull(rsTmp.Fields("FA67")) = False Then: textFA67 = rsTmp.Fields("FA67")
      If IsNull(rsTmp.Fields("FA68")) = False Then: textFA68 = rsTmp.Fields("FA68")
      If IsNull(rsTmp.Fields("FA69")) = False Then: cboStatus = rsTmp.Fields("FA69") 'Modify by Amy 2015/08/24 原:textFA69
      If IsNull(rsTmp.Fields("FA71")) = False Then: textFA71 = rsTmp.Fields("FA71")
      If IsNull(rsTmp.Fields("FA72")) = False Then: textFA72 = rsTmp.Fields("FA72")
        'Add By Cheng 2003/11/17
      If IsNull(rsTmp.Fields("FA73")) = False Then: textFA73 = rsTmp.Fields("FA73")
      If IsNull(rsTmp.Fields("FA74")) = False Then: textFA74 = rsTmp.Fields("FA74")
      ' 全部折扣起始日
      If IsNull(rsTmp.Fields("FA75")) = False Then
         If rsTmp.Fields("FA75") <> "0" Then
            textFA75 = TAIWANDATE(rsTmp.Fields("FA75"))
         End If
      End If
      
      'Add By Sindy 2025/3/10
      If IsNull(rsTmp.Fields("FA137")) = False Then: textFA137 = rsTmp.Fields("FA137")
      If IsNull(rsTmp.Fields("FA138")) = False Then: textFA138 = rsTmp.Fields("FA138")
      ' 全部折扣終止日
      If IsNull(rsTmp.Fields("FA139")) = False Then
         If rsTmp.Fields("FA139") <> "0" Then
            textFA139 = TAIWANDATE(rsTmp.Fields("FA139"))
         End If
      End If
      '2025/3/10 END
      
      'add by nickc 2005/12/02
      If IsNull(rsTmp.Fields("FA76")) = False Then: TextFA76 = rsTmp.Fields("FA76")
      If IsNull(rsTmp.Fields("FA77")) = False Then: textFA77 = rsTmp.Fields("FA77")
      If textFA77 = "Y" Then textFA01.ForeColor = &HFF&: textFA02.ForeColor = &HFF& Else textFA01.ForeColor = &H80000008: textFA02.ForeColor = &H80000008
      'End
      If IsNull(rsTmp.Fields("FA78")) = False Then: textFA78 = rsTmp.Fields("FA78") 'Add by Morgan 2006/10/18
      If IsNull(rsTmp.Fields("FA85")) = False Then: textFA85 = rsTmp.Fields("FA85") 'Add by Morgan 2007/10/26
      textFA85.Tag = textFA85.Text 'Added by Lydia 2019/05/27
      If IsNull(rsTmp.Fields("FA79")) = False Then: textFA79 = rsTmp.Fields("FA79") 'Add by Morgan 2008/1/16
      If IsNull(rsTmp.Fields("FA80")) = False Then: textFA80 = rsTmp.Fields("FA80") 'Add by Morgan 2008/1/16
      If IsNull(rsTmp.Fields("FA81")) = False Then: textFA81 = rsTmp.Fields("FA81") 'Add by Morgan 2008/1/16
      If IsNull(rsTmp.Fields("FA82")) = False Then: textFA82 = rsTmp.Fields("FA82") 'Add by Morgan 2008/1/16
      If IsNull(rsTmp.Fields("FA87")) = False Then: textFA87 = rsTmp.Fields("FA87") 'Add by Morgan 2008/3/13
      If IsNull(rsTmp.Fields("FA88")) = False Then: textFA88 = rsTmp.Fields("FA88") 'Add by Morgan 2008/3/13
      If IsNull(rsTmp.Fields("FA89")) = False Then: textFA89 = rsTmp.Fields("FA89") 'Add by Morgan 2008/3/13
      If IsNull(rsTmp.Fields("FA90")) = False Then: textFA90 = rsTmp.Fields("FA90") 'Add by Morgan 2008/3/13
      If IsNull(rsTmp.Fields("FA92")) = False Then: textFA92 = rsTmp.Fields("FA92") 'Add by Morgan 2008/6/3
      
      If IsNull(rsTmp.Fields("FA93")) = False Then: TextFA93 = rsTmp.Fields("FA93") 'add by Toni 2008/10/21
      If IsNull(rsTmp.Fields("FA97")) = False Then: textFA97 = rsTmp.Fields("FA97") '2008/12/9 add by sonia
      
      If IsNull(rsTmp.Fields("FA100")) = False Then: textFA100 = rsTmp.Fields("FA100") 'Add By Sindy 2011/3/10
      
      'Add By Sindy 2011/3/4
      If IsNull(rsTmp.Fields("FA106")) = False Then: textFA106 = rsTmp.Fields("FA106")
      If IsNull(rsTmp.Fields("FA107")) = False Then: textFA107 = rsTmp.Fields("FA107")
      'Modify By Sindy 2013/1/17
'      If IsNull(rsTmp.Fields("FA108")) = False Then: textFA108 = rsTmp.Fields("FA108")
      If IsNull(rsTmp.Fields("FA108")) = False Then
         For i = 0 To Combo2(1).ListCount - 1
            Combo2(1).ListIndex = i
            If InStr(Combo2(1).Text, rsTmp.Fields("FA108")) > 0 Then
               Exit For
            End If
         Next
      Else
         Combo2(1).ListIndex = 0
      End If
      '2013/1/17 End
      If IsNull(rsTmp.Fields("FA109")) = False Then: textFA109 = rsTmp.Fields("FA109")
      If IsNull(rsTmp.Fields("FA110")) = False Then: textFA110 = rsTmp.Fields("FA110")
      If IsNull(rsTmp.Fields("FA111")) = False Then: textFA111 = rsTmp.Fields("FA111")
      If IsNull(rsTmp.Fields("FA112")) = False Then: textFA112 = rsTmp.Fields("FA112")
      '2011/3/4 End
      If IsNull(rsTmp.Fields("FA117")) = False Then: textFA117 = rsTmp.Fields("FA117") 'Add By Sindy 2013/8/15
      If IsNull(rsTmp.Fields("FA119")) = False Then: Combo4.Text = rsTmp.Fields("FA119") 'Add By Sindy 2016/12/5
      'Add By Sindy 2021/3/3
      If IsNull(rsTmp.Fields("FA126")) = False Then
         Combo5.ListIndex = rsTmp.Fields("FA126")
      End If
      'Add by Amy 2017/01/05
      If IsNull(rsTmp.Fields("FA120")) = False Then
        textFA120 = rsTmp.Fields("FA120")
        textFA120.Tag = textFA120 'Add by Amy 2017/03/10
      End If
      
      'Add By Sindy 2012/6/5
      If IsNull(rsTmp.Fields("FA113")) = False Then
         For i = 0 To Combo1.ListCount - 1
            Combo1.ListIndex = i
            If InStr(Combo1.Text, rsTmp.Fields("FA113")) > 0 Then
               Exit For
            End If
         Next
      Else
         Combo1.ListIndex = 0
      End If
      '2012/6/5 End
      'Add By Sindy 2013/1/17
      If IsNull(rsTmp.Fields("FA115")) = False Then
         Combo3(0).ListIndex = rsTmp.Fields("FA115")
      Else
         Combo3(0).ListIndex = 0
      End If
      If IsNull(rsTmp.Fields("FA116")) = False Then
         Combo3(1).ListIndex = rsTmp.Fields("FA116")
      Else
         Combo3(1).ListIndex = 0
      End If
      '2013/1/17 End
      
      'Add by Morgan 2008/11/13 改用陣列以免控制項超過且以後再新增欄位也不必改
      For Each m_Txt In txtFA
         If IsNull(rsTmp.Fields("FA" & m_Txt.Index)) = False Then: m_Txt = rsTmp.Fields("FA" & m_Txt.Index)
      Next
      'Modified by Lydia 2021/12/14 改為Form 2.0元件
      PUB_SetUserList lstDeveloper, "" & rsTmp.Fields("FA94"), True
      'end 2008/11/13
      
      'Add By Sindy 2025/1/7
      If IsNull(rsTmp.Fields("FA135")) = False Then
         arrID = Split(rsTmp.Fields("FA135"), ",")
         For intI = UBound(arrID) To LBound(arrID) Step -1
            Chk1K(Val(arrID(intI)) - 1).Value = 1
         Next intI
      End If
      '2025/1/7 END
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
      textFA10_Validate False
      textFA30_Validate False
      textFA38_Validate False
      textFA55_Validate False
      textFA59_Validate False
      textFA61_Validate False
      textFA62_Validate False
      textFA66_Validate False
      textFA67_Validate False
      textFA71_Validate False
      textFA72_Validate False
      'Add By Sindy 2011/3/4
      textFA107_Validate False
      textFA111_Validate False
      textFA112_Validate False
      '2011/3/4 End
      textFA120_Validate False 'Add by Amy 2017/01/05
      'Add by Amy 2022/11/25
      If strSrvDate(1) >= 代理人來源啟用日 Then
            If IsNull(rsTmp.Fields("FA127")) Then
               cboSource.ListIndex = -1
            Else
              cboSource.ListIndex = rsTmp.Fields("FA127")
            End If
            Call Pub_GetXYSource(1, textFA01, strTp(0), strTp(1), strTp(2))
            txtXYS02 = strTp(0)
            LblSourceN.Caption = strTp(1)
            txtXYS03.Text = strTp(2)
            'Add by Amy 2022/12/28
            txtXYS02.Tag = txtXYS02
            txtXYS03.Tag = txtXYS03
            'Add by Amy 2023/07/12 從Form_Load及 onWork 搬過來,解決上下一筆 按鈕顏色不會重抓
            cmdIntroduce.BackColor = &H8000000F
            If Pub_GetXYSource(2, Left(textFA01, 8)) = True Then
               cmdIntroduce.BackColor = &HFFFF80
            End If
      End If
      'end 2022/11/25
   End If
   rsTmp.Close
   'Add by Morgan 2006/1/10
   textFA05.Tag = textFA05.Text
   textFA63.Tag = textFA63.Text
   textFA64.Tag = textFA64.Text
   textFA65.Tag = textFA65.Text
   '2006/1/10 end
   'add by nickc 2006/03/16
   textFA03.Tag = textFA03.Text
   
   'add by nickc 2006/12/26
   textFA04.Tag = textFA04.Text
   textFA06.Tag = textFA06.Text
   textFA12.Tag = textFA12.Text
   textFA13.Tag = textFA13.Text
   textFA14.Tag = textFA14.Text
   textFA15.Tag = textFA15.Text
   textFA17.Tag = textFA17.Text
   textFA18.Tag = textFA18.Text
   textFA19.Tag = textFA19.Text
   textFA20.Tag = textFA20.Text
   textFA21.Tag = textFA21.Text
   textFA22.Tag = textFA22.Text
   textFA23.Tag = textFA23.Text
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'Modify By Sindy 2021/12/7 因改Form2.0無法用 p_Listbox.ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
'先 Public 改 Private
'Mark by Lydia 2021/12/14 改basQuery的共用模組
'Private Sub PUB_SetUserList(p_Listbox As Object, p_stNums As String)
'   Dim arrID, stSQL As String, intR As Integer, rstTmp As ADODB.Recordset
'   p_Listbox.Clear
'   If p_stNums <> "" Then
'      stSQL = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
'      intR = 1
'      Set rstTmp = ClsLawReadRstMsg(intR, stSQL)
'      If intR = 1 Then
'         arrID = Split(p_stNums, ",")
'         With rstTmp
'         '照原順序排
'         For intI = UBound(arrID) To LBound(arrID) Step -1
'            .MoveFirst
'            Do While Not .EOF
'               If .Fields("st01") = arrID(intI) Then
'                  p_Listbox.AddItem "" & .Fields(1), 0
'                  '2012/2/14 MODIFY BY SONIA 員工編號已可非數字需做轉換
'                  'Modify By Sindy 2021/12/7 Mark
'                  'p_Listbox.ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
'                  '2021/12/7 END
'                  .MoveLast
'               End If
'               .MoveNext
'            Loop
'         Next
'         End With
'      End If
'   End If
'   Set rstTmp = Nothing
'End Sub
'end 2021/12/14

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("FA46")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FA46")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("FA46"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FA47")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FA47")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("FA47"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FA48")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FA48")) = False Then
         strTemp = rsSrcTmp.Fields("FA48")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FA49")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FA49")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("FA49"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FA50")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FA50")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("FA50"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("FA51")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("FA51")) = False Then
         strTemp = rsSrcTmp.Fields("FA51")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   'Modified by Morgan 2014/6/4 內容太長無法完全顯示,去掉欄位間的冒號
   textCUID = "CREATE : " & strCName & " " & _
              strCDate & " " & _
              strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              strUDate & " " & _
              strUTime
              
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT FA01,FA02 FROM FAGENT " & _
               "WHERE FA01 = '" & m_CurrKEY(0) & "' AND " & _
                     "FA02 = (SELECT MIN(FA02) FROM FAGENT " & _
                             "WHERE FA01 = '" & m_CurrKEY(0) & "' AND " & _
                                   "FA02 > '" & m_CurrKEY(1) & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("FA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("FA01")
         If IsNull(rsTmp.Fields("FA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("FA02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT FA01,FA02 FROM FAGENT " & _
               "WHERE FA01 = (SELECT MIN(FA01) FROM FAGENT " & _
                              "WHERE FA01 > '" & m_CurrKEY(0) & "') AND " & _
                     "FA02 = (SELECT MIN(FA02) FROM FAGENT " & _
                              "WHERE FA01 = (SELECT MIN(FA01) FROM FAGENT " & _
                                             "WHERE FA01 > '" & m_CurrKEY(0) & "')) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("FA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("FA01")
         If IsNull(rsTmp.Fields("FA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("FA02")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT FA01,FA02 FROM FAGENT " & _
            "WHERE FA01 = '" & m_CurrKEY(0) & "' AND " & _
                  "FA02 = (SELECT MAX(FA02) FROM FAGENT " & _
                          "WHERE FA01 = '" & m_CurrKEY(0) & "' AND " & _
                                "FA02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("FA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("FA01")
      If IsNull(rsTmp.Fields("FA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("FA02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT FA01,FA02 FROM FAGENT " & _
            "WHERE FA01 = (SELECT MAX(FA01) FROM FAGENT " & _
                           "WHERE FA01 < '" & m_CurrKEY(0) & "') AND " & _
                  "FA02 = (SELECT MAX(FA02) FROM FAGENT " & _
                           "WHERE FA01 = (SELECT MAX(FA01) FROM FAGENT " & _
                                          "WHERE FA01 < '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("FA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("FA01")
      If IsNull(rsTmp.Fields("FA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("FA02")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT FA01,FA02 FROM FAGENT " & _
            "WHERE FA01 = '" & m_CurrKEY(0) & "' AND " & _
                  "FA02 = (SELECT MIN(FA02) FROM FAGENT " & _
                          "WHERE FA01 = '" & m_CurrKEY(0) & "' AND " & _
                                "FA02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("FA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("FA01")
      If IsNull(rsTmp.Fields("FA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("FA02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT FA01,FA02 FROM FAGENT " & _
            "WHERE FA01 = (SELECT MIN(FA01) FROM FAGENT " & _
                           "WHERE FA01 > '" & m_CurrKEY(0) & "') AND " & _
                  "FA02 = (SELECT MIN(FA02) FROM FAGENT " & _
                           "WHERE FA01 = (SELECT MIN(FA01) FROM FAGENT " & _
                                          "WHERE FA01 > '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("FA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("FA01")
      If IsNull(rsTmp.Fields("FA02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("FA02")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
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

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
'edit by nickc 2006/11/13
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
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
   'Add By Sindy 2014/8/29 當focus在備註欄時按enter鍵維持換行功能而不是存檔功能
   'Modify by Amy 2022/12/28 +txtXYS03
   If KeyAscii = 13 And (UCase(Me.ActiveControl.Name) = UCase("textFA29") Or UCase(Me.ActiveControl.Name) = UCase("txtXYS03")) Then
      'If Me.ActiveControl.Index = 15 Then
         Exit Sub
      'End If
   End If
   '2014/8/29 END
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTp As String 'Add by Amy 2022/11/25
   
   m_SubMode = 0
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         textFA77.Locked = False 'Add by Amy 2013/11/04
         textFA77.Enabled = True
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         'Add by Amy 2013/10/29 FA77為B.宣告破產 則不可修改呆帳記錄
         strExc(0) = "Select FA103 From Fagent Where FA01 = '" & textFA01 & "' AND " & _
                          "FA02 = '" & textFA02 & "' And FA103='B' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
             textFA77.Locked = True
             textFA77.Enabled = False
         Else
             textFA77.Locked = False
             textFA77.Enabled = True
         End If
        'end 2013/10/29
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            'Add by Amy 2018/07/18 有往來記錄不可刪除
            strExc(0) = "Select CR03 From ContactRecord Where CR03 = '" & textFA01 & textFA02 & "'  " & _
                    "Union Select COR03 From ContactRecord1 Where COR03 = '" & textFA01 & textFA02 & "'   "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                MsgBox "有往來記錄不可刪除！"
                Exit Sub
            End If
            'end 2018/07/18
            'Add by Amy 2022/11/25 若存在XYS02介紹來源編號,則不可刪
            'Modify by 2024/11/29 考慮多筆,改訊息至共用
            If textFA02 = "0" And Pub_GetXYSource(2, textFA01, , , , Me.Name, strTp) = True Then
               MsgBox strTp, vbOKOnly, "注意"
               Exit Sub
            End If
            m_EditMode = 3
            OnWork
            UpdateToolbarState
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
         'Modify By Sindy 2014/8/29 Mark
         'PUB_FilterFormText Me 'Add by Morgan 2008/6/20 修正畫面所有含跳行符號的文字框
         '2014/8/29 END
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         'Addedd by Lydia 2017/03/31 修正畫面所有含跳行符號的文字框
         'Modify by 2024/11/29 +txtXYS03
         PUB_FilterFormText Me, "textFA45,textFA110,textFA92,textFA29,txtXYS03"
         'end 2017/03/31
         If CheckDataValid() = True Then
            UpdateFieldNewData
            OnWork
            UpdateToolbarState
        End If
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
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
      'SSTab1.Tab = 0
   End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2024/01/22  國外潛在客戶維護轉號存檔切至此畫面-陳金蓮
   If m_PrevForm Is Nothing = False Then
      If UCase(m_PrevForm.Name) = "FRM140402" Then
         m_PrevForm.AfterTransfer
      End If
      m_PrevForm.Show 'Added by Lydia 2024/02/22
   End If
   
   'Add By Cheng 2002/07/18
   Set frm050705 = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   'Modify By Sindy 2025/3/10 mark:因畫面上移位置了
'   'Modify by 2024/11/29 從Form_Load搬過來,否則切其他頁籤會有殘影
'   If SSTab1.Tab = 3 Then
'      'Add by Amy 2024/03/08 隱藏延展單筆不跑,將FCT註冊費自動代繳移位
'      Label8(2).Left = 120
'      TextFA93.Left = 1900
'   End If
End Sub

'add by nickc 2008/01/11 把駐點給前面的
Private Sub textFA02_LostFocus()
   Select Case m_EditMode
      Case 1:
         If IsRecordExist(textFA01, textFA02) = True Then
            textFA01.SetFocus
            Exit Sub
         End If
      Case Else:
   End Select
End Sub

'Added by Morgan 2011/12/30
Private Sub textFA10_Change()
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   SetFA100
End Sub

'Add By Sindy 2011/1/14
Private Sub textFA100_GotFocus()
   CloseIme
   TextInverse textFA100
End Sub
Private Sub textFA100_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub textFA100_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   'Modified by Morgan 2011/12/30 專利雙週報欄位改放 N(不寄)
   'If textFA100.Text = "" Then Exit Sub
   'If textFA100.Text <> "Y" Then
   '   ShowMsg "輸入錯誤 !"
   '   Cancel = True
   'Else
   '   If textFA10 <> "020" Then
   '      ShowMsg "代理人國籍不是大陸，不可設為要寄專利雙週報 !"
   '      Cancel = True
   '   Else
   '      If TextFA76 = "C" Then
   '         ShowMsg "代理人性質為其他時，不可設為要寄專利雙週報 !"
   '         Cancel = True
   '      End If
   '   End If
   'End If
   Cancel = FA100CheckError()
   'end 2011/12/30
End Sub
'2011/1/14 End

'Add By Sindy 2011/3/4
Private Sub textFA117_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   Select Case textFA117
      Case "Y", "":
      Case Else:
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "不催延展只可輸入Y或空白"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA117_GotFocus
         End Select
   End Select
End Sub
'2011/3/4 End

'Add by Amy 2017/01/05 +管控智權人員
Private Sub textFA120_GotFocus()
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    
    textFA120.Locked = True
    If Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "P2" Then
        textFA120.Locked = False
        textFA120.TabStop = True
        CloseIme
        TextInverse textFA120
    Else
        textFA120.TabStop = False
    End If
End Sub

Private Sub textFA120_KeyPress(KeyAscii As Integer)
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA120_Validate(Cancel As Boolean)
    Dim stFa120No As String, stFA120Name As String

    If textFA10 < "010" Then Exit Sub

    lblFA120 = ""
    If textFA120 <> MsgText(601) Then
        If m_EditMode <> 1 And m_EditMode <> 2 Then
            lblFA120 = GetStaffName(textFA120, True)
        ElseIf Left(textFA120, 4) <> "MCTF" Then
            MsgBox Left(Label53, Len(Label53) - 1) & " 只可輸入MCTF開頭編號！", vbExclamation
            Cancel = True
            SSTab1.Tab = 3
            textFA120.SetFocus
            textFA120_GotFocus
            Exit Sub
        Else
            lblFA120 = GetStaffName(textFA120)
            If lblFA120 = MsgText(601) Then
                MsgBox Left(Label53, Len(Label53) - 1) & " 輸入錯誤請確認！", vbExclamation
                Cancel = True
                SSTab1.Tab = 3
                textFA120.SetFocus
                textFA120_GotFocus
                Exit Sub
            End If
        End If
    End If
End Sub
'end 2017/01/05

Private Sub textFA16_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub

Private Sub textFA16_Validate(Cancel As Boolean)
   If textFA16.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textFA16.Text)
End Sub

'Mark by Amy 2016/08/05 因會先輸地址才輸國籍所以此不檢查
''Add by Amy 2016/06/30 +依國籍確認臺灣地址格式
'Private Sub textFA17_LostFocus()
'    Dim strTit As String
'    Dim strMsg As String, strAddr As String
'    Dim nResponse
'    Dim strTemp As String 'Add by Amy 2016/06/30
'
'    If textFA55 >= "010" Or IsEmptyText(textFA17) = True Then Exit Sub
'
'    If CheckAddrData(textFA17, strTemp) = False Then
'        strTit = "檢核資料"
'        strMsg = "代理人地址(中)" & strTemp
'        If InStr(strTemp, "格式") > 0 Then
'            If bolShow100135 = False Then
'                nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'                bolShow100135 = True
'                frm100135.Show vbModal
'                bolShow100135 = False
'                textFA17_GotFocus
'                textFA17.SetFocus
'                Exit Sub
'            Else
'                Exit Sub
'            End If
'        Else
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textFA17_GotFocus
'            textFA17.SetFocus
'            Exit Sub
'        End If
'    End If
'End Sub

'add by nickc 2005/06/15 日文地址要轉全形
Private Sub textFA23_KeyPress(KeyAscii As ReturnInteger)
KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub textFA30_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA107_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA38_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA39_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA40_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA41_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA42_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Private Sub textFA43_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''Add By Sindy 2011/3/4
'Private Sub textFA108_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

Private Sub textFA44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA109_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2013/8/15
Private Sub textFA117_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA59_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA61_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA62_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA66_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA67_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA68_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA71_GotFocus()
   InverseTextBox textFA71
End Sub

Private Sub textFA71_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA71_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA71_2 = Empty
   If IsEmptyText(textFA71) = False Then
      Select Case Mid(textFA71, 1, 1)
         Case "X":
            textFA71_2 = GetCustomerName(textFA71, 0)
         Case "Y":
            textFA71_2 = GetFAgentName(textFA71)
         Case Else:
            textFA71_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "專利D/N固定列印對象代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA71_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA71_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "專利D/N固定列印對象代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA71_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA111_GotFocus()
   InverseTextBox textFA111
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA111_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA111_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA111_2 = Empty
   If IsEmptyText(textFA111) = False Then
      Select Case Mid(textFA111, 1, 1)
         Case "X":
            textFA111_2 = GetCustomerName(textFA111, 0)
         Case "Y":
            textFA111_2 = GetFAgentName(textFA111)
         Case Else:
            textFA111_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "商標D/N固定列印對象代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA111_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA111_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商標D/N固定列印對象代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA111_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

Private Sub textFA72_GotFocus()
   InverseTextBox textFA72
End Sub

Private Sub textFA72_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA72_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA72_2 = Empty
   If IsEmptyText(textFA72) = False Then
      Select Case Mid(textFA72, 1, 1)
         Case "X":
            textFA72_2 = GetCustomerName(textFA72, 0)
         Case "Y":
            textFA72_2 = GetFAgentName(textFA72)
         Case Else:
            textFA72_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "年費D/N列印對象代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA72_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA72_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "年費D/N列印對象代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA72_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA112_GotFocus()
   InverseTextBox textFA112
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA112_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA112_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA112_2 = Empty
   If IsEmptyText(textFA112) = False Then
      Select Case Mid(textFA112, 1, 1)
         Case "X":
            textFA112_2 = GetCustomerName(textFA112, 0)
         Case "Y":
            textFA112_2 = GetFAgentName(textFA112)
         Case Else:
            textFA112_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "延展D/N列印對象代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA112_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA112_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "延展D/N列印對象代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA112_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

Private Sub textFA73_GotFocus()
    TextInverse Me.textFA73
End Sub

Private Sub textFA73_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA73) = False Then
      If IsNumeric(textFA73) = False Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商標全部折扣只可輸入數值資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA73_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

Private Sub textFA74_GotFocus()
    TextInverse Me.textFA74
End Sub

Private Sub textFA74_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA74) = False Then
      If IsNumeric(textFA74) = False Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商標申請/翻議折扣只可輸入數值資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA74_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

Private Sub textFA75_GotFocus()
    TextInverse Me.textFA75
End Sub

Private Sub textFA75_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA75) = False Then
      If CheckIsTaiwanDate(textFA75, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商標全部折扣起始日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA75_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2025/3/10
Private Sub textFA137_GotFocus()
    TextInverse Me.textFA137
End Sub
Private Sub textFA137_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA137) = False Then
      If IsNumeric(textFA137) = False Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商標全部折扣只可輸入數值資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA137_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub
Private Sub textFA138_GotFocus()
    TextInverse Me.textFA138
End Sub
Private Sub textFA138_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA138) = False Then
      If IsNumeric(textFA138) = False Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商標全部折扣只可輸入數值資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA138_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub
Private Sub textFA139_GotFocus()
    TextInverse Me.textFA139
End Sub

Private Sub textFA139_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA139) = False Then
      If CheckIsTaiwanDate(textFA139, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商標全部折扣終止日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA139_GotFocus
      End If
   End If
End Sub
'2025/3/10 END

'Added by Morgan 2011/12/30
Private Sub TextFA76_Change()
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   SetFA100
End Sub

Private Sub TextFA76_GotFocus()
   CloseIme
   TextInverse Me.TextFA76
End Sub

Private Sub TextFA76_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'僅能輸入 A or B OR C
If (KeyAscii > 67 Or KeyAscii < 65) And KeyAscii <> 8 And KeyAscii <> 44 Then
    KeyAscii = 0
End If
End Sub

Private Sub TextFA76_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(TextFA76) = True Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "性質不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      TextFA76.SetFocus
      TextFA76_GotFocus
      Exit Sub
   End If
   'Add By Sindy 2011/3/10
   '若為新增或修改狀態, 預設下列欄位值
   If m_EditMode = "1" Or m_EditMode = "2" Then
       '預設是否寄發專利雙週報
       'Modified by Morgan 2011/12/30 專利雙週報欄位改放 N(不寄)
       'If IsEmptyText(textFA100) = True Then
       '   If textFA10 = "020" And TextFA76 <> "C" Then
       '      textFA100 = "Y"
       '   End If
       'End If
       Cancel = FA100CheckError()
       'end 2011/12/30
   End If
End Sub

Private Sub textFA77_GotFocus()
   TextInverse Me.textFA77
End Sub

Private Sub textFA77_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'僅能輸入 Y or space
If KeyAscii <> 89 And KeyAscii <> 8 And KeyAscii <> 44 Then
    KeyAscii = 0
End If
End Sub

Private Sub textFA79_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub

Private Sub textFA80_GotFocus()
   CloseIme
   TextInverse textFA80
End Sub

Private Sub textFA80_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub

Private Sub textFA80_Validate(Cancel As Boolean)
   If textFA80.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textFA80.Text)
End Sub

Private Sub textFA81_GotFocus()
   CloseIme
   TextInverse textFA81
End Sub

Private Sub textFA81_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub

Private Sub textFA81_Validate(Cancel As Boolean)
   If textFA81.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textFA81.Text)
End Sub

Private Sub textFA82_GotFocus()
   CloseIme
   TextInverse textFA82
End Sub

Private Sub textFA82_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Added by Morgan 2011/11/30 Email輸入字元檢查
End Sub

Private Sub textFA82_Validate(Cancel As Boolean)
   If textFA82.Text = "" Or m_EditMode = "0" Then Exit Sub
   Cancel = Not PUB_CheckMail(textFA82.Text)
End Sub

Private Sub textFA85_GotFocus()
   InverseTextBox textFA85
End Sub

Private Sub textFA85_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textFA87_GotFocus()
   TextInverse textFA87
End Sub

Private Sub textFA87_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textFA88_GotFocus()
   TextInverse textFA88
End Sub

Private Sub textFA88_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textFA89_GotFocus()
   TextInverse textFA89
End Sub

Private Sub textFA89_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textFA90_GotFocus()
   TextInverse textFA90
End Sub


Private Sub textFA90_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textFA92_GotFocus()
   InverseTextBox textFA92
   OpenIme
End Sub

Private Sub TextFA93_GotFocus()
   TextInverse TextFA93
End Sub

Private Sub TextFA93_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub TextFA93_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   Cancel = False
   If IsEmptyText(TextFA93) = False Or TextFA93 = " " Then
      Select Case TextFA93
         Case "Y"
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "請輸入Y,不可輸入空白"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  TextFA93_GotFocus
            End Select
      End Select
   End If
End Sub
'2008/12/9 add by sonia
Private Sub textFA97_GotFocus()
   InverseTextBox textFA97
End Sub

Private Sub textFA97_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textFA97_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'If IsEmptyText(textFA97) = False Then  '2008/12/9 cancel by sonia 否則空白會存入
      Select Case textFA97
         Case "N", "":
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "是否寄電子報只可輸入N"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA97_GotFocus
            End Select
      End Select
   'End If
End Sub

'2008/12/9 end
' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
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

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM FAGENT " & _
            "WHERE FA01 = '" & strKEY01 & "' AND " & _
                  "FA02 = '" & strKEY02 & "' "
                  
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
   Dim strFA01 As String
   Dim strFA02 As String
   
   strFA01 = textFA01 & String(8 - Len(textFA01), "0")
   strFA02 = textFA02 & String(1 - Len(textFA02), "0")
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strFA01, strFA02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      'Modify by Amy 2024/03/19
      'GoTo EXITSUB
      Exit Sub
      'end 2024/03/19
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO FAGENT ("
   For nIndex = 0 To TF_FA - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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
   For nIndex = 0 To TF_FA - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            ' 90.12.18 modify by louis 字串中有單引號的處理
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
   
'Modify by Amy 2024/03/19 +BeginTrans
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   'add by nickc 2006/12/20
   Pub_SeekTbLog strSql
   
   cnnConnection.Execute strSql
   textFA03.Tag = ""  '2012/2/29 ADD BY SONIA
   MODCUSTOMER  '2005/12/21 ADD BY SONIA
   
   'Add by Amy 2022/11/25 +客戶代理人來源資料檔
   'If cboSource <> MsgText(601) And (txtXYS02 <> MsgText(601) Or txtXYS03 <> MsgText(601)) Then
      'Modify by 2024/11/29 拿掉[代理人來源啟用日]日期控制,並調整共用函數,避免有未改到
      'Modify by Amy 2023/05/08 改為共用 +Me.Name
      strMsg = SaveXYNoSource(1, Me.Name, textFA01, txtXYS02, txtXYS03, Left(cboSource, 2))
      If Len(strMsg) > 1 Then
         GoTo ErrHand
      End If
   'End If
   strMsg = ""
   'end 2024/11/29

   
   cnnConnection.CommitTrans
'end 20224/03/19
   
   If ((strFA01 & strFA02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strFA01 & strFA02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   
   ShowCurrRecord strFA01, strFA02
'Modify by Amy 2024/03/19
   Exit Sub

'EXITSUB:
ErrHand:
    cnnConnection.RollbackTrans
    'Modify by 2024/11/29 SaveXYNoSource有誤回傳其錯誤
    If strMsg = MsgText(601) Then
      strMsg = " 新增失敗！" & vbCrLf & Err.Description
    End If
    MsgBox strMsg
'end 2024/03/19
End Sub

' 修改記錄
Private Sub ModRecord()
   Dim strSql As String, strTmp As String, strTit As String, strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean, bFirst As Boolean
   Dim strFA01 As String, strFA02 As String
   Dim stMsg As String 'Add by Amy 2019/06/25
   'Modify by Amy 2022/12/28
   Dim intXYNoChoose As Integer, strUpdFCP605 As String
   Dim bolBeginTrans As Boolean, bolXYNoSourceData As Boolean '有下Trans/有更新XYNoSource 資料
   
   strFA01 = m_CurrKEY(0)
   strFA02 = m_CurrKEY(1)
   '910910  nick tigger
   '***** start
   'strSQL = "UPDATE FAGENT SET "
   strSql = "begin user_data.user_enabled:=1; UPDATE FAGENT SET "
   '***** end
   bFirst = True
   bDifference = False
   For nIndex = 0 To TF_FA - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
      strTmp = Empty
      '92.05.22 nick 跳過 create & update
      If nIndex < 45 Or nIndex > 50 Then
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     ' 90.12.18 modify by louis 字串中有單引號的處理
                     'strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
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
   '910910 nick tigger
   '***** start
   'strSQL = strSQL & " " & _
                  "WHERE FA01 = '" & strFA01 & "' AND " & _
                        "FA02 = '" & strFA02 & "' "
   strSql = strSql & " " & _
                  "WHERE FA01 = '" & strFA01 & "' AND " & _
                        "FA02 = '" & strFA02 & "'; end; "
    '***** end
'910910 nick tigger
'***** start
On Error GoTo ErrHand
'***** end
   If bDifference = True Then
      '910910 nick tigger
      '**** start
      cnnConnection.BeginTrans
      bolBeginTrans = True 'Add by Amy 2022/12/28
      '***** end
      'add by nickc 2006/12/20
      Pub_SeekTbLog strSql, , , True 'Modified by Morgan 2019/10/4 +第4參數
      
      cnnConnection.Execute strSql
      MODCUSTOMER  '2005/12/21 ADD BY SONIA
      '910910 nick tigger
      'Add by Amy 2017/03/10 由空白改MCTF時更新客戶檔及下一程序
      'Memo by Amy 2017/03/22 修改UpdMCTF_NP拿掉申請國家是台灣的判斷
      If textFA120.Tag = MsgText(601) And Left(textFA120, 4) = "MCTF" Then
        'Modify by Amy 2019/06/26 +stMsg
        stMsg = "Y" 'Add by Amy 2019/06/26 UpdMCTF_NP(textFA01 & textFA02, textFA120, stMsg)
        'Modify by Amy 2019/09/04 bug原:UpdMCTF_NP(textFA01 & textFA02, textFA120,, stMsg)
        If UpdMCTF_NP(textFA01 & textFA02, textFA120, , stMsg) = False Then GoTo ErrHand
        'end 2019/0626
      End If
      'Add by Amy 2019/06/25 修改控管制權人員欄位,更新當日AB類收文之 收文MCTF組別
      If textFA120.Tag <> textFA120 Then
        stMsg = "Y"
        If UpdCP161(textFA01 & textFA02, textFA120, stMsg) = False Then GoTo ErrHand
        If UpdMCTF_NP(textFA01 & textFA02, textFA120, , stMsg) = False Then GoTo ErrHand  'add by sonia 2020/3/23 Y22238020改MCTF管制人
      End If
      
      'Added by Lydia 2019/11/27 年費不續辦FA41=N => 目前案件的年費期限自動上不續辦
      strMsg = "": strTmp = ""  'Added by Lydia 2020/03/17
      If textFA41.Text = "N" And textFA41.Tag <> textFA41.Text Then
          'Modified by Lydia 2020/03/17 回傳FMP案範圍，發清單通知程序
          'Call Pub_AutoUpdFCP605(textFA01 & textFA02)
          'Modify by Amy 2022/12/28 避免strTmp變數被使用導致,ModRecordMail有問題,故改變數
          strUpdFCP605 = ""
          'If Pub_AutoUpdFCP605(textFA01 & textFA02, strTmp, strMsg) = False Then
          If Pub_AutoUpdFCP605(textFA01 & textFA02, strUpdFCP605, strMsg) = False Then
               GoTo ErrHand
          End If
          'end 2020/03/17
      End If
      'end 2019/11/27
      'Memo by Amy 2022/12/28 若只改txtXYS02 or txtXYS03 會沒更新到,故搬至外面
   
      'Add by Amy 2022/12/05 修改母號,更名前一併更新
      strSql = Left(cboSource, 2)
      If textFA02 = "0" And Left(cboSource, 2) <> m_FieldList(126).fiOldData Then
            strSql = "Update Fagent Set fa127=" & CNULL(strSql) & " Where fa01='" & textFA01 & "' And fa02<>'0' "
            cnnConnection.Execute strSql
      End If
      
      '***** start
      'Mark by Amy 2022/12/28 改後面做,原PUB_GetP605Email程式搬至 ModRecordMail
      'cnnConnection.CommitTrans
      '***** end
      'ShowCurrRecord strFA01, strFA02
      'end 2022/12/28
   End If
   'Add by Amy 2022/11/25 +客戶代理人來源資料檔
   'Modify by Amy 2022/12/28 拿掉代理人來源啟用日,若只修改  txtXYS02 或 txtXYS03 也要更新「客戶代理人來源資料檔」
   'If strSrvDate(1) >= 代理人來源啟用日 Then
        'Modify by 2024/11/29 依 來所原因/txtXYS02/txtXYS03 資料,如何改「客戶代理人來源資料檔」之判斷改於SaveXYNoSource,讓其他支程式也可用
        If textFA02 = "0" Then
            stMsg = SaveXYNoSource(2, Me.Name, textFA01, txtXYS02, txtXYS03, Left(cboSource, 2), m_FieldList(126).fiOldData)
            If Len(stMsg) > 1 Then
                GoTo ErrHand
            Else
                bolXYNoSourceData = True
            End If
        End If
        stMsg = ""
        'end 2024/11/29
        
        If bolBeginTrans = True Then cnnConnection.CommitTrans
        If bDifference = True Or bolXYNoSourceData = True Then
            If bDifference = True Then
                Call ModRecordMail(strUpdFCP605, strMsg)
            End If
            ShowCurrRecord strFA01, strFA02
        End If
    'End If
    'end 2022/12/28
   
'910910 nick tigger
'***** start
   Exit Sub
ErrHand:
    'Add by Amy 2019/06/25 +if
    If stMsg = MsgText(601) Then
        stMsg = Err.Description
    End If
    'MsgBox (Err.Description)
    'Modified by Lydia 2020/03/17
    'MsgBox stMsg
    MsgBox stMsg & vbCrLf & strMsg
    'end 2019/06/25
    'Modify by Amy 2022/12/28
    If bolBeginTrans = True Then cnnConnection.RollbackTrans
'******* end
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strFA01 As String, strFA02 As String
   Dim stMsg As String 'Add by 2024/11/29
   
   strFA01 = m_CurrKEY(0)
   strFA02 = m_CurrKEY(1)
   
   'Added by Lydia 2023/01/03
On Error GoTo ErrHandle
   cnnConnection.BeginTrans
   'end 2023/01/03
   
   strSql = "DELETE FROM FAGENT " & _
            "WHERE FA01 = '" & strFA01 & "' AND " & _
                  "FA02 = '" & strFA02 & "' "

   'add by nickc 2006/12/20
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If strFA02 = "0" Then 'Added by Lydia 2021/10/27 增加判斷更名後的號碼不可刪除;  ex.刪除代理人Y34013002一併刪除各項指示
       '2012/11/9 ADD BY SONIA 同時刪除接洽人
       strSql = "DELETE FROM POTCUSTCONT " & _
                "WHERE PCC01 = '" & strFA01 & "' "
       Pub_SeekTbLog strSql
       cnnConnection.Execute strSql
       '2012/11/9 END
    
       'Added by Lydia 2016/10/28 一併刪除申請人指定國外代理人檔
       strSql = "delete from CustAssignAgent where caa04=" & CNULL(strFA01)
       Pub_SeekTbLog strSql
       cnnConnection.Execute strSql
       'end 2016/10/28
       
       'Added by Lydia 2016/11/22 一併刪除國外固定寄催款單代理人檔
       strSql = "delete from Acc225 where a2251=" & CNULL(strFA01)
       Pub_SeekTbLog strSql
       cnnConnection.Execute strSql
       'end 2016/11/22
       
       'Added by Lydia 2016/11/24 一併刪除各項指示
       strSql = "DELETE FROM INSTRUCTIONS WHERE ITS01=" & CNULL(Pub_GetITS01Type(strFA01)) & " AND ITS02=" & CNULL(strFA01)
       Pub_SeekTbLog strSql
       cnnConnection.Execute strSql
       'end 2016/11/24
        
       'Added by Lydia 2016/11/30 一併刪除國外部關聯企業資料
       strSql = "delete from frelation where fr01=" & CNULL(strFA01) & " or fr02=" & CNULL(strFA01)
       Pub_SeekTbLog strSql
       cnnConnection.Execute strSql
       'end 2016/11/30
       
       'Add by Amy 2022/11/25 一併刪除 客戶代理人來源資料檔 的 被介紹 資料
       'Modify by 2024/11/29 改抓共用函數,避免有未改到
       stMsg = SaveXYNoSource(3, Me.Name, strFA01)
      If Len(stMsg) > 1 Then
         GoTo ErrHandle
      End If
      stMsg = ""
      'end 2024/11/29
       
       'Added by Lydia 2023/01/03 刪除外專特殊設定備註; ex.Y53912直接刪除代理人編號
       '下一程序固定備註(NpMemo)
       If ChkExistSpec("NPMEMO", strFA01, 8) = True Then
          strSql = "Delete From NPMEMO WHERE NM04='" & Left(strFA01, 8) & "' AND NM05 IS NULL "
          Pub_SeekTbLog strSql
          cnnConnection.Execute strSql
       End If
       If ChkExistSpec("NPMEMO", strFA01, 6) = True Then
          strSql = "Delete From NPMEMO WHERE NM04='" & Left(strFA01, 6) & "' AND NM05 IS NULL "
          Pub_SeekTbLog strSql
          cnnConnection.Execute strSql
       End If

      '核准函輸入備註(ApprovalMemo2)
      If ChkExistSpec("APPROVALMEMO2", strFA01, 8) = True Then
          strSql = "Delete From ApprovalMemo2 WHERE AM04='" & Left(strFA01, 8) & "' AND AM05 IS NULL "
          Pub_SeekTbLog strSql
          cnnConnection.Execute strSql
       End If
       If ChkExistSpec("APPROVALMEMO2", strFA01, 6) = True Then
          strSql = "Delete From ApprovalMemo2 WHERE AM04='" & Left(strFA01, 6) & "' AND AM05 IS NULL "
          Pub_SeekTbLog strSql
          cnnConnection.Execute strSql
       End If
       '核駁及審查意見通知函備註(IncomMemo)
       If ChkExistSpec("INCOMMEMO", strFA01, 8) = True Then
           strSql = "Delete From IncomMemo WHERE IM04='" & Left(strFA01, 8) & "' AND IM05 IS NULL "
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
       End If
       If ChkExistSpec("INCOMMEMO", strFA01, 6) = True Then
          strSql = "Delete From IncomMemo WHERE IM04='" & Left(strFA01, 6) & "' AND IM05 IS NULL "
          Pub_SeekTbLog strSql
          cnnConnection.Execute strSql
       End If
       '請款函預設備註維護檔(DebitNotePS)
       If ChkExistSpec("DEBITNOTEPS", strFA01, 8) = True Then
           strSql = "Delete From DEBITNOTEPS WHERE DNPS04='" & Left(strFA01, 8) & "' AND DNPS05 IS NULL "
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
       End If
       If ChkExistSpec("DEBITNOTEPS", strFA01, 6) = True Then
           strSql = "Delete From DEBITNOTEPS WHERE DNPS04='" & Left(strFA01, 6) & "' AND DNPS05 IS NULL "
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
       End If
       'FCP承辦單設定維護(FcpEMPbill)
       If ChkExistSpec("FCPEMPBILL", strFA01, 8) = True Then
           strSql = "Delete From FcpEMPbill WHERE FEB04='" & Left(strFA01, 8) & "' AND FEB05 IS NULL "
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
       End If
       If ChkExistSpec("FCPEMPBILL", strFA01, 6) = True Then
           strSql = "Delete From FcpEMPbill WHERE FEB04='" & Left(strFA01, 6) & "' AND FEB05 IS NULL "
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
       End If
       '通知告准加註(ApprovalPS)
       If ChkExistSpec("APPROVALPS", strFA01, 8) = True Then
           strSql = "Delete From APPROVALPS WHERE APS04='" & Left(strFA01, 8) & "' AND APS05 IS NULL "
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
       End If
       If ChkExistSpec("APPROVALPS", strFA01, 6) = True Then
           strSql = "Delete From APPROVALPS WHERE APS04='" & Left(strFA01, 6) & "' AND APS05 IS NULL "
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
       End If
       'end 2023/01/03
   End If 'Added by Lydia 2021/10/27
   
    'Added by Lydia 2022/03/28 一併刪除DHL輸入資料
    strSql = "delete from dhl_input_data where did01=" & CNULL(strFA01) & " and did02=" & CNULL(strFA02)
    cnnConnection.Execute strSql
    'end 2022/03/28
    
   '93.10.7 ADD BY SONIA
   If textFA03 <> "" Then
      strSql = "UPDATE CUSTOMER SET CU03=NULL WHERE CU01='" & textFA03 & "'"
      'add by nickc 2006/12/20
      Pub_SeekTbLog strSql
      
      cnnConnection.Execute strSql
   End If
   
   cnnConnection.CommitTrans 'Added by Lydia 2023/01/03
   
   '93.10.7 END
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strFA01 = m_LastKEY(0) And strFA02 = m_LastKEY(1)) Or (strFA01 = m_FirstKEY(0) And strFA02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   
   ShowCurrRecord strFA01, strFA02
   
EXITSUB:

   'Added by Lydia 2023/01/03
   Exit Sub
   
ErrHandle:
   cnnConnection.RollbackTrans
   'Modify by 2024/11/29 SaveXYNoSource有誤回傳其錯誤
    If stMsg = MsgText(601) Then
      stMsg = "刪除失敗！" & vbCrLf & Err.Description
    End If
   MsgBox stMsg
   'end 2024/11/29
End Sub


' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strFA01 As String
   Dim strFA02 As String
   
   QueryRecord = False
   strFA01 = textFA01 & String(8 - Len(textFA01), "0")
   strFA02 = textFA02 & String(1 - Len(textFA02), "0")
   'add by nickc 2006/03/17
   textCUID = ""
   If IsRecordExist(strFA01, strFA02) = True Then
      m_CurrKEY(0) = strFA01
      m_CurrKEY(1) = strFA02
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If
   
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Dim bolChk2nd As Boolean 'Added by Lydia 2019/05/27 是否修改FCP是否核對已准專利=N
   
   Select Case m_EditMode
      Case 1: '新增
'         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            AddRecord
            RefreshRange
'         Else
'            GoTo EXITSUB
'         End If
      Case 2: '修改
'         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'Added by Lydia 2019/05/27
            bolChk2nd = False
            If textFA85.Tag <> textFA85.Text And textFA85.Text = "N" Then
                bolChk2nd = True
            End If
            
            ModRecord
            
            'Added by Lydia 2019/05/27 設定"FCP是否核對已准專利"上" N"，則出"核對已准專利"未發文之清單
            If bolChk2nd = True Then
                If Pub_GetFA85CU122List(textFA01 & textFA02) = True Then
                End If
            End If
            'end 2019/05/27
'         Else
'            GoTo EXITSUB
'         End If
      Case 3: '刪除
         DelRecord
         RefreshRange
      Case 4: '查詢
'         If CheckDataValid() = True Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
'         Else
'            GoTo EXITSUB
'         End If
   End Select
  
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textFA01.SetFocus
      Case 2:
         'Modify by Amy 2024/01/22 +if 國外潛在客戶維護轉號存檔切至此畫面欄位鎖住會錯
         If textFA03.Enabled = False Then
            textFA03.SetFocus
         End If
      Case 4: textFA01.SetFocus
   End Select
End Sub

Private Sub textFA01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 代理人編號
Private Sub textFA01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   '若有輸入代理人編號
   If IsEmptyText(textFA01) = False Then
      Select Case m_EditMode
         Case 1, 4:
            If Len(textFA01) < 6 Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "代理人編號請至少輸入六碼"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA01_GotFocus
               GoTo EXITSUB
            End If
      End Select
      ' 補滿八碼
      textFA01 = textFA01 & String(8 - Len(textFA01), "0")
      Select Case m_EditMode
         Case 1, 4:
            If Mid(textFA01, 1, 1) <> "Y" Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "代理人編號必須為Y開頭"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA01_GotFocus
            End If
      End Select
        'Add By Cheng 2003/03/27
        '在新增時輸入的代理人編號, 不可大於自動編號檔的流水號
        Select Case m_EditMode
        Case 1 '新增
            If GetAutoNumY(Mid(Me.textFA01.Text, 2, 5)) <> "" Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "代理人編號不可大於" & GetAutoNumY(Mid(Me.textFA01.Text, 2, 5))
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA01_GotFocus
            End If
        End Select
   End If
EXITSUB:
End Sub

' 代理人編號
Private Sub textFA02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA02 = textFA02 & String(1 - Len(textFA02), "0")
   Select Case m_EditMode
      Case 1:
         If IsRecordExist(textFA01, textFA02) = True Then
            'edit by nickc 2008/01/11
            'Cancel = True
            strTit = "檢核資料"
            strMsg = "該筆代理人已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            'edit by nickc 2008/01/11
            'textFA01_GotFocus
            textFA01.SetFocus
            Exit Sub
         End If
      Case Else:
   End Select
End Sub

Private Sub textFA03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 客戶編號
Private Sub textFA03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTemp As String
   Cancel = False
   If IsEmptyText(textFA03) = False Then
      strTemp = GetCustomerName(textFA03, 0)
      If IsEmptyText(strTemp) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "客戶編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA03_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

' 代理人名稱(中)
Private Sub textFA04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA04) = False Then
      'Modified by Lydia 2021/01/07 額外判斷字串個數
      'If GetTextLength(textFA04) > 80 Then
      If GetTextLength(textFA04) > 80 And Len(textFA04) > 80 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人名稱(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA04_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA04.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代理人名稱(日)
Private Sub textFA06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA06) = False Then
      'Modified by Lydia 2021/01/07 額外判斷字串個數
      'If GetTextLength(textFA06) > 80 Then
      If GetTextLength(textFA06) > 80 And Len(textFA06) > 80 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人名稱(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA06_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA06.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 聯絡人1(中)
Private Sub textFA07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA07) = False Then
      'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
      'If gettextlength(textFA07) > 10 Then
      If GetTextLength(textFA07) > 30 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聯絡人1(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA07_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA07.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 聯絡人1(日)
Private Sub textFA09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA09) = False Then
      If GetTextLength(textFA09) > textFA09.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聯絡人1(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA09_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA09.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代理人國籍
Private Sub textFA10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA10_2 = Empty
   If IsEmptyText(textFA10) = False Then
      textFA10_2 = GetNationName(textFA10, 0)
      'Add by Amy 2015/09/09 +不可輸入000
      If (m_EditMode = 1 Or m_EditMode = 2) And textFA10 = 台灣國家代號 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "代理人國籍不可輸000台灣"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFA10_GotFocus
            Exit Sub
      End If
      If IsEmptyText(textFA10_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "代理人國籍不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA10_GotFocus
            Case Else:
         End Select
      Else
         If IsEmptyText(textFA55) = True Then
            textFA55 = textFA10
            textFA55_2 = GetNationName(textFA10, 0)
         End If
      End If
        'Modify By Cheng 2004/02/03
        '若為新增或修改狀態, 預設下列欄位值
        If m_EditMode = "1" Or m_EditMode = "2" Then
            If IsEmptyText(textFA31) = True Then
               '93.3.17 MODIFY BY SONIA 預設定稿語文
               'If textFA10 < "010" Then
               If textFA10 < "010" Or textFA10 = "020" Then
               '93.3.17 END
                  textFA31 = "1"
               '2012/4/13 ADD BY SONIA
               ElseIf Left(textFA10.Text, 3) = "011" Then
                  textFA31 = "3"
               '2012/4/13 END
               Else
                  textFA31 = "2"
               End If
            End If
            'Add By Sindy 2011/3/10 預設是否寄發專利雙週報
            'Modified by Morgan 2011/12/30 專利雙週報欄位改放 N(不寄)
            'If IsEmptyText(textFA100) = True Then
            '   If textFA10 = "020" And TextFA76 <> "C" Then
            '      textFA100 = "Y"
            '   End If
            'End If
            Cancel = FA100CheckError()
            'end 2011/12/30
        End If
        'End
   End If
End Sub

' 開發日期
Private Sub textFA11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA11) = False Then
      If CheckIsTaiwanDate(textFA11, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "開發日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA11_GotFocus
      End If
   End If
End Sub

' 代理人地址(中)
Private Sub textFA17_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textFA17) = False Then
      'Modified by Lydia 2020/11/16 O12用Char儲存，所以額外判斷字串個數; ex.Y55054的中文地址字元長度超過70，但是字數未超出
      'If GetTextLength(textFA17) > 70 Then
      If GetTextLength(textFA17) > 70 And Len(textFA17) >= 70 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人地址(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA17_GotFocus
         Exit Sub
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA17.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代理人地址(日)
Private Sub textFA23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA23) = False Then
      If GetTextLength(textFA23) > 70 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人地址(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA17_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA23.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textFA24_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否寄台一雜誌
Private Sub textFA24_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'If IsEmptyText(textFA24) = False Then  '2008/12/9 cancel by sonia 否則空白會存入
      Select Case textFA24
         Case "N", "":
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "是否寄台一雜誌只可輸入N"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA24_GotFocus
            End Select
      End Select
   'End If
End Sub

' 全部折扣
Private Sub textFA25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA25) = False Then
      If IsNumeric(textFA25) = False Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "專利全部折扣只可輸入數值資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA25_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

' 申請/翻議折扣
Private Sub textFA26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA26) = False Then
      If IsNumeric(textFA26) = False Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "專利申請/翻議折扣只可輸入數值資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA26_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

' 全部折扣起始日
Private Sub textFA27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA27) = False Then
      If CheckIsTaiwanDate(textFA27, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "專利全部折扣起始日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA27_GotFocus
      End If
   End If
End Sub


' 代理人備註
Private Sub textFA29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA29) = False Then
      'Modfiy by Amy 2015/08/28 原:2000
      If GetTextLength(textFA29) > textFA29.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人備註內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA29_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA29.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 專利固定請款對象
Private Sub textFA30_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA30_2 = Empty
   If IsEmptyText(textFA30) = False Then
      If (textFA30 & String(9 - Len(textFA30), "0")) = (textFA01 & String(8 - Len(textFA01), "0") & textFA02 & String(1 - Len(textFA02), "0")) Or Mid(textFA30 & String(9 - Len(textFA30), "0"), 1, 8) = textFA03 & String(8 - Len(textFA03), "0") Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "專利固定請款對象不可為該筆資料的代理人"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA30_GotFocus
            Case Else:
         End Select
      End If
      
      Select Case Mid(textFA30, 1, 1)
         Case "X":
            textFA30_2 = GetCustomerName(textFA30, 0)
         Case "Y":
            textFA30_2 = GetFAgentName(textFA30)
         Case Else:
            textFA30_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "專利固定請款對項代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA30_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA30_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "專利固定請款對項代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA30_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

'Add By Sindy 2011/3/4
' 商標固定請款對象
Private Sub textFA107_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA107_2 = Empty
   If IsEmptyText(textFA107) = False Then
      If (textFA107 & String(9 - Len(textFA107), "0")) = (textFA01 & String(8 - Len(textFA01), "0") & textFA02 & String(1 - Len(textFA02), "0")) Or Mid(textFA107 & String(9 - Len(textFA107), "0"), 1, 8) = textFA03 & String(8 - Len(textFA03), "0") Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商標固定請款對象不可為該筆資料的代理人"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA107_GotFocus
            Case Else:
         End Select
      End If
      
      Select Case Mid(textFA107, 1, 1)
         Case "X":
            textFA107_2 = GetCustomerName(textFA107, 0)
         Case "Y":
            textFA107_2 = GetFAgentName(textFA107)
         Case Else:
            textFA107_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "商標固定請款對項代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA107_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA107_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商標固定請款對項代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA107_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

' 定稿語文
Private Sub textFA31_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA31) = False Then
      Select Case textFA31
         Case "1", "2", "3":
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "定稿語文只可輸入1,2或3"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA31_GotFocus
            End Select
      End Select
   End If
End Sub

' 副本收受人
Private Sub textFA38_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA38_2 = Empty
   If IsEmptyText(textFA38) = False Then
      If (textFA38 & String(9 - Len(textFA38), "0")) = (textFA01 & String(8 - Len(textFA01), "0") & textFA02 & String(1 - Len(textFA02), "0")) Or Mid(textFA38 & String(9 - Len(textFA38), "0"), 1, 8) = textFA03 & String(8 - Len(textFA03), "0") Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "副本收受人不可為該筆資料的代理人"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA38_GotFocus
            Case Else:
         End Select
      End If
      Select Case Mid(textFA38, 1, 1)
         Case "X":
            textFA38_2 = GetCustomerName(textFA38, 0)
         Case "Y":
            textFA38_2 = GetFAgentName(textFA38)
         Case Else:
            textFA38_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "副本收受人代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA38_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA38_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "副本收受人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA38_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

' 收款後辦案
Private Sub textFA39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'If IsEmptyText(textFA39) = False Then  '2008/12/9 cancel by sonia 否則空白會存入
      Select Case textFA39
         Case "Y", "":
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "收款後辦案只可輸入Y"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA39_GotFocus
            End Select
      End Select
   'End If
End Sub

' FCP年費通知函單筆不跑
Private Sub textFA40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'If IsEmptyText(textFA40) = False Then  '2008/12/9 cancel by sonia 否則空白會存入
      Select Case textFA40
         Case "Y", "":
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "FCP年費通知函單筆不跑只可輸入Y"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA40_GotFocus
            End Select
      End Select
   'End If
End Sub

' FCP年費自動代繳
Private Sub textFA41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'If IsEmptyText(textFA41) = False Then  '2008/12/9 cancel by sonia 否則空白會存入
      Select Case textFA41
         'Modified by Lydia 2016/08/15 +N
         Case "Y", "N", "":
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  'Modified by Lydia 2016/08/15 +N
                  'strMsg = "FCP年費自動代繳只可輸入Y"
                  strMsg = "只可輸入FCP年費自動代繳(Y)或寄證書後年費不續辦(N)"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA41_GotFocus
            End Select
      End Select
   'End If
End Sub

' FCP領證自動代繳
Private Sub textFA42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'If IsEmptyText(textFA42) = False Then  '2008/12/9 cancel by sonia 否則空白會存入
      Select Case textFA42
         Case "Y", "":
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "FCP領證自動代繳只可輸入Y"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA42_GotFocus
            End Select
      End Select
   'End If
End Sub

'' 專利D/N幣別
'Private Sub textFA43_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textFA43) = False Then
'      Select Case textFA43
'         Case "U", "N", "R":
'         Case Else:
'            Select Case m_EditMode
'               Case 1, 2:
'                  Cancel = True
'                  strTit = "檢核資料"
'                  strMsg = "專利D/N幣別只可輸入U、N、R或空白"
'                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'                  textFA43_GotFocus
'            End Select
'      End Select
'   End If
'End Sub

''Add By Sindy 2011/3/4
'' 商標D/N幣別
'Private Sub textFA108_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textFA108) = False Then
'      Select Case textFA108
'         Case "U", "N", "R":
'         Case Else:
'            Select Case m_EditMode
'               Case 1, 2:
'                  Cancel = True
'                  strTit = "檢核資料"
'                  strMsg = "商標D/N幣別只可輸入U、N、R或空白"
'                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'                  textFA108_GotFocus
'            End Select
'      End Select
'   End If
'End Sub

' 專利D/N是否列印申請人
Private Sub textFA44_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'If IsEmptyText(textFA44) = False Then  '2008/12/9 cancel by sonia 否則空白會存入
      Select Case textFA44
         Case "Y", "":
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "專利D/N是否列印申請人只可輸入Y"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA44_GotFocus
            End Select
      End Select
   'End If
End Sub

'Add By Sindy 2011/3/4
' 商標D/N是否列印申請人
Private Sub textFA109_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   Select Case textFA109
      Case "Y", "":
      Case Else:
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商標D/N是否列印申請人只可輸入Y"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA109_GotFocus
         End Select
   End Select
End Sub
'2011/3/4 End

' 聯絡人2(中)
Private Sub textFA52_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA52) = False Then
      'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
      'If gettextlength(textFA52) > 10 Then
      If GetTextLength(textFA52) > 30 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聯絡人2(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA52_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA52.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 聯絡人2(日)
Private Sub textFA54_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA54) = False Then
      If GetTextLength(textFA54) > textFA54.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聯絡人2(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA54_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA54.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 實體聯絡人中文名稱
Private Sub textFA56_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA56) = False Then
      If GetTextLength(textFA56) > 10 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "實體聯絡人中文名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA56_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA56.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 實體聯絡人日文名稱
Private Sub textFA58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA58) = False Then
      If GetTextLength(textFA58) > 20 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "實體聯絡人日文名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA58_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA58.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 地址國籍
Private Sub textFA55_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA55_2 = Empty
   If IsEmptyText(textFA55) = False Then
      textFA55_2 = GetNationName(textFA55, 0)
      'Add by Amy 2015/09/09 +不可輸入000
      If (m_EditMode = 1 Or m_EditMode = 2) And textFA55 = 台灣國家代號 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "地址國籍不可輸000台灣"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFA55_GotFocus
            Exit Sub
      End If
      If IsEmptyText(textFA55_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "地址國籍國籍不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA55_GotFocus
         End Select
      End If
   End If
End Sub

' 實體副本收受人
Private Sub textFA59_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA59_2 = Empty
   If IsEmptyText(textFA59) = False Then
      If (textFA59 & String(9 - Len(textFA59), "0")) = (textFA01 & String(8 - Len(textFA01), "0") & textFA02 & String(1 - Len(textFA02), "0")) Or Mid(textFA59 & String(9 - Len(textFA59), "0"), 1, 8) = textFA03 & String(8 - Len(textFA03), "0") Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "實體副本收受人不可為該筆資料的代理人"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA59_GotFocus
               GoTo EXITSUB
            Case Else:
         End Select
      End If
      Select Case Mid(textFA59, 1, 1)
         Case "X":
            textFA59_2 = GetCustomerName(textFA59, 0)
         Case "Y":
            textFA59_2 = GetFAgentName(textFA59)
         Case Else:
            textFA59_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "實體副本收受人代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA59_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA59_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "實體副本收受人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA59_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

' 年費代理人
Private Sub textFA61_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA61_2 = Empty
   If IsEmptyText(textFA61) = False Then
      If (textFA61 & String(9 - Len(textFA61), "0")) = (textFA01 & String(8 - Len(textFA01), "0") & textFA02 & String(1 - Len(textFA02), "0")) Or Mid(textFA61 & String(9 - Len(textFA61), "0"), 1, 8) = textFA03 & String(8 - Len(textFA03), "0") Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "年費代理人不可為該筆資料的代理人"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA61_GotFocus
               GoTo EXITSUB
            Case Else:
         End Select
      End If
      Select Case Mid(textFA61, 1, 1)
         Case "X":
            textFA61_2 = GetCustomerName(textFA61, 0)
         Case "Y":
            textFA61_2 = GetFAgentName(textFA61)
         Case Else:
            textFA61_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "年費代理人代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA61_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA61_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "年費代理人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA61_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

' 年費請款對象
Private Sub textFA62_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA62_2 = Empty
   If IsEmptyText(textFA62) = False Then
      If (textFA62 & String(9 - Len(textFA62), "0")) = (textFA01 & String(8 - Len(textFA01), "0") & textFA02 & String(1 - Len(textFA02), "0")) Or Mid(textFA62 & String(9 - Len(textFA62), "0"), 1, 8) = textFA03 & String(8 - Len(textFA03), "0") Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "年費請款對象不可為該筆資料的代理人"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA62_GotFocus
               GoTo EXITSUB
            Case Else:
         End Select
      End If
      Select Case Mid(textFA62, 1, 1)
         Case "X":
            textFA62_2 = GetCustomerName(textFA62, 0)
         Case "Y":
            textFA62_2 = GetFAgentName(textFA62)
         Case Else:
            textFA62_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "年費請款對象代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA62_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA62_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "年費請款對象代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA62_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

' 延展通知人
Private Sub textFA66_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA66_2 = Empty
   If IsEmptyText(textFA66) = False Then
      If (textFA66 & String(9 - Len(textFA66), "0")) = (textFA01 & String(8 - Len(textFA01), "0") & textFA02 & String(1 - Len(textFA02), "0")) Or Mid(textFA66 & String(9 - Len(textFA66), "0"), 1, 8) = textFA03 & String(8 - Len(textFA03), "0") Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "延展通知人不可為該筆資料的代理人"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA66_GotFocus
               GoTo EXITSUB
            Case Else:
         End Select
      End If
      Select Case Mid(textFA66, 1, 1)
         Case "X":
            textFA66_2 = GetCustomerName(textFA66, 0)
         Case "Y":
            textFA66_2 = GetFAgentName(textFA66)
         Case Else:
            textFA66_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "延展通知人代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA66_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA66_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "延展通知人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA66_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

' 延展請款對象
Private Sub textFA67_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textFA67_2 = Empty
   If IsEmptyText(textFA67) = False Then
      If (textFA67 & String(9 - Len(textFA67), "0")) = (textFA01 & String(8 - Len(textFA01), "0") & textFA02 & String(1 - Len(textFA02), "0")) Or Mid(textFA67 & String(9 - Len(textFA67), "0"), 1, 8) = textFA03 & String(8 - Len(textFA03), "0") Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "延展請款對象不可為該筆資料的代理人"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA67_GotFocus
               GoTo EXITSUB
            Case Else:
         End Select
      End If
      Select Case Mid(textFA67, 1, 1)
         Case "X":
            textFA67_2 = GetCustomerName(textFA67, 0)
         Case "Y":
            textFA67_2 = GetFAgentName(textFA67)
         Case Else:
            textFA67_2 = Empty
            Select Case m_EditMode
               Case 1, 2
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "延展請款對象代號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA67_GotFocus
                  GoTo EXITSUB
               Case Else:
            End Select
      End Select
      If IsEmptyText(textFA67_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "延展請款對象代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA67_GotFocus
            Case Else:
         End Select
      End If
   End If
EXITSUB:
End Sub

' 延展單筆不跑
Private Sub textFA68_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'If IsEmptyText(textFA68) = False Then  '2008/12/9 cancel by sonia 否則空白會存入
      Select Case textFA68
         Case "Y", "":
         Case Else:
            Select Case m_EditMode
               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "延展單筆不跑只可輸入或Y"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textFA68_GotFocus
            End Select
      End Select
   'End If
End Sub

'Mark by Amy 2015/08/24 改為下拉選單
'Private Sub textFA69_GotFocus()
'   InverseTextBox textFA69
'   'edit by nickc 2007/06/06 切換輸入法改用API
'   'textFA69.IMEMode = 1
'   OpenIme
'End Sub

'' 代理人狀態
'Private Sub textFA69_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textFA69) = False Then
'      If gettextlength(textFA69) > textFA69.MaxLength Then
'         Cancel = True
'         strTit = "檢核資料"
'         strMsg = "代理人狀態內容太長"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textFA69_GotFocus
'      End If
'   End If
'   'edit by nickc 2007/06/06 切換輸入法改用API
'   'If Cancel = False Then textFA69.IMEMode = 2
'   If Cancel = False Then CloseIme
'End Sub

' 聯絡人部門(日)
Private Sub textFA78_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA78) = False Then
      If GetTextLength(textFA78) > textFA78.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聯絡人部門(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA78_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textFA78.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim strTmp As String
   Dim nResponse
   Dim strTemp As String 'Add by Amy 2016/06/30
   Dim iRtn As Integer 'Add by Amy 2021/11/26
   
   CheckDataValid = False

   Select Case m_EditMode
      Case 4:
         ' 代理人編號不可空白
         If IsEmptyText(textFA01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入代理人編號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFA01.SetFocus
            GoTo EXITSUB
         End If
         If IsEmptyText(textFA02) = True Then
            textFA02 = "0"
         End If
        ' 代理人編號尾碼不可空白
         'If IsEmptyText(textFA02) = True Then
         '   strTit = "檢核資料"
         '   strMsg = "請輸入代理人編號尾碼"
         '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '   textFA02.SetFocus
         '   GoTo ExitSub
         'End If
      Case Else:
   End Select
      
   Select Case m_EditMode
      Case 1, 2:
         'add by nickc 2008/03/12
         If textFA03 <> "" Then
            If Mid(textFA03, 1, 1) <> "X" Then
               strTit = "檢核資料"
               strMsg = "客戶編號應為 X 開頭"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textFA03.SetFocus
               GoTo EXITSUB
            End If
         End If
         ' 中文名稱, 英文名稱, 日文名稱不可全為空白
         If IsEmptyText(textFA04) = True And IsEmptyText(textFA05) = True And IsEmptyText(textFA06) = True Then
            strTit = "檢核資料"
            strMsg = "中文名稱, 英文名稱, 日文名稱不可全為空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFA04.SetFocus
            GoTo EXITSUB
         End If
         ' 代理人國籍
         If IsEmptyText(textFA10) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入代理人國籍"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFA10.SetFocus
            GoTo EXITSUB
         End If
         ' 開發日期
         If IsEmptyText(textFA11) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入開發日期"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFA11.SetFocus
            GoTo EXITSUB
         End If
         'Add by Amy 2016/06/30 +地址國籍為臺灣判斷地址格式是否正確
         If IsEmptyText(textFA17) = False And textFA55 < "010" Then
            If CheckAddrData(textFA17, strTemp) = False Then
                strTit = "檢核資料"
                strMsg = "代理人地址(中)" & strTemp
                If InStr(strTemp, "格式") > 0 Then
                    If bolShow100135 = False Then
                        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                        bolShow100135 = True
                        frm100135.Show vbModal
                        bolShow100135 = False
                        textFA17_GotFocus
                        textFA17.SetFocus
                        GoTo EXITSUB
                    Else
                        GoTo EXITSUB
                    End If
                Else
                    nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                    textFA17_GotFocus
                    textFA17.SetFocus
                    GoTo EXITSUB
                End If
            Else
                strMsg = "代理人地址(中)"
                If CheckTaiwanAddr(textFA17, "000", strMsg) = False Then
                    textFA17_GotFocus
                    textFA17.SetFocus
                    GoTo EXITSUB
                End If
            End If
         End If
         'end 2016/06/30
         'Add by Amy 2021/11/26 國籍為台灣且中文名稱有「事務所」字樣且狀態 非「國內同業」,彈 國內同業控制
         If textFA10 < "010" And InStr(textFA04, "事務所") > 0 And cboStatus <> "國內同業" Then
            iRtn = MsgBox("國籍在台灣之事務所，請確認是否為國內同業？" & vbCrLf & _
                                        "是:為國內同業　否:非國內同業", vbYesNoCancel + vbDefaultButton3)
             '取消
            If iRtn = 2 Then
               Exit Function
            '是
            ElseIf iRtn = 6 Then
                cboStatus = "國內同業"
            End If 'iRtn
         End If
         
         If cboStatus = "國內同業" Then
            '非財務Email不可輸
            'Modify by Amy 2022/09/23 訊息一次顯示 原:ShowMsg
            strMsg = ""
            If textFA16 <> MsgText(601) Then
                strMsg = strMsg & "此為國內同業,不可輸入E-Mail(代表)以免誤發電子郵件, 如有需要請加註於備註欄 ！" & vbCrLf
'                textFA16.SetFocus
'                textFA16_GotFocus
'                SSTab1.Tab = 1
'                Exit Function
            End If
            If textFA80 <> MsgText(601) Then
                strMsg = strMsg & "此為國內同業,不可輸入E-Mail(其他1)以免誤發電子郵件, 如有需要請加註於備註欄 ！" & vbCrLf
'                textFA80.SetFocus
'                textFA80_GotFocus
'                SSTab1.Tab = 1
'                Exit Function
            End If
            If textFA81 <> MsgText(601) Then
                strMsg = strMsg & "此為國內同業,不可輸入E-Mail(其他2)以免誤發電子郵件, 如有需要請加註於備註欄 ！" & vbCrLf
'                textFA81.SetFocus
'                textFA81_GotFocus
'                SSTab1.Tab = 1
'                Exit Function
            End If
            If textFA82 <> MsgText(601) Then
                strMsg = strMsg & "此為國內同業,不可輸入E-Mail(其他3)以免誤發電子郵件, 如有需要請加註於備註欄 ！" & vbCrLf
'                textFA82.SetFocus
'                textFA82_GotFocus
'                SSTab1.Tab = 1
'                Exit Function
            End If
            '電子報要設定不寄
            If textFA100 <> "N" Then
                strMsg = strMsg & "此為國內同業, 不可寄專利雙週報 ！" & vbCrLf
'                textFA100.SetFocus
'                textFA100_GotFocus
'                SSTab1.Tab = 0
'                Exit Function
            End If
            If textFA24 <> "N" Then
                strMsg = strMsg & "此為國內同業, 不可寄台一雜誌 ！" & vbCrLf
'                textFA24.SetFocus
'                textFA24_GotFocus
'                SSTab1.Tab = 4
'                Exit Function
            End If
            If textFA97 <> "N" Then
                strMsg = strMsg & "此為國內同業, 不可寄電子報 ！" & vbCrLf
'                textFA97.SetFocus
'                textFA97_GotFocus
'                SSTab1.Tab = 4
'                Exit Function
            End If
            If txtFA(121) <> "N" Then
                strMsg = strMsg & "此為國內同業, 不可寄竹曆！" & vbCrLf
'                txtFA(121).SetFocus
'                txtFA_GotFocus (121)
'                SSTab1.Tab = 4
'                Exit Function
            End If
            If txtFA(122) <> "N" Then
                strMsg = strMsg & "此為國內同業, 不可寄促銷信！" & vbCrLf
'                txtFA(122).SetFocus
'                txtFA_GotFocus (122)
'                SSTab1.Tab = 4
'                Exit Function
            End If
            If strMsg <> MsgText(601) Then
                MsgBox strMsg, vbCritical + vbOKOnly, MsgText(9001)
                Exit Function
            End If
            'end 2022/09/23
         End If
         'end 2021/11/26
         
         'add by nickc 2006/03/10 檢查用客戶和代理人是否與客戶檔相同
         If textFA03 <> "" Then
            'add by nickc 2008/03/12 控制性質
            If TextFA76 <> "B" Then
               ShowMsg "代理人性質輸入錯誤 !"
               TextFA76.SetFocus
               GoTo EXITSUB
            End If
         Else
            '2008/7/17 modify by sonia 改為詢問方式
            If TextFA76 = "B" Then
               strTmp = "代理人性質為 B ! 是否要輸入相對應之客戶編號 ?"
               If MsgBox(strTmp, vbYesNo + vbCritical) = vbYes Then
                  textFA03.SetFocus
                  GoTo EXITSUB
               End If
            End If
            '2008/7/17 END
         End If
         'Add By Sindy 2013/1/17
         If Trim(Me.Combo2(0).Text) <> "" Then
            '若輸入幣別就一定要選格式
            If Trim(Me.Combo3(0).Text) = "" Then
               ShowMsg "專利請款單列印幣別格式不可空白 !"
               Me.Combo3(0).SetFocus
               GoTo EXITSUB
            End If
            '請款幣別<>NTD時不可輸入1
            If Trim(Me.Combo2(0).Text) <> "NTD" And Me.Combo3(0).ListIndex = 1 Then
               ShowMsg "專利請款幣別<>NTD時，專利請款單列印幣別格式不可選純台幣 !"
               Me.Combo3(0).SetFocus
               GoTo EXITSUB
            End If
         End If
         If Trim(Me.Combo2(1).Text) <> "" Then
            '若輸入幣別就一定要選格式
            If Trim(Me.Combo3(1).Text) = "" Then
               ShowMsg "商標請款單列印幣別格式不可空白 !"
               Me.Combo3(1).SetFocus
               GoTo EXITSUB
            End If
            '請款幣別<>NTD時不可輸入1
            If Trim(Me.Combo2(1).Text) <> "NTD" And Me.Combo3(1).ListIndex = 1 Then
               ShowMsg "商標請款幣別<>NTD時，商標請款單列印幣別格式不可選純台幣 !"
               Me.Combo3(1).SetFocus
               GoTo EXITSUB
            End If
         End If
         '2013/1/17 End
      Case Else:
   End Select
   
'edit by nickc 2008/05/08 改成共用  function
'   'add by nickc 2006/06/15 加入檢查英文名稱第一碼
'   If Mid(textFA10, 1, 3) = "101" Then
'       '2008/1/4 MODIFY BY SONIA 原為A~I為101,J~Z為1011,2008年改為分四段
'       If Mid(UCase(LTrim(textFA05)), 1, 1) >= "A" And Mid(UCase(LTrim(textFA05)), 1, 1) <= "E" Then
'            If Trim(textFA10) <> "101" Then
'                ShowMsg "代理人英文名稱第一碼介於 A~E 之間，代理人國籍應該為 101 !"
'                GoTo EXITSUB
'            End If
'       ElseIf Mid(UCase(LTrim(textFA05)), 1, 1) >= "F" And Mid(UCase(LTrim(textFA05)), 1, 1) <= "I" Then
'            If Trim(textFA10) <> "1011" Then
'                ShowMsg "代理人英文名稱第一碼介於 F~I 之間，代理人國籍應該為 1011 !"
'                GoTo EXITSUB
'            End If
'       ElseIf Mid(UCase(LTrim(textFA05)), 1, 1) >= "J" And Mid(UCase(LTrim(textFA05)), 1, 1) <= "N" Then
'            If Trim(textFA10) <> "1012" Then
'                ShowMsg "代理人英文名稱第一碼介於 J~N 之間，代理人國籍應該為 1012 !"
'                GoTo EXITSUB
'            End If
'       ElseIf Mid(UCase(LTrim(textFA05)), 1, 1) >= "O" And Mid(UCase(LTrim(textFA05)), 1, 1) <= "Z" Then
'            If Trim(textFA10) <> "1013" Then
'                ShowMsg "代理人英文名稱第一碼介於 O~Z 之間，代理人國籍應該為 1013 !"
'                GoTo EXITSUB
'            End If
'       '2008/1/9 add by sonia
'       Else
'            If Trim(textFA10) <> "1013" Then
'                ShowMsg "代理人英文名稱第一碼非英文字母或無英文名稱，代理人國籍應該為 1013 !"
'                GoTo EXITSUB
'            End If
'       '2008/1/9 end
'       End If
'   ElseIf Mid(textFA10, 1, 3) = "011" Then
'       '2008/4/21 MODIFY BY SONIA 原為A~L為011,M~Z為0111,2008/4/22改為分三段(將M~Z再細分成二段)
'       If Mid(UCase(LTrim(textFA05)), 1, 1) >= "A" And Mid(UCase(LTrim(textFA05)), 1, 1) <= "L" Then
'            If Trim(textFA10) <> "011" Then
'                ShowMsg "代理人英文名稱第一碼介於 A~L 之間，代理人國籍應該為 011 !"
'                GoTo EXITSUB
'            End If
'       ElseIf Mid(UCase(LTrim(textFA05)), 1, 1) >= "M" And Mid(UCase(LTrim(textFA05)), 1, 1) <= "O" Then
'            If Trim(textFA10) <> "0111" Then
'                ShowMsg "代理人英文名稱第一碼介於 M~O 之間，代理人國籍應該為 0111 !"
'                GoTo EXITSUB
'            End If
'       ElseIf Mid(UCase(LTrim(textFA05)), 1, 1) >= "P" And Mid(UCase(LTrim(textFA05)), 1, 1) <= "Z" Then
'            If Trim(textFA10) <> "0112" Then
'                ShowMsg "代理人英文名稱第一碼介於 P~Z 之間，代理人國籍應該為 0112 !"
'                GoTo EXITSUB
'            End If
'       '2008/1/9 modify by sonia
'       'ElseIf Trim(textFA05) = "" Then
'       Else
'            If Trim(textFA10) <> "0112" Then
'                ShowMsg "代理人英文名稱第一碼非英文字母或無英文名稱，代理人國籍應該為 0112 !"
'                GoTo EXITSUB
'            End If
'       End If
'   End If
    If Trim(textFA10) <> pub_NationByName(textFA05 & textFA63 & textFA64 & textFA65, Trim(textFA10), True, "代理人") Then
       'Added by Lydia 2016/08/10
        If Me.ActiveControl = textFA10 Then
           textFA10_GotFocus
        Else
           textFA10.SetFocus
        End If
        SSTab1.Tab = 0
        'end 2016/08/10
        GoTo EXITSUB
    End If
    
    'Add by Amy2021/11/26 以國籍判斷定稿語文彈提醒-外商阿蓮
    If m_EditMode <> 4 Then
        If textFA31 = MsgText(601) Then
            ShowMsg "定稿語文不可為空白 !"
            textFA31.SetFocus
            textFA31_GotFocus
            SSTab1.Tab = 4
            GoTo EXITSUB
        Else
            'Modify by Amy 2021/12/09 原:textFA10
            If Left(textFA10, 3) = "011" And textFA31 <> "3" Then
                If MsgBox("國籍為「日本」，定稿語文確定「不是」日文？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                    textFA31.SetFocus
                    textFA31_GotFocus
                    SSTab1.Tab = 4
                    GoTo EXITSUB
                End If
            ElseIf Left(textFA10, 3) <> "011" And textFA31 = "3" Then
                If MsgBox("國籍為「不是」日本，定稿語文確定是「日文」？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                    textFA31.SetFocus
                    textFA31_GotFocus
                    SSTab1.Tab = 4
                    GoTo EXITSUB
                End If
            End If
        End If
    End If
    'end 2021/11/26
    
    'Added by Lydia 2019/12/04 年費不續辦衝突管制：
                                     '若設申請人設定年費自動代繳/年費不續辦，與代理人有衝突，發email通知承辦CC程序管制
    If m_EditMode = 2 Then
      If textFA41.Text <> "" And textFA41.Tag <> textFA41.Text Then
          '保留
          'strExc(0) = "SELECT PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) AS CASENO" & _
                           " FROM PATENT,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5" & _
                           " WHERE SUBSTR(PA75,1,8)='" & Left(ChangeCustomerL(textFA01), 8) & "' " & _
                           " AND SUBSTR(PA26,1,8)=C1.CU01(+) AND SUBSTR(PA26,9,1)=C1.CU02(+)" & _
                           " AND SUBSTR(PA27,1,8)=C2.CU01(+) AND SUBSTR(PA27,9,1)=C2.CU02(+)" & _
                           " AND SUBSTR(PA28,1,8)=C3.CU01(+) AND SUBSTR(PA28,9,1)=C3.CU02(+)" & _
                           " AND SUBSTR(PA29,1,8)=C4.CU01(+) AND SUBSTR(PA29,9,1)=C4.CU02(+)" & _
                           " AND SUBSTR(PA30,1,8)=C5.CU01(+) AND SUBSTR(PA30,9,1)=C5.CU02(+)" & _
                           " AND NVL(X1.CU74,NVL(X2.CU74,NVL(X3.CU74,NVL(X4.CU74,X5.CU74))))=" & CNULL(IIf(textFA41.Text = "Y", "N", "Y"))
          strExc(0) = "SELECT COUNT(*) AS CNT" & _
                           " FROM PATENT,CUSTOMER C1,CUSTOMER C2,CUSTOMER C3,CUSTOMER C4,CUSTOMER C5" & _
                           " WHERE SUBSTR(PA75,1,8)='" & Left(ChangeCustomerL(textFA01), 8) & "' " & _
                           " AND SUBSTR(PA26,1,8)=C1.CU01(+) AND SUBSTR(PA26,9,1)=C1.CU02(+)" & _
                           " AND SUBSTR(PA27,1,8)=C2.CU01(+) AND SUBSTR(PA27,9,1)=C2.CU02(+)" & _
                           " AND SUBSTR(PA28,1,8)=C3.CU01(+) AND SUBSTR(PA28,9,1)=C3.CU02(+)" & _
                           " AND SUBSTR(PA29,1,8)=C4.CU01(+) AND SUBSTR(PA29,9,1)=C4.CU02(+)" & _
                           " AND SUBSTR(PA30,1,8)=C5.CU01(+) AND SUBSTR(PA30,9,1)=C5.CU02(+)" & _
                           " AND NVL(C1.CU74,NVL(C2.CU74,NVL(C3.CU74,NVL(C4.CU74,C5.CU74))))=" & CNULL(IIf(textFA41.Text = "Y", "N", "Y"))
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
              '保留
              'strExc(1) = ""
              'RsTemp.MoveFirst
              'Do While Not RsTemp.EOF
              '     strExc(1) = strExc(1) & vbCrLf & convForm("" & RsTemp.Fields("caseno"), 15)
              '     RsTemp.MoveNext
              'Loop
              'If strExc(1) <> "" Then strExc(1) = convForm("本所案號", 15) & vbCrLf & strExc(1)
              If Val("" & RsTemp.Fields("CNT")) > 0 Then
                    strExc(2) = "目前案件有申請人設為" & IIf(textFA41.Text = "N", "年費自動代繳=Y", "年費不續辦=N") & _
                                 "，此代理人不可設定" & IIf(textFA41.Text = "Y", "年費自動代繳=Y", "年費不續辦=N") & "，請改在個案設定！"
                    MsgBox strExc(2), vbExclamation
                    
                    If textFA10 <> "" Then
                        strExc(0) = "select NA16,NA51 from nation where na01='" & textFA10 & "' "
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                            strExc(3) = "" & RsTemp.Fields("NA16") 'FCP管制
                            strExc(4) = "" & RsTemp.Fields("NA51") 'FCP承辦
                            If strExc(4) = "" Then
                                strExc(4) = strExc(3): strExc(3) = ""
                            End If
                            '保留  'Y與X設定有衝突，請確認。
                            'PUB_SendMail strUserNum, strExc(4), "", textFA01 & textFA02 & "，此代理人不可設定" & IIf(textFA41.Text = "Y", "年費自動代繳=Y", "年費不續辦=N") & "，請改在個案設定！", strExc(2) & vbCrLf & vbCrLf & strExc(1), , , , , , strExc(3)
                            PUB_SendMail strUserNum, strExc(4), "", textFA01 & textFA02 & "代理人設定與申請人設定有衝突，請確認", "同主旨", , , , , , strExc(3)
                        End If
                    End If
                    textFA41.Text = ""
              End If
          End If
      End If
    End If
    'end 2019/12/04
      
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textFA01_GotFocus()
   InverseTextBox textFA01
   CloseIme
End Sub

Private Sub textFA02_GotFocus()
   InverseTextBox textFA02
End Sub

Private Sub textFA03_GotFocus()
   InverseTextBox textFA03
End Sub

Private Sub textFA04_GotFocus()
   InverseTextBox textFA04
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA04.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA05_GotFocus()
   InverseTextBox textFA05
End Sub

Private Sub textFA06_GotFocus()
   InverseTextBox textFA06
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA06.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA07_GotFocus()
   InverseTextBox textFA07
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA07.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA08_GotFocus()
   InverseTextBox textFA08
End Sub

Private Sub textFA09_GotFocus()
   InverseTextBox textFA09
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA09.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA10_GotFocus()
   InverseTextBox textFA10
End Sub

Private Sub textFA11_GotFocus()
   InverseTextBox textFA11
End Sub

Private Sub textFA12_GotFocus()
   InverseTextBox textFA12
End Sub

Private Sub textFA13_GotFocus()
   InverseTextBox textFA13
End Sub

Private Sub textFA14_GotFocus()
   InverseTextBox textFA14
End Sub

Private Sub textFA15_GotFocus()
   InverseTextBox textFA15
End Sub

Private Sub textFA16_GotFocus()
   InverseTextBox textFA16
End Sub

Private Sub textFA17_GotFocus()
   InverseTextBox textFA17
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA17.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA18_GotFocus()
   InverseTextBox textFA18
End Sub

Private Sub textFA19_GotFocus()
   InverseTextBox textFA19
End Sub

Private Sub textFA20_GotFocus()
   InverseTextBox textFA20
End Sub

Private Sub textFA21_GotFocus()
   InverseTextBox textFA21
End Sub

Private Sub textFA22_GotFocus()
   InverseTextBox textFA22
End Sub

Private Sub textFA70_GotFocus()
   InverseTextBox textFA70
End Sub

Private Sub textFA23_GotFocus()
   InverseTextBox textFA23
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA23.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA24_GotFocus()
   InverseTextBox textFA24
End Sub

Private Sub textFA25_GotFocus()
   InverseTextBox textFA25
End Sub

Private Sub textFA26_GotFocus()
   InverseTextBox textFA26
End Sub

Private Sub textFA27_GotFocus()
   InverseTextBox textFA27
End Sub

Private Sub textFA28_GotFocus()
   InverseTextBox textFA28
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA106_GotFocus()
   InverseTextBox textFA106
End Sub

Private Sub textFA29_GotFocus()
   InverseTextBox textFA29
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA29.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA30_GotFocus()
   InverseTextBox textFA30
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA107_GotFocus()
   InverseTextBox textFA107
End Sub

Private Sub textFA31_GotFocus()
   InverseTextBox textFA31
End Sub

Private Sub textFA32_GotFocus()
   InverseTextBox textFA32
End Sub

Private Sub textFA33_GotFocus()
   InverseTextBox textFA33
End Sub

Private Sub textFA34_GotFocus()
   InverseTextBox textFA34
End Sub

Private Sub textFA35_GotFocus()
   InverseTextBox textFA35
End Sub

Private Sub textFA36_GotFocus()
   InverseTextBox textFA36
End Sub

Private Sub textFA37_GotFocus()
   InverseTextBox textFA37
End Sub

Private Sub textFA38_GotFocus()
   InverseTextBox textFA38
End Sub

Private Sub textFA39_GotFocus()
   InverseTextBox textFA39
End Sub

Private Sub textFA40_GotFocus()
   InverseTextBox textFA40
End Sub

Private Sub textFA41_GotFocus()
   InverseTextBox textFA41
End Sub

Private Sub textFA42_GotFocus()
   InverseTextBox textFA42
End Sub

'Private Sub textFA43_GotFocus()
'   InverseTextBox textFA43
'End Sub
'
''Add By Sindy 2011/3/4
'Private Sub textFA108_GotFocus()
'   InverseTextBox textFA108
'End Sub

Private Sub textFA44_GotFocus()
   InverseTextBox textFA44
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA109_GotFocus()
   InverseTextBox textFA109
End Sub

'Add By Sindy 2013/8/15
Private Sub textFA117_GotFocus()
   InverseTextBox textFA117
End Sub

Private Sub textFA45_GotFocus()
   InverseTextBox textFA45
End Sub

'Add By Sindy 2011/3/4
Private Sub textFA110_GotFocus()
   InverseTextBox textFA110
End Sub

Private Sub textFA52_GotFocus()
   InverseTextBox textFA52
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA52.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA53_GotFocus()
   InverseTextBox textFA53
End Sub

Private Sub textFA54_GotFocus()
   InverseTextBox textFA54
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA54.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA55_GotFocus()
   InverseTextBox textFA55
End Sub

Private Sub textFA56_GotFocus()
   InverseTextBox textFA56
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA56.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA57_GotFocus()
   InverseTextBox textFA57
End Sub

Private Sub textFA58_GotFocus()
   InverseTextBox textFA58
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA58.IMEMode = 1
   OpenIme
End Sub

Private Sub textFA59_GotFocus()
   InverseTextBox textFA59
End Sub

Private Sub textFA60_GotFocus()
   InverseTextBox textFA60
End Sub

Private Sub textFA61_GotFocus()
   InverseTextBox textFA61
End Sub

Private Sub textFA62_GotFocus()
   InverseTextBox textFA62
End Sub

Private Sub textFA63_GotFocus()
   InverseTextBox textFA63
End Sub

Private Sub textFA64_GotFocus()
   InverseTextBox textFA64
End Sub

Private Sub textFA65_GotFocus()
   InverseTextBox textFA65
End Sub

Private Sub textFA66_GotFocus()
   InverseTextBox textFA66
End Sub

Private Sub textFA67_GotFocus()
   InverseTextBox textFA67
End Sub

Private Sub textFA68_GotFocus()
   InverseTextBox textFA68
End Sub

Private Sub textFA78_GotFocus()
   InverseTextBox textFA78
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textFA78.IMEMode = 1
   OpenIme
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strMsg As String 'Add by Amy 2017/03/10
   
TxtValidate = False

'Add by Sindy 2021/12/07 檢查畫面上的物件是否含有Unicode文字
If PUB_ChkUniText(Me, True, True) = False Then
   Exit Function
End If

'Add by Morgan 2009/10/16
For Each objTxt In Me.txtFA
   If objTxt.Enabled = True Then
      Cancel = False
      txtFA_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next
'end 2009/10/16

If Me.textFA01.Enabled = True Then
   Cancel = False
   textFA01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA02.Enabled = True Then
   Cancel = False
   textFA02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA03.Enabled = True Then
   Cancel = False
   textFA03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA04.Enabled = True Then
   Cancel = False
   textFA04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA06.Enabled = True Then
   Cancel = False
   textFA06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA07.Enabled = True Then
   Cancel = False
   textFA07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA09.Enabled = True Then
   Cancel = False
   textFA09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA10.Enabled = True Then
   Cancel = False
   textFA10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA11.Enabled = True Then
   Cancel = False
   textFA11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA17.Enabled = True Then
   Cancel = False
   textFA17_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA23.Enabled = True Then
   Cancel = False
   textFA23_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA24.Enabled = True Then
   Cancel = False
   textFA24_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA25.Enabled = True Then
   Cancel = False
   textFA25_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA26.Enabled = True Then
   Cancel = False
   textFA26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA27.Enabled = True Then
   Cancel = False
   textFA27_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA29.Enabled = True Then
   Cancel = False
   textFA29_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA30.Enabled = True Then
   Cancel = False
   textFA30_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2011/3/4
If Me.textFA107.Enabled = True Then
   Cancel = False
   textFA107_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textFA31.Enabled = True Then
   Cancel = False
   textFA31_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA38.Enabled = True Then
   Cancel = False
   textFA38_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA39.Enabled = True Then
   Cancel = False
   textFA39_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA40.Enabled = True Then
   Cancel = False
   textFA40_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA41.Enabled = True Then
   Cancel = False
   textFA41_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA42.Enabled = True Then
   Cancel = False
   textFA42_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'If Me.textFA43.Enabled = True Then
'   Cancel = False
'   textFA43_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If
'
''Add By Sindy 2011/3/4
'If Me.textFA108.Enabled = True Then
'   Cancel = False
'   textFA108_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If

'Add By Sindy 2013/1/17
For i = 0 To 1
   If Me.Combo2(i).Enabled = True Then
      Cancel = False
      Combo2_Validate i, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next i
'2013/1/17 End

If Me.textFA44.Enabled = True Then
   Cancel = False
   textFA44_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2011/3/4
If Me.textFA109.Enabled = True Then
   Cancel = False
   textFA109_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textFA52.Enabled = True Then
   Cancel = False
   textFA52_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA54.Enabled = True Then
   Cancel = False
   textFA54_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA55.Enabled = True Then
   Cancel = False
   textFA55_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA56.Enabled = True Then
   Cancel = False
   textFA56_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA58.Enabled = True Then
   Cancel = False
   textFA58_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA59.Enabled = True Then
   Cancel = False
   textFA59_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA61.Enabled = True Then
   Cancel = False
   textFA61_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA62.Enabled = True Then
   Cancel = False
   textFA62_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA66.Enabled = True Then
   Cancel = False
   textFA66_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA67.Enabled = True Then
   Cancel = False
   textFA67_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA68.Enabled = True Then
   Cancel = False
   textFA68_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'Modify by Amy 2015/08/24 改為下拉選單 原:textFA69
If Me.cboStatus.Enabled = True Then
   Cancel = False
   cboStatus_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA71.Enabled = True Then
   Cancel = False
   textFA71_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2011/3/4
If Me.textFA111.Enabled = True Then
   Cancel = False
   textFA111_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textFA72.Enabled = True Then
   Cancel = False
   textFA72_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2011/3/4
If Me.textFA112.Enabled = True Then
   Cancel = False
   textFA112_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Cheng 2003/11/17
If Me.textFA73.Enabled = True Then
   Cancel = False
   textFA73_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA74.Enabled = True Then
   Cancel = False
   textFA74_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA75.Enabled = True Then
   Cancel = False
   textFA75_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2025/3/10
If Me.textFA137.Enabled = True Then
   Cancel = False
   textFA137_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA138.Enabled = True Then
   Cancel = False
   textFA138_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textFA139.Enabled = True Then
   Cancel = False
   textFA139_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'2025/3/10 END

'add by nickc 2005/12/02
If Me.TextFA76.Enabled = True Then
   Cancel = False
   TextFA76_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'End
If Me.textFA78.Enabled = True Then
   Cancel = False
   textFA78_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'2008/10/21 add by Toni
If Me.TextFA93.Enabled = True Then
   Cancel = False
   TextFA93_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'end 2008/10/21
'2008/12/9 add by sonia
If Me.textFA97.Enabled = True Then
   Cancel = False
   textFA97_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'2008/12/9 end

'Add By Sindy 2011/3/10
If Me.textFA100.Enabled = True Then
   Cancel = False
   textFA100_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2013/8/15
If Me.textFA117.Enabled = True Then
   Cancel = False
   textFA117_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add by Amy 2022/11/25 來所原因(原名稱:代理人來源)
'Modify by 2024/11/29 改抓共用函數,避免沒改到
strMsg = ChkXYSourceReason(0, Me.Name, m_EditMode, cboSource, txtXYS02, _
                           , m_FieldList(10).fiOldData, m_FieldList(46).fiOldData, txtXYS03, textFA01)
If strMsg <> MsgText(601) Then
   MsgBox strMsg, vbInformation
   SSTab1.Tab = 0
   If InStr(strMsg, "來所原因 不可為空") > 0 Then
      cboSource.SetFocus
   ElseIf InStr(strMsg, "介紹者編號") > 0 Then
      txtXYS02.SetFocus
   ElseIf InStr(strMsg, "其他說明") > 0 Then
      txtXYS03.SetFocus
   End If
   Exit Function
End If
strMsg = ChkXYSourceReason(2, Me.Name, m_EditMode, cboSource, txtXYS02, m_FieldList(126).fiOldData)
If strMsg <> MsgText(601) Then
   If MsgBox(strMsg, vbYesNo + vbCritical) = vbNo Then
      SSTab1.Tab = 0
      txtXYS02.SetFocus
      Exit Function
   End If
End If
strMsg = ""
'end 2024/11/29

'Add by Morgan 2010/7/16
'新增關係企業時檢查母號特定欄位設定並提醒
If m_EditMode = 1 And Mid(textFA01, 7) > "00" Then
   strExc(0) = "select * from fagent where fa01='" & Left(textFA01, 6) & "00' and fa02='0'"
   strExc(1) = ""
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         If textFA85 = "" And Not IsNull(.Fields("fa85")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "FCP是否核對已准專利：" & .Fields("fa85")
         End If
         If textFA41 = "" And Not IsNull(.Fields("fa41")) Then
            'Modified by Lydia 2016/08/18
            'strExc(1) = strExc(1) & vbCrLf & vbTab & "FCP年費自動代繳：" & .Fields("fa41")
            strExc(1) = strExc(1) & vbCrLf & vbTab & "FCP年費自動代繳(Y)/寄證書後年費不續辦(N)：" & .Fields("fa41")
         End If
         If textFA28 = "" And Not IsNull(.Fields("fa28")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "代理人專利財務編號：" & .Fields("fa28")
         End If
         'Add By Sindy 2011/3/4
         If textFA106 = "" And Not IsNull(.Fields("fa106")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "代理人商標財務編號：" & .Fields("fa106")
         End If
         '2011/3/4 End
         'Modified by Lydia 2016/08/18 Debug
         'If textFA41 = "" And Not IsNull(.Fields("fa42")) Then
         If textFA42 = "" And Not IsNull(.Fields("fa42")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "FCP領證自動代繳：" & .Fields("fa42")
         End If
         If txtFA(95) = "" And Not IsNull(.Fields("fa95")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利領證折扣：" & .Fields("fa95") & "%"
         End If
         If txtFA(96) = "" And Not IsNull(.Fields("fa96")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利年費折扣：" & .Fields("fa96") & "%"
         End If
         If textFA25 = "" And Not IsNull(.Fields("fa25")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利全部折扣：" & .Fields("fa25") & "%"
         End If
         If textFA26 = "" And Not IsNull(.Fields("fa26")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利申請/翻譯折扣：" & .Fields("fa26") & "%"
         End If
         If textFA26 = "" And Not IsNull(.Fields("fa27")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利全部折扣起始日：" & TransDate(.Fields("fa27"), 1)
         End If
         If textFA30 = "" And Not IsNull(.Fields("fa30")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利固定請款對象：" & .Fields("fa30")
         End If
         'Add By Sindy 2011/3/4
         If textFA107 = "" And Not IsNull(.Fields("fa107")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "商標固定請款對象：" & .Fields("fa107")
         End If
         '2011/3/4 End
         If textFA71 = "" And Not IsNull(.Fields("fa71")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利D/N固定列印對象：" & .Fields("fa71")
         End If
         'Add By Sindy 2011/3/4
         If textFA111 = "" And Not IsNull(.Fields("fa111")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "商標D/N固定列印對象：" & .Fields("fa111")
         End If
         '2011/3/4 End
         If textFA72 = "" And Not IsNull(.Fields("fa72")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "年費D/N列印對象：" & .Fields("fa72")
         End If
         'Add By Sindy 2011/3/4
         If textFA112 = "" And Not IsNull(.Fields("fa112")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "延展D/N列印對象：" & .Fields("fa112")
         End If
         '2011/3/4 End
         If textFA87 = "" And Not IsNull(.Fields("fa87")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利定稿份數：" & .Fields("fa87")
         End If
         If textFA89 = "" And Not IsNull(.Fields("fa89")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利請款單份數：" & .Fields("fa89")
         End If
         If txtFA(86) = "" And Not IsNull(.Fields("fa86")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利以 EMail 通知：" & .Fields("fa86")
         End If
         If txtFA(98) = "" And Not IsNull(.Fields("fa98")) Then
            strExc(1) = strExc(1) & vbCrLf & vbTab & "專利 Email 同時寄紙本：" & .Fields("fa98")
         End If
         
         If strExc(1) <> "" Then
            If MsgBox("下列欄位母號有設定但本關係企業並未設定，請確認是否要繼續？" & vbCrLf & strExc(1), vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
      End With
   End If
End If
'end 2010/7/16

'Added by Morgan 2022/9/13
'E化及份數檢查
If txtFA(86) <> "" And txtFA(98) = "" Then
   If textFA87 <> "" Then
      MsgBox "當設定【專利以 EMail 通知】，需同時設定【專利 Email 同時寄紙本】才可指定【專利定稿份數】！", vbExclamation
      SSTab1.Tab = 2
      txtFA(98).SetFocus
      Exit Function
   ElseIf textFA89 <> "" Then
      MsgBox "當設定【專利以 EMail 通知】，需同時設定【專利 Email 同時寄紙本】才可指定【專利請款單份數】！", vbExclamation
      SSTab1.Tab = 2
      txtFA(98).SetFocus
      Exit Function
   End If
End If
If txtFA(91) <> "" And txtFA(99) = "" Then
   If textFA88 <> "" Then
      MsgBox "當設定【商標以 EMail 通知】，需同時設定【商標 Email 同時寄紙本】才可指定【商標定稿份數】！", vbExclamation
      SSTab1.Tab = 3
      txtFA(99).SetFocus
      Exit Function
   ElseIf textFA90 <> "" Then
      MsgBox "當設定【商標以 EMail 通知】，需同時設定【商標 Email 同時寄紙本】才可指定【商標請款單份數】！", vbExclamation
      SSTab1.Tab = 3
      txtFA(99).SetFocus
      Exit Function
   End If
End If
'end 2022/9/13

'Add by Amy 2017/01/05 +管控智權人員控制
If textFA10 > "010" And (Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "P2") Then
    If Left(Pub_StrUserSt03, 2) = "P2" And textFA120 = MsgText(601) Then
        MsgBox Left(Label53, Len(Label53) - 1) & " 不可為空！", vbInformation
        SSTab1.Tab = 3
        textFA120.SetFocus
        textFA120_GotFocus
        Exit Function
    End If
    Cancel = False
    textFA120_Validate Cancel
    If Cancel = True Then
        Exit Function
    End If
End If
'end 2017/01/05
'Add by Amy 2017/03/10 內容有?彈訊息詢問
If InStr(textFA17, "?") > 0 Then strMsg = strMsg & "、" & Left(Label18, Len(Label18) - 1)
If InStr(textFA32, "?") > 0 Then strMsg = strMsg & "、" & "POB" & Left(Label41(7), Label41(7) - 1)
If InStr(textFA33, "?") > 0 Then strMsg = strMsg & "、" & "POB" & Left(Label41(12), Label41(12) - 1)
If InStr(textFA34, "?") > 0 Then strMsg = strMsg & "、" & "POB" & Left(Label41(8), Label41(8) - 1)
If InStr(textFA35, "?") > 0 Then strMsg = strMsg & "、" & "POB" & Left(Label41(11), Label41(11) - 1)
If InStr(textFA36, "?") > 0 Then strMsg = strMsg & "、" & "POB" & Left(Label41(9), Label41(9) - 1)
If InStr(textFA16, "?") > 0 Then strMsg = strMsg & "、" & "Email(代表)"
If InStr(textFA79, "?") > 0 Then strMsg = strMsg & "、" & "Email(財務)"
If InStr(textFA80, "?") > 0 Then strMsg = strMsg & "、" & "Email(其他1)"
If InStr(textFA81, "?") > 0 Then strMsg = strMsg & "、" & "Email(其他2)"
If InStr(textFA82, "?") > 0 Then strMsg = strMsg & "、" & "Email(其他3)"
If strMsg <> MsgText(601) Then
    If MsgBox("輸入的" & Mid(strMsg, 2) & "含有問號，" & vbCrLf & "　　　　若是正常請按　是　繼續" & vbCrLf & "　　　　若不正常請按　否　修正！", vbYesNo) = vbNo Then
        Exit Function
    End If
End If

TxtValidate = True
End Function

'Add By Cheng 2003/03/27
Private Function GetAutoNumY(strFA01 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetAutoNumY = ""
StrSQLa = "Select * From AutoNumber Where AU01='Y' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If "" & rsA("AU03").Value < strFA01 Then
        GetAutoNumY = "" & rsA("AU03").Value
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'2005/12/21 ADD BY SONIA
Private Function MODCUSTOMER() As Boolean
Dim strTit As String
Dim strTmp As String

   If textFA03 <> "" Then
      strTit = "select * from customer where cu01='" & Left(textFA03 & "00000000", 8) & "' and cu02='0' "
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strTit, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount <> 0 Then
         If Mid(CheckStr(AdoRecordSet3.Fields("cu03")), 1, 8) <> textFA01 And CheckStr(AdoRecordSet3.Fields("cu03")) <> "" Then
            '2008/7/17 modify by sonia 改為詢問是否更新
            strTmp = "代理人編號與客戶檔之代理人編號(" & Mid(CheckStr(AdoRecordSet3.Fields("cu03")), 1, 8) & ")不同 ! 是否更新客戶檔之代理人編號 ?"
            If MsgBox(strTmp, vbYesNo + vbCritical) = vbYes Then
               strSql = "UPDATE CUSTOMER SET CU03='" & textFA01 & "' WHERE CU01='" & Left(textFA03 & "00000000", 8) & "'"
               cnnConnection.Execute strSql
            End If
         '2008/7/17 add by sonia
         ElseIf CheckStr(AdoRecordSet3.Fields("cu03")) = "" Then
            strSql = "UPDATE CUSTOMER SET CU03='" & textFA01 & "' WHERE CU01='" & Left(textFA03 & "00000000", 8) & "'"
            cnnConnection.Execute strSql
         '2008/7/17 end
         End If
      End If
      '2008/7/17 ADD BY SONIA
      If textFA03.Tag <> "" And textFA03.Tag <> textFA03 Then
         strTmp = "原客戶編號(" & textFA03.Tag & ")之客戶檔的代理人編號是否清除 ?"
         If MsgBox(strTmp, vbYesNo + vbCritical) = vbYes Then
            strSql = "UPDATE CUSTOMER SET CU03=NULL WHERE CU01='" & textFA03.Tag & "' and cu02='0' "
            cnnConnection.Execute strSql
         End If
      End If
      '2008/7/17 END
   'add by nickc 2006/03/16 若清空，拿除關聯
   ElseIf Trim(textFA03.Text) = "" And textFA03.Tag <> "" Then
      strSql = "UPDATE CUSTOMER SET CU03=null WHERE CU01='" & textFA03.Tag & "' and cu02='0' "
      cnnConnection.Execute strSql
   End If
End Function
'2005/12/21 END

'Add by Morgan 2008/11/13
Private Sub txtFA_GotFocus(Index As Integer)
   If txtFA(Index).Enabled = True And txtFA(Index).Locked = False Then
      InverseTextBox txtFA(Index)
      CloseIme
   End If
End Sub
'Add by Morgan 2008/11/13
Private Sub txtFA_KeyPress(Index As Integer, KeyAscii As Integer)
   If txtFA(Index).Enabled = True And txtFA(Index).Locked = False Then
      Select Case Index
         Case 95, 96
            If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         'Add by Morgan 2009/9/15
         'Modified by Morgan 2014/6/4
         'Case 86, 91, 98, 99, 102
         'Modified by Lydia 2017/11/30 +FCP是否電子送件(FA104)
         'Modified by Morgan 2019/1/18 +124
         'Modified by Morgan 2025/2/10 +136
         Case 98, 99, 102, 104, 124, 136
            KeyAscii = UpperCase(KeyAscii)
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         'Added by Morgan 2014/6/4
         Case 86, 91
            KeyAscii = UpperCase(KeyAscii)
            If KeyAscii <> 89 And KeyAscii <> 68 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         'Added by Lydia 2018/01/19 增加是否寄發竹曆與促銷信,只可輸入Y/N
         'Modified by Lydia 2018/06/28 +FA123 是否同意歐盟通用資料保護規範(GDPR)
         Case 121, 122, 123
            KeyAscii = UpperCase(KeyAscii)
            'Added by Lydia 2018/08/10 GDPR增加輸入W
            If Index = 123 Then
                If KeyAscii <> 87 And KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
                   KeyAscii = 0
                   Beep
                End If
            Else
            'end 2018/08/10
                If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
                   KeyAscii = 0
                   Beep
                End If
            End If 'end 2018/08/10
         'end 2018/01/19
         'Added by Morgan 2022/12/1
         Case 128, 129
            If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
      End Select
   End If
End Sub

'Add by Morgan 2009/10/16
Private Sub txtFA_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 86, 98
         If (txtFA(86) = "" And txtFA(98) = "Y") Then
            MsgBox "【專利 EMail 同時寄紙本】為 Y 時，【專利以 EMail 通知】欄位也必須為 Y！"
            Cancel = True
            Exit Sub
         End If
      Case 91, 99
         If (txtFA(91) = "" And txtFA(99) = "Y") Then
            MsgBox "【商標 EMail 同時寄紙本】為 Y 時，【商標以 EMail 通知】欄位也必須為 Y！"
            Cancel = True
            Exit Sub
         End If
   End Select
End Sub
'Added by Morgan 2011/12/30
Private Function FA100CheckError(Optional bolNoMsg As Boolean) As Boolean
   If textFA100.Text = "" Then
      If textFA10 <> "020" Then
         FA100CheckError = True
         If bolNoMsg = False Then ShowMsg "代理人國籍不是大陸，不可設為要寄專利雙週報 !"
      Else
         If TextFA76 = "C" Then
            FA100CheckError = True
            If bolNoMsg = False Then ShowMsg "代理人性質為其他時，不可設為要寄專利雙週報 !"
         End If
      End If
      
   ElseIf textFA100.Text <> "N" Then
      FA100CheckError = True
      If bolNoMsg = False Then ShowMsg "輸入錯誤 !"
   End If
End Function
'Added by Morgan 2011/12/30
Private Sub SetFA100()
   If textFA10 = "020" And TextFA76 <> "C" Then
      textFA100 = ""
   Else
      textFA100 = "N"
   End If
End Sub

'Add by Amy 2016/06/30
Private Function CheckAddrData(ByRef objTxt As Object, ByRef stMsg As String) As Boolean
    Dim strZipCode As String, strAddr As String, strCountry As String, strIndArea As String, strROC As String
    Dim bolMany As Boolean, intArea As Integer
    
    CheckAddrData = False: stMsg = ""
        
    objTxt.Text = ReplaceAddrTW(objTxt.Text)
    strROC = ""
    strAddr = objTxt.Text
    strAddr = AddrToZipAddr(strAddr) 'Add by Amy 2016/07/06 +去除有郵遞區號
    If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
    If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
    If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
    '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
    strIndArea = "True"
    strAddr = ReplaceIndArea(strAddr, strIndArea)
    If strIndArea = "True" Then strIndArea = MsgText(601)
    If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
        strIndArea = "新竹" & strIndArea
        strAddr = Mid(strAddr, 3)
    End If
    '** 第3個字是 縣 / 市
    If Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣" Or Mid(strAddr, 1, 3) = "釣魚臺" Or Mid(strAddr, 1, 3) = "海南島" Then
       'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
       If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
          Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
          Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
            '傳入地址前7個字抓到郵遞區號
            intArea = 7
            strZipCode = GetPostZip(Left(strAddr, 7), 7, , strCountry, bolMany)
            '傳入地址前6個字抓到郵遞區號
            If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 6), 6, , strCountry, bolMany): intArea = 6
             'end 2018/12/19
            '傳入地址前5個字取郵遞區號
            If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 5), 5, , strCountry, bolMany): intArea = 5
            '抓到郵遞區號
            If strZipCode = MsgText(601) Then
                stMsg = "臺灣地址有誤！"
                Exit Function
            End If
        '無鄉/鎮/市/區
        Else
            stMsg = "臺灣地址格式有誤，無鄉/鎮/市/區！"
            Exit Function
        End If
     '** 第三3個字無 縣 / 市
    Else
        stMsg = "臺灣地址格式有誤，第三3個字無 縣 / 市！"
        Exit Function
    End If
        
    CheckAddrData = True
End Function

'Add by Amy 2016/07/06 將地址拆成郵遞區號及地址
Private Function AddrToZipAddr(ByVal strCAddr As String) As String
    Dim strZip As String
 
    Do While Left(strCAddr, 1) >= "０" And Left(strCAddr, 1) <= "９"
       strZip = strZip & Left(strCAddr, 1)
       strCAddr = Mid(strCAddr, 2)
    Loop
    AddrToZipAddr = strCAddr
End Function

'Added by Lydia 2017/03/31 去掉跳行符號
'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA04_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA04 = PUB_StringFilter(textFA04)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA05_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA05 = PUB_StringFilter(textFA05)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA06_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA06 = PUB_StringFilter(textFA06)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA63_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA63 = PUB_StringFilter(textFA63)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA64_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA64 = PUB_StringFilter(textFA64)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA65_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA65 = PUB_StringFilter(textFA65)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA17_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA17 = PUB_StringFilter(textFA17)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA18_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA18 = PUB_StringFilter(textFA18)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA19_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA19 = PUB_StringFilter(textFA19)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA20_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA20 = PUB_StringFilter(textFA20)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA21_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA21 = PUB_StringFilter(textFA21)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA22_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA22 = PUB_StringFilter(textFA22)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA23_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA23 = PUB_StringFilter(textFA23)
End Sub

'Modified by Lydia 2017/04/05 從lostfocus改成change
Private Sub textFA70_Change()
   'Modified by Lydia 2017/04/05 限定在新增或修改
   If m_EditMode = 1 Or m_EditMode = 2 Then textFA70 = PUB_StringFilter(textFA70)
End Sub
'end 2017/03/31

'Add by Amy 2022/11/25
Private Sub txtXYS02_GotFocus()
    InverseTextBox txtXYS02
End Sub

Private Sub txtXYS02_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'Memo 使用bCancel避免彈訊息後無法跳離 ex:來源選04 輸了Y編號,需刪Y編號,再重選
Private Sub txtXYS02_Validate(Cancel As Boolean)
    'Modify by 2024/11/29 改抓共用函數,避免有未改到
    Dim stName As String, stMsg As String
    
    If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
    If txtXYS02 = MsgText(601) Then LblSourceN.Caption = "": Exit Sub
    
    'Modify by Amy 2022/12/28 直接按 Enter鍵存檔,部分資料未正常檢查,並調整訊息
    bCancel = False
    LblSourceN.Caption = ""
    txtXYS02 = Left(ChangeCustomerL(txtXYS02), 8) '補滿8碼
    stMsg = ChkXYSourceReason(1, Me.Name, m_EditMode, cboSource, txtXYS02, , , , , textFA01, stName)
    If stMsg <> MsgText(601) Then
         MsgBox stMsg, vbInformation
         SSTab1.Tab = 0
         'Memo 使用bCancel避免彈訊息後無法跳離 ex:來源選04 輸了Y編號,需刪Y編號,再重選
         bCancel = True
         txtXYS02_GotFocus
         Exit Sub
    End If
    LblSourceN.Caption = stName
   'end 2024/11/29
End Sub

Private Sub txtXYS03_GotFocus()
    InverseTextBox txtXYS03
End Sub

'Add by Amy 2022/12/28 將ModRecord 發mail 的部分拆出
Private Sub ModRecordMail(ByRef strTmp As String, ByRef strMsg As String)
    'Added by Lydia 2020/03/17 FMP案件不自動上年費不續辦，改發清單給程序，由各區程序逐筆產生定稿通知大陸代理人
     If strTmp <> "" Then
        If PUB_GetP605Email("1", strTmp, strMsg) = False Then
           If strMsg <> "" Then
               MsgBox strMsg, vbCritical
           End If
        End If
     End If
     'end 2020/03/17
     
      'Add by Morgan 2006/1/10 若代理人有修改且有相對的匯款銀行資料時需發Mail通知婧瑄
      '2011/12/26 MODIFY BY SONIA 加入中日文名稱欄位
      'If (textFA05.Tag & textFA63.Tag & textFA64.Tag & textFA65.Tag <> "") And (textFA05.Tag & textFA63.Tag & textFA64.Tag & textFA65.Tag <> textFA05 & textFA63 & textFA64 & textFA65) Then
         'PUB_AccDataCheck textFA01 & textFA02, "英：" & textFA05.Tag & " " & textFA63.Tag & " " & textFA64.Tag & " " & textFA65.Tag & " --> " & textFA05 & " " & textFA63 & " " & textFA64 & " " & textFA65
      'End If
      If (textFA04.Tag & textFA06.Tag & textFA05.Tag & textFA63.Tag & textFA64.Tag & textFA65.Tag <> "") And (textFA04.Tag & textFA06.Tag & textFA05.Tag & textFA63.Tag & textFA64.Tag & textFA65.Tag <> textFA04 & textFA06 & textFA05 & textFA63 & textFA64 & textFA65) Then
         PUB_AccDataCheck textFA01 & textFA02, _
         "中：" & textFA04.Tag & " --> " & textFA04 & Chr(13) & _
         "英：" & textFA05.Tag & " " & textFA63.Tag & " " & textFA64.Tag & " " & textFA65.Tag & " --> " & textFA05 & " " & textFA63 & " " & textFA64 & " " & textFA65 & Chr(13) & _
         "日：" & textFA06.Tag & " --> " & textFA06
      End If
      '2011/12/26 END
      '2006/1/10 end
      
      'add by nickc 2006/12/26 若有改名稱，地址，電話，傳真將列出該代理人國籍，編號，名稱，3個月內期限，6個月內期限，所有案件數
      'Mark by Sonia 2017/08/11
'      Dim strScanFagent As String
'      'add by nickc 2006/12/28
'      Dim intLine As Integer
'      Dim intLineCnt As Integer
'      Dim nowCnt As Integer
'      Dim Seek01 As String
'      intLineCnt = 4
'      'Modify by Morgan 2010/11/24 排除電話傳真從無到有的修改
'      'If textFA04.Tag <> textFA04.Text Or textFA05.Tag <> textFA05.Text Or textFA06.Tag <> textFA06.Text Or _
'         textFA63.Tag <> textFA63.Text Or textFA64.Tag <> textFA64.Text Or textFA65.Tag <> textFA65.Text Or _
'         textFA12.Tag <> textFA12.Text Or textFA13.Tag <> textFA13.Text Or textFA14.Tag <> textFA14.Text Or _
'         textFA15.Tag <> textFA15.Text Or textFA17.Tag <> textFA17.Text Or textFA18.Tag <> textFA18.Text Or _
'         textFA19.Tag <> textFA19.Text Or textFA20.Tag <> textFA20.Text Or textFA21.Tag <> textFA21.Text Or _
'         textFA22.Tag <> textFA22.Text Or textFA23.Tag <> textFA23.Text Then
'
'      If textFA04.Tag <> textFA04.Text Or textFA05.Tag <> textFA05.Text Or textFA06.Tag <> textFA06.Text Or _
'         textFA63.Tag <> textFA63.Text Or textFA64.Tag <> textFA64.Text Or textFA65.Tag <> textFA65.Text Or _
'         (textFA12.Tag <> "" And textFA12.Tag <> textFA12.Text) Or (textFA13.Tag <> "" And textFA13.Tag <> textFA13.Text) Or _
'         (textFA14.Tag <> "" And textFA14.Tag <> textFA14.Text) Or (textFA15.Tag <> "" And textFA15.Tag <> textFA15.Text) Or _
'         textFA17.Tag <> textFA17.Text Or textFA18.Tag <> textFA18.Text Or _
'         textFA19.Tag <> textFA19.Text Or textFA20.Tag <> textFA20.Text Or textFA21.Tag <> textFA21.Text Or _
'         textFA22.Tag <> textFA22.Text Or textFA23.Tag <> textFA23.Text Then
'
'            strScanFagent = " select fnation,fnum,fname,sum(c3) as C3,sum(c6) as C6,ccount from ("
'            strScanFagent = strScanFagent & " select distinct 'ok' as key, 1 as c3,0 as c6,cp01||cp02||cp03||cp04 from caseprogress where (cp01,cp02,cp03,cp04) in ("
'            strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'            strScanFagent = strScanFagent & " and cp06>=to_number(to_char(sysdate,'YYYYMMDD')) and cp06<=to_number(to_char(add_months(sysdate,3),'YYYYMMDD')) and cp27 is null and cp57 is null"
'            strScanFagent = strScanFagent & " union select 'ok' as key,1 as c3,0 as c6,np02||np03||np04||np05 from nextprogress where (np02,np03,np04,np05) in ("
'            strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'            'Modify By Sindy 2009/07/24 增加LIN系統類別
'            '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
'            strScanFagent = strScanFagent & " and np06 is null and np08>=to_number(to_char(sysdate,'YYYYMMDD')) and np08<=to_number(to_char(add_months(sysdate,3),'YYYYMMDD')) " & strNpSqlOfNoSalesDuty
'            strScanFagent = strScanFagent & " union select 'ok' as key,0 as c3,1 as c6,cp01||cp02||cp03||cp04 from caseprogress where (cp01,cp02,cp03,cp04) in ("
'            strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'            strScanFagent = strScanFagent & " and cp06>=to_number(to_char(sysdate,'YYYYMMDD')) and cp06<=to_number(to_char(add_months(sysdate,6),'YYYYMMDD')) and cp27 is null and cp57 is null"
'            strScanFagent = strScanFagent & " union select 'ok' as key,0 as c3,1 as c6,np02||np03||np04||np05 from nextprogress where (np02,np03,np04,np05) in ("
'            strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'            'Modify By Sindy 2009/07/24 增加LIN系統類別
'            '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
'            strScanFagent = strScanFagent & " and np06 is null and np08>=to_number(to_char(sysdate,'YYYYMMDD')) and np08<=to_number(to_char(add_months(sysdate,6),'YYYYMMDD')) " & strNpSqlOfNoSalesDuty
'            strScanFagent = strScanFagent & " )B,(select 'ok' as key,count(*) as CCount from (select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'            strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')CA"
'            strScanFagent = strScanFagent & " )C,(select 'ok' as key,nvl(na03,na04) as Fnation,fa01||fa02 as fnum,nvl(fa05,nvl(na04,na06)) as fname"
'            strScanFagent = strScanFagent & " from fagent,nation where fa01='" & strFA01 & "' and fa02='" & strFA02 & "' and fa10=na01(+)"
'            strScanFagent = strScanFagent & " )D where D.key=B.key(+) and D.key=C.key(+) group by fnation,fnum,fname,ccount"
'            CheckOC3
'            AdoRecordSet3.CursorLocation = adUseClient
'            AdoRecordSet3.Open strScanFagent, cnnConnection, adOpenStatic, adLockReadOnly
'            If AdoRecordSet3.RecordCount <> 0 Then
'                'add by nickc 2006/12/28 沒資料不印
'                'edit by nickc 2007/01/17 有請作單說有 6 個月案件才印
'                'If Val(CheckStr(AdoRecordSet3.Fields("C3"))) = 0 And Val(CheckStr(AdoRecordSet3.Fields("C6"))) = 0 And Val(CheckStr(AdoRecordSet3.Fields("Ccount"))) = 0 Then
'                If Val(CheckStr(AdoRecordSet3.Fields("C3"))) = 0 And Val(CheckStr(AdoRecordSet3.Fields("C6"))) = 0 Then
'                    MsgBox "無六個月期限案件！", vbInformation, "不印明細表"
'                Else
'                    Printer.Font.Size = "20"
'                    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("修改代理人資料")) / 2
'                    Printer.CurrentY = 0
'                    Printer.Print "修改代理人資料"
'                    Printer.Font.Size = 12
'                    Printer.CurrentX = 0
'                    Printer.CurrentY = 600
'                    Printer.Print "修改人員：" & GetStaffName(strUserNum)
'                    Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 300
'                    Printer.CurrentY = 600
'                    Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'                    Printer.CurrentX = 0
'                    Printer.CurrentY = 900
'                    Printer.Print String(150, "=")
'                    Printer.CurrentX = 200
'                    Printer.CurrentY = 1200
'                    Printer.Print "國  籍"
'                    Printer.CurrentX = 1500
'                    Printer.CurrentY = 1200
'                    Printer.Print "編   號"
'                    Printer.CurrentX = 3000
'                    Printer.CurrentY = 1200
'                    Printer.Print "名         稱"
'                    Printer.CurrentX = 6500
'                    Printer.CurrentY = 1200
'                    Printer.Print "3個月期限"
'                    Printer.CurrentX = 7700
'                    Printer.CurrentY = 1200
'                    Printer.Print "6個月期限"
'                    Printer.CurrentX = 8900
'                    Printer.CurrentY = 1200
'                    Printer.Print "所有案件數"
'                    Printer.CurrentX = 0
'                    Printer.CurrentY = 1500
'                    Printer.Print String(150, "=")
'                    Printer.CurrentX = 200
'                    Printer.CurrentY = 1800
'                    Printer.Print StrToStr(CheckStr(AdoRecordSet3.Fields("fnation")), 5)
'                    Printer.CurrentX = 1500
'                    Printer.CurrentY = 1800
'                    Printer.Print StrToStr(CheckStr(AdoRecordSet3.Fields("fnum")), 5)
'                    Printer.CurrentX = 3000
'                    Printer.CurrentY = 1800
'                    Printer.Print StrToStr(CheckStr(AdoRecordSet3.Fields("fname")), 10)
'                    Printer.CurrentX = 6500 + Printer.TextWidth("3個月期限") - Printer.TextWidth(Format(Val(CheckStr(AdoRecordSet3.Fields("C3"))), "0"))
'                    Printer.CurrentY = 1800
'                    Printer.Print Format(Val(CheckStr(AdoRecordSet3.Fields("C3"))), "0")
'                    Printer.CurrentX = 7700 + Printer.TextWidth("6個月期限") - Printer.TextWidth(Format(Val(CheckStr(AdoRecordSet3.Fields("C6"))), "0"))
'                    Printer.CurrentY = 1800
'                    Printer.Print Format(Val(CheckStr(AdoRecordSet3.Fields("C6"))), "0")
'                    Printer.CurrentX = 8900 + Printer.TextWidth("所有案件數") - Printer.TextWidth(Format(Val(CheckStr(AdoRecordSet3.Fields("ccount"))), "0"))
'                    Printer.CurrentY = 1800
'                    Printer.Print CheckStr(AdoRecordSet3.Fields("ccount"))
'
'                    strScanFagent = "select Cp01||'-'||cp02||'-'||cp03||'-'||cp04,cp01,cp02,cp03,cp04 from caseprogress where (cp01,cp02,cp03,cp04) in ("
'                    strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'                    strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'                    strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'                    strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'                    strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'                    strScanFagent = strScanFagent & " and cp06>=to_number(to_char(sysdate,'YYYYMMDD')) and cp06<=to_number(to_char(add_months(sysdate,6),'YYYYMMDD')) and cp27 is null and cp57 is null"
'                    strScanFagent = strScanFagent & " union select np02||'-'||np03||'-'||np04||'-'||np05,np02 as cp01,np03 as cp02,np04 as cp03,np05 as cp04 from nextprogress where (np02,np03,np04,np05) in ("
'                    strScanFagent = strScanFagent & " select pa01,pa02,pa03,pa04 from patent where pa75='" & strFA01 & strFA02 & "'"
'                    strScanFagent = strScanFagent & " union select tm01,tm02,tm03,tm04 from trademark where tm44='" & strFA01 & strFA02 & "'"
'                    strScanFagent = strScanFagent & " union select sp01,sp02,sp03,sp04 from servicepractice where sp26='" & strFA01 & strFA02 & "'"
'                    strScanFagent = strScanFagent & " union select lc01,lc02,lc03,lc04 from lawcase where lc22='" & strFA01 & strFA02 & "'"
'                    strScanFagent = strScanFagent & " union select cp01,cp02,cp03,cp04 from caseprogress where cp44='" & strFA01 & strFA02 & "')"
'                    'Modify By Sindy 2009/07/24 增加LIN系統類別
'                    '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
'                    strScanFagent = strScanFagent & " and np06 is null and np08>=to_number(to_char(sysdate,'YYYYMMDD')) and np08<=to_number(to_char(add_months(sysdate,6),'YYYYMMDD')) " & strNpSqlOfNoSalesDuty
'                    strScanFagent = strScanFagent & " order by cp01,cp02,cp03,cp04  "
'                    CheckOC3
'
'                    AdoRecordSet3.CursorLocation = adUseClient
'                    AdoRecordSet3.Open strScanFagent, cnnConnection, adOpenStatic, adLockReadOnly
'                    If AdoRecordSet3.RecordCount <> 0 Then
'                        Printer.CurrentX = 0
'                        Printer.CurrentY = 2700
'                        Printer.Print "六個月本所期限案件明細："
'                        intLine = 3000
'                        AdoRecordSet3.MoveFirst
'                        nowCnt = 0
'                        Seek01 = SystemNumber(CheckStr(AdoRecordSet3.Fields(0).Value), 1)
'                        Do While Not AdoRecordSet3.EOF
'                            nowCnt = nowCnt + 1
'                            If nowCnt > intLineCnt Then
'                                nowCnt = 1
'                            End If
'                            If Seek01 <> SystemNumber(CheckStr(AdoRecordSet3.Fields(0).Value), 1) Then
'                                nowCnt = 1
'                                intLine = intLine + 600
'                                Seek01 = SystemNumber(CheckStr(AdoRecordSet3.Fields(0).Value), 1)
'                            End If
'                            Select Case nowCnt '(AdoRecordSet3.AbsolutePosition Mod intLineCnt)
'                            Case intLineCnt
'                                 Printer.CurrentX = Printer.ScaleWidth - ((1 / intLineCnt) * Printer.ScaleWidth)
'                                 Printer.CurrentY = intLine
'                                 Printer.Print CheckStr(AdoRecordSet3.Fields(0).Value)
'                                 intLine = intLine + 300
'                            Case Else
'                                 Printer.CurrentX = Printer.ScaleWidth * (((nowCnt Mod intLineCnt) - 1) / intLineCnt)
'                                 Printer.CurrentY = intLine
'                                 Printer.Print CheckStr(AdoRecordSet3.Fields(0).Value)
'                            End Select
'                            'add by nickc 2007/01/24
'                            If intLine > Printer.ScaleHeight - 1500 Then
'                                Printer.NewPage
'                                intLine = 300
'                                Printer.CurrentX = 0
'                                Printer.CurrentY = intLine
'                                Printer.Print "六個月本所期限案件明細："
'                                intLine = intLine + 600
'                            End If
'                            AdoRecordSet3.MoveNext
'                        Loop
''edit by nickc 2007/01/17 填請作單改有 6 個月期限案件才印
'                    'add by nickc 2007/01/02
''                    Else
''                        Printer.CurrentX = 0
''                        Printer.CurrentY = 2700
''                        Printer.Print "無六個月本所期限案件!!"
'                    End If
'                    Printer.Print
'                    Printer.Print "PS：1.若為修改名稱、地址、電話或傳真時，請將此報表及異動資料交給檔案室歸卷。"
'                    Printer.Print "　　2.若非聯絡資料的變動，則可忽略此報表。"
'                    Printer.EndDoc
'                End If
'            End If
'      End If
'      'end 2017/08/11
End Sub

'Added by Lydia 2023/01/03
Private Function ChkExistSpec(ByVal pTBL As String, ByVal pNo As String, ByVal pLen As Integer) As Boolean
Dim rsQD As New ADODB.Recordset
Dim strQ1 As String, intQ As Integer
   
   ChkExistSpec = False
   Select Case UCase(pTBL)
       Case "NPMEMO"
           strQ1 = "SELECT NM01 FROM NPMEMO WHERE NM04='" & Left(pNo, pLen) & "' AND NM05 IS NULL "
       Case "APPROVALMEMO2"
           strQ1 = "SELECT AM01 FROM APPROVALMEMO2 WHERE AM04='" & Left(pNo, pLen) & "' AND AM05 IS NULL "
       Case "INCOMMEMO"
           strQ1 = "SELECT IM01 FROM INCOMMEMO WHERE IM04='" & Left(pNo, pLen) & "' AND IM05 IS NULL "
       Case "DEBITNOTEPS"
           strQ1 = "SELECT DNPS01 FROM DEBITNOTEPS WHERE DNPS04='" & Left(pNo, pLen) & "' AND DNPS05 IS NULL "
       Case "FCPEMPBILL"
           strQ1 = "SELECT FEB01 FROM FCPEMPBILL WHERE FEB04='" & Left(pNo, pLen) & "' AND FEB05 IS NULL "
       Case "APPROVALPS"
           strQ1 = "SELECT APS01 FROM APPROVALPS WHERE APS04='" & Left(pNo, pLen) & "' AND APS05 IS NULL "
   End Select
   If strQ1 <> "" Then
       intQ = 1
       Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
       If intQ = 1 Then
          ChkExistSpec = True
       End If
       Set rsQD = Nothing
   End If
   
End Function


