VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_10 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人資料"
   ClientHeight    =   6190
   ClientLeft      =   280
   ClientTop       =   3540
   ClientWidth     =   9260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6190
   ScaleWidth      =   9260
   Begin VB.CommandButton CmdOk1 
      Caption         =   "被介紹者"
      Height          =   400
      Index           =   4
      Left            =   4072
      Style           =   1  '圖片外觀
      TabIndex        =   223
      Top             =   70
      Width           =   950
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "各項指示"
      Height          =   400
      Index           =   3
      Left            =   3090
      TabIndex        =   215
      Top             =   70
      Width           =   950
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "合約資料查詢"
      Height          =   400
      Index           =   5
      Left            =   5054
      TabIndex        =   214
      Top             =   70
      Width           =   1275
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "平台帳號"
      Height          =   400
      Index           =   2
      Left            =   6361
      TabIndex        =   185
      Top             =   70
      Width           =   950
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束"
      Height          =   400
      Index           =   1
      Left            =   8325
      TabIndex        =   21
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "回前畫面"
      Height          =   400
      Index           =   0
      Left            =   7343
      TabIndex        =   20
      Top             =   70
      Width           =   950
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5640
      Left            =   60
      TabIndex        =   0
      Top             =   525
      Width           =   9132
      _ExtentX        =   16104
      _ExtentY        =   9948
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   7
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm100101_10.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label30"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label29"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label27"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label16"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label18"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl1(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl1(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl1(41)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label7(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label46"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl1(47)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label47"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label48"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl1(48)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label50"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label79"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl1(66)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label69"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl1(67)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txt1(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txt1(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txt1(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txt1(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txt1(4)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txt1(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label1(4)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label1(2)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label1(5)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lblFA127"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lblXYS02"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lblXYS03"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "lblXYS02_N"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "聯絡資料"
      TabPicture(1)   =   "frm100101_10.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label24(2)"
      Tab(1).Control(1)=   "Label24(1)"
      Tab(1).Control(2)=   "Label24(0)"
      Tab(1).Control(3)=   "Label25"
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(5)=   "Label36"
      Tab(1).Control(6)=   "Label37"
      Tab(1).Control(7)=   "Label38"
      Tab(1).Control(8)=   "Label39"
      Tab(1).Control(9)=   "Label40"
      Tab(1).Control(10)=   "Label41"
      Tab(1).Control(11)=   "Label4"
      Tab(1).Control(12)=   "lbl1(4)"
      Tab(1).Control(13)=   "lbl1(5)"
      Tab(1).Control(14)=   "lbl1(6)"
      Tab(1).Control(15)=   "lbl1(8)"
      Tab(1).Control(16)=   "lbl1(9)"
      Tab(1).Control(17)=   "lbl1(10)"
      Tab(1).Control(18)=   "lbl1(11)"
      Tab(1).Control(19)=   "lbl1(12)"
      Tab(1).Control(20)=   "lbl1(13)"
      Tab(1).Control(21)=   "Label66"
      Tab(1).Control(22)=   "Label65"
      Tab(1).Control(23)=   "Label64"
      Tab(1).Control(24)=   "Label63"
      Tab(1).Control(25)=   "Label54"
      Tab(1).Control(26)=   "lbl1(7)"
      Tab(1).Control(27)=   "Label68"
      Tab(1).Control(28)=   "Label89"
      Tab(1).Control(29)=   "txt1(7)"
      Tab(1).Control(30)=   "txt1(8)"
      Tab(1).Control(31)=   "txt1(9)"
      Tab(1).Control(32)=   "txt1(10)"
      Tab(1).Control(33)=   "txt1(11)"
      Tab(1).Control(34)=   "txt1(15)"
      Tab(1).Control(35)=   "txt1(16)"
      Tab(1).Control(36)=   "txt1(17)"
      Tab(1).Control(37)=   "txt1(18)"
      Tab(1).Control(38)=   "txt1(19)"
      Tab(1).Control(39)=   "txt1(22)"
      Tab(1).ControlCount=   40
      TabCaption(2)   =   "專利"
      TabPicture(2)   =   "frm100101_10.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label93"
      Tab(2).Control(1)=   "Label91"
      Tab(2).Control(2)=   "Label90"
      Tab(2).Control(3)=   "Label60(1)"
      Tab(2).Control(4)=   "Label8(0)"
      Tab(2).Control(5)=   "Label9"
      Tab(2).Control(6)=   "Label10"
      Tab(2).Control(7)=   "Label11"
      Tab(2).Control(8)=   "Label14"
      Tab(2).Control(9)=   "Label15"
      Tab(2).Control(10)=   "Label19"
      Tab(2).Control(11)=   "Label20"
      Tab(2).Control(12)=   "Label21"
      Tab(2).Control(13)=   "Label22"
      Tab(2).Control(14)=   "Label31"
      Tab(2).Control(15)=   "Label52"
      Tab(2).Control(16)=   "Label60(0)"
      Tab(2).Control(17)=   "lbl1(16)"
      Tab(2).Control(18)=   "lbl1(17)"
      Tab(2).Control(19)=   "lbl1(18)"
      Tab(2).Control(20)=   "lbl1(30)"
      Tab(2).Control(21)=   "lbl1(49)"
      Tab(2).Control(22)=   "lbl1(19)"
      Tab(2).Control(23)=   "lbl1(20)"
      Tab(2).Control(24)=   "lbl1(21)"
      Tab(2).Control(25)=   "lbl1(29)"
      Tab(2).Control(26)=   "lbl1(31)"
      Tab(2).Control(27)=   "lblFA(95)"
      Tab(2).Control(28)=   "lblFA(96)"
      Tab(2).Control(29)=   "Label35(0)"
      Tab(2).Control(30)=   "Label34(0)"
      Tab(2).Control(31)=   "Label23"
      Tab(2).Control(32)=   "lbl1(22)"
      Tab(2).Control(33)=   "lbl1(32)"
      Tab(2).Control(34)=   "lbl1(33)"
      Tab(2).Control(35)=   "Label1(0)"
      Tab(2).Control(36)=   "Label42"
      Tab(2).Control(37)=   "lbl1(42)"
      Tab(2).Control(38)=   "lbl1(43)"
      Tab(2).Control(39)=   "Label62"
      Tab(2).Control(40)=   "Label12"
      Tab(2).Control(41)=   "Label53"
      Tab(2).Control(42)=   "lbl1(56)"
      Tab(2).Control(43)=   "Label55"
      Tab(2).Control(44)=   "lbl1(58)"
      Tab(2).Control(45)=   "lbl1(63)"
      Tab(2).Control(46)=   "lbl1(51)"
      Tab(2).Control(47)=   "Label80(30)"
      Tab(2).Control(48)=   "Label71"
      Tab(2).Control(49)=   "Label85"
      Tab(2).Control(50)=   "lbl1(70)"
      Tab(2).Control(51)=   "lblFA(124)"
      Tab(2).Control(52)=   "txt1(12)"
      Tab(2).Control(53)=   "txt1(13)"
      Tab(2).Control(54)=   "lblFA(128)"
      Tab(2).Control(55)=   "lblFA(136)"
      Tab(2).Control(56)=   "Combo3(0)"
      Tab(2).ControlCount=   57
      TabCaption(3)   =   "商標"
      TabPicture(3)   =   "frm100101_10.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label92"
      Tab(3).Control(1)=   "Label8(1)"
      Tab(3).Control(2)=   "lbl1(36)"
      Tab(3).Control(3)=   "lbl1(35)"
      Tab(3).Control(4)=   "lbl1(34)"
      Tab(3).Control(5)=   "Label35(1)"
      Tab(3).Control(6)=   "Label34(1)"
      Tab(3).Control(7)=   "Label8(2)"
      Tab(3).Control(8)=   "lbl1(61)"
      Tab(3).Control(9)=   "Label58"
      Tab(3).Control(10)=   "Label56"
      Tab(3).Control(11)=   "Label57"
      Tab(3).Control(12)=   "Label67"
      Tab(3).Control(13)=   "Label45"
      Tab(3).Control(14)=   "Label43"
      Tab(3).Control(15)=   "Label44"
      Tab(3).Control(16)=   "lbl1(57)"
      Tab(3).Control(17)=   "lbl1(59)"
      Tab(3).Control(18)=   "lbl1(45)"
      Tab(3).Control(19)=   "lbl1(46)"
      Tab(3).Control(20)=   "lbl1(44)"
      Tab(3).Control(21)=   "lbl1(64)"
      Tab(3).Control(22)=   "lbl1(60)"
      Tab(3).Control(23)=   "Label73"
      Tab(3).Control(24)=   "Label74"
      Tab(3).Control(25)=   "Label75"
      Tab(3).Control(26)=   "Label76"
      Tab(3).Control(27)=   "Label77"
      Tab(3).Control(28)=   "lbl1(54)"
      Tab(3).Control(29)=   "Label1(1)"
      Tab(3).Control(30)=   "Label78"
      Tab(3).Control(31)=   "lbl1(55)"
      Tab(3).Control(32)=   "lbl1(65)"
      Tab(3).Control(33)=   "lbl1(53)"
      Tab(3).Control(34)=   "lbl1(52)"
      Tab(3).Control(35)=   "Label80(29)"
      Tab(3).Control(36)=   "Label1(28)"
      Tab(3).Control(37)=   "lbl1(68)"
      Tab(3).Control(38)=   "Label81"
      Tab(3).Control(39)=   "lbl1(69)"
      Tab(3).Control(40)=   "Label82"
      Tab(3).Control(41)=   "lblFA120"
      Tab(3).Control(42)=   "Label83"
      Tab(3).Control(43)=   "Label80(0)"
      Tab(3).Control(44)=   "txt1(20)"
      Tab(3).Control(45)=   "txt1(21)"
      Tab(3).Control(46)=   "lblFA(129)"
      Tab(3).Control(47)=   "lbl1(71)"
      Tab(3).Control(48)=   "Label94"
      Tab(3).Control(49)=   "Label95"
      Tab(3).Control(50)=   "lbl1(72)"
      Tab(3).Control(51)=   "Label96"
      Tab(3).Control(52)=   "lbl1(73)"
      Tab(3).Control(53)=   "Combo3(1)"
      Tab(3).Control(54)=   "Combo5"
      Tab(3).ControlCount=   55
      TabCaption(4)   =   "其他"
      TabPicture(4)   =   "frm100101_10.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label72"
      Tab(4).Control(1)=   "Label1(3)"
      Tab(4).Control(2)=   "lblFA(102)"
      Tab(4).Control(3)=   "lbl1(40)"
      Tab(4).Control(4)=   "lbl1(39)"
      Tab(4).Control(5)=   "lbl1(38)"
      Tab(4).Control(6)=   "lbl1(37)"
      Tab(4).Control(7)=   "lbl1(28)"
      Tab(4).Control(8)=   "lbl1(27)"
      Tab(4).Control(9)=   "Label51"
      Tab(4).Control(10)=   "Label49"
      Tab(4).Control(11)=   "Label61"
      Tab(4).Control(12)=   "lbl1(62)"
      Tab(4).Control(13)=   "lbl1(23)"
      Tab(4).Control(14)=   "lbl1(24)"
      Tab(4).Control(15)=   "lbl1(25)"
      Tab(4).Control(16)=   "lbl1(26)"
      Tab(4).Control(17)=   "lbl1(50)"
      Tab(4).Control(18)=   "Label59"
      Tab(4).Control(19)=   "Label17"
      Tab(4).Control(20)=   "Label84"
      Tab(4).Control(21)=   "Label33"
      Tab(4).Control(22)=   "Label32"
      Tab(4).Control(23)=   "Label26"
      Tab(4).Control(24)=   "Label28"
      Tab(4).Control(25)=   "Label2(2)"
      Tab(4).Control(26)=   "lbl1(14)"
      Tab(4).Control(27)=   "lbl1(15)"
      Tab(4).Control(28)=   "Label6"
      Tab(4).Control(29)=   "Label70"
      Tab(4).Control(30)=   "lblFA(83)"
      Tab(4).Control(31)=   "lblFA(101)"
      Tab(4).Control(32)=   "Label86"
      Tab(4).Control(33)=   "Label87"
      Tab(4).Control(34)=   "lblFA(121)"
      Tab(4).Control(35)=   "lblFA(122)"
      Tab(4).Control(36)=   "Label88"
      Tab(4).Control(37)=   "lblFA(123)"
      Tab(4).Control(38)=   "txt1(14)"
      Tab(4).Control(39)=   "lstDeveloper"
      Tab(4).Control(40)=   "Frame1K"
      Tab(4).ControlCount=   41
      TabCaption(5)   =   "參考備註"
      TabPicture(5)   =   "frm100101_10.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txt1(6)"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame1K 
         Enabled         =   0   'False
         Height          =   280
         Left            =   -70800
         TabIndex        =   232
         Top             =   2040
         Width           =   4870
         Begin VB.CheckBox Chk1K 
            Caption         =   "月帳單"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   235
            Top             =   60
            Width           =   910
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "上傳平台"
            Height          =   180
            Index           =   1
            Left            =   2790
            TabIndex        =   234
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "帳單另寄"
            Height          =   180
            Index           =   0
            Left            =   1740
            TabIndex        =   233
            Top             =   60
            Width           =   1030
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "請款單寄送類型："
            Height          =   180
            Index           =   34
            Left            =   150
            TabIndex        =   236
            Top             =   60
            Width           =   1440
         End
      End
      Begin VB.ComboBox Combo5 
         Height          =   260
         ItemData        =   "frm100101_10.frx":00A8
         Left            =   -68250
         List            =   "frm100101_10.frx":00B8
         Style           =   2  '單純下拉式
         TabIndex        =   216
         Top             =   2210
         Width           =   2010
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   1
         ItemData        =   "frm100101_10.frx":00F4
         Left            =   -70800
         List            =   "frm100101_10.frx":0107
         Style           =   2  '單純下拉式
         TabIndex        =   188
         Top             =   2520
         Width           =   1560
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         Index           =   0
         ItemData        =   "frm100101_10.frx":013B
         Left            =   -70860
         List            =   "frm100101_10.frx":014E
         Style           =   2  '單純下拉式
         TabIndex        =   186
         Top             =   330
         Width           =   1470
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   73
         Left            =   -66840
         TabIndex        =   243
         Top             =   1250
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label96 
         AutoSize        =   -1  'True
         Caption         =   "延展折扣：        （％）"
         Height          =   180
         Left            =   -67740
         TabIndex        =   244
         Top             =   1280
         Width           =   1840
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   72
         Left            =   -68730
         TabIndex        =   241
         Top             =   1250
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label95 
         AutoSize        =   -1  'True
         Caption         =   "繳註冊費折扣：        （％）"
         Height          =   180
         Left            =   -69990
         TabIndex        =   242
         Top             =   1280
         Width           =   2200
      End
      Begin VB.Label Label94 
         AutoSize        =   -1  'True
         Caption         =   "商標全部折扣終止日："
         Height          =   180
         Left            =   -70740
         TabIndex        =   240
         Top             =   1580
         Width           =   1800
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   71
         Left            =   -68940
         TabIndex        =   239
         Top             =   1580
         Width           =   1400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2469;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFA 
         Height          =   252
         Index           =   136
         Left            =   -67236
         TabIndex        =   237
         Top             =   2880
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
      Begin MSForms.Label lblXYS02_N 
         Height          =   255
         Left            =   2040
         TabIndex        =   230
         Top             =   5160
         Width           =   2500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblXYS02_N"
         Size            =   "4410;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblXYS03 
         Height          =   600
         Left            =   5640
         TabIndex        =   229
         Top             =   4860
         Width           =   3300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblXYS03"
         Size            =   "5821;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblXYS02 
         Height          =   255
         Left            =   1200
         TabIndex        =   228
         Top             =   5160
         Width           =   800
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblXYS02"
         Size            =   "1411;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFA127 
         Height          =   255
         Left            =   1200
         TabIndex        =   227
         Top             =   4860
         Width           =   3000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFA127"
         Size            =   "5292;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "其他     說明："
         Height          =   495
         Index           =   5
         Left            =   4980
         TabIndex        =   226
         Top             =   4830
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "來所原因："
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   225
         Top             =   4860
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹者編號："
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   224
         Top             =   5145
         Width           =   1080
      End
      Begin MSForms.Label lblFA 
         Height          =   260
         Index           =   129
         Left            =   -72810
         TabIndex        =   221
         Top             =   5340
         Width           =   410
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   128
         Left            =   -68565
         TabIndex        =   219
         Top             =   1890
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.ListBox lstDeveloper 
         Height          =   315
         Left            =   -69360
         TabIndex        =   218
         Top             =   1590
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
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   22
         Left            =   -69300
         TabIndex        =   17
         Top             =   3360
         Width           =   3330
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5874;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   5085
         Index           =   6
         Left            =   -74910
         TabIndex        =   192
         Top             =   360
         Width           =   8940
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "15769;8969"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   21
         Left            =   -73020
         TabIndex        =   173
         Top             =   4170
         Width           =   1700
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "2990;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   530
         Index           =   20
         Left            =   -73660
         TabIndex        =   167
         Top             =   2780
         Width           =   7520
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13256;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   13
         Left            =   -68580
         TabIndex        =   19
         Top             =   2205
         Width           =   2325
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4101;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   525
         Index           =   12
         Left            =   -73650
         TabIndex        =   18
         Top             =   735
         Width           =   7605
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13414;917"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   19
         Left            =   -73560
         TabIndex        =   16
         Top             =   3360
         Width           =   3330
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5874;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   18
         Left            =   -69300
         TabIndex        =   15
         Top             =   2976
         Width           =   3330
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5874;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   17
         Left            =   -73560
         TabIndex        =   14
         Top             =   2976
         Width           =   3330
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5874;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   16
         Left            =   -69300
         TabIndex        =   13
         Top             =   2598
         Width           =   3330
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5874;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   15
         Left            =   -73560
         TabIndex        =   12
         Top             =   2598
         Width           =   3330
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5874;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   645
         Index           =   14
         Left            =   -73920
         TabIndex        =   100
         Top             =   2370
         Width           =   7620
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13441;1129"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   11
         Left            =   -70095
         TabIndex        =   10
         Top             =   738
         Width           =   4170
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "7355;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   10
         Left            =   -70095
         TabIndex        =   8
         Top             =   360
         Width           =   4170
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "7355;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   645
         Index           =   9
         Left            =   -74280
         TabIndex        =   11
         Top             =   1116
         Width           =   8355
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14737;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   8
         Left            =   -74280
         TabIndex        =   9
         Top             =   738
         Width           =   4170
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "7355;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   7
         Left            =   -74280
         TabIndex        =   7
         Top             =   360
         Width           =   4170
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "7355;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   450
         Index           =   5
         Left            =   1560
         TabIndex        =   6
         Top             =   3360
         Width           =   7470
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13176;794"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   645
         Index           =   4
         Left            =   1560
         TabIndex        =   5
         Top             =   2700
         Width           =   7470
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13176;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   450
         Index           =   2
         Left            =   1560
         TabIndex        =   3
         Top             =   1770
         Width           =   7470
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13176;794"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   645
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   1110
         Width           =   7470
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13176;1138"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   450
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Top             =   645
         Width           =   7470
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13176;794"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   450
         Index           =   3
         Left            =   1545
         TabIndex        =   4
         Top             =   2250
         Width           =   7470
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13176;794"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單之本所帳戶："
         Height          =   180
         Index           =   0
         Left            =   -70290
         TabIndex        =   217
         Top             =   2214
         Width           =   1980
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   124
         Left            =   -73290
         TabIndex        =   212
         Top             =   2250
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
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         Caption         =   "(財務CF)："
         Height          =   180
         Left            =   -70200
         TabIndex        =   211
         Top             =   3450
         Width           =   870
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   123
         Left            =   -71520
         TabIndex        =   210
         Top             =   3920
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "是否同意歐盟通用資料保護規範(GDPR)：        （W:待回覆 Y:同意  N:不同意）"
         Height          =   180
         Left            =   -74790
         TabIndex        =   209
         Top             =   3915
         Width           =   6105
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   122
         Left            =   -69360
         TabIndex        =   208
         Top             =   3360
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   121
         Left            =   -69360
         TabIndex        =   207
         Top             =   3090
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label87 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發促銷信：　　（Y：一定要寄，N：一定不要寄）"
         Height          =   180
         Left            =   -70800
         TabIndex        =   206
         Top             =   3360
         Width           =   4560
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發竹曆：　　    （Y：一定要寄，N：一定不要寄）"
         Height          =   180
         Left            =   -70800
         TabIndex        =   205
         Top             =   3090
         Width           =   4560
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   70
         Left            =   -66900
         TabIndex        =   204
         Top             =   1635
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
      Begin VB.Label Label85 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否電子送件 :                 (Y:是)"
         Height          =   180
         Left            =   -68730
         TabIndex        =   203
         Top             =   1635
         Width           =   2700
      End
      Begin VB.Label Label83 
         Caption         =   "管控智權人員："
         Height          =   260
         Left            =   -70290
         TabIndex        =   202
         Top             =   300
         Width           =   1310
      End
      Begin MSForms.Label lblFA120 
         Height          =   260
         Left            =   -69000
         TabIndex        =   201
         Top             =   300
         Width           =   1130
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1984;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "陸代定稿加註："
         Height          =   180
         Left            =   -74850
         TabIndex        =   200
         Top             =   4530
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   780
         Index           =   69
         Left            =   -73590
         TabIndex        =   199
         Top             =   4530
         Width           =   7490
         BackColor       =   16777215
         Caption         =   "lblFM2"
         Size            =   "13203;1376"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "(僅提醒)"
         Height          =   180
         Left            =   -74400
         TabIndex        =   198
         Top             =   3020
         Width           =   660
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "(僅提醒)"
         Height          =   180
         Left            =   -74445
         TabIndex        =   197
         Top             =   1020
         Width           =   660
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   101
         Left            =   -72765
         TabIndex        =   196
         Top             =   3645
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   83
         Left            =   -72765
         TabIndex        =   195
         Top             =   3360
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "財務處是否寄發催款單：              (1. 每月寄對帳單　2. 客戶要求不寄對帳單　3. 其他)"
         Height          =   180
         Left            =   -74790
         TabIndex        =   194
         Top             =   3645
         Width           =   6690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "財務處是否寄發FC收據：             (N：不寄)"
         Height          =   180
         Left            =   -74790
         TabIndex        =   193
         Top             =   3360
         Width           =   3375
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   68
         Left            =   -69030
         TabIndex        =   191
         Top             =   4215
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "不催延展：               (Y:不催)"
         Height          =   180
         Index           =   28
         Left            =   -70080
         TabIndex        =   190
         Top             =   4215
         Width           =   2220
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單列印幣別格式："
         Height          =   180
         Index           =   29
         Left            =   -72960
         TabIndex        =   189
         Top             =   2520
         Width           =   2160
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "專利請款單列印幣別格式："
         Height          =   180
         Index           =   30
         Left            =   -73020
         TabIndex        =   187
         Top             =   390
         Width           =   2160
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   67
         Left            =   5640
         TabIndex        =   183
         Top             =   4560
         Width           =   1125
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1984;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "帳單幣別："
         Height          =   180
         Left            =   4620
         TabIndex        =   184
         Top             =   4560
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   66
         Left            =   1950
         TabIndex        =   181
         Top             =   4560
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
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發專利雙週報：           (N:不寄)"
         Height          =   180
         Left            =   120
         TabIndex        =   182
         Top             =   4560
         Width           =   2940
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   52
         Left            =   -67260
         TabIndex        =   168
         Top             =   2520
         Width           =   380
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "661;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   53
         Left            =   -73590
         TabIndex        =   169
         Top             =   2520
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
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   65
         Left            =   -73270
         TabIndex        =   180
         Top             =   3890
         Width           =   6000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10583;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   55
         Left            =   -72930
         TabIndex        =   179
         Top             =   3600
         Width           =   6000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10583;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "延展D/N列印對象："
         Height          =   180
         Left            =   -74850
         TabIndex        =   178
         Top             =   3890
         Width           =   1550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N固定列印對象："
         Height          =   180
         Index           =   1
         Left            =   -74850
         TabIndex        =   177
         Top             =   3620
         Width           =   1910
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   54
         Left            =   -73230
         TabIndex        =   176
         Top             =   3330
         Width           =   6000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10583;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "商標固定請款對象："
         Height          =   180
         Left            =   -74850
         TabIndex        =   175
         Top             =   3350
         Width           =   1620
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "代理人商標財務編號："
         Height          =   180
         Left            =   -74850
         TabIndex        =   174
         Top             =   4220
         Width           =   1800
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "商標請款幣別："
         Height          =   180
         Left            =   -74850
         TabIndex        =   172
         Top             =   2520
         Width           =   1260
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N是否印申請人：         （Y：印）"
         Height          =   180
         Left            =   -69210
         TabIndex        =   171
         Top             =   2520
         Width           =   3200
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "商標D/N備註："
         Height          =   180
         Left            =   -74850
         TabIndex        =   170
         Top             =   2790
         Width           =   1190
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   60
         Left            =   -73230
         TabIndex        =   119
         Top             =   1910
         Width           =   410
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
         Height          =   260
         Index           =   64
         Left            =   -72900
         TabIndex        =   118
         Top             =   2210
         Width           =   410
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
         Height          =   260
         Index           =   44
         Left            =   -73560
         TabIndex        =   133
         Top             =   1250
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   46
         Left            =   -73050
         TabIndex        =   131
         Top             =   1580
         Width           =   1400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2461;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   45
         Left            =   -70950
         TabIndex        =   132
         Top             =   1250
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   59
         Left            =   -67620
         TabIndex        =   124
         Top             =   910
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   57
         Left            =   -69870
         TabIndex        =   121
         Top             =   910
         Width           =   440
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "776;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   -73305
         TabIndex        =   128
         Top             =   390
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   -73860
         TabIndex        =   127
         Top             =   675
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "代理人狀態："
         Height          =   180
         Index           =   2
         Left            =   -74790
         TabIndex        =   166
         Top             =   2100
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   180
         Left            =   -74790
         TabIndex        =   165
         Top             =   1245
         Width           =   1080
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人："
         Height          =   180
         Left            =   -74790
         TabIndex        =   164
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "實體副本收受人："
         Height          =   180
         Left            =   -74790
         TabIndex        =   163
         Top             =   1530
         Width           =   1440
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "實體副本聯絡人："
         Height          =   180
         Left            =   -74790
         TabIndex        =   162
         Top             =   1815
         Width           =   1440
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文：           （1.中文  2.英文  3.日文）"
         Height          =   180
         Left            =   -74790
         TabIndex        =   161
         Top             =   675
         Width           =   3420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "是否寄台一雜誌：             （N：不寄）"
         Height          =   180
         Left            =   -74790
         TabIndex        =   160
         Top             =   390
         Width           =   3045
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "帳單備註："
         Height          =   180
         Left            =   -74790
         TabIndex        =   159
         Top             =   2370
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   50
         Left            =   -73680
         TabIndex        =   158
         Top             =   2100
         Width           =   3140
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5539;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   26
         Left            =   -73305
         TabIndex        =   157
         Top             =   1815
         Width           =   2775
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4895;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   25
         Left            =   -73305
         TabIndex        =   156
         Top             =   1530
         Width           =   2775
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4895;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   24
         Left            =   -73635
         TabIndex        =   155
         Top             =   1245
         Width           =   6045
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10663;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   -73635
         TabIndex        =   154
         Top             =   960
         Width           =   6045
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10663;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   62
         Left            =   -69030
         TabIndex        =   126
         Top             =   390
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   51
         Left            =   -68805
         TabIndex        =   148
         Top             =   4590
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
         Index           =   63
         Left            =   -68490
         TabIndex        =   146
         Top             =   4890
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
         Index           =   58
         Left            =   -73395
         TabIndex        =   153
         Top             =   4890
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "專利請款單份數："
         Height          =   180
         Left            =   -74865
         TabIndex        =   152
         Top             =   4890
         Width           =   1440
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   56
         Left            =   -73590
         TabIndex        =   151
         Top             =   4590
         Width           =   495
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "專利以 EMail 通知：         （Y：是  D：僅D/N）"
         Height          =   180
         Left            =   -70380
         TabIndex        =   150
         Top             =   4590
         Width           =   3690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "專利定稿份數："
         Height          =   180
         Left            =   -74865
         TabIndex        =   149
         Top             =   4590
         Width           =   1260
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "專利 Email 同時寄紙本：        （Y：是）"
         Height          =   180
         Left            =   -70380
         TabIndex        =   147
         Top             =   4890
         Width           =   3135
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   43
         Left            =   -73290
         TabIndex        =   145
         Top             =   4305
         Width           =   6000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10583;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   42
         Left            =   -72945
         TabIndex        =   144
         Top             =   4020
         Width           =   6000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10583;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "年費D/N列印對象："
         Height          =   180
         Left            =   -74865
         TabIndex        =   143
         Top             =   4305
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N固定列印對象："
         Height          =   180
         Index           =   0
         Left            =   -74865
         TabIndex        =   142
         Top             =   4020
         Width           =   1905
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   33
         Left            =   -73710
         TabIndex        =   141
         Top             =   3750
         Width           =   6000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10583;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   32
         Left            =   -73530
         TabIndex        =   140
         Top             =   3450
         Width           =   6000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10583;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   22
         Left            =   -73185
         TabIndex        =   139
         Top             =   3180
         Width           =   6000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10583;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "專利固定請款對象："
         Height          =   180
         Left            =   -74865
         TabIndex        =   138
         Top             =   3180
         Width           =   1620
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "年費代理人："
         Height          =   180
         Index           =   0
         Left            =   -74865
         TabIndex        =   137
         Top             =   3750
         Width           =   1080
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "年費請款對象："
         Height          =   180
         Index           =   0
         Left            =   -74865
         TabIndex        =   136
         Top             =   3450
         Width           =   1260
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "商標申請/翻譯折扣：        （％）"
         Height          =   180
         Left            =   -72630
         TabIndex        =   135
         Top             =   1280
         Width           =   2550
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "商標全部折扣：        （％）"
         Height          =   180
         Left            =   -74850
         TabIndex        =   134
         Top             =   1280
         Width           =   2200
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "商標全部折扣起始日："
         Height          =   180
         Left            =   -74850
         TabIndex        =   130
         Top             =   1580
         Width           =   1800
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   96
         Left            =   -70665
         TabIndex        =   77
         Top             =   2565
         Width           =   450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "794;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   95
         Left            =   -73590
         TabIndex        =   76
         Top             =   2565
         Width           =   450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "794;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   31
         Left            =   -67056
         TabIndex        =   78
         Top             =   2565
         Width           =   456
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "804;444"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   252
         Index           =   29
         Left            =   -70068
         TabIndex        =   74
         Top             =   2880
         Width           =   1128
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1984;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   -73590
         TabIndex        =   73
         Top             =   2880
         Width           =   450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "右-lblFM2"
         Size            =   "794;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   20
         Left            =   -73290
         TabIndex        =   72
         Top             =   1942
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
         Index           =   19
         Left            =   -73290
         TabIndex        =   71
         Top             =   1635
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
         Index           =   49
         Left            =   -66900
         TabIndex        =   87
         Top             =   1320
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
         Index           =   30
         Left            =   -73800
         TabIndex        =   75
         Top             =   1320
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   18
         Left            =   -70470
         TabIndex        =   70
         Top             =   1320
         Width           =   375
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "661;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   17
         Left            =   -67350
         TabIndex        =   69
         Top             =   390
         Width           =   500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "882;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   -73620
         TabIndex        =   68
         Top             =   390
         Width           =   500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "882;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報 ：            （N：不寄）"
         Height          =   180
         Left            =   -70350
         TabIndex        =   129
         Top             =   390
         Width           =   2865
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "商標 Email 同時寄紙本：        （Y：是）"
         Height          =   180
         Left            =   -74850
         TabIndex        =   125
         Top             =   2210
         Width           =   3140
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "商標請款單份數："
         Height          =   180
         Left            =   -69120
         TabIndex        =   123
         Top             =   940
         Width           =   1440
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "商標定稿份數："
         Height          =   180
         Left            =   -71220
         TabIndex        =   122
         Top             =   940
         Width           =   1260
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "商標以 EMail 通知：         （Y：是  D：僅D/N）"
         Height          =   180
         Left            =   -74850
         TabIndex        =   120
         Top             =   1910
         Width           =   3690
      End
      Begin VB.Label Label49 
         Caption         =   "Create ID："
         Height          =   180
         Left            =   -74820
         TabIndex        =   117
         Top             =   4590
         Width           =   945
      End
      Begin VB.Label Label51 
         Caption         =   "Update ID："
         Height          =   180
         Left            =   -74820
         TabIndex        =   116
         Top             =   4860
         Width           =   945
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   27
         Left            =   -73860
         TabIndex        =   115
         Top             =   4590
         Width           =   1995
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3528;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   28
         Left            =   -73860
         TabIndex        =   114
         Top             =   4860
         Width           =   2000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3528;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   37
         Left            =   -71805
         TabIndex        =   113
         Top             =   4590
         Width           =   1995
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3519;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   38
         Left            =   -69060
         TabIndex        =   112
         Top             =   4590
         Width           =   1995
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3519;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   39
         Left            =   -71805
         TabIndex        =   111
         Top             =   4860
         Width           =   2000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3528;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   40
         Left            =   -69045
         TabIndex        =   110
         Top             =   4860
         Width           =   2000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3528;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日)："
         Height          =   180
         Left            =   -74865
         TabIndex        =   109
         Top             =   5280
         Width           =   1380
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   -73455
         TabIndex        =   108
         Top             =   5280
         Width           =   6405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11289;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFA 
         Height          =   255
         Index           =   102
         Left            =   -72765
         TabIndex        =   106
         Top             =   3090
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "開發人員："
         Height          =   180
         Index           =   3
         Left            =   -70350
         TabIndex        =   105
         Top             =   1590
         Width           =   900
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "專利領證折扣：         （％）"
         Height          =   180
         Index           =   0
         Left            =   -74865
         TabIndex        =   103
         Top             =   2565
         Width           =   2205
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   61
         Left            =   -73200
         TabIndex        =   102
         Top             =   910
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "FCT註冊自動代繳 ：         （Y：單筆不跑）"
         Height          =   180
         Index           =   2
         Left            =   -74850
         TabIndex        =   101
         Top             =   940
         Width           =   3410
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "延展代理人："
         Height          =   180
         Index           =   1
         Left            =   -74850
         TabIndex        =   98
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "延展請款對象："
         Height          =   180
         Index           =   1
         Left            =   -74850
         TabIndex        =   97
         Top             =   300
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   34
         Left            =   -73500
         TabIndex        =   96
         Top             =   300
         Width           =   2990
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5265;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   35
         Left            =   -73500
         TabIndex        =   95
         Top             =   600
         Width           =   6020
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10619;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   36
         Left            =   -73320
         TabIndex        =   94
         Top             =   760
         Visible         =   0   'False
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail(代表)："
         Height          =   180
         Left            =   -74865
         TabIndex        =   93
         Top             =   2688
         Width           =   1140
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "(財務)："
         Height          =   180
         Left            =   -69990
         TabIndex        =   92
         Top             =   2688
         Width           =   660
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "(其他1)："
         Height          =   180
         Left            =   -74370
         TabIndex        =   91
         Top             =   3066
         Width           =   750
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "(其他2)："
         Height          =   180
         Left            =   -70080
         TabIndex        =   90
         Top             =   3066
         Width           =   750
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "(其他3)："
         Height          =   180
         Left            =   -74370
         TabIndex        =   89
         Top             =   3450
         Width           =   750
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "FCP是否核對已准專利:      （N:否）"
         Height          =   180
         Left            =   -68730
         TabIndex        =   88
         Top             =   1320
         Width           =   2760
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Left            =   8520
         TabIndex        =   86
         Top             =   3930
         Width           =   465
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   48
         Left            =   8115
         TabIndex        =   85
         Top             =   3930
         Width           =   270
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "476;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "呆帳紀錄："
         Height          =   180
         Left            =   7200
         TabIndex        =   84
         Top             =   3930
         Width           =   900
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "(A:律師事務所 B:公司直接委辦 C:其他)"
         Height          =   180
         Left            =   5610
         TabIndex        =   83
         Top             =   4230
         Width           =   3045
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   47
         Left            =   5235
         TabIndex        =   82
         Top             =   4230
         Width           =   240
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "423;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "性質："
         Height          =   180
         Left            =   4620
         TabIndex        =   81
         Top             =   4230
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "代理人編號："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   41
         Left            =   1560
         TabIndex        =   79
         Top             =   360
         Width           =   1245
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2196;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   -73650
         TabIndex        =   67
         Top             =   5020
         Width           =   6405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11298;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   -73650
         TabIndex        =   66
         Top             =   4763
         Width           =   6405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11298;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   -73650
         TabIndex        =   65
         Top             =   4506
         Width           =   6405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11298;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   -73650
         TabIndex        =   64
         Top             =   4249
         Width           =   6405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11298;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   -73650
         TabIndex        =   63
         Top             =   3992
         Width           =   6405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11298;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   -73650
         TabIndex        =   62
         Top             =   3735
         Width           =   6405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11289;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   6
         Left            =   -73080
         TabIndex        =   61
         Top             =   2325
         Width           =   6195
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10927;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   -73080
         TabIndex        =   60
         Top             =   2052
         Width           =   6200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10936;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   -73080
         TabIndex        =   59
         Top             =   1779
         Width           =   6200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10936;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   58
         Top             =   4230
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
         Height          =   255
         Index           =   2
         Left            =   5580
         TabIndex        =   57
         Top             =   3930
         Width           =   1215
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2143;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   56
         Top             =   3930
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
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   55
         Top             =   360
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FAX："
         Height          =   180
         Left            =   -74865
         TabIndex        =   54
         Top             =   738
         Width           =   510
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人２(日)："
         Height          =   180
         Left            =   -74865
         TabIndex        =   53
         Top             =   4950
         Width           =   1200
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人２(英)："
         Height          =   180
         Left            =   -74865
         TabIndex        =   52
         Top             =   4710
         Width           =   1200
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人２(中)："
         Height          =   180
         Left            =   -74865
         TabIndex        =   51
         Top             =   4455
         Width           =   1200
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人１(日)："
         Height          =   180
         Left            =   -74865
         TabIndex        =   50
         Top             =   4215
         Width           =   1200
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人１(英)："
         Height          =   180
         Left            =   -74865
         TabIndex        =   49
         Top             =   3975
         Width           =   1200
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人１(中)："
         Height          =   180
         Left            =   -74865
         TabIndex        =   48
         Top             =   3735
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TEL："
         Height          =   180
         Left            =   -74865
         TabIndex        =   47
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "POB："
         Height          =   180
         Left            =   -74865
         TabIndex        =   46
         Top             =   1116
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "開發日期："
         Height          =   180
         Left            =   4620
         TabIndex        =   45
         Top             =   3930
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代理人地址(中)："
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   2310
         Width           =   1380
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "代理人地址(英)："
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   2730
         Width           =   1380
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代理人地址(日)："
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   3390
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "代理人國籍："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   3930
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   0
         Left            =   4485
         TabIndex        =   40
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "代理人名稱(中)："
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   645
         Width           =   1380
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "代理人名稱(英)："
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   1140
         Width           =   1380
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "代理人名稱(日)："
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "地址國籍:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   4230
         Width           =   765
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人中文名稱:"
         Height          =   180
         Index           =   0
         Left            =   -74865
         TabIndex        =   35
         Top             =   1779
         Width           =   1665
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人英文名稱:"
         Height          =   180
         Index           =   1
         Left            =   -74865
         TabIndex        =   34
         Top             =   2052
         Width           =   1665
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人日文名稱:"
         Height          =   180
         Index           =   2
         Left            =   -74865
         TabIndex        =   33
         Top             =   2325
         Width           =   1665
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "收款後辦案:        （ Y:先收）"
         Height          =   180
         Left            =   -74865
         TabIndex        =   32
         Top             =   1320
         Width           =   2235
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "代理人專利財務編號："
         Height          =   180
         Left            =   -70410
         TabIndex        =   31
         Top             =   2235
         Width           =   1800
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "專利全部折扣起始日："
         Height          =   180
         Left            =   -71940
         TabIndex        =   30
         Top             =   2880
         Width           =   1800
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "專利申請/翻譯折扣：         （％）"
         Height          =   180
         Left            =   -68724
         TabIndex        =   29
         Top             =   2565
         Width           =   2616
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "專利全部折扣：         （％）"
         Height          =   180
         Left            =   -74865
         TabIndex        =   28
         Top             =   2880
         Width           =   2205
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N備註："
         Height          =   180
         Left            =   -74865
         TabIndex        =   27
         Top             =   750
         Width           =   1185
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "專利D/N是否印申請人：            （Y：印）"
         Height          =   180
         Left            =   -69285
         TabIndex        =   26
         Top             =   390
         Width           =   3285
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "專利請款幣別："
         Height          =   180
         Left            =   -74865
         TabIndex        =   25
         Top             =   390
         Width           =   1260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "FCP領證自動代繳：             （Y：自動代繳）"
         Height          =   180
         Left            =   -74865
         TabIndex        =   24
         Top             =   1942
         Width           =   3525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費自動代繳：             （Y：自動代繳 / N：寄證書後年費不續辦)"
         Height          =   180
         Left            =   -74865
         TabIndex        =   23
         Top             =   1635
         Width           =   5460
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "FCP年費通知函單筆不跑:       （Y:單筆不跑）"
         Height          =   180
         Index           =   0
         Left            =   -72480
         TabIndex        =   22
         Top             =   1320
         Width           =   3525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "延展單筆不跑：              （Y：單筆不跑）"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   99
         Top             =   760
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "專利年費折扣：         （％）"
         Height          =   180
         Index           =   1
         Left            =   -71940
         TabIndex        =   104
         Top             =   2565
         Width           =   2205
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "是否用LEDES電子帳單：              (Y：是)"
         Height          =   180
         Left            =   -74790
         TabIndex        =   107
         Top             =   3090
         Width           =   3195
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         Caption         =   "FCP實審自動代繳：             （Y：自動代繳）"
         Height          =   180
         Left            =   -74865
         TabIndex        =   213
         Top             =   2250
         Width           =   3525
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "台灣案專利證書形式：             (1:電子 2:紙本)"
         Height          =   180
         Left            =   -70410
         TabIndex        =   220
         Top             =   1920
         Width           =   3540
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "台灣案商標註冊證形式：             (1:電子 2:紙本)"
         Height          =   180
         Left            =   -74850
         TabIndex        =   222
         Top             =   5370
         Width           =   3720
      End
      Begin VB.Label Label93 
         AutoSize        =   -1  'True
         Caption         =   "專利不得請雜費 :              (Y:是)"
         Height          =   180
         Left            =   -68748
         TabIndex        =   238
         Top             =   2880
         Width           =   2496
      End
   End
   Begin VB.Label SpecCU 
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   30
      TabIndex        =   231
      Top             =   60
      Width           =   3465
   End
End
Attribute VB_Name = "frm100101_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/13 改成Form2.0 ; lbl1(index)、txt1(index)、lblFA(index)、lstDeveloper、lblFA120
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit

'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim Lbl As Object 'Add by Amy 2023/07/25
Public m_pub_QL05 As String 'Add By Sindy 2025/8/28 只記錄於此Form


'92.04.16 nick
Public Sub PubShowNextData()
Dim stName As String 'Add by Amy 2022/12/06

Select Case cmdState
   Case 0
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Case 1
      fnCloseAllFrm100
   'Add By Sindy 2012/9/20
   Case 2 '平台帳號
      cmdState = -1
      Me.Enabled = False
      If fnSaveParentForm(Me) = False Then
         Me.Enabled = True
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      frm100101_26.Show
      frm100101_26.Tag = Trim(Lbl1(41))
      'frm100101_26.Tag = "X23450000"
      frm100101_26.StrMenu
      Screen.MousePointer = vbDefault
      Me.Enabled = True
   'Added by Lydia 2016/11/11
   Case 3 '各項指示
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
      frm12040159.SetParent "Q", Trim(Lbl1(41)), Me
      frm12040159.Show
      Screen.MousePointer = vbDefault
      Me.Enabled = True
   'Add by Amy 2019/05/08 合約資料查詢
   Case 5
    cmdState = -1
    Me.Enabled = False
    If fnSaveParentForm(Me) = False Then
       Me.Enabled = True
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    frm100101_N.lblCT01 = Left(Lbl1(41).Caption, 8)
    frm100101_N.Show
    Call frm100101_N.QueryData(Left(Lbl1(41).Caption, 8))
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub
   Case 4 'Add by Amy 2022/12/06 被介紹者
    If CmdOk1(4).BackColor <> &HFFFF80 Then
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
    frm050705_1.txtNo = Left(Lbl1(0), 8)
    frm050705_1.Lbl1(0) = Mid(Lbl1(1), 1, InStr(Lbl1(1), " "))
    frm050705_1.Lbl1(1) = Mid(Lbl1(1), InStr(Lbl1(1), " ") + 1)
    frm050705_1.Lbl1(3) = stName
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
'     Me.Hide
'Case 1
'   bolToEndByNick = True
'     Unload Me
'     Exit Sub
'Case Else
'End Select
End Sub

Sub StrMenu()
Dim strSql  As String, i As Integer
Dim Str01 As String    ', str02 As String, str03 As String, str04 As String
'Dim strArr(0 To 74) As String, StrOk(45) As String, StrOkTxt(13) As String
'edit by nickc 2005/10/24
'Dim strArr(75) As String, StrOk(46) As String, StrOkTxt(13) As String
'Dim strArr(76) As String, StrOk(47) As String, StrOkTxt(13) As String
'Modify by Morgan 2007/10/29
'Dim strArr(77) As String, StrOk(48) As String, StrOkTxt(13) As String
'2007/11/1 modify by sonia 加代理人狀態
'Dim strArr() As String, StrOk(49) As String, StrOkTxt(13) As String
'Modify by Morgan 2008/1/17
'Dim strArr() As String, StrOk(50) As String, StrOkTxt(13) As String
'Modify by Morgan 2008/3/13
'Dim strArr() As String, StrOk(55) As String, StrOkTxt(13) As String
'2008/12/9 modify by sonia
'Dim strArr() As String, StrOk(60) As String, StrOkTxt(14) As String
'Modify by Morgan 2009/10/16
'Dim strArr() As String, StrOk(62) As String, StrOkTxt(14) As String
'2010/9/29 STROKTXT(14)->(19)
'Modified by Lydia 2017/11/30
'Dim strArr() As String, StrOk(69) As String, StrOkTxt(21) As String
'Modified by Lydia 2018/07/20
'Dim strArr() As String, StrOk(70) As String, StrOkTxt(21) As String
Dim strArr() As String, StrOk(73) As String, StrOkTxt(22) As String
ReDim strArr(TF_FA) As String
Dim strTp(2) As String 'Add by Amy 2022/12/07
'Add by Morgan 2008/11/13
Dim oLbl As Object
'end 2007/10/29
Dim arrID 'Add By Sindy 2025/1/7

Str01 = Me.Tag

'Added by Lydia 2021/12/13
For Each oLbl In Lbl1
   oLbl.BackColor = &H8000000F
Next
LblFA120.BackColor = &H8000000F
'end 2021/12/13

'add by sonia 2024/8/13 增加檢查有無國外代理人查詢權限
If CheckUse("frm100114_1", strExec, False) Then
Else
   i = MsgBox("您沒有查詢代理人資料的權限！", , "沒權限")
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
pub_QL05 = m_pub_QL05 & ";代理人編號：" & IIf(Str01 = "0", " ", Str01) & "(基本資料)" 'Add By Sindy 2025/8/13

'欲搜尋的SQL字串
If Len(Str01) = 9 Then
    strSql = "SELECT * FROM FAGENT WHERE FA01='" & Left(Str01, 8) & "' AND FA02='" & Right(Str01, 1) & "'"
Else
    strSql = "SELECT * FROM FAGENT WHERE FA01='" & Str01 & "' AND FA02='0'"
End If
CheckOC

adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/13
   'For i = 0 To (74 - 1)
   'exit by  nickc  2005/10/24
   'For i = 0 To 74
   'Modify by Morgan 2007/10/29
   'For i = 0 To 76
   For i = 0 To TF_FA - 1
      Select Case i
      'Modify By Sindy 2013/1/30 +115,116
      'Modified by Lydia 2018/
      Case 10, 24, 25, 26, 46, 47, 49, 50, 72, 73, 74, 115 '數值
           If IsNull(adoRecordset.Fields(i)) Then
               strArr(i + 1) = ""
           Else
               strArr(i + 1) = str(adoRecordset.Fields(i))
           End If
      Case Else '文字
           If IsNull(adoRecordset.Fields(i)) Then
                strArr(i + 1) = ""
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
'For i = 0 To 68
'For i = 1 To 74
For i = 1 To UBound(strArr)
    Select Case i
    Case 1
         StrOk(41) = strArr(i) & strArr(2)
    Case 3
         StrOk(0) = strArr(i)
    Case 4
         StrOkTxt(0) = strArr(i)
    Case 5
         'Modify by Amy 2017/07/11 +IIF 避免太多個換行,游標跳入TextBox會跑到最後,不到文字
         StrOkTxt(1) = strArr(i) & IIf(Len(Trim(strArr(63))) = 0, "", Chr(13) & Chr(10) & strArr(63)) & IIf(Len(Trim(strArr(64))) = 0, "", Chr(13) & Chr(10) & strArr(64)) & _
                    IIf(Len(Trim(strArr(65))) = 0, "", Chr(13) & Chr(10) & strArr(65))
    Case 6
         StrOkTxt(2) = strArr(i)
    Case 17
         StrOkTxt(3) = strArr(i)
    Case 18
        'Modify by Amy 2017/07/11 +IIF 避免太多個換行,游標跳入TextBox會跑到最後,不到文字
         StrOkTxt(4) = strArr(i) & IIf(Len(Trim(strArr(19))) = 0, "", Chr(13) & Chr(10) & strArr(19)) & IIf(Len(Trim(strArr(20))) = 0, "", Chr(13) & Chr(10) & strArr(20)) & _
                    IIf(Len(Trim(strArr(21))) = 0, "", Chr(13) & Chr(10) & strArr(21)) & IIf(Len(Trim(strArr(22))) = 0, "", Chr(13) & Chr(10) & strArr(22)) & _
                    IIf(Len(Trim(strArr(70))) = 0, "", Chr(13) & Chr(10) & strArr(70))
    Case 23
         StrOkTxt(5) = strArr(i)
    Case 10
        'Modify By Cheng 2002/12/19
'         strSQL = "SELECT NA03 FROM NATION WHERE NA01='" & StrArr(1) & "'"
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               StrOk(1) = strArr(i) & ""
            Else
               StrOk(1) = strArr(i) & "  " & adoRecordset.Fields(0)
            End If
            'Add by Morgan 2004/1/19
            Lbl1(1).ForeColor = vbBlack
         Else
            StrOk(1) = ""
            'Add by Morgan 2004/1/19
            Lbl1(1).ForeColor = vbRed
            StrOk(1) = strArr(i)
         End If
         CheckOC
    Case 11
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(2) = ""
         Else
             StrOk(2) = ChangeWStringToTString(strArr(i))
         End If
    Case 55
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             If IsNull(adoRecordset.Fields(0)) Then
                 StrOk(3) = strArr(i) & ""
             Else
                 StrOk(3) = strArr(i) & "  " & adoRecordset.Fields(0)
             End If
             'Add by Morgan 2004/1/19
             Lbl1(3).ForeColor = vbBlack
         Else
             StrOk(3) = ""
             'Add by Morgan 2004/1/19
             Lbl1(3).ForeColor = vbRed
             StrOk(3) = strArr(i)
         End If
         CheckOC
    Case 29
         StrOkTxt(6) = strArr(i)
    'Add by Morgan 2008/6/3
    Case 92
         StrOkTxt(14) = strArr(i)
    Case 12
         StrOkTxt(7) = strArr(i)
    Case 14
         StrOkTxt(8) = strArr(i)
    Case 32
         StrOkTxt(9) = strArr(i) & Chr(13) & Chr(10) & strArr(33) & Chr(13) & Chr(10) & strArr(34) & Chr(13) & Chr(10) & strArr(35) & Chr(13) & Chr(10) & strArr(36)
    Case 13
         StrOkTxt(10) = strArr(i)
    Case 15
         StrOkTxt(11) = strArr(i)
    Case 56
         StrOk(4) = strArr(i)
    Case 57
         StrOk(5) = strArr(i)
    Case 58
         StrOk(6) = strArr(i)
    Case 16
         '2010/9/29 MODIFY BY SONIA 改為TXT欄使用者可複製
         'StrOk(7) = strArr(i)
         StrOkTxt(15) = strArr(i)
    Case 7
         StrOk(8) = strArr(i)
    Case 8
         StrOk(9) = strArr(i)
    Case 9
         StrOk(10) = strArr(i)
    Case 52
         StrOk(11) = strArr(i)
    Case 53
         StrOk(12) = strArr(i)
    Case 54
         StrOk(13) = strArr(i)
    Case 31
         StrOk(14) = strArr(i)
    Case 24
         StrOk(15) = strArr(i)
    Case 43
         StrOk(16) = strArr(i)
    Case 44
         StrOk(17) = strArr(i)
    Case 45
         StrOkTxt(12) = strArr(i)
    Case 40
         StrOk(18) = strArr(i)
    Case 41
         StrOk(19) = strArr(i)
    Case 42
         StrOk(20) = strArr(i)
    Case 25
         StrOk(21) = strArr(i)
    Case 30
         If Left$(strArr(i), 1) = "X" Then
            StrOk(22) = GetCustomerName(strArr(i), 0)
         Else
            StrOk(22) = GetFAgentName(strArr(i))
         End If
          If StrOk(22) <> Empty Then
            StrOk(22) = strArr(i) & "  " & StrOk(22)
            Lbl1(22).ForeColor = vbBlack
         Else
            StrOk(22) = ""
            Lbl1(22).ForeColor = vbRed
            StrOk(22) = strArr(i)
         End If
         CheckOC
   Case 38
        If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
                StrOk(23) = GetCustomerName(strArr(i), 0)
         Else
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
            StrOk(23) = GetFAgentName(strArr(i))
         End If
'edit by nickc 2007/08/24 改成同 fa 維護
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            StrOk(23) = strArr(i) & "   " & adoRecordset.Fields(0)
         If StrOk(23) <> Empty Then
            StrOk(23) = strArr(i) & "  " & StrOk(23)
            'Add by Morgan 2004/1/19
            Lbl1(23).ForeColor = vbBlack
         Else
            StrOk(23) = strArr(i) & ""
            'Add by Morgan 2004/1/19
            Lbl1(23).ForeColor = vbRed
         End If
         CheckOC
    Case 37
         StrOk(24) = strArr(i)
    Case 59
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
            StrOk(25) = GetCustomerName(strArr(i), 0)
         Else
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
                StrOk(25) = GetFAgentName(strArr(i))
         End If
'edit by nickc 2007/08/24 改成同 fa 維護
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            StrOk(25) = strArr(i) & "   " & adoRecordset.Fields(0)
         If StrOk(25) <> Empty Then
            StrOk(25) = strArr(i) & "  " & StrOk(25)
            'Add by Morgan 2004/1/19
            Lbl1(25).ForeColor = vbBlack
         Else
            StrOk(25) = strArr(i) & ""
            'Add by Morgan 2004/1/19
            Lbl1(25).ForeColor = vbRed
         End If
         CheckOC
    Case 60
         StrOk(26) = strArr(i)
    Case 46
         StrOk(27) = GetPrjSalesNM(strArr(i))
    Case 49
         StrOk(28) = GetPrjSalesNM(strArr(i))
    Case 27
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(29) = ""
         Else
             StrOk(29) = ChangeWStringToTString(strArr(i))
         End If
    Case 28
         StrOkTxt(13) = strArr(i)
    Case 39
         StrOk(30) = strArr(i)
    Case 26
         StrOk(31) = strArr(i)
    Case 62
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                 strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                 strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
            StrOk(32) = GetCustomerName(strArr(i), 0)
         Else
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
            StrOk(32) = GetFAgentName(strArr(i))
         End If
'edit by nickc 2007/08/24 改成同 fa 維護
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If Not IsNull(adoRecordset.Fields(0)) Then
'               StrOk(32) = strArr(i) & "   " & adoRecordset.Fields(0)
'            Else
'               StrOk(32) = strArr(i)
'            End If
         If StrOk(32) <> Empty Then
            StrOk(32) = strArr(i) & "  " & StrOk(32)
            'Add by Morgan 2004/1/19
            Lbl1(32).ForeColor = vbBlack
         Else
            StrOk(32) = strArr(i) & ""
            'Add by Morgan 2004/1/19
            Lbl1(32).ForeColor = vbRed
         End If
         CheckOC
    Case 61
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                 strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                 strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
            StrOk(33) = GetCustomerName(strArr(i), 0)
         Else
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                   strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                   strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
            StrOk(33) = GetFAgentName(strArr(i))
         End If
'edit by nickc 2007/08/24 改成同 fa 維護
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            StrOk(33) = strArr(i) & "   " & adoRecordset.Fields(0)
         If StrOk(33) <> Empty Then
            StrOk(33) = strArr(i) & "  " & StrOk(33)
            'Add by Morgan 2004/1/19
            Lbl1(33).ForeColor = vbBlack
         Else
            StrOk(33) = strArr(i) & ""
            'Add by Morgan 2004/1/19
            Lbl1(33).ForeColor = vbRed
         End If
         CheckOC
    Case 67
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/24 改成同 fa 維護
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
            StrOk(34) = GetCustomerName(strArr(i), 0)
         Else
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
            StrOk(34) = GetFAgentName(strArr(i))
         End If
'edit by nickc 2007/08/24 改成同 fa 維護
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            StrOk(34) = strArr(i) & "   " & adoRecordset.Fields(0)
         If StrOk(34) <> Empty Then
            StrOk(34) = strArr(i) & "  " & StrOk(34)
            'Add by Morgan 2004/1/19
            Lbl1(34).ForeColor = vbBlack
         Else
            StrOk(34) = strArr(i) & ""
            'Add by Morgan 2004/1/19
            Lbl1(34).ForeColor = vbRed
         End If
         CheckOC
    Case 66
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'              Else
'                  strSQL = "SELECT CU04 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'              End If
            StrOk(35) = GetCustomerName(strArr(i), 0)
         Else
'edit by nickc 2007/08/24 改成同 fa 維護
'              If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'              Else
'                  strSQL = "SELECT FA04,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'              End If
            StrOk(35) = GetFAgentName(strArr(i))
         End If
'edit by nickc 2007/08/24 改成同 fa 維護
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            StrOk(35) = strArr(i) & "   " & adoRecordset.Fields(0)
         If StrOk(35) <> Empty Then
            StrOk(35) = strArr(i) & "  " & StrOk(35)
            'Add by Morgan 2004/1/19
            Lbl1(35).ForeColor = vbBlack
         Else
            StrOk(35) = strArr(i) & ""
            'Add by Morgan 2004/1/19
            Lbl1(35).ForeColor = vbRed
         End If
         CheckOC
    Case 68
         StrOk(36) = strArr(i)
         
    Case 47
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(37) = ""
         Else
             'edit by nick 2004/10/05
             'StrOk(37) = ChangeWStringToTString(strArr(i))
             StrOk(37) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 48
         'edit by nick 2004/10/05
         'StrOk(38) = strArr(i)
         StrOk(38) = Format(strArr(i), "##:##")
    Case 50
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(39) = ""
         Else
             'edit by nick 2004/10/05
             'StrOk(39) = ChangeWStringToTString(strArr(i))
             StrOk(39) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 51
         'edit by nick 2004/10/05
         'StrOk(40) = strArr(i)
         StrOk(40) = Format(strArr(i), "##:##")
    Case 71 '專利D/N固定列印對象
         If Left$(strArr(i), 1) = "X" Then
            StrOk(42) = GetCustomerName(strArr(i), 0)
         Else
            StrOk(42) = GetFAgentName(strArr(i))
         End If
         If StrOk(42) <> Empty Then
            StrOk(42) = strArr(i) & "  " & StrOk(42)
         Else
            StrOk(42) = strArr(i) & ""
         End If
         CheckOC
    Case 72 '年費D/N列印對象
         If Left$(strArr(i), 1) = "X" Then
            StrOk(43) = GetCustomerName(strArr(i), 0)
         Else
            StrOk(43) = GetFAgentName(strArr(i))
         End If
         If StrOk(43) <> Empty Then
            StrOk(43) = strArr(i) & "  " & StrOk(43)
         Else
            StrOk(43) = strArr(i) & ""
         End If
         CheckOC
    Case 73 '商標全部折扣
        StrOk(44) = strArr(i)
    Case 74 '商標申請/翻譯折扣
        StrOk(45) = strArr(i)
    Case 75 '商標全部折扣起始日
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(46) = ""
         Else
             StrOk(46) = ChangeWStringToTString(strArr(i))
         End If
    'Add By Sindy 2025/3/10
    Case 137 '繳註冊費折扣
        StrOk(72) = strArr(i)
    Case 138 '延展折扣
        StrOk(73) = strArr(i)
    Case 139 '商標全部折扣終止日
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(71) = ""
         Else
             StrOk(71) = ChangeWStringToTString(strArr(i))
         End If
    '2025/3/10 END
    'add  by nickc 2005/10/24
    Case 76
         StrOk(47) = strArr(i)
    Case 77
         StrOk(48) = strArr(i)
         If StrOk(48) = "Y" Then
            Lbl1(41).ForeColor = &HFF&
         Else
            Lbl1(41).ForeColor = &H80000012
         End If
   'Add by Morgan 2007/10/29
    Case 85
         StrOk(49) = strArr(i)
   'Add by Morgan 2008/1/17
    Case 86
         StrOk(51) = strArr(i)
   'Add by Morgan 2008/1/17
    Case 79, 80, 81, 82
         '2010/9/29 MODIFY BY SONIA 改為TXT欄使用者可複製
         'StrOk(i - 27) = strArr(i)
         StrOkTxt(i - 63) = strArr(i)
    '2007/11/1 ADD BY SONIA
    Case 69
         StrOk(50) = strArr(i)
    'Add by Morgan 2008/3/13
    Case 87 To 90
         StrOk(i - 31) = strArr(i)
    'Add by Morgan 2008/5/26
    Case 91
         StrOk(60) = strArr(i)
    'add by Toni 2008/10/21
    Case 93
      StrOk(61) = strArr(i)
    'end 2008/10/21
    '2008/12/9 add by sonia
    Case 97
       StrOk(62) = strArr(i)
    '2008/12/9 end
    'Add by Morgan 2009/10/16
    Case 98
       StrOk(63) = strArr(i)
    Case 99
       StrOk(64) = strArr(i)
    'end 2009/10/16
    'Add By Sindy 2011/3/4
    Case 78
         StrOk(7) = strArr(i)
    'Added by Lydia 2018/07/20 財務信箱(CF)
    Case 105
         StrOkTxt(22) = strArr(i)
    'end 2018/07/20
    Case 106
         StrOkTxt(21) = strArr(i)
    Case 107
         If Left$(strArr(i), 1) = "X" Then
            StrOk(54) = GetCustomerName(strArr(i), 0)
         Else
            StrOk(54) = GetFAgentName(strArr(i))
         End If
          If StrOk(54) <> Empty Then
            StrOk(54) = strArr(i) & "  " & StrOk(54)
            Lbl1(54).ForeColor = vbBlack
         Else
            StrOk(54) = ""
            Lbl1(54).ForeColor = vbRed
            StrOk(54) = strArr(i)
         End If
         CheckOC
    Case 108
         StrOk(53) = strArr(i)
    Case 109
         StrOk(52) = strArr(i)
    Case 110
         StrOkTxt(20) = strArr(i)
    Case 111 '商標D/N固定列印對象
         If Left$(strArr(i), 1) = "X" Then
            StrOk(55) = GetCustomerName(strArr(i), 0)
         Else
            StrOk(55) = GetFAgentName(strArr(i))
         End If
         If StrOk(55) <> Empty Then
            StrOk(55) = strArr(i) & "  " & StrOk(55)
         Else
            StrOk(55) = strArr(i) & ""
         End If
         CheckOC
    Case 112 '延展D/N列印對象
         If Left$(strArr(i), 1) = "X" Then
            StrOk(65) = GetCustomerName(strArr(i), 0)
         Else
            StrOk(65) = GetFAgentName(strArr(i))
         End If
         If StrOk(65) <> Empty Then
            StrOk(65) = strArr(i) & "  " & StrOk(65)
         Else
            StrOk(65) = strArr(i) & ""
         End If
         CheckOC
   '2011/3/4 End
   'Add By Sindy 2011/3/10
   Case 100
         StrOk(66) = strArr(i)
   '2011/3/10 End
   'Added by Lydia 2017/11/30 FCP是否電子送件
   Case 104
         StrOk(70) = strArr(i)
   'end 2017/11/30
   'Add By Sindy 2012/6/25
   Case 113
         strExc(0) = "SELECT A1Y01||'-'||A1Y02 FROM ACC1Y0 WHERE A1Y01='" & strArr(i) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            StrOk(67) = RsTemp.Fields(0)
         Else
            StrOk(67) = strArr(i)
         End If
   '2012/6/25 End
   'Add By Sindy 2013/1/30 +115,116
   Case 115
      Combo3(0).ListIndex = 0
      If Not IsNull(strArr(i)) And Trim(strArr(i)) <> "" Then
         Combo3(0).ListIndex = strArr(i)
      End If
    Case 116
      Combo3(1).ListIndex = 0
      If Not IsNull(strArr(i)) And Trim(strArr(i)) <> "" Then
         Combo3(1).ListIndex = strArr(i)
      End If
    '2013/1/30 End
   'Add By Sindy 2013/8/26
   Case 117
      StrOk(68) = strArr(i)
   '2013/8/26 End
   'Add By Sindy 2016/12/5
   Case 119
      StrOk(69) = strArr(i)
   '2016/12/5 End
   'Add by Amy 2017/03/06
   Case 120
      LblFA120 = strArr(i)
   'Add By Sindy 2021/3/3 +126
   Case 126
      Combo5.ListIndex = -1
      If Not IsNull(strArr(i)) And Trim(strArr(i)) <> "" Then
         Combo5.ListIndex = strArr(i)
      End If
   'Add by Amy 2022/12/07
   Case 127 '代理人來源
      lblFA127.BackColor = &H8000000F:  lblXYS03.BackColor = &H8000000F
      lblXYS02.BackColor = &H8000000F: lblXYS02_N.BackColor = &H8000000F
      lblFA127 = "": lblXYS02 = "": lblXYS02_N = "": lblXYS03 = ""
      If strArr(i) <> MsgText(601) Then
            lblFA127 = GetSourceName(strArr(i))
            Call Pub_GetXYSource(1, Left(Str01, 8), strTp(0), strTp(1), strTp(2))
            lblXYS02 = strTp(0)
            lblXYS02_N = strTp(1)
            lblXYS03 = strTp(2)
      End If
   'Add by Sindy 2025/1/7
    Case 135
         If Trim(strArr(i)) <> "" Then
            arrID = Split(strArr(i), ",")
            For intI = UBound(arrID) To LBound(arrID) Step -1
               Chk1K(Val(arrID(intI)) - 1).Value = 1
            Next intI
         End If
    '2025/1/7 END
   Case Else
   End Select
Next i
For i = 0 To UBound(StrOk)
'   Select Case i
'      Case 7, 52, 53, 54, 55 '2010/9/29 ADD BY SONIA E-MAIL欄移至txt1
'      Case Else
         Lbl1(i) = StrOk(i)
'   End Select
Next i
For i = 0 To UBound(StrOkTxt)
    txt1(i) = StrOkTxt(i)
Next i

'Add by Morgan 2008/11/13 改index與欄位序次相同的陣列，將來再新增欄位時只需加畫面的物件並指定相同的index就好
For Each oLbl In lblFa
   oLbl = strArr(oLbl.Index)
   oLbl.BackColor = &H8000000F
Next
'Modified by Lydia 2021/12/13 改為Form 2.0元件
PUB_SetUserList lstDeveloper, strArr(94), True
'end 2008/11/13

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
If PUB_ChkCustWebIDUserRights(Lbl1(41), strUserNum) = True Then
   CmdOk1(2).Visible = True
Else
   CmdOk1(2).Visible = False
End If
'2012/10/2 End
'Add by Amy 2019/05/08
If frm100101_10.CmdOk1(2).Visible = False Then
    frm100101_10.CmdOk1(5).Left = 5850
    'Modify by Amy 2022/12/08 各項指示放最左邊-外專(Morgan通知)
    frm100101_10.CmdOk1(3).Left = 3420 'Added by Lydia 2020/09/18
    frm100101_10.CmdOk1(4).Left = 4630 'Add by Amy 2022/12/06 被介紹者
End If
'Add by Amy 2022/12/06
If strSrvDate(1) >= 代理人來源啟用日 Then
    CmdOk1(4).Visible = True
    CmdOk1(4).BackColor = &H8000000F
    If Pub_GetXYSource(2, Left(Lbl1(41), 8)) = True Then
        CmdOk1(4).BackColor = &HFFFF80
    End If
End If
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   
   'Mark by Lydia 2021/12/15 已在外層控制權限，所以不用限制顯示
   'If bolFNation = False Then
   '    SSTab1.TabVisible(0) = False
   '    SSTab1.TabVisible(1) = False
   '    SSTab1.TabVisible(2) = False
   'End If
   'end 2021/12/15
   
   SSTab1.Tab = 0 'Added by Lydia 2016/11/11

   'Add by Morgan 2007/4/25
   '考慮共榮(X22558000)的案件需要於客戶檔設定年費代理人，為避免邏輯過於複雜故取消代理人檔的年費代理人(FA61)&年費請款對象(FA62)
   Label35(0).Visible = False: Lbl1(32).Visible = False
   Label34(0).Visible = False: Lbl1(33).Visible = False
   'end 2007/4/25

   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      CmdOk1(3).Visible = True
   Else
      CmdOk1(3).Visible = False
      'Mark by Lydia 2020/09/18 按鈕移到最上方
      'txt1(6).Top = 330
      'txt1(6).Height = 4740
      'end 2020/09/18
   End If
   'end 2020/05/05
   'Add by Amy 2023/07/25 Label沒清,當案件沒代理人時會顯示Label名稱 ex:T-239854
   For Each Lbl In Lbl1
      Lbl.Caption = ""
   Next
   lblFA127.Caption = ""
   lblXYS02.Caption = ""
   lblXYS02_N.Caption = ""
   lblXYS03.Caption = ""
   'end 2023/07/25
   
   'Added by Lydia 2021/12/13 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstDeveloper.Height = 720
   lstDeveloper.Width = 1500
   'Memo by Amy 2025/02/11 FCT註冊費自動代繳移位,切其他頁籤會有殘影搬至按頁籤
   
   '92.04.16 nick
   cmdState = -1
   
   Frame1K.BorderStyle = 0 'Add By Sindy 2025/1/7
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100101_10 = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   'Modify By Sindy 2025/3/10 mark,因畫面上移位置了
'   'Add by Amy 2025/02/11 從Form_Load搬過來,否則切其他頁籤會有殘影
'   If SSTab1.Tab = 3 Then
'      'Add by Amy 2024/03/08 隱藏延展單筆不跑,將FCT註冊費自動代繳移位
'      Label8(2).Left = 240
'      lbl1(36).Left = 2000
'   End If
End Sub

'Added by Lydia 2016/10/29 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
End Sub

'Add by Amy 2022/12/07
Private Function GetSourceName(ByVal stNo As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    If stNo = MsgText(601) Then Exit Function
    strQ = "Select ac02||' '||ac03 as stName From AllCode Where ac01='13' And ac02='" & stNo & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetSourceName = "" & RsQ.Fields("stName")
    End If
    Set RsQ = Nothing
End Function
