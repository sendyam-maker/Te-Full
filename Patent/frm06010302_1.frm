VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010302_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-讓與, 合併"
   ClientHeight    =   6490
   ClientLeft      =   630
   ClientTop       =   1100
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6490
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdOK 
      Caption         =   "變更事項(&R)"
      Height          =   400
      Index           =   3
      Left            =   5670
      TabIndex        =   0
      Top             =   70
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   525
      Left            =   8460
      TabIndex        =   128
      Top             =   630
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   129
         Top             =   210
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   130
         Top             =   0
         Width           =   765
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5025
      Left            =   60
      TabIndex        =   16
      Top             =   1440
      Width           =   9225
      _ExtentX        =   16281
      _ExtentY        =   8872
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm06010302_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNameAgent"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label22"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label24"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label25"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label26"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label27"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label28"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lstNameAgent"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "SSTab2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCP84"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text10(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text10(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text10(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text10(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text5"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text14"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "變更事項"
      TabPicture(1)   =   "frm06010302_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkAtt(26)"
      Tab(1).Control(1)=   "chkAtt(27)"
      Tab(1).Control(2)=   "chkAtt(28)"
      Tab(1).Control(3)=   "chkAtt(29)"
      Tab(1).Control(4)=   "chkAtt(30)"
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(6)=   "Label12"
      Tab(1).Control(7)=   "Label18(0)"
      Tab(1).Control(8)=   "Label8"
      Tab(1).ControlCount=   9
      Begin VB.Frame Frame3 
         BorderStyle     =   0  '沒有框線
         Height          =   315
         Left            =   3960
         TabIndex        =   184
         Top             =   4470
         Visible         =   0   'False
         Width           =   2775
         Begin VB.TextBox TextPA178 
            Height          =   270
            Left            =   960
            MaxLength       =   1
            TabIndex        =   185
            Top             =   0
            Width           =   300
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "證書形式:       （1: 電子 2: 紙本）"
            Height          =   180
            Index           =   5
            Left            =   150
            TabIndex        =   186
            Top             =   30
            Width           =   2565
         End
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更原申請人之地址"
         Height          =   195
         Index           =   26
         Left            =   -74490
         TabIndex        =   40
         Top             =   720
         Width           =   1950
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更原申請人之代理人"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   27
         Left            =   -74490
         TabIndex        =   39
         Top             =   960
         Width           =   2430
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更原申請人之代表人"
         Height          =   195
         Index           =   28
         Left            =   -74490
         TabIndex        =   38
         Top             =   1200
         Width           =   2430
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更原申請人之姓名或名稱"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   29
         Left            =   -74490
         TabIndex        =   37
         Top             =   1440
         Width           =   2580
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更原申請人之國籍"
         Height          =   195
         Index           =   30
         Left            =   -74490
         TabIndex        =   36
         Top             =   1680
         Width           =   2430
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Left            =   2310
         MaxLength       =   1
         TabIndex        =   22
         Top             =   4470
         Width           =   492
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   21
         Top             =   3870
         Width           =   1155
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   0
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   17
         Top             =   3870
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   1
         Left            =   2310
         MaxLength       =   2
         TabIndex        =   18
         Top             =   3870
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   2
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   19
         Top             =   4170
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Index           =   3
         Left            =   2310
         MaxLength       =   2
         TabIndex        =   20
         Top             =   4170
         Width           =   495
      End
      Begin VB.TextBox txtCP84 
         Height          =   270
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   23
         Top             =   4170
         Width           =   1155
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3465
         Left            =   210
         TabIndex        =   45
         Top             =   360
         Width           =   8805
         _ExtentX        =   15522
         _ExtentY        =   6121
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         BackColor       =   -2147483644
         TabCaption(0)   =   "受讓人1"
         TabPicture(0)   =   "frm06010302_1.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label5(8)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label5(7)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label5(6)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label5(5)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label5(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label5(3)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label14(1)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label18(2)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label38"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label39"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label40"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtAppName(1)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txtAppName(2)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txtAppName(3)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtCaseField(39)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txtCaseField(40)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtCaseField(41)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txtCaseField(42)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txtCaseField(43)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txtCaseField(44)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Combo2(1)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Combo2(0)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtAppNew(1)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).ControlCount=   23
         TabCaption(1)   =   "受讓人2"
         TabPicture(1)   =   "frm06010302_1.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label5(24)"
         Tab(1).Control(1)=   "Label5(25)"
         Tab(1).Control(2)=   "Label5(26)"
         Tab(1).Control(3)=   "Label5(27)"
         Tab(1).Control(4)=   "Label5(28)"
         Tab(1).Control(5)=   "Label5(29)"
         Tab(1).Control(6)=   "Label14(2)"
         Tab(1).Control(7)=   "Label18(1)"
         Tab(1).Control(8)=   "Label41"
         Tab(1).Control(9)=   "Label42"
         Tab(1).Control(10)=   "Label43"
         Tab(1).Control(11)=   "txtAppName(4)"
         Tab(1).Control(12)=   "txtAppName(5)"
         Tab(1).Control(13)=   "txtAppName(6)"
         Tab(1).Control(14)=   "txtCaseField(45)"
         Tab(1).Control(15)=   "txtCaseField(46)"
         Tab(1).Control(16)=   "txtCaseField(47)"
         Tab(1).Control(17)=   "txtCaseField(48)"
         Tab(1).Control(18)=   "txtCaseField(49)"
         Tab(1).Control(19)=   "txtCaseField(50)"
         Tab(1).Control(20)=   "Combo2(3)"
         Tab(1).Control(21)=   "Combo2(2)"
         Tab(1).Control(22)=   "txtAppNew(2)"
         Tab(1).ControlCount=   23
         TabCaption(2)   =   "受讓人3"
         TabPicture(2)   =   "frm06010302_1.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label5(18)"
         Tab(2).Control(1)=   "Label5(17)"
         Tab(2).Control(2)=   "Label5(16)"
         Tab(2).Control(3)=   "Label14(6)"
         Tab(2).Control(4)=   "Label5(33)"
         Tab(2).Control(5)=   "Label5(34)"
         Tab(2).Control(6)=   "Label5(35)"
         Tab(2).Control(7)=   "Label14(3)"
         Tab(2).Control(8)=   "Label44"
         Tab(2).Control(9)=   "Label45"
         Tab(2).Control(10)=   "Label46"
         Tab(2).Control(11)=   "txtAppName(7)"
         Tab(2).Control(12)=   "txtAppName(8)"
         Tab(2).Control(13)=   "txtAppName(9)"
         Tab(2).Control(14)=   "txtCaseField(51)"
         Tab(2).Control(15)=   "txtCaseField(52)"
         Tab(2).Control(16)=   "txtCaseField(53)"
         Tab(2).Control(17)=   "txtCaseField(54)"
         Tab(2).Control(18)=   "txtCaseField(55)"
         Tab(2).Control(19)=   "txtCaseField(56)"
         Tab(2).Control(20)=   "Combo2(5)"
         Tab(2).Control(21)=   "Combo2(4)"
         Tab(2).Control(22)=   "txtAppNew(3)"
         Tab(2).ControlCount=   23
         TabCaption(3)   =   "受讓人4"
         TabPicture(3)   =   "frm06010302_1.frx":008C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label5(21)"
         Tab(3).Control(1)=   "Label5(20)"
         Tab(3).Control(2)=   "Label5(19)"
         Tab(3).Control(3)=   "Label18(4)"
         Tab(3).Control(4)=   "Label5(12)"
         Tab(3).Control(5)=   "Label5(11)"
         Tab(3).Control(6)=   "Label5(10)"
         Tab(3).Control(7)=   "Label14(5)"
         Tab(3).Control(8)=   "Label47"
         Tab(3).Control(9)=   "Label48"
         Tab(3).Control(10)=   "Label49"
         Tab(3).Control(11)=   "txtAppName(10)"
         Tab(3).Control(12)=   "txtAppName(11)"
         Tab(3).Control(13)=   "txtAppName(12)"
         Tab(3).Control(14)=   "txtCaseField(57)"
         Tab(3).Control(15)=   "txtCaseField(58)"
         Tab(3).Control(16)=   "txtCaseField(59)"
         Tab(3).Control(17)=   "txtCaseField(60)"
         Tab(3).Control(18)=   "txtCaseField(61)"
         Tab(3).Control(19)=   "txtCaseField(62)"
         Tab(3).Control(20)=   "Combo2(7)"
         Tab(3).Control(21)=   "Combo2(6)"
         Tab(3).Control(22)=   "txtAppNew(4)"
         Tab(3).ControlCount=   23
         TabCaption(4)   =   "受讓人5"
         TabPicture(4)   =   "frm06010302_1.frx":00A8
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label5(15)"
         Tab(4).Control(1)=   "Label5(14)"
         Tab(4).Control(2)=   "Label5(13)"
         Tab(4).Control(3)=   "Label18(3)"
         Tab(4).Control(4)=   "Label5(9)"
         Tab(4).Control(5)=   "Label5(2)"
         Tab(4).Control(6)=   "Label5(1)"
         Tab(4).Control(7)=   "Label14(4)"
         Tab(4).Control(8)=   "Label50"
         Tab(4).Control(9)=   "Label51"
         Tab(4).Control(10)=   "Label52"
         Tab(4).Control(11)=   "txtAppName(13)"
         Tab(4).Control(12)=   "txtAppName(14)"
         Tab(4).Control(13)=   "txtAppName(15)"
         Tab(4).Control(14)=   "txtCaseField(63)"
         Tab(4).Control(15)=   "txtCaseField(64)"
         Tab(4).Control(16)=   "txtCaseField(65)"
         Tab(4).Control(17)=   "txtCaseField(66)"
         Tab(4).Control(18)=   "txtCaseField(67)"
         Tab(4).Control(19)=   "txtCaseField(68)"
         Tab(4).Control(20)=   "Combo2(9)"
         Tab(4).Control(21)=   "Combo2(8)"
         Tab(4).Control(22)=   "txtAppNew(5)"
         Tab(4).ControlCount=   23
         Begin VB.TextBox txtAppNew 
            Height          =   270
            Index           =   1
            Left            =   360
            MaxLength       =   9
            TabIndex        =   60
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtAppNew 
            Height          =   270
            Index           =   2
            Left            =   -74640
            MaxLength       =   9
            TabIndex        =   59
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtAppNew 
            Height          =   270
            Index           =   3
            Left            =   -74640
            MaxLength       =   9
            TabIndex        =   58
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtAppNew 
            Height          =   270
            Index           =   4
            Left            =   -74610
            MaxLength       =   9
            TabIndex        =   57
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtAppNew 
            Height          =   270
            Index           =   5
            Left            =   -74640
            MaxLength       =   9
            TabIndex        =   56
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   0
            ItemData        =   "frm06010302_1.frx":00C4
            Left            =   945
            List            =   "frm06010302_1.frx":00C6
            Style           =   2  '單純下拉式
            TabIndex        =   55
            Top             =   1140
            Width           =   6225
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   1
            ItemData        =   "frm06010302_1.frx":00C8
            Left            =   945
            List            =   "frm06010302_1.frx":00CA
            Style           =   2  '單純下拉式
            TabIndex        =   54
            Top             =   2250
            Width           =   6225
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   2
            ItemData        =   "frm06010302_1.frx":00CC
            Left            =   -74055
            List            =   "frm06010302_1.frx":00CE
            Style           =   2  '單純下拉式
            TabIndex        =   53
            Top             =   1140
            Width           =   6225
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   3
            ItemData        =   "frm06010302_1.frx":00D0
            Left            =   -74055
            List            =   "frm06010302_1.frx":00D2
            Style           =   2  '單純下拉式
            TabIndex        =   52
            Top             =   2265
            Width           =   6225
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   4
            ItemData        =   "frm06010302_1.frx":00D4
            Left            =   -74055
            List            =   "frm06010302_1.frx":00D6
            Style           =   2  '單純下拉式
            TabIndex        =   51
            Top             =   1140
            Width           =   6225
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   5
            ItemData        =   "frm06010302_1.frx":00D8
            Left            =   -74055
            List            =   "frm06010302_1.frx":00DA
            Style           =   2  '單純下拉式
            TabIndex        =   50
            Top             =   2280
            Width           =   6225
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   6
            ItemData        =   "frm06010302_1.frx":00DC
            Left            =   -74055
            List            =   "frm06010302_1.frx":00DE
            Style           =   2  '單純下拉式
            TabIndex        =   49
            Top             =   1140
            Width           =   6225
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   7
            ItemData        =   "frm06010302_1.frx":00E0
            Left            =   -74055
            List            =   "frm06010302_1.frx":00E2
            Style           =   2  '單純下拉式
            TabIndex        =   48
            Top             =   2265
            Width           =   6225
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   8
            ItemData        =   "frm06010302_1.frx":00E4
            Left            =   -74055
            List            =   "frm06010302_1.frx":00E6
            Style           =   2  '單純下拉式
            TabIndex        =   47
            Top             =   1140
            Width           =   6225
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   9
            ItemData        =   "frm06010302_1.frx":00E8
            Left            =   -74055
            List            =   "frm06010302_1.frx":00EA
            Style           =   2  '單純下拉式
            TabIndex        =   46
            Top             =   2250
            Width           =   6225
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   68
            Left            =   -74055
            TabIndex        =   183
            Top             =   3090
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   67
            Left            =   -74055
            TabIndex        =   182
            Top             =   2820
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   66
            Left            =   -74055
            TabIndex        =   181
            Top             =   2550
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   65
            Left            =   -74055
            TabIndex        =   180
            Top             =   1980
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   64
            Left            =   -74055
            TabIndex        =   179
            Top             =   1710
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   63
            Left            =   -74055
            TabIndex        =   178
            Top             =   1440
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   62
            Left            =   -74055
            TabIndex        =   177
            Top             =   3090
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   61
            Left            =   -74055
            TabIndex        =   176
            Top             =   2820
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   60
            Left            =   -74055
            TabIndex        =   175
            Top             =   2550
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   59
            Left            =   -74055
            TabIndex        =   174
            Top             =   1980
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   58
            Left            =   -74055
            TabIndex        =   173
            Top             =   1710
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   57
            Left            =   -74055
            TabIndex        =   172
            Top             =   1440
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   56
            Left            =   -74055
            TabIndex        =   171
            Top             =   3150
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   55
            Left            =   -74055
            TabIndex        =   170
            Top             =   2850
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   54
            Left            =   -74055
            TabIndex        =   169
            Top             =   2580
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   53
            Left            =   -74055
            TabIndex        =   168
            Top             =   2010
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   52
            Left            =   -74055
            TabIndex        =   167
            Top             =   1710
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   51
            Left            =   -74055
            TabIndex        =   166
            Top             =   1440
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   50
            Left            =   -74055
            TabIndex        =   165
            Top             =   3090
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   49
            Left            =   -74055
            TabIndex        =   164
            Top             =   2820
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   48
            Left            =   -74055
            TabIndex        =   163
            Top             =   2550
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   47
            Left            =   -74055
            TabIndex        =   162
            Top             =   1980
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   46
            Left            =   -74055
            TabIndex        =   161
            Top             =   1710
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   45
            Left            =   -74055
            TabIndex        =   160
            Top             =   1440
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   44
            Left            =   945
            TabIndex        =   159
            Top             =   3090
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   43
            Left            =   945
            TabIndex        =   158
            Top             =   2820
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   42
            Left            =   945
            TabIndex        =   157
            Top             =   2550
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   41
            Left            =   945
            TabIndex        =   156
            Top             =   1980
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   40
            Left            =   945
            TabIndex        =   155
            Top             =   1710
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtCaseField 
            Height          =   285
            Index           =   39
            Left            =   945
            TabIndex        =   154
            Top             =   1440
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483644
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   15
            Left            =   -72870
            TabIndex        =   153
            Top             =   870
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   14
            Left            =   -72870
            TabIndex        =   152
            Top             =   600
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   13
            Left            =   -72870
            TabIndex        =   151
            Top             =   330
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   12
            Left            =   -72840
            TabIndex        =   150
            Top             =   870
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   11
            Left            =   -72840
            TabIndex        =   149
            Top             =   600
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   10
            Left            =   -72840
            TabIndex        =   148
            Top             =   330
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   9
            Left            =   -72870
            TabIndex        =   147
            Top             =   870
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   8
            Left            =   -72870
            TabIndex        =   146
            Top             =   600
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   7
            Left            =   -72870
            TabIndex        =   145
            Top             =   330
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   6
            Left            =   -72870
            TabIndex        =   144
            Top             =   870
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   5
            Left            =   -72870
            TabIndex        =   143
            Top             =   600
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   4
            Left            =   -72870
            TabIndex        =   142
            Top             =   330
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   3
            Left            =   2160
            TabIndex        =   141
            Top             =   870
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   2
            Left            =   2160
            TabIndex        =   140
            Top             =   600
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAppName 
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   139
            Top             =   330
            Width           =   6225
            VariousPropertyBits=   679495707
            BackColor       =   -2147483648
            MaxLength       =   80
            Size            =   "10980;503"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "中:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   127
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "英:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   126
            Top             =   630
            Width           =   225
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "日:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   125
            Top             =   900
            Width           =   225
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "中:"
            Height          =   180
            Left            =   -73140
            TabIndex        =   124
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "英:"
            Height          =   210
            Left            =   -73140
            TabIndex        =   123
            Top             =   630
            Width           =   225
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "日:"
            Height          =   210
            Left            =   -73140
            TabIndex        =   122
            Top             =   900
            Width           =   225
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "中:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   121
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "英:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   120
            Top             =   630
            Width           =   225
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "日:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   119
            Top             =   900
            Width           =   225
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "中:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   118
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "英:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   117
            Top             =   630
            Width           =   225
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "日:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   116
            Top             =   900
            Width           =   225
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "中:"
            Height          =   180
            Left            =   1830
            TabIndex        =   115
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "英:"
            Height          =   180
            Left            =   1830
            TabIndex        =   114
            Top             =   630
            Width           =   225
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "日:"
            Height          =   180
            Left            =   1830
            TabIndex        =   113
            Top             =   900
            Width           =   225
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "中:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   112
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "英:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   111
            Top             =   600
            Width           =   225
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "日:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   110
            Top             =   840
            Width           =   225
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "中:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   109
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "英:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   108
            Top             =   600
            Width           =   225
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "日:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   107
            Top             =   840
            Width           =   225
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "中:"
            Height          =   180
            Left            =   -73140
            TabIndex        =   106
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "英:"
            Height          =   180
            Left            =   -73140
            TabIndex        =   105
            Top             =   600
            Width           =   225
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "日:"
            Height          =   180
            Left            =   -73140
            TabIndex        =   104
            Top             =   840
            Width           =   225
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "中:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   103
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "英:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   102
            Top             =   600
            Width           =   225
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "日:"
            Height          =   180
            Left            =   -73170
            TabIndex        =   101
            Top             =   840
            Width           =   225
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "代表人2"
            Height          =   180
            Index           =   2
            Left            =   90
            TabIndex        =   100
            Top             =   2340
            Width           =   630
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "代表人1"
            Height          =   180
            Index           =   1
            Left            =   90
            TabIndex        =   99
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   3
            Left            =   450
            TabIndex        =   98
            Top             =   1470
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   4
            Left            =   450
            TabIndex        =   97
            Top             =   1710
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   5
            Left            =   450
            TabIndex        =   96
            Top             =   1950
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   6
            Left            =   450
            TabIndex        =   95
            Top             =   2610
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   7
            Left            =   450
            TabIndex        =   94
            Top             =   2850
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   8
            Left            =   450
            TabIndex        =   93
            Top             =   3090
            Width           =   345
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "代表人4"
            Height          =   180
            Index           =   1
            Left            =   -74910
            TabIndex        =   92
            Top             =   2295
            Width           =   630
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "代表人3"
            Height          =   180
            Index           =   2
            Left            =   -74910
            TabIndex        =   91
            Top             =   1185
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   29
            Left            =   -74550
            TabIndex        =   90
            Top             =   1515
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   28
            Left            =   -74550
            TabIndex        =   89
            Top             =   1755
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   27
            Left            =   -74550
            TabIndex        =   88
            Top             =   1995
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   26
            Left            =   -74550
            TabIndex        =   87
            Top             =   2595
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   25
            Left            =   -74550
            TabIndex        =   86
            Top             =   2835
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   24
            Left            =   -74550
            TabIndex        =   85
            Top             =   3075
            Width           =   345
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "代表人5"
            Height          =   180
            Index           =   3
            Left            =   -74910
            TabIndex        =   84
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   35
            Left            =   -74550
            TabIndex        =   83
            Top             =   1500
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   34
            Left            =   -74550
            TabIndex        =   82
            Top             =   1770
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   33
            Left            =   -74550
            TabIndex        =   81
            Top             =   2040
            Width           =   345
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "代表人6"
            Height          =   180
            Index           =   6
            Left            =   -74910
            TabIndex        =   80
            Top             =   2340
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   16
            Left            =   -74550
            TabIndex        =   79
            Top             =   2610
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   17
            Left            =   -74550
            TabIndex        =   78
            Top             =   2850
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   18
            Left            =   -74550
            TabIndex        =   77
            Top             =   3090
            Width           =   345
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "代表人8"
            Height          =   180
            Index           =   5
            Left            =   -74910
            TabIndex        =   76
            Top             =   2310
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   10
            Left            =   -74550
            TabIndex        =   75
            Top             =   2610
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   11
            Left            =   -74550
            TabIndex        =   74
            Top             =   2850
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   12
            Left            =   -74550
            TabIndex        =   73
            Top             =   3090
            Width           =   345
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "代表人7"
            Height          =   180
            Index           =   4
            Left            =   -74910
            TabIndex        =   72
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   19
            Left            =   -74550
            TabIndex        =   71
            Top             =   1470
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   20
            Left            =   -74550
            TabIndex        =   70
            Top             =   1740
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   21
            Left            =   -74550
            TabIndex        =   69
            Top             =   1980
            Width           =   345
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "代表人10"
            Height          =   180
            Index           =   4
            Left            =   -74910
            TabIndex        =   68
            Top             =   2310
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   1
            Left            =   -74550
            TabIndex        =   67
            Top             =   2580
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   2
            Left            =   -74550
            TabIndex        =   66
            Top             =   2850
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   9
            Left            =   -74550
            TabIndex        =   65
            Top             =   3150
            Width           =   345
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "代表人9"
            Height          =   180
            Index           =   3
            Left            =   -74910
            TabIndex        =   64
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(中):"
            Height          =   180
            Index           =   13
            Left            =   -74550
            TabIndex        =   63
            Top             =   1470
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(英):"
            Height          =   180
            Index           =   14
            Left            =   -74550
            TabIndex        =   62
            Top             =   1770
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(日):"
            Height          =   180
            Index           =   15
            Left            =   -74550
            TabIndex        =   61
            Top             =   2040
            Width           =   345
         End
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   7470
         TabIndex        =   24
         Top             =   3870
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;556"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "同時辦理事項"
         Height          =   180
         Left            =   -74850
         TabIndex        =   44
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label12 
         Caption         =   "+ 300"
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   -73650
         TabIndex        =   43
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label18 
         Caption         =   "客戶資料維護修改地址再產生申請書)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   -71790
         TabIndex        =   42
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label Label8 
         Caption         =   "(請先至"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -72450
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "(Y:WORD)"
         Height          =   180
         Left            =   2910
         TabIndex        =   35
         Top             =   4470
         Width           =   810
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "是否修改申請書內容"
         Height          =   180
         Left            =   270
         TabIndex        =   34
         Top             =   4470
         Width           =   1620
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "年年費"
         Height          =   180
         Left            =   2910
         TabIndex        =   33
         Top             =   4170
         Width           =   540
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   2070
         TabIndex        =   32
         Top             =   4140
         Width           =   180
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "繳納第"
         Height          =   180
         Left            =   270
         TabIndex        =   31
         Top             =   4170
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "年年費"
         Height          =   180
         Left            =   2910
         TabIndex        =   30
         Top             =   3870
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   2070
         TabIndex        =   29
         Top             =   3840
         Width           =   180
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "領證及繳納第"
         Height          =   180
         Left            =   270
         TabIndex        =   28
         Top             =   3870
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "申請書日期:"
         Height          =   180
         Left            =   3960
         TabIndex        =   27
         Top             =   3870
         Width           =   945
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   6540
         TabIndex        =   26
         Top             =   3870
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "繳費金額:"
         Height          =   180
         Left            =   4110
         TabIndex        =   25
         Top             =   4200
         Width           =   765
      End
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2805
      MaxLength       =   2
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2565
      MaxLength       =   1
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1725
      MaxLength       =   6
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   3
      Top             =   120
      Width           =   550
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7725
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6900
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin MSForms.Label Label4 
      Height          =   195
      Index           =   6
      Left            =   4080
      TabIndex        =   132
      Top             =   990
      Width           =   1035
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1826;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   195
      Index           =   7
      Left            =   6090
      TabIndex        =   138
      Top             =   1230
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   195
      Index           =   5
      Left            =   6090
      TabIndex        =   137
      Top             =   990
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   195
      Index           =   4
      Left            =   1230
      TabIndex        =   136
      Top             =   1230
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   195
      Index           =   3
      Left            =   1230
      TabIndex        =   135
      Top             =   990
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   195
      Index           =   0
      Left            =   1230
      TabIndex        =   134
      Top             =   450
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   195
      Index           =   1
      Left            =   4080
      TabIndex        =   133
      Top             =   450
      Width           =   1785
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3149;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   1200
      TabIndex        =   131
      Top             =   660
      Width           =   7245
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "12779;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   210
      TabIndex        =   15
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Index           =   0
      Left            =   3270
      TabIndex        =   14
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   13
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   12
      Top             =   180
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   210
      TabIndex        =   11
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   5250
      TabIndex        =   10
      Top             =   990
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   3090
      TabIndex        =   9
      Top             =   990
      Width           =   945
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   5250
      TabIndex        =   8
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   210
      TabIndex        =   7
      Top             =   990
      Width           =   765
   End
End
Attribute VB_Name = "frm06010302_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/13 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/8/8 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, cp() As String
Dim m_CP110 As String, m_CP110_2 As String, m_AgentName As String
Dim intWhere As Integer
' 90.07.05 modify by louis 儲存應繳年費的資料
Dim m_CaseFee(1 To 2) As String
Dim m_OldCaseFee As String
'Add By Cheng 2003/03/07
Dim m_CP55 As String '原讓與人
'Add by Morgan 2010/5/24
Dim m_CP(93 To 96) As String
Dim m_CaseNo As String
Dim m_Giver(1 To 5) As String


'Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
' Dim strTxt(1 To 10) As String, strTmp(1 To 2) As String, i As Integer
'   EndLetter ET01, strReceiveNo, ET03, strUserNum
'   If Text10(0) <> "" And Text10(1) <> "" Then
'      strTmp(1) = Text10(0).Text
'      strTmp(2) = Text10(1).Text
'   Else
'      strTmp(1) = Text10(2).Text
'      strTmp(2) = Text10(3).Text
'   End If
'   For i = 1 To 2
'      Select Case strTmp(i)
'         Case "1"
'            strTmp(i) = "一"
'         Case "2"
'            strTmp(i) = "二"
'         Case "3"
'            strTmp(i) = "三"
'         Case "4"
'            strTmp(i) = "四"
'         Case "5"
'            strTmp(i) = "五"
'         Case "6"
'            strTmp(i) = "六"
'         Case "7"
'            strTmp(i) = "七"
'         Case "8"
'            strTmp(i) = "八"
'         Case "9"
'            strTmp(i) = "九"
'         Case "10"
'            strTmp(i) = "十"
'         Case "11"
'            strTmp(i) = "十一"
'         Case "12"
'            strTmp(i) = "十二"
'         Case "13"
'            strTmp(i) = "十三"
'         Case "14"
'            strTmp(i) = "十四"
'         Case "15"
'            strTmp(i) = "十五"
'         Case "16"
'            strTmp(i) = "十六"
'         Case "17"
'            strTmp(i) = "十七"
'         Case "18"
'            strTmp(i) = "十八"
'         Case "19"
'            strTmp(i) = "十九"
'         Case "20"
'            strTmp(i) = "二十"
'      End Select
'   Next
'   If strTmp(1) = strTmp(2) Then
'      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費','第" & strTmp(1) & "年年費')"
'   Else
'      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費','第" & strTmp(1) & "年至第" & strTmp(2) & "年年費')"
'   End If
'   'edit by nickc 2007/02/05 不用 dll 了
'   'If Not objLawDll.ExecSQL(1, strTxt) Then
'   If Not ClsLawExecSQL(1, strTxt) Then
'      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
'   End If
'End Sub

Private Sub chkAtt_Click(Index As Integer)
   If chkAtt(27).Value = 1 Or _
      chkAtt(29).Value = 1 Then
      Label12.Visible = True
   Else
      Label12.Visible = False
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean
Dim strFolder As String
Dim strFileName As String
   
   Select Case Index
      Case 0
         If Text5 = "" Then
            MsgBox "申請書日期不可空白 !", vbCritical
            Text5.SetFocus
            Exit Sub
         End If
         'Modify By Sindy 2018/6/13
'         If Text6 <> "" Then
'            If Text7(0) = "" And Text7(1) = "" And Text7(2) = "" Then
'               MsgBox "受讓人名稱不可同時空白 !", vbCritical
'               'Text7(0).SetFocus
'               Exit Sub
'            End If
'         Else
'            MsgBox "受讓人不可空白 !", vbCritical
'            Text6.SetFocus
'            Text6_GotFocus
'            Exit Sub
'         End If
         If txtAppNew(1) = "" Then
            MsgBox "受讓人不可空白 !", vbCritical
            txtAppNew(1).SetFocus
            txtAppNew_GotFocus 1
            Exit Sub
         End If
         '2018/6/13 END
         If Text10(0) <> "" And Text10(1) <> "" And Text10(2) <> "" And Text10(3) <> "" Then
            MsgBox "領證及繳納年費與繳納年費不可同時輸入 !", vbCritical
            Text10(0).SetFocus
            Exit Sub
         End If
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
        'Added by Lydia 2020/02/21 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
        If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
            MsgBox MsgText(1111), vbInformation
            If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                Exit Sub
            End If
        End If
        'end 2020/02/21
            
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
         'If Pub_StrUserSt03 = "M51" Then
         If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
            strFolder = PUB_Getdesktop
         Else
            strFolder = FCP電子送件檔案存放路徑
         End If
         strFolder = strFolder & "\" & m_CaseNo
         If Dir(strFolder, vbDirectory) = "" Then
            MkDir strFolder
         End If
         
         strLetterDate = Text5.Text
         If Text14 = "Y" Then
            bolChk = True
         Else
            bolChk = False
         End If
         
         'Add By Sindy 2018/6/15 + 產生申請書
         If frm060103_1.Text6 = "3" Then '電子送件
            '1.基本資料
            'StartLetterPA_EData "01", "13", strReceiveNo, pa, cp, False ', IIf(chkAtt(26).Value = 1, True, False)
            If StartLetter1("01", "13") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "13", False, strUserNum, , , True, strExc(9)
            strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9), strFileName)
            
            '2.申請書
            If Trim(pa(22)) = "" Then '判斷專利號數
               If StartLetter2("01", "11") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "11", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "申請權讓與登記申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            Else
               If StartLetter2("01", "12") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "12", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利權讓與登記申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            End If
         '紙本申請書
         Else
            Call GetApplBook
         End If
         
         frm060103_1.Show
         ' 90.08.27 modify by louis
         frm060103_1.ClearForm
         Unload Me
      Case 1
         frm060103_1.Show
         ' 90.08.27 modify by louis (回到原畫面要清除畫面)
         frm060103_1.ClearForm
         Unload Me
      Case 3
         Set frm06010303_1.oParent = Me 'Add by Morgan 2011/10/5
         frm06010303_1.LoadMe strReceiveNo, pa(1), pa(2), pa(3), pa(4), 62
         Me.Hide
   End Select
End Sub

'基本資料表
Private Function StartLetter1(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(110) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim strInventor As String
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   '受讓人--畫面上
   Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), txtAppNew(1), txtAppNew(2), txtAppNew(3), txtAppNew(4), txtAppNew(5), , , , cp(10), txtCaseField(39), txtCaseField(40), txtCaseField(42), txtCaseField(43), txtCaseField(45), txtCaseField(46), txtCaseField(48), txtCaseField(49), txtCaseField(51), txtCaseField(52), txtCaseField(54), txtCaseField(55), txtCaseField(57), txtCaseField(58), txtCaseField(60), txtCaseField(61), txtCaseField(63), txtCaseField(64), txtCaseField(66), txtCaseField(67), "E", ET01, ET03, strReceiveNo)
   '讓與人--CP讓與人或發文前申請人
   Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), , , , , , , , , cp(10), , , , , , , , , , , , , , , , , , , , , "E", ET01, ET03, strReceiveNo)
   
   '受讓人之代理人(出名代理人)
   'Modify By Sindy 2021/6/21 cp(110)=>m_CP110
   strExc(0) = "select oa08,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & m_CP110 & "',oa02)>0 and st01(+)=oa02 order by OA03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','受讓人代理人" & jj & "-證書字號','" & .Fields("oa08") & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','受讓人代理人" & jj & "-ID','" & .Fields("ST26") & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','受讓人代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   '讓與人之代理人
   Call GetCP110_2
   If m_CP110_2 <> "" Then
      strExc(0) = "select oa08,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & m_CP110_2 & "',oa02)>0 and st01(+)=oa02 order by OA03"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         jj = 1
         Do While Not .EOF
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','讓與人代理人" & jj & "-證書字號','" & .Fields("oa08") & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','讓與人代理人" & jj & "-ID','" & .Fields("ST26") & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','讓與人代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
            jj = jj + 1
            .MoveNext
         Loop
         End With
      End If
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter1 = True
   End If
End Function

'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(200) As String, strTmp As String
Dim ii As Integer, jj As Integer
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   'Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa(), IIf(chkAtt(26).Value = 1, False, True))
   '受讓人--畫面上
   Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), txtAppNew(1), txtAppNew(2), txtAppNew(3), txtAppNew(4), txtAppNew(5), , , , cp(10), txtCaseField(39), txtCaseField(40), txtCaseField(42), txtCaseField(43), txtCaseField(45), txtCaseField(46), txtCaseField(48), txtCaseField(49), txtCaseField(51), txtCaseField(52), txtCaseField(54), txtCaseField(55), txtCaseField(57), txtCaseField(58), txtCaseField(60), txtCaseField(61), txtCaseField(63), txtCaseField(64), txtCaseField(66), txtCaseField(67), "E", ET01, ET03, strReceiveNo)
   '讓與人--CP讓與人或發文前申請人
   Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), , , , , , , , , cp(10), , , , , , , , , , , , , , , , , , , , , "E", ET01, ET03, strReceiveNo)
   
   '受讓人之代理人(出名代理人)
   'Modify By Sindy 2021/6/21 cp(110)=>m_CP110
   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & m_CP110 & "',oa02)>0 and st01(+)=oa02 order by OA03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','受讓人之代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   '讓與人之代理人
   Call GetCP110_2
   If m_CP110_2 <> "" Then
      strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & m_CP110_2 & "',oa02)>0 and st01(+)=oa02 order by OA03"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         jj = 1
         Do While Not .EOF
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','讓與人之代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
            jj = jj + 1
            .MoveNext
         Loop
         End With
      End If
   End If
   
   If chkAtt(26).Value = 1 Or chkAtt(27).Value = 1 Or chkAtt(28).Value = 1 Or chkAtt(29).Value = 1 Or chkAtt(30).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','同時辦理事項','♀')"
   End If
   If chkAtt(26).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & chkAtt(26).Caption & "','是')"
   End If
   If chkAtt(27).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & chkAtt(27).Caption & "','是')"
   End If
   If chkAtt(28).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & chkAtt(28).Caption & "','是')"
   End If
   If chkAtt(29).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & chkAtt(29).Caption & "','是')"
   End If
   If chkAtt(30).Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & chkAtt(30).Caption & "','是')"
   End If
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

Private Function FormSave() As Boolean
Dim stUpdate As String
Dim ii As Integer
   
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
  
   stUpdate = ""
   '若有輸入讓與申請人
   If Me.txtAppNew(1).Text <> "" Then
      stUpdate = stUpdate & ",CP56=" & CNULL(ChangeCustomerL(txtAppNew(1).Text))
      
      '若原讓與人與原申請人不同
      If ChangeCustomerL(m_CP55) <> ChangeCustomerL(pa(26)) Then
         stUpdate = stUpdate & ",CP55=" & CNULL(ChangeCustomerL(pa(26)))
      End If
   End If
   
   For intI = 0 To 3
      If Me.txtAppNew(intI + 2).Text <> "" Then
         stUpdate = stUpdate & ",CP" & (89 + intI) & "=" & CNULL(ChangeCustomerL(txtAppNew(intI + 2).Text))
         '若原讓與人與原申請人不同
         If ChangeCustomerL(m_CP(93 + intI)) <> ChangeCustomerL(pa(27 + intI)) Then
            stUpdate = stUpdate & ",CP" & (93 + intI) & "=" & CNULL(ChangeCustomerL(pa(27 + intI)))
         End If
      End If
   Next
   
'   cp(110) = ""
'   For ii = 0 To lstNameAgent.ListCount - 1
'      If lstNameAgent.Selected(ii) = True Then
'         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
'         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
'         cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
'      End If
'   Next
'   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   stUpdate = stUpdate & ",cp110=" & CNULL(m_CP110)
   
   stUpdate = stUpdate & ",cp84=" & Val(txtCP84) '發文規費
   If frm060103_1.Text6 = "3" Then '電子送件
      stUpdate = stUpdate & ",cp118='A'"
   Else
      stUpdate = stUpdate & ",cp118=null"
   End If
   
   If stUpdate <> "" Then
      stUpdate = Mid(stUpdate, 2)
      'Modify By Sindy 2018/6/19 + and cp158=0 and cp159=0
      strSql = " UPDATE CASEPROGRESS SET " & stUpdate & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql, intI
   End If
   
   'Added by Morgan 2022/12/27
   '證書形式
   If Frame3.Visible = True Then
      strSql = "Update patent Set pa178='" & TextPA178 & "' " & _
            "WHERE pa01='" & pa(1) & "' and pa02='" & pa(2) & "'" & _
             " and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
      cnnConnection.Execute strSql
   End If
   'end 2022/12/27
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
   End If
End Function


'Add By Sindy 2016/4/28
Private Sub Combo2_Click(Index As Integer)

   Dim i As Integer, strTmp As String
   
   If Combo2(Index).Text = "" Then
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

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/8/8
   ReDim pa(TF_PA)
   ReDim cp(TF_CP) 'Add By Sindy 2018/6/15
   ReadPatent
   'Add by Morgan 2005/8/8
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
   
   Label12.Visible = False
   SSTab1.Tab = 0
   SSTab2.Tab = 0
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010302_1 = Nothing
End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
Dim i As Integer, j As Integer
Dim strKey(0 To 5) As String
Dim nIndex As Integer

   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   For Each Lbl In Label4
      Lbl = ""
   Next
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label4(0) = pa(11)
      Label4(1) = pa(22)
      AddCboName Combo1, pa(5), pa(6), pa(7)
          j = 0
      For i = 26 To 30
         If pa(i) <> "" Then j = j + 1
      Next
      If pa(79) <> "" Then j = j + 1
      If pa(82) <> "" Then j = j + 1
      If j > 1 Then Text14 = "Y"
   End If
   
   If Trim(pa(22)) = "" Then '申請權
      chkAtt(26).Caption = "變更原申請人之地址"
      chkAtt(27).Caption = "變更原申請人之代理人"
      chkAtt(28).Caption = "變更原申請人之代表人"
      chkAtt(29).Caption = "變更原申請人之姓名或名稱"
      chkAtt(30).Caption = "變更原申請人之國籍"
   Else '專利權
      chkAtt(26).Caption = "變更專利權人之地址"
      chkAtt(27).Caption = "變更專利權人之代理人"
      chkAtt(28).Caption = "變更專利權人之代表人"
      chkAtt(29).Caption = "變更專利權人之姓名或名稱"
      chkAtt(30).Caption = "變更專利權人之國籍"
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
   End If
   
   ' 原已繳費的年度
   m_OldCaseFee = pa(72)
   ' 設定本所案號
   For nIndex = 1 To 4
      strKey(nIndex) = pa(nIndex)
   Next
   ' 取得繳年費的資料
   If GetMoneyDate(pa(8), pa(9), strKey, m_CaseFee(1), m_CaseFee(2)) = True Then
   End If
   'Modify by Amy 2014/08/14 +CP10
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,CP43,CP56,CP55,CP110,CP89,CP90,CP91,CP92,CP93,CP94,CP95,CP96,CP10,CP17,CP84,cp05" & _
      " from caseprogress,casepropertymap,staff,staff staff1 where " & _
      "cp09='" & strReceiveNo & "' and cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)"
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         m_CP110 = "" & .Fields("CP110")
         If Not IsNull(.Fields(0)) Then Label4(3) = .Fields(0)
         If Not IsNull(.Fields(1)) Then Label4(4) = .Fields(1)
         'Modify By Cheng 2003/01/01
'         If Not IsNull(.Fields(2)) Then Label4(5) = .Fields(6)
         If Not IsNull(.Fields(2)) Then Label4(5) = .Fields(2)
         If Not IsNull(.Fields(3)) Then
            strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
            intI = 1
            Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not IsNull(rsTemp1.Fields(0)) Then Label4(6) = TransDate(rsTemp1.Fields(0), 1)
               If Not IsNull(rsTemp1.Fields(1)) Then Label4(7) = rsTemp1.Fields(1)
            End If
         End If
         
         'Modify By Sindy 2018/6/25
'         If Not IsNull(.Fields(4)) Then txtAppNew(1).Text = ChangeCustomerS(.Fields(4)): txtAppNew_Validate 1, False
         '記錄讓與人
         m_CP55 = "" & .Fields("CP55").Value
         'Add by Morgan 2010/5/24 +受讓人,讓與人 2~5
         For intI = 0 To 3
'            If Not IsNull(.Fields("CP" & (89 + intI))) Then Text6(intI + 1).Text = ChangeCustomerS(.Fields("CP" & (89 + intI))): Text6_Validate intI + 1, False
            m_CP(93 + intI) = "" & .Fields("CP" & (93 + intI))
         Next
         'END 2010/5/24
         If Not IsNull(.Fields("CP56")) Then txtAppNew(1) = .Fields("CP56"): ChgType 11
         If Not IsNull(.Fields("CP89")) Then txtAppNew(2) = .Fields("CP89"): ChgType 12
         If Not IsNull(.Fields("CP90")) Then txtAppNew(3) = .Fields("CP90"): ChgType 13
         If Not IsNull(.Fields("CP91")) Then txtAppNew(4) = .Fields("CP91"): ChgType 14
         If Not IsNull(.Fields("CP92")) Then txtAppNew(5) = .Fields("CP92"): ChgType 15
         For i = 1 To 5 '讓與申請人代表
            Call SetCombo2(i)
         Next
         '2018/6/25 END
         
         'Add by Morgan 2006/6/8 讓與人
         m_Giver(1) = ChangeCustomerS("" & .Fields("cp55"))
         m_Giver(2) = ChangeCustomerS("" & .Fields("cp93"))
         m_Giver(3) = ChangeCustomerS("" & .Fields("cp94"))
         m_Giver(4) = ChangeCustomerS("" & .Fields("cp95"))
         m_Giver(5) = ChangeCustomerS("" & .Fields("cp96"))
         
'         If "" & .Fields("CP84") = 0 Then
'            txtCP84 = IIf("" & .Fields("CP17") > 0, .Fields("CP17"), 0)
'         Else
'            txtCP84 = "" & .Fields("CP84")
'         End If
         txtCP84.Tag = cp(17)
         txtCP84.Text = txtCP84.Tag
      End If
   End With
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
   
   'Added by Morgan 2022/12/27
   'Modify By Sindy 2024/8/19 + Or cp(10) = 合併
   If frm060103_1.Text6 = "3" And (cp(10) = 讓與 Or cp(10) = 合併) _
      And pa(22) <> "" And strSrvDate(1) >= "20230101" Then
      Frame3.Visible = True
   Else
      Frame3.Visible = False
   End If
   'end 2022/12/27
   
End Sub

'Add By Sindy 2018/6/25
Private Sub SetCombo2(Index As Integer)
Dim i As Integer, j As Integer
   
   Combo2((Index - 1) * 2).Clear
   Combo2((Index - 1) * 2).AddItem ""
   Combo2((Index - 1) * 2 + 1).Clear
   Combo2((Index - 1) * 2 + 1).AddItem ""
   
   If txtAppNew(Index) <> "" Then
      strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(txtAppNew(Index))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         For j = 1 To 6
            If IsNull(RsTemp.Fields(j - 1)) Then
               strExc(0) = ""
            Else
               strExc(0) = "-" & RsTemp.Fields(j - 1)
            End If
            Combo2((Index - 1) * 2).AddItem txtAppNew(Index) & "-" & j & strExc(0)
            Combo2((Index - 1) * 2 + 1).AddItem txtAppNew(Index) & "-" & j & strExc(0)
         Next
      End If
   End If
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String, j As Integer
   ChgType = False
   Select Case i
      Case 11, 12, 13, 14, 15
         If ClsLawGetCusCAJnam(txtAppNew(i - 10).Text, strExc(1), strExc(2), strExc(3)) = True Then
            ChgType = True
            For j = 1 To 3
               txtAppName(3 * (i - 11) + j) = strExc(j)
            Next
         End If
   End Select
End Function

Private Sub Text10_GotFocus(Index As Integer)
  TextInverse Text10(Index)
End Sub

Private Sub Text10_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      If KeyAscii <> 49 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   Else
      If (KeyAscii > 57 Or KeyAscii < 49) And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub Text10_LostFocus(Index As Integer)
 Dim i As Integer, bolChk As Boolean
   
   Select Case Index
      Case 1
         If Text10(0) = "" And Text10(1) = "" Then Exit Sub
         If Text10(0) <> "" And Text10(1) <> "" Then
            If ChkRange(Text10(0), Text10(1), "繳費年度") = True Then
               For i = Text10(0) To Text10(1)
                  If InStr(pa(72), Format(i)) > 0 Then
                     bolChk = True
                     Exit For
                  End If
               Next
               If bolChk = True Then
                  MsgBox "繳費年度錯誤，請查明後再輸入 !", vbCritical
                  Text10(0).SetFocus
               End If
            Else
               Text10(0).SetFocus
            End If
         Else
            MsgBox "繳費年度只可為空白或為 1 至 N 年 !", vbCritical
            Text10(0).SetFocus
         End If
      Case 3
         If Text10(2) = "" And Text10(3) = "" Then Exit Sub
         If Text10(0) <> "" And Text10(1) <> "" And Text10(3) <> "" And Text10(3) <> "" Then
            MsgBox "與領證及繳納年費不可同時輸入 !", vbCritical
            Text10(2) = ""
            Text10(3) = ""
            Exit Sub
         End If
         If Text10(2) <> "" And Text10(3) <> "" Then
            If ChkRange(Text10(2), Text10(3), "繳費年度") = True Then
               For i = Text10(2) To Text10(3)
                  If InStr(pa(72), Format(i)) > 0 Then
                     bolChk = True
                     Exit For
                  End If
               Next
               If bolChk = True Then
                  MsgBox "繳費年度錯誤，請查明後再輸入 !", vbCritical
                  Text10(3).SetFocus
               End If
            Else
               Text10(3).SetFocus
            End If
         Else
            MsgBox "繳費年度只可為空白或為 1 至 N 年 !", vbCritical
            Text10(3).SetFocus
         End If
   End Select
End Sub

Private Sub Text14_GotFocus()
  TextInverse Text14
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If Me.Text5.Enabled = True Then
      Cancel = False
      Text5_Validate Cancel
      If Cancel = True Then
         Me.Text5.SetFocus
         Text5_GotFocus
         Exit Function
      End If
   End If
   
   For ii = 1 To 5
      If Me.txtAppNew(ii).Enabled = True Then
         Cancel = False
         txtAppNew_Validate ii, Cancel
         If Cancel = True Then
            SSTab1.Tab = ii - 1
            Me.txtAppNew(ii).SetFocus
            txtAppNew_GotFocus ii
            Exit Function
         End If
      End If
   Next
   
'   For Each objTxt In Text10
'      If objTxt.Enabled = True Then
'         Cancel = False
'         Text10_LostFocus objTxt.Index, Cancel
'         If Cancel = True Then
'            Me.Text10(objTxt.Index).SetFocus
'            Text10_GotFocus objTxt.Index
'            Exit Function
'         End If
'      End If
'   Next
   
   'Add by Morgan 2005/8/8
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Morgan 2022/12/27
   If Frame3.Visible Then
      If TextPA178 = "" Then
         MsgBox "請輸入證書形式！", vbExclamation
         TextPA178.SetFocus
         Exit Function
      End If
   End If
   'end 2022/12/27
   
   'Memo by Morgan 2022/12/29 下面的程式固定放最後以免重複+300
   If Label12.Visible = True Then
      txtCP84 = Val(txtCP84) + 300
      txtCP84.Tag = txtCP84.Text
   End If
   If Val(txtCP84) > 0 And txtCP84.Enabled Then
      Cancel = False
      txtCP84_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         If Label12.Visible = True Then txtCP84 = Val(txtCP84) - 300
         txtCP84.SetFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

'Add by Morgan 2005/8/8
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer

   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

Private Sub txtAppNew_GotFocus(Index As Integer)
    TextInverse txtAppNew(Index)
End Sub

Private Sub txtAppNew_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtAppNew_Validate(Index As Integer, Cancel As Boolean)
   If Trim(txtAppNew(Index)) = Empty Then
      If Index = 1 Then
         MsgBox "受讓人1不可為空白 !", vbCritical
         Cancel = True
      End If
   Else
      txtAppNew(Index) = ChangeCustomerL(Trim(txtAppNew(Index)))
      If Not ChgType(Index + 10) Then
         Cancel = True
      Else
         If PUB_CheckStatus(txtAppNew(Index).Text) = False Then
            Cancel = True
            Exit Sub
         End If
      End If
      Call SetCombo2(Index) 'Add By Sindy 2016/4/29
   End If
End Sub

'Add By Sindy 2016/4/29 產生申請書
Private Function GetApplBook() As Boolean
Dim m_FileName As String
Dim i As Integer, int_TotCnt As Integer
Dim strName As String, strText As String, strLineText As String
Dim strApplLineText As String, strApplLineText2 As String
Dim bolFileExist As Boolean
Dim intCP56Cnt As Integer, intApplCnt As Integer
Dim strCP56Text As String, strApplText As String
Dim bolHad1603 As Boolean 'Add By Sindy 2016/5/11
Dim strTemp As String
   
On Error GoTo ErrHand
   
   GetApplBook = False
   
   m_MySt(1) = pa(1)
   m_MySt(2) = pa(2)
   m_MySt(3) = pa(3)
   m_MySt(4) = pa(4)
   m_SysKind = CheckSys(m_MySt(1))
   SetLetterSt
   
   'Modify By Sindy 2016/5/11 改檢查是否有專利證書
   strExc(0) = "select cp09" & _
               " from caseprogress" & _
               " where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
               " and cp10='1603'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   bolHad1603 = False
   If intI = 1 Then
      bolHad1603 = True
   End If
   '2016/5/11 END
   
   '取得樣本檔
   Select Case cp(10)
      'Modify By Sindy 2017/3/16 合併案和讓與同申請書樣本
      Case 讓與, 合併
         'Modify By Sindy 2016/5/11
         'If Trim(pa(16)) = "" Then '無核准結果
         'If bolHad1603 = False Then
         '2016/5/11 END
         If Trim(pa(22)) = "" Then 'Modify By Sindy 2016/5/26 改判斷專利號數
            m_FileName = "$$申請權讓與_樣本.doc"
            Call PUB_GetSampleFile(m_FileName, "M51-000200-0-05", bolFileExist)
            If bolFileExist = True Then MsgBox "請將已開啟的申請書檔案，關閉後再重新執行！": Exit Function
            int_TotCnt = 11
'            '受讓人--畫面上
'            strCP56Text = PUB_GetApplData(pa(1), pa(2), pa(3), pa(4), txtAppNew(1), txtAppNew(2), txtAppNew(3), txtAppNew(4), txtAppNew(5), strApplLineText, intCP56Cnt, intApplCnt, m_CP10, txtCaseField(39), txtCaseField(40), txtCaseField(42), txtCaseField(43), txtCaseField(45), txtCaseField(46), txtCaseField(48), txtCaseField(49), txtCaseField(51), txtCaseField(52), txtCaseField(54), txtCaseField(55), txtCaseField(57), txtCaseField(58), txtCaseField(60), txtCaseField(61), txtCaseField(63), txtCaseField(64), txtCaseField(66), txtCaseField(67))
'            '讓與人--CP
'            strApplText = PUB_GetApplData(pa(1), pa(2), pa(3), pa(4), , , , , , strApplLineText2, , , m_CP10)
         Else
            m_FileName = "$$專利權讓與_樣本.doc"
            Call PUB_GetSampleFile(m_FileName, "M51-000200-0-06", bolFileExist)
            If bolFileExist = True Then MsgBox "請將已開啟的申請書檔案，關閉後再重新執行！": Exit Function
            int_TotCnt = 11
'            '受讓人
'            Call GetCP56Data(intCP56Cnt, strCP56Text)
'            '讓與人
'            Call GetApplData(intApplCnt, strApplText)
         End If
         'Modify By Sindy 2017/5/11
         '受讓人--畫面上
         strCP56Text = PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), txtAppNew(1), txtAppNew(2), txtAppNew(3), txtAppNew(4), txtAppNew(5), strApplLineText, intCP56Cnt, intApplCnt, cp(10), txtCaseField(39), txtCaseField(40), txtCaseField(42), txtCaseField(43), txtCaseField(45), txtCaseField(46), txtCaseField(48), txtCaseField(49), txtCaseField(51), txtCaseField(52), txtCaseField(54), txtCaseField(55), txtCaseField(57), txtCaseField(58), txtCaseField(60), txtCaseField(61), txtCaseField(63), txtCaseField(64), txtCaseField(66), txtCaseField(67))
         '讓與人--CP讓與人或發文前申請人
         strApplText = PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), , , , , , strApplLineText2, , , cp(10))
         '2017/5/11 END
   End Select
   
   If Dir(App.path & "\" & m_FileName) <> "" Then
      Screen.MousePointer = vbHourglass
      '判斷word是否已開啟
      If g_WordAp Is Nothing Then
RestarWord:
         Set g_WordAp = New Word.Application
         g_WordAp.Visible = True 'False
      End If
'         If Dir(PUB_Getdesktop & "\" & m_TempFileName) <> "" Then
'            Kill PUB_Getdesktop & "\" & m_TempFileName
'         End If
      g_WordAp.Documents.Open App.path & "\" & m_FileName
'         g_WordAp.ActiveDocument.SaveAs PUB_Getdesktop & "\" & m_TempFileName
'         g_WordAp.ActiveDocument.Close
'         g_WordAp.Documents.Open PUB_Getdesktop & "\" & m_TempFileName
      With g_WordAp
         .Selection.WholeStory
         .Selection.Copy
         For i = 0 To int_TotCnt
            strName = ""
            strText = ""
            strLineText = ""
            If i = 0 Then
               strName = "申請案號"
               If pa(11) <> "" Then
                  strText = pa(11)
               Else
                  strText = "         "
               End If
            ElseIf i = 1 Then
               strName = "案號"
               strText = pa(2)
            ElseIf i = 2 Then
               strName = "公告號"
               'Modify By Sindy 2016/11/25 FCP-55818 目前證書號和公告號是一樣的,但舊資料二者還是不同,因此改抓證書號
               'strText = IIf(pa(8) = "1", "■發明", "□發明") & "　" & IIf(pa(8) = "2", "■新型", "□新型") & "　" & IIf(pa(8) <> "1" And pa(8) <> "2", "■新式樣/設計/衍生設計", "□新式樣/設計/衍生設計") & " 第 " & pa(15) & " 號"
               strText = IIf(pa(8) = "1", "■發明", "□發明") & "　" & IIf(pa(8) = "2", "■新型", "□新型") & "　" & IIf(pa(8) <> "1" And pa(8) <> "2", "■新式樣/設計/衍生設計", "□新式樣/設計/衍生設計") & " 第 " & pa(22) & " 號"
            ElseIf i = 3 Then
               strName = "受讓申請人"
               If txtAppNew(1) <> "" Then
                  If Trim(txtAppName(1)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & GetPrjNationName(GetPrjNationNumber1(txtAppNew(1), "CU10", "1"), "NA81", pa(1)) & Trim(txtAppName(1))
                  If Trim(txtAppName(2)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & Trim(txtAppName(2))
               End If
               If txtAppNew(2) <> "" Then
                  If Trim(txtAppName(4)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & GetPrjNationName(GetPrjNationNumber1(txtAppNew(2), "CU10", "1"), "NA81", pa(1)) & Trim(txtAppName(4))
                  If Trim(txtAppName(5)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & Trim(txtAppName(5))
               End If
               If txtAppNew(3) <> "" Then
                  If Trim(txtAppName(7)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & GetPrjNationName(GetPrjNationNumber1(txtAppNew(3), "CU10", "1"), "NA81", pa(1)) & Trim(txtAppName(7))
                  If Trim(txtAppName(8)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & Trim(txtAppName(8))
               End If
               If txtAppNew(4) <> "" Then
                  If Trim(txtAppName(10)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & GetPrjNationName(GetPrjNationNumber1(txtAppNew(4), "CU10", "1"), "NA81", pa(1)) & Trim(txtAppName(10))
                  If Trim(txtAppName(11)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & Trim(txtAppName(11))
               End If
               If txtAppNew(5) <> "" Then
                  If Trim(txtAppName(13)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & GetPrjNationName(GetPrjNationNumber1(txtAppNew(5), "CU10", "1"), "NA81", pa(1)) & Trim(txtAppName(13))
                  If Trim(txtAppName(14)) <> "" Then strText = IIf(strText <> "", strText & vbCrLf, "") & Trim(txtAppName(14))
               End If
            ElseIf i = 4 Then
               strName = "共幾人1"
               strText = intCP56Cnt
            ElseIf i = 5 Then
               strName = "受讓人"
               strText = strCP56Text
               'strLineText = strApplLineText
            ElseIf i = 6 Then
               strName = "出名代理人1"
               strText = PUB_GetAgentCP110(strReceiveNo, m_CP110, pa(1))
            ElseIf i = 7 Then
               strName = "共幾人2"
               strText = intApplCnt
            ElseIf i = 8 Then
               strName = "讓與人"
               strText = strApplText
               'strLineText = strApplLineText2
            ElseIf i = 9 Then
               strName = "出名代理人2"
               strText = GetCP110_2
'               '最近一筆A,B類收文已發文,有主管機關者
'               strExc(0) = "select cp09,cp110" & _
'                           " from caseprogress" & _
'                           " where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                           " and cp110 is not null and cp130 is not null and cp27 is not null" & _
'                           " and cp57 is null" & _
'                           " order by cp27 desc,cp09 asc"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strText = PUB_GetAgentCP110(RsTemp.Fields("cp09"))
'               Else
'                  strText = PUB_GetAgentCP110("")
'               End If
            ElseIf i = 10 Then
               strName = "發文字號"
               If cp(27) = "" Then
                  strText = "發文字號： " & Val(Left(strSrvDate(1), 4)) - 1911 & " 年 " & _
                                         Mid(strSrvDate(1), 5, 2) & " 月 " & _
                                         Right(strSrvDate(1), 2) & " 日(" & Left(strSrvDate(1), 4) - 1911 & ")"
               Else
                  strText = "發文字號： " & Val(Left(DBDATE(cp(27)), 4)) - 1911 & " 年 " & _
                                         Mid(DBDATE(cp(27)), 5, 2) & " 月 " & _
                                         Right(DBDATE(cp(27)), 2) & " 日(" & Left(DBDATE(cp(27)), 4) - 1911 & ")"
               End If
               strText = strText & "晉專外字第　　　　　　　號"
            ElseIf i = 11 Then
               strName = "發文規費"
               strText = IIf(Val(txtCP84) = 0, "　　　", Format(Val(txtCP84), "#,##0"))
            End If
            'Find並且置換
            'Modify By Sindy 2016/5/26 改判斷專利號數
            'If Trim(strName) <> "" And Not (strName = "公告號" And bolHad1603 = False) Then
            If Trim(strName) <> "" And Not (strName = "公告號" And Trim(pa(22)) = "") Then
               .Selection.Find.ClearFormatting
               .Selection.Find.Text = "|#" & strName & "#|"
               .Selection.Find.Replacement.Text = ""
               .Selection.Find.Forward = True
               .Selection.Find.Wrap = wdFindContinue
               .Selection.Find.Format = False
               .Selection.Find.MatchCase = False
               .Selection.Find.MatchWholeWord = False
               .Selection.Find.MatchWildcards = False
               .Selection.Find.MatchSoundsLike = False
               .Selection.Find.MatchAllWordForms = False
               .Selection.Find.MatchByte = True
               .Selection.Find.Execute
               .Selection.Delete
               .Selection.TypeText strText
               If strLineText <> "" Then
                  .Selection.HomeKey
                  .Selection.Find.ClearFormatting
                  With .Selection.Find
                      .Text = strLineText
                      .Replacement.Text = ""
                      .Forward = True
                      .Wrap = wdFindContinue
                      .Format = False
                      .MatchCase = False
                      .MatchWholeWord = False
                      .MatchWildcards = False
                      .MatchSoundsLike = False
                      .MatchAllWordForms = False
                      .MatchByte = True
                  End With
                  .Selection.Find.Execute
                  .Selection.Font.Underline = wdUnderlineSingle
               End If
               If InStr("出名代理人1;出名代理人2", strName) > 0 Then
                  ChgWordFormat g_WordAp.Application, strText
               End If
            End If
ReadNext:
         Next i
      End With
      Screen.MousePointer = vbDefault
'         g_WordAp.ActiveDocument.Save
'         g_WordAp.ActiveDocument.Close
'         MsgBox "檔案已存放在：" & PUB_Getdesktop & "\" & m_TempFileName
      MsgBox "資料已產生完畢!!!"
      'cmdOK(4).Tag = "1" '已執行過
      GetApplBook = True
   Else
      MsgBox "無申請書的樣本!!!"
   End If

   Exit Function
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   End If
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'讓與人之代理人
Private Function GetCP110_2() As String
   m_CP110_2 = ""
   GetCP110_2 = ""
   '最近一筆A,B類收文已發文,有主管機關者
   strExc(0) = "select cp09,cp110" & _
               " from caseprogress" & _
               " where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
               " and cp110 is not null and cp130 is not null and cp27 is not null" & _
               " and cp57 is null" & _
               " order by cp27 desc,cp09 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_CP110_2 = "" & RsTemp.Fields("cp110")
      GetCP110_2 = PUB_GetAgentCP110(RsTemp.Fields("cp09"))
   Else
      m_CP110_2 = ""
      GetCP110_2 = PUB_GetAgentCP110("")
   End If
End Function

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP84_Validate(Cancel As Boolean)
   '台灣
   If pa(9) = "000" Then
      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            txtCP84_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub

'Add by Morgan 2006/6/8
Public Function StopMe() As Boolean
   '若申請人與讓與人不同時提示並結束
   For intI = 1 To 5
      If pa(25 + intI) <> m_Giver(intI) Then
         MsgBox "讓與人(移轉人)與申請人不同，請更正資料後再行發文！", vbCritical
         StopMe = True
         Exit For
      End If
   Next
End Function

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
