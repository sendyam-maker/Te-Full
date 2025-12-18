VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_c 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-讓與"
   ClientHeight    =   6480
   ClientLeft      =   -2328
   ClientTop       =   3456
   ClientWidth     =   8748
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8748
   Begin VB.TextBox txtEmail 
      Height          =   270
      Left            =   8010
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "Y"
      Top             =   4830
      Width           =   345
   End
   Begin VB.TextBox txtRecDate 
      Height          =   270
      Left            =   8010
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4350
      Width           =   345
   End
   Begin VB.TextBox txtPayToday 
      Height          =   264
      Left            =   5115
      MaxLength       =   1
      TabIndex        =   5
      Top             =   5685
      Width           =   255
   End
   Begin VB.TextBox txtCP118 
      Height          =   270
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   4
      Top             =   5670
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "產生申請書"
      Height          =   375
      Index           =   4
      Left            =   3120
      TabIndex        =   80
      Top             =   45
      Visible         =   0   'False
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3525
      Left            =   30
      TabIndex        =   44
      Top             =   2100
      Width           =   7155
      _ExtentX        =   12637
      _ExtentY        =   6223
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "受讓人1"
      TabPicture(0)   =   "frm060104_c.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label24"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label25"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label14(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label5(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCaseField(39)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCaseField(40)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCaseField(41)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCaseField(42)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCaseField(43)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCaseField(44)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtAppName(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtAppName(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtAppName(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtAppNew(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Combo2(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Combo2(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "受讓人2"
      TabPicture(1)   =   "frm060104_c.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "Label18(1)"
      Tab(1).Control(4)=   "Label14(2)"
      Tab(1).Control(5)=   "Label5(29)"
      Tab(1).Control(6)=   "Label5(28)"
      Tab(1).Control(7)=   "Label5(27)"
      Tab(1).Control(8)=   "Label5(26)"
      Tab(1).Control(9)=   "Label5(25)"
      Tab(1).Control(10)=   "Label5(24)"
      Tab(1).Control(11)=   "txtCaseField(45)"
      Tab(1).Control(12)=   "txtCaseField(46)"
      Tab(1).Control(13)=   "txtCaseField(47)"
      Tab(1).Control(14)=   "txtCaseField(48)"
      Tab(1).Control(15)=   "txtCaseField(49)"
      Tab(1).Control(16)=   "txtCaseField(50)"
      Tab(1).Control(17)=   "txtAppName(6)"
      Tab(1).Control(18)=   "txtAppName(5)"
      Tab(1).Control(19)=   "txtAppName(4)"
      Tab(1).Control(20)=   "txtAppNew(2)"
      Tab(1).Control(21)=   "Combo2(2)"
      Tab(1).Control(22)=   "Combo2(3)"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "受讓人3"
      TabPicture(2)   =   "frm060104_c.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(2)=   "Label17"
      Tab(2).Control(3)=   "Label14(3)"
      Tab(2).Control(4)=   "Label5(35)"
      Tab(2).Control(5)=   "Label5(34)"
      Tab(2).Control(6)=   "Label5(33)"
      Tab(2).Control(7)=   "Label14(6)"
      Tab(2).Control(8)=   "Label5(16)"
      Tab(2).Control(9)=   "Label5(17)"
      Tab(2).Control(10)=   "Label5(18)"
      Tab(2).Control(11)=   "txtCaseField(51)"
      Tab(2).Control(12)=   "txtCaseField(52)"
      Tab(2).Control(13)=   "txtCaseField(53)"
      Tab(2).Control(14)=   "txtCaseField(54)"
      Tab(2).Control(15)=   "txtCaseField(55)"
      Tab(2).Control(16)=   "txtCaseField(56)"
      Tab(2).Control(17)=   "txtAppName(9)"
      Tab(2).Control(18)=   "txtAppName(8)"
      Tab(2).Control(19)=   "txtAppName(7)"
      Tab(2).Control(20)=   "txtAppNew(3)"
      Tab(2).Control(21)=   "Combo2(4)"
      Tab(2).Control(22)=   "Combo2(5)"
      Tab(2).ControlCount=   23
      TabCaption(3)   =   "受讓人4"
      TabPicture(3)   =   "frm060104_c.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label19"
      Tab(3).Control(1)=   "Label22"
      Tab(3).Control(2)=   "Label23"
      Tab(3).Control(3)=   "Label14(5)"
      Tab(3).Control(4)=   "Label5(10)"
      Tab(3).Control(5)=   "Label5(11)"
      Tab(3).Control(6)=   "Label5(12)"
      Tab(3).Control(7)=   "Label18(4)"
      Tab(3).Control(8)=   "Label5(19)"
      Tab(3).Control(9)=   "Label5(20)"
      Tab(3).Control(10)=   "Label5(21)"
      Tab(3).Control(11)=   "txtCaseField(57)"
      Tab(3).Control(12)=   "txtCaseField(58)"
      Tab(3).Control(13)=   "txtCaseField(59)"
      Tab(3).Control(14)=   "txtCaseField(60)"
      Tab(3).Control(15)=   "txtCaseField(61)"
      Tab(3).Control(16)=   "txtCaseField(62)"
      Tab(3).Control(17)=   "txtAppName(12)"
      Tab(3).Control(18)=   "txtAppName(11)"
      Tab(3).Control(19)=   "txtAppName(10)"
      Tab(3).Control(20)=   "txtAppNew(4)"
      Tab(3).Control(21)=   "Combo2(6)"
      Tab(3).Control(22)=   "Combo2(7)"
      Tab(3).ControlCount=   23
      TabCaption(4)   =   "受讓人5"
      TabPicture(4)   =   "frm060104_c.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label27"
      Tab(4).Control(1)=   "Label28"
      Tab(4).Control(2)=   "Label29"
      Tab(4).Control(3)=   "Label14(4)"
      Tab(4).Control(4)=   "Label5(1)"
      Tab(4).Control(5)=   "Label5(2)"
      Tab(4).Control(6)=   "Label5(9)"
      Tab(4).Control(7)=   "Label18(3)"
      Tab(4).Control(8)=   "Label5(13)"
      Tab(4).Control(9)=   "Label5(14)"
      Tab(4).Control(10)=   "Label5(15)"
      Tab(4).Control(11)=   "txtCaseField(63)"
      Tab(4).Control(12)=   "txtCaseField(64)"
      Tab(4).Control(13)=   "txtCaseField(65)"
      Tab(4).Control(14)=   "txtCaseField(66)"
      Tab(4).Control(15)=   "txtCaseField(67)"
      Tab(4).Control(16)=   "txtCaseField(68)"
      Tab(4).Control(17)=   "txtAppName(15)"
      Tab(4).Control(18)=   "txtAppName(14)"
      Tab(4).Control(19)=   "txtAppName(13)"
      Tab(4).Control(20)=   "txtAppNew(5)"
      Tab(4).Control(21)=   "Combo2(8)"
      Tab(4).Control(22)=   "Combo2(9)"
      Tab(4).ControlCount=   23
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   9
         ItemData        =   "frm060104_c.frx":008C
         Left            =   -74055
         List            =   "frm060104_c.frx":008E
         Style           =   2  '單純下拉式
         TabIndex        =   149
         Top             =   2280
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   8
         ItemData        =   "frm060104_c.frx":0090
         Left            =   -74055
         List            =   "frm060104_c.frx":0092
         Style           =   2  '單純下拉式
         TabIndex        =   145
         Top             =   1140
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   7
         ItemData        =   "frm060104_c.frx":0094
         Left            =   -74055
         List            =   "frm060104_c.frx":0096
         Style           =   2  '單純下拉式
         TabIndex        =   133
         Top             =   2265
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   6
         ItemData        =   "frm060104_c.frx":0098
         Left            =   -74055
         List            =   "frm060104_c.frx":009A
         Style           =   2  '單純下拉式
         TabIndex        =   129
         Top             =   1140
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   5
         ItemData        =   "frm060104_c.frx":009C
         Left            =   -74055
         List            =   "frm060104_c.frx":009E
         Style           =   2  '單純下拉式
         TabIndex        =   121
         Top             =   2310
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   4
         ItemData        =   "frm060104_c.frx":00A0
         Left            =   -74055
         List            =   "frm060104_c.frx":00A2
         Style           =   2  '單純下拉式
         TabIndex        =   113
         Top             =   1140
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   3
         ItemData        =   "frm060104_c.frx":00A4
         Left            =   -74055
         List            =   "frm060104_c.frx":00A6
         Style           =   2  '單純下拉式
         TabIndex        =   101
         Top             =   2265
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   2
         ItemData        =   "frm060104_c.frx":00A8
         Left            =   -74055
         List            =   "frm060104_c.frx":00AA
         Style           =   2  '單純下拉式
         TabIndex        =   97
         Top             =   1110
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   1
         ItemData        =   "frm060104_c.frx":00AC
         Left            =   945
         List            =   "frm060104_c.frx":00AE
         Style           =   2  '單純下拉式
         TabIndex        =   85
         Top             =   2280
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Index           =   0
         ItemData        =   "frm060104_c.frx":00B0
         Left            =   945
         List            =   "frm060104_c.frx":00B2
         Style           =   2  '單純下拉式
         TabIndex        =   81
         Top             =   1140
         Width           =   6135
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   5
         Left            =   -74640
         MaxLength       =   9
         TabIndex        =   73
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   4
         Left            =   -74610
         MaxLength       =   9
         TabIndex        =   66
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   3
         Left            =   -74640
         MaxLength       =   9
         TabIndex        =   59
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   2
         Left            =   -74640
         MaxLength       =   9
         TabIndex        =   52
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtAppNew 
         Height          =   270
         Index           =   1
         Left            =   360
         MaxLength       =   9
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   13
         Left            =   -72810
         TabIndex        =   76
         Top             =   330
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "7726;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   14
         Left            =   -72810
         TabIndex        =   75
         Top             =   600
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   15
         Left            =   -72810
         TabIndex        =   74
         Top             =   870
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   10
         Left            =   -72780
         TabIndex        =   69
         Top             =   330
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   11
         Left            =   -72780
         TabIndex        =   68
         Top             =   600
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   12
         Left            =   -72780
         TabIndex        =   67
         Top             =   870
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   7
         Left            =   -72810
         TabIndex        =   62
         Top             =   330
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   8
         Left            =   -72810
         TabIndex        =   61
         Top             =   600
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   9
         Left            =   -72810
         TabIndex        =   60
         Top             =   870
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   4
         Left            =   -72810
         TabIndex        =   55
         Top             =   330
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   5
         Left            =   -72810
         TabIndex        =   54
         Top             =   600
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   6
         Left            =   -72810
         TabIndex        =   53
         Top             =   870
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   1
         Left            =   2190
         TabIndex        =   48
         Top             =   330
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   2
         Left            =   2190
         TabIndex        =   47
         Top             =   600
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAppName 
         Height          =   285
         Index           =   3
         Left            =   2190
         TabIndex        =   46
         Top             =   870
         Width           =   4380
         VariousPropertyBits=   679493661
         MaxLength       =   60
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   68
         Left            =   -74055
         TabIndex        =   152
         Top             =   3135
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   67
         Left            =   -74055
         TabIndex        =   151
         Top             =   2850
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   66
         Left            =   -74055
         TabIndex        =   150
         Top             =   2565
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   65
         Left            =   -74055
         TabIndex        =   148
         Top             =   1995
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   64
         Left            =   -74055
         TabIndex        =   147
         Top             =   1710
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   63
         Left            =   -74055
         TabIndex        =   146
         Top             =   1425
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   62
         Left            =   -74055
         TabIndex        =   136
         Top             =   3120
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   61
         Left            =   -74055
         TabIndex        =   135
         Top             =   2835
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   60
         Left            =   -74055
         TabIndex        =   134
         Top             =   2550
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   59
         Left            =   -74055
         TabIndex        =   132
         Top             =   1980
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   58
         Left            =   -74055
         TabIndex        =   131
         Top             =   1695
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   57
         Left            =   -74055
         TabIndex        =   130
         Top             =   1425
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   56
         Left            =   -74055
         TabIndex        =   124
         Top             =   3165
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   55
         Left            =   -74055
         TabIndex        =   123
         Top             =   2880
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   54
         Left            =   -74055
         TabIndex        =   122
         Top             =   2595
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   53
         Left            =   -74055
         TabIndex        =   116
         Top             =   2010
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   52
         Left            =   -74055
         TabIndex        =   115
         Top             =   1710
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   51
         Left            =   -74055
         TabIndex        =   114
         Top             =   1425
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   50
         Left            =   -74055
         TabIndex        =   104
         Top             =   3120
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   49
         Left            =   -74055
         TabIndex        =   103
         Top             =   2835
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   48
         Left            =   -74055
         TabIndex        =   102
         Top             =   2550
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   47
         Left            =   -74055
         TabIndex        =   100
         Top             =   1980
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   46
         Left            =   -74055
         TabIndex        =   99
         Top             =   1680
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   45
         Left            =   -74055
         TabIndex        =   98
         Top             =   1395
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   44
         Left            =   945
         TabIndex        =   88
         Top             =   3120
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   43
         Left            =   945
         TabIndex        =   87
         Top             =   2835
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   42
         Left            =   945
         TabIndex        =   86
         Top             =   2565
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   41
         Left            =   945
         TabIndex        =   84
         Top             =   1995
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   40
         Left            =   945
         TabIndex        =   83
         Top             =   1710
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   60
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   39
         Left            =   945
         TabIndex        =   82
         Top             =   1425
         Width           =   6135
         VariousPropertyBits=   679493659
         BackColor       =   -2147483648
         ForeColor       =   -2147483630
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   15
         Left            =   -74550
         TabIndex        =   160
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   14
         Left            =   -74550
         TabIndex        =   159
         Top             =   1770
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   13
         Left            =   -74550
         TabIndex        =   158
         Top             =   1470
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人9"
         Height          =   180
         Index           =   3
         Left            =   -74910
         TabIndex        =   157
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   9
         Left            =   -74550
         TabIndex        =   156
         Top             =   3180
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   2
         Left            =   -74550
         TabIndex        =   155
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   1
         Left            =   -74550
         TabIndex        =   154
         Top             =   2580
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人10"
         Height          =   180
         Index           =   4
         Left            =   -74910
         TabIndex        =   153
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   210
         Index           =   21
         Left            =   -74550
         TabIndex        =   144
         Top             =   2010
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   210
         Index           =   20
         Left            =   -74550
         TabIndex        =   143
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   210
         Index           =   19
         Left            =   -74550
         TabIndex        =   142
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人7"
         Height          =   210
         Index           =   4
         Left            =   -74910
         TabIndex        =   141
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   210
         Index           =   12
         Left            =   -74550
         TabIndex        =   140
         Top             =   3120
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   210
         Index           =   11
         Left            =   -74550
         TabIndex        =   139
         Top             =   2850
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   210
         Index           =   10
         Left            =   -74550
         TabIndex        =   138
         Top             =   2550
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人8"
         Height          =   210
         Index           =   5
         Left            =   -74910
         TabIndex        =   137
         Top             =   2280
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   18
         Left            =   -74550
         TabIndex        =   128
         Top             =   3210
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   17
         Left            =   -74550
         TabIndex        =   127
         Top             =   2910
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   16
         Left            =   -74550
         TabIndex        =   126
         Top             =   2610
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人6"
         Height          =   180
         Index           =   6
         Left            =   -74910
         TabIndex        =   125
         Top             =   2370
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   33
         Left            =   -74550
         TabIndex        =   120
         Top             =   2010
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   34
         Left            =   -74550
         TabIndex        =   119
         Top             =   1740
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   35
         Left            =   -74550
         TabIndex        =   118
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人5"
         Height          =   180
         Index           =   3
         Left            =   -74910
         TabIndex        =   117
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   24
         Left            =   -74550
         TabIndex        =   112
         Top             =   3135
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   25
         Left            =   -74550
         TabIndex        =   111
         Top             =   2865
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   26
         Left            =   -74550
         TabIndex        =   110
         Top             =   2565
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   27
         Left            =   -74550
         TabIndex        =   109
         Top             =   1995
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   28
         Left            =   -74550
         TabIndex        =   108
         Top             =   1695
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   29
         Left            =   -74550
         TabIndex        =   107
         Top             =   1395
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人3"
         Height          =   180
         Index           =   2
         Left            =   -74910
         TabIndex        =   106
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人4"
         Height          =   180
         Index           =   1
         Left            =   -74910
         TabIndex        =   105
         Top             =   2295
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   8
         Left            =   450
         TabIndex        =   96
         Top             =   3150
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   7
         Left            =   450
         TabIndex        =   95
         Top             =   2850
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   6
         Left            =   450
         TabIndex        =   94
         Top             =   2580
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   5
         Left            =   450
         TabIndex        =   93
         Top             =   2040
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   4
         Left            =   450
         TabIndex        =   92
         Top             =   1740
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   3
         Left            =   450
         TabIndex        =   91
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人1"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   90
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人2"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   89
         Top             =   2310
         Width           =   630
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "日:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   79
         Top             =   870
         Width           =   225
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "英:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   78
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "中:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   77
         Top             =   330
         Width           =   225
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "日:"
         Height          =   180
         Left            =   -73140
         TabIndex        =   72
         Top             =   870
         Width           =   225
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "英:"
         Height          =   180
         Left            =   -73140
         TabIndex        =   71
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "中:"
         Height          =   180
         Left            =   -73140
         TabIndex        =   70
         Top             =   330
         Width           =   225
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "日:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   65
         Top             =   870
         Width           =   225
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "英:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   64
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "中:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   63
         Top             =   330
         Width           =   225
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "日:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   58
         Top             =   870
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "英:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   57
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "中:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   56
         Top             =   330
         Width           =   225
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "日:"
         Height          =   180
         Left            =   1830
         TabIndex        =   51
         Top             =   870
         Width           =   225
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "英:"
         Height          =   180
         Left            =   1830
         TabIndex        =   50
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "中:"
         Height          =   180
         Left            =   1830
         TabIndex        =   49
         Top             =   330
         Width           =   225
      End
   End
   Begin VB.TextBox txtCP84 
      Height          =   288
      Left            =   4770
      TabIndex        =   2
      Top             =   1785
      Width           =   1005
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1785
      Width           =   1005
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   750
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1785
      Width           =   1005
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   6615
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1785
      Width           =   675
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "變更事項(&R)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   10
      Top             =   45
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   6540
      TabIndex        =   11
      Top             =   45
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   7380
      TabIndex        =   12
      Top             =   45
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   375
      Index           =   3
      Left            =   4230
      TabIndex        =   9
      Top             =   45
      Width           =   1110
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060104_c.frx":00B4
      Left            =   990
      List            =   "frm060104_c.frx":00C1
      Style           =   2  '單純下拉式
      TabIndex        =   13
      Top             =   1380
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   990
      MaxLength       =   3
      TabIndex        =   17
      Top             =   150
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   16
      Top             =   150
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2310
      MaxLength       =   1
      TabIndex        =   15
      Top             =   150
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2550
      MaxLength       =   2
      TabIndex        =   14
      Top             =   150
      Width           =   375
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email維護:        (Y:是)"
      Height          =   180
      Left            =   7200
      TabIndex        =   165
      Top             =   4875
      Width           =   1635
   End
   Begin VB.Label lblRecDate 
      AutoSize        =   -1  'True
      Caption         =   "當天報告:         (Y:是)"
      Height          =   180
      Left            =   7200
      TabIndex        =   164
      Top             =   4395
      Width           =   1635
   End
   Begin MSForms.TextBox Text14 
      Height          =   435
      Left            =   930
      TabIndex        =   6
      Top             =   5940
      Width           =   7410
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13070;767"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7200
      TabIndex        =   163
      Top             =   2430
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
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:         (Y/N)"
      Height          =   180
      Left            =   3180
      TabIndex        =   162
      Top             =   5715
      Width           =   2655
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:          (Y: 是)"
      Height          =   180
      Left            =   90
      TabIndex        =   161
      Top             =   5715
      Width           =   2085
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人:"
      Height          =   180
      Left            =   7200
      TabIndex        =   43
      Top             =   2220
      Width           =   945
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   3975
      TabIndex        =   42
      Top             =   1815
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   1965
      TabIndex        =   41
      Top             =   1815
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   135
      X2              =   8600
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   135
      X2              =   8600
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   135
      TabIndex        =   40
      Top             =   1815
      Width           =   585
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   6000
      TabIndex        =   39
      Top             =   1815
      Width           =   585
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   90
      TabIndex        =   38
      Top             =   6000
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   8
      Left            =   1650
      TabIndex        =   37
      Top             =   1380
      Width           =   7020
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "12382;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   7
      Left            =   990
      TabIndex        =   36
      Top             =   1110
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   6
      Left            =   4830
      TabIndex        =   35
      Top             =   900
      Width           =   3840
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6773;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   990
      TabIndex        =   34
      Top             =   900
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   4830
      TabIndex        =   33
      Top             =   690
      Width           =   3840
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6773;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   990
      TabIndex        =   32
      Top             =   690
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   6990
      TabIndex        =   31
      Top             =   450
      Width           =   1230
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2170;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   4830
      TabIndex        =   30
      Top             =   450
      Width           =   1230
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2170;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   990
      TabIndex        =   29
      Top             =   450
      Width           =   1860
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3281;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   150
      TabIndex        =   28
      Top             =   450
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3990
      TabIndex        =   27
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   150
      TabIndex        =   26
      Top             =   180
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   6180
      TabIndex        =   25
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   24
      Top             =   690
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   3990
      TabIndex        =   23
      Top             =   690
      Width           =   675
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   150
      TabIndex        =   22
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Index           =   0
      Left            =   3990
      TabIndex        =   21
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   150
      TabIndex        =   20
      Top             =   1110
      Width           =   675
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   19
      Top             =   1380
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   9
      Left            =   7320
      TabIndex        =   18
      Top             =   1815
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2328;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm060104_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/17 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
Dim pa() As String, m_CP110 As String, m_AgentName As String
Dim intWhere As Integer
Dim m_CP10 As String
Dim m_CP14 As String 'Add By Sindy 2016/11/16
Dim m_blnClkChgEvnBtn As Boolean '是否按下變更事項按鈕
Dim m_CP17 As String '收文規費
Dim m_Giver(1 To 5) As String
Dim m_PA143 As String 'Add by Morgan 2008/3/18
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/20 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_CP60 As String 'Added by Lydia 2015/02/26
Dim m_CP142 As String 'Add by Sindy 2015/12/17
Dim m_CP164 As String 'Add By Sindy 2021/4/20
'Added by Lydia 2018/09/11
Dim m_CP118 As String '是否電子送件
Dim m_CP82 As String '發文時間


Private Sub cmdok_Click(Index As Integer)
 'Added by Lydia 2018/09/11
 Dim strFilePath As String '記錄智慧局收文文號
 Dim strNewCP64 As String '保留進度備註
 'end 2018/09/11
 
   Select Case Index
      Case 0, 3 '確定,同時發文
         If TxtValidate = False Then Exit Sub
         
         'Add by Morgan 2008/3/18
         '若未辦或不辦重新委任時不可發文
         If PUB_Check928NotOk(pa) = True Then
            MsgBox "本案下一程序有重新委任之補文件未辦理，不可發文！"
            Exit Sub
         End If
         '若基本檔年費申請人是否出名為N時提醒存檔將取消
         m_PA143 = pa(143)
         If pa(143) = "N" Then
            MsgBox "年費申請人是否出名現為【N】，存檔時將自動取消！"
            m_PA143 = ""
         End If
         'end 2008/3/18
         
         'Added by Lydia 2018/09/11 是否電子送件
         strNewCP64 = Text14
         If txtCP118 = "Y" Then
            '電子送件也要記錄主管機關
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text9, , True) = False Then
               Exit Sub
            End If
            strExc(0) = InputBox("請輸入智慧局收文文號!!")
            If strExc(0) = "" Then
               Exit Sub
            Else
               strFilePath = strExc(0)  '記錄智慧局收文文號
               strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Text14
            End If
         Else
         'end 2018/09/11
            'Add by Morgan 2009/4/28
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text9) = False Then
               Exit Sub
            End If
            If m_CP123s = "Y" Then
            'end 2009/4/28
               'Add by Morgan 2009/3/20 設定是否算發文室案件
               'modify by sonia 2014/6/23 加傳發文規費, P-108903
               If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, txtCP84, Text9) = False Then
                   Exit Sub
               End If
               'end 2009/3/20
            End If
         End If 'end 2018/09/11
         
'         'Add By Sindy 2016/5/5 檢查是否執行過產生申請書
'         If cmdOK(4).Enabled = True And cmdOK(4).Tag <> "1" Then '1.代表已執行過
'            If MsgBox("尚未產生申請書，確定不出申請書要直接存檔嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
'               Exit Sub
'            End If
'         End If
'         '2016/5/5 END
         
         'Added by Lydia 2018/09/11 依據輸入的智慧局收文號(受理號,ex: 1073066637-0)，將本機C:\E-SET\RdcDocDir\(收文號ex: 1073066637-0)的pdf檔自動搬移到卷宗區(by Phoebe);
         If txtCP118.Text = "Y" And strFilePath <> "" Then
             strExc(1) = m_CP82
             If Val(m_CP82) > 0 Then
                 If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                      strExc(1) = ""
                 End If
             End If
             If Val(strExc(1)) = 0 Then
                'Modified by Lydia 2019/03/22 +傳入發文日
                If Pub_AutoEsetToCpp(True, pa(1), pa(2), pa(3), pa(4), pa(8), Label2(0).Caption, m_CP10, strFilePath, Text9.Text) = False Then
                      Exit Sub
                End If
             End If
         End If
         'end 2018/09/11
         
         'Added by Lydia 2018/09/11 檢查完畢，更新備註欄位
         Text14.Text = strNewCP64
         
         'Add by Sindy 2021/11/17 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         Screen.MousePointer = vbHourglass
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Add by Morgan 2008/2/20 檢查代理人Email
         PUB_CheckEMail pa(75), pa(144)
         If pa(145) <> "" Then
            PUB_CheckEMail pa(75), pa(145)
         End If
         'end 2008/2/20
         
         If pa(1) = "FCP" Then
            'Add By Sindy 2016/11/16 特殊代理人彈訊息提醒
            If (PUB_GetST03(m_CP14) = "F21" Or PUB_GetST03(m_CP14) = "F51" Or PUB_GetST03(m_CP14) = "F52") And _
               Not (m_CP10 = "901" And m_CP10 = "902" And m_CP10 = "1202" And m_CP10 = "1002") Then
               strExc(0) = "select cp130 from caseprogress where cp09='" & strReceiveNo & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If "" & RsTemp.Fields(0) <> "" Then
                     If ChangeCustomerL(pa(75)) = "Y34440B30" Then
                        MsgBox "請當日優先請款報告!!"
                     End If
                  End If
               End If
            End If
         End If
         '2016/11/16 END
            
         Screen.MousePointer = vbDefault
         
         If Index = 0 Then
            'Add By Sindy 2023/11/9
            If frm060104_1.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2023/11/9 End
            '若有未發文資料顯示警告
            'Modify By Sindy 2023/11/9
            If PUB_GetCPunIssueDatas("" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)) Then
               frm060104_1.Show
               frm060104_1.ReQuery
            Else
               'Add By Sindy 2023/11/9
               If frm060104_1.bolIsEMPFlow = True Then
                  Unload frm060104_1
               Else
               '2023/11/9 End
                  frm060104_1.Show
                  frm060104_1.Clear
               End If
            End If
         Else
            frm060104_1.Show
            frm060104_1.ReQuery
         End If
         
         'Add By Sindy 2022/5/12
         If txtEmail.Text = "Y" Then
            frm060104_k.m_CP09 = strReceiveNo 'cp(9)
            frm060104_k.m_strRecDate = txtRecDate
            frm060104_k.Hide
            frm060104_k.cmdOK(0) = 1
            Unload frm060104_k
         End If
         '2022/5/12 END
         
         Unload Me
         
      Case 1 '回前畫面
         frm060104_1.Show
         Unload Me
      Case 2 '變更事項
         Me.Hide
         frm060104_5.LoadMe strReceiveNo, pa(1), pa(2), pa(3), pa(4), 11
         m_blnClkChgEvnBtn = True
      
      'Add By Sindy 2016/4/29 +產生申請書
      Case 4
         If TxtValidate = False Then Exit Sub
         Call GetApplBook
   End Select
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
   Select Case m_CP10
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
         strCP56Text = PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), txtAppNew(1), txtAppNew(2), txtAppNew(3), txtAppNew(4), txtAppNew(5), strApplLineText, intCP56Cnt, intApplCnt, m_CP10, txtCaseField(39), txtCaseField(40), txtCaseField(42), txtCaseField(43), txtCaseField(45), txtCaseField(46), txtCaseField(48), txtCaseField(49), txtCaseField(51), txtCaseField(52), txtCaseField(54), txtCaseField(55), txtCaseField(57), txtCaseField(58), txtCaseField(60), txtCaseField(61), txtCaseField(63), txtCaseField(64), txtCaseField(66), txtCaseField(67))
         '讓與人--CP讓與人或發文前申請人
         strApplText = PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), , , , , , strApplLineText2, , , m_CP10)
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
                  strText = PUB_GetAgentCP110(RsTemp.Fields("cp09"))
               Else
                  strText = PUB_GetAgentCP110("")
               End If
            ElseIf i = 10 Then
               strName = "發文字號"
               If Text9 = "" Then
                  strText = "發文字號： 　 年 　 月 　 日(" & Left(strSrvDate(2), 3) & ")"
               Else
                  strText = "發文字號： " & Val(Left(DBDATE(Text9), 4)) - 1911 & " 年 " & _
                                         Mid(DBDATE(Text9), 5, 2) & " 月 " & _
                                         Right(DBDATE(Text9), 2) & " 日(" & Left(DBDATE(Text9), 4) - 1911 & ")"
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
      cmdOK(4).Tag = "1" '已執行過
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

'Add By Sindy 2016/4/29
'受讓人
Private Function GetCP56Data(intCP56Cnt As Integer, strCP56Text As String) As String
   intCP56Cnt = 0
   strCP56Text = ""
   If txtAppNew(1) <> "" Then
      intCP56Cnt = intCP56Cnt + 1
      strCP56Text = GetCustData(txtAppNew(1), "0", 1, strCP56Text)
   End If
   If txtAppNew(2) <> "" Then
      intCP56Cnt = intCP56Cnt + 1
      strCP56Text = GetCustData(txtAppNew(2), "0", 2, strCP56Text)
   End If
   If txtAppNew(3) <> "" Then
      intCP56Cnt = intCP56Cnt + 1
      strCP56Text = GetCustData(txtAppNew(3), "0", 3, strCP56Text)
   End If
   If txtAppNew(4) <> "" Then
      intCP56Cnt = intCP56Cnt + 1
      strCP56Text = GetCustData(txtAppNew(4), "0", 4, strCP56Text)
   End If
   If txtAppNew(5) <> "" Then
      intCP56Cnt = intCP56Cnt + 1
      strCP56Text = GetCustData(txtAppNew(5), "0", 5, strCP56Text)
   End If
End Function
'讓與人
Private Function GetApplData(intApplCnt As Integer, strApplText As String) As String
   intApplCnt = 0
   strApplText = ""
   If pa(26) <> "" Then
      intApplCnt = intApplCnt + 1
      strApplText = GetCustData(pa(26), "1", 1, strApplText)
   End If
   If pa(27) <> "" Then
      intApplCnt = intApplCnt + 1
      strApplText = GetCustData(pa(27), "1", 2, strApplText)
   End If
   If pa(28) <> "" Then
      intApplCnt = intApplCnt + 1
      strApplText = GetCustData(pa(28), "1", 3, strApplText)
   End If
   If pa(29) <> "" Then
      intApplCnt = intApplCnt + 1
      strApplText = GetCustData(pa(29), "1", 4, strApplText)
   End If
   If pa(30) <> "" Then
      intApplCnt = intApplCnt + 1
      strApplText = GetCustData(pa(30), "1", 5, strApplText)
   End If
End Function
'抓客戶資料
'strType = 0:受讓人 1:讓與人
Private Function GetCustData(strCustNo As String, strType As String, _
                             intNumber As Integer, strText As String) As String
   strCustNo = ChangeCustomerL(strCustNo)
   GetCustData = strText & IIf(strText <> "", vbCrLf, "")
   strExc(0) = "select cu04,cu05||' '||cu88||' '||cu89||' '||cu90 cu05," & _
               "cu23,cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 cu24," & _
               "cu11,cu10,na03,na04" & _
               " From customer,Nation" & _
               " where cu01='" & Left(strCustNo, 8) & "' and cu02='" & Mid(strCustNo, 9, 1) & "' and cu10=na01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetCustData = GetCustData & "姓名或名稱：(中文/英文)　　　ID：" & RsTemp.Fields("cu11") & vbCrLf
      GetCustData = GetCustData & IIf("" & RsTemp.Fields("cu04") <> "", RsTemp.Fields("cu04") & vbCrLf, "")
      GetCustData = GetCustData & IIf(Trim("" & RsTemp.Fields("cu05")) <> "", RsTemp.Fields("cu05") & vbCrLf, "")
      GetCustData = GetCustData & "代表人：(中文/英文)" & vbCrLf
      If strType = "0" Then '受讓人
         If intNumber = 1 Then
            If txtCaseField(39) <> "" Then GetCustData = GetCustData & txtCaseField(39) & vbCrLf
            If txtCaseField(40) <> "" Then GetCustData = GetCustData & txtCaseField(40) & vbCrLf
            If txtCaseField(42) <> "" Then GetCustData = GetCustData & txtCaseField(42) & vbCrLf
            If txtCaseField(43) <> "" Then GetCustData = GetCustData & txtCaseField(43) & vbCrLf
         ElseIf intNumber = 2 Then
            If txtCaseField(45) <> "" Then GetCustData = GetCustData & txtCaseField(45) & vbCrLf
            If txtCaseField(46) <> "" Then GetCustData = GetCustData & txtCaseField(46) & vbCrLf
            If txtCaseField(48) <> "" Then GetCustData = GetCustData & txtCaseField(48) & vbCrLf
            If txtCaseField(49) <> "" Then GetCustData = GetCustData & txtCaseField(49) & vbCrLf
         ElseIf intNumber = 3 Then
            If txtCaseField(51) <> "" Then GetCustData = GetCustData & txtCaseField(51) & vbCrLf
            If txtCaseField(52) <> "" Then GetCustData = GetCustData & txtCaseField(52) & vbCrLf
            If txtCaseField(54) <> "" Then GetCustData = GetCustData & txtCaseField(54) & vbCrLf
            If txtCaseField(55) <> "" Then GetCustData = GetCustData & txtCaseField(55) & vbCrLf
         ElseIf intNumber = 4 Then
            If txtCaseField(57) <> "" Then GetCustData = GetCustData & txtCaseField(57) & vbCrLf
            If txtCaseField(58) <> "" Then GetCustData = GetCustData & txtCaseField(58) & vbCrLf
            If txtCaseField(60) <> "" Then GetCustData = GetCustData & txtCaseField(60) & vbCrLf
            If txtCaseField(61) <> "" Then GetCustData = GetCustData & txtCaseField(61) & vbCrLf
         ElseIf intNumber = 5 Then
            If txtCaseField(63) <> "" Then GetCustData = GetCustData & txtCaseField(63) & vbCrLf
            If txtCaseField(64) <> "" Then GetCustData = GetCustData & txtCaseField(64) & vbCrLf
            If txtCaseField(66) <> "" Then GetCustData = GetCustData & txtCaseField(66) & vbCrLf
            If txtCaseField(67) <> "" Then GetCustData = GetCustData & txtCaseField(67) & vbCrLf
         End If
      Else '讓與人
         If intNumber = 1 Then
            If pa(79) <> "" Then GetCustData = GetCustData & pa(79) & vbCrLf
            If pa(80) <> "" Then GetCustData = GetCustData & pa(80) & vbCrLf
            If pa(82) <> "" Then GetCustData = GetCustData & pa(82) & vbCrLf
            If pa(83) <> "" Then GetCustData = GetCustData & pa(83) & vbCrLf
         ElseIf intNumber = 2 Then
            If pa(109) <> "" Then GetCustData = GetCustData & pa(109) & vbCrLf
            If pa(110) <> "" Then GetCustData = GetCustData & pa(110) & vbCrLf
            If pa(112) <> "" Then GetCustData = GetCustData & pa(112) & vbCrLf
            If pa(113) <> "" Then GetCustData = GetCustData & pa(113) & vbCrLf
         ElseIf intNumber = 3 Then
            If pa(115) <> "" Then GetCustData = GetCustData & pa(115) & vbCrLf
            If pa(116) <> "" Then GetCustData = GetCustData & pa(116) & vbCrLf
            If pa(118) <> "" Then GetCustData = GetCustData & pa(118) & vbCrLf
            If pa(119) <> "" Then GetCustData = GetCustData & pa(119) & vbCrLf
         ElseIf intNumber = 4 Then
            If pa(121) <> "" Then GetCustData = GetCustData & pa(121) & vbCrLf
            If pa(122) <> "" Then GetCustData = GetCustData & pa(122) & vbCrLf
            If pa(124) <> "" Then GetCustData = GetCustData & pa(124) & vbCrLf
            If pa(125) <> "" Then GetCustData = GetCustData & pa(125) & vbCrLf
         ElseIf intNumber = 5 Then
            If pa(127) <> "" Then GetCustData = GetCustData & pa(127) & vbCrLf
            If pa(128) <> "" Then GetCustData = GetCustData & pa(128) & vbCrLf
            If pa(130) <> "" Then GetCustData = GetCustData & pa(130) & vbCrLf
            If pa(131) <> "" Then GetCustData = GetCustData & pa(131) & vbCrLf
         End If
      End If
      GetCustData = GetCustData & "住居所或營業所地址：(中文/英文)" & vbCrLf
      GetCustData = GetCustData & IIf("" & RsTemp.Fields("CU23") <> "", RsTemp.Fields("CU23") & vbCrLf, "")
      GetCustData = GetCustData & IIf(Trim("" & RsTemp.Fields("CU24")) <> "", RsTemp.Fields("CU24") & vbCrLf, "")
      GetCustData = GetCustData & "國籍：(中文/英文)" & RsTemp.Fields("NA03") & "/" & RsTemp.Fields("NA04") & vbCrLf
      GetCustData = GetCustData & "電話及分機："
   End If
End Function
'2016/4/29 END

Private Function FormSave() As Boolean
   Dim intMax As Long
   Dim stUpdateSQL1 As String
   Dim stUpdateSQL2 As String
   Dim stCuID As String, iIdx As Integer
   'Add by Morgan 2011/6/13
   Dim bolChkMemo605 As Boolean, bolChkMemo416 As Boolean
   Dim strOldMemo605 As String, strOldMemo416 As String
   Dim strNewMemo605 As String, strNewMemo416 As String
   '2011/6/13 END
   Dim i As Integer 'Add By Sindy 2016/4/29
   Dim stCP118 As String, stCP152 As String 'Added by Lydia 2018/09/11

On Error GoTo CheckingErr

   cnnConnection.BeginTrans
   
   'Add by Morgan 2011/6/13
   '若有期限則於更新資料前紀錄原來的備註
   'Modified by Morgan 2012/2/2 +pa75
   strExc(0) = "select pa26,pa27,pa28,pa29,pa30,pa75,np07 from nextprogress,patent" & _
      " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07 in (416,605)" & _
      " and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp("np07") = "416" Then
         bolChkMemo416 = True
         'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
         'For intI = 0 To 4
         '   If Not IsNull(RsTemp(intI)) Then
         '      'Modified by Morgan 2012/2/2 +pa75
         '      'Modified by Morgan 2013/9/11 改抓設定檔
         '      'strOldMemo416 = PUB_Get416Memo(ChangeCustomerL(RsTemp(intI)), ChangeCustomerL("" & RsTemp("pa75")))
         '      strOldMemo416 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp("pa75")), ChangeCustomerL(RsTemp(intI)))
         '      If strOldMemo416 <> "" Then Exit For
         '   End If
         'Next
         strOldMemo416 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp.Fields("pa75")), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
         'end 2022/08/02
         
      ElseIf RsTemp("np07") = "605" Then
         bolChkMemo605 = True
         strExc(9) = PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1")
         'Modified by Morgan 2012/6/4 +pa26
         'Modified by Morgan 2013/9/11 改抓設定檔
         'strOldMemo605 = PUB_Get605Memo(strExc(9), RsTemp("pa26"), pa(1) & pa(2) & pa(3) & pa(4))
         'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
         'strOldMemo605 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp("pa26"))
         strOldMemo605 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
      End If
   End If
   'end 2011/6/13
   
   '受讓人1
   stCuID = ChangeCustomerL(txtAppNew(1))
   '讓與申請人1
   stUpdateSQL1 = ",CP56=" & CNULL(stCuID)
   '若原申請人與讓與申請人不同時, 才更新讓與人(避免讓與人欄位被蓋掉)
   If ChangeCustomerL(pa(26)) <> stCuID Then
      '讓與人1
      stUpdateSQL1 = stUpdateSQL1 & ",CP55=" & CNULL(ChangeCustomerL(pa(26)))
      '新申請人1
      stUpdateSQL2 = stUpdateSQL2 & ",PA26=" & CNULL(stCuID)
      '新申請人1地址
      'edit by nickc 2007/02/02 不用 dll 了
      'Call objPublicData.GetCustomerNameAndAddress(stCuID, strExc(0), strExc(1), strExc(2), strExc(3))
      Call ClsPDGetCustomerNameAndAddress(stCuID, strExc(0), strExc(1), strExc(2), strExc(3))
      stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(31) & "=" & CNULL(ChgSQL(Trim(strExc(1))))
      stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(36) & "=" & CNULL(ChgSQL(Trim(strExc(2))))
      stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(41) & "=" & CNULL(ChgSQL(Trim(strExc(3))))
   'Add by Morgan 2006/3/23
   '讓與人與受讓人相同時判斷CP有資料時才不更新
   Else
      '讓與人1
      stUpdateSQL1 = stUpdateSQL1 & ",CP55=NVL(CP55," & CNULL(ChangeCustomerL(pa(26))) & ")"
      'Add By Sindy 2018/7/18 ex:FCP-041158
      '新申請人1地址
      Call ClsPDGetCustomerNameAndAddress(stCuID, strExc(0), strExc(1), strExc(2), strExc(3))
      stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(31) & "=" & CNULL(ChgSQL(Trim(strExc(1))))
      stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(36) & "=" & CNULL(ChgSQL(Trim(strExc(2))))
      stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(41) & "=" & CNULL(ChgSQL(Trim(strExc(3))))
      '2018/7/18 END
   '2006/3/23 end
   End If
   
   '受讓人2~5
   For iIdx = 2 To 5
      stCuID = ChangeCustomerL(txtAppNew(iIdx))
      '讓與申請人
      stUpdateSQL1 = stUpdateSQL1 & ",CP" & Format(87 + iIdx) & "=" & CNULL(stCuID)
      '(原申請人與讓與申請人不同時, 才更新)
      If ChangeCustomerL(pa(25 + iIdx)) <> stCuID Then
         '讓與人
         stUpdateSQL1 = stUpdateSQL1 & ",CP" & Format(91 + iIdx) & "=" & CNULL(ChangeCustomerL(pa(25 + iIdx)))
         '新申請人
         stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(25 + iIdx) & "=" & CNULL(stCuID)
         '新申請人地址
         If stCuID <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'Call objPublicData.GetCustomerNameAndAddress(stCuID, strExc(0), strExc(1), strExc(2), strExc(3))
            Call ClsPDGetCustomerNameAndAddress(stCuID, strExc(0), strExc(1), strExc(2), strExc(3))
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(30 + iIdx) & "=" & CNULL(ChgSQL(Trim(strExc(1))))
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(35 + iIdx) & "=" & CNULL(ChgSQL(Trim(strExc(2))))
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(40 + iIdx) & "=" & CNULL(ChgSQL(Trim(strExc(3))))
         Else
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(30 + iIdx) & "=NULL"
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(35 + iIdx) & "=NULL"
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(40 + iIdx) & "=NULL"
         End If
      'Add by Morgan 2006/3/23
      '讓與人與受讓人相同時判斷CP有資料時才不更新
      Else
         '讓與人
         stUpdateSQL1 = stUpdateSQL1 & ",CP" & Format(91 + iIdx) & "=NVL(CP" & Format(91 + iIdx) & "," & CNULL(ChangeCustomerL(pa(25 + iIdx))) & ")"
         'Add By Sindy 2018/7/18 ex:FCP-041158
         '新申請人地址
         If stCuID <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'Call objPublicData.GetCustomerNameAndAddress(stCuID, strExc(0), strExc(1), strExc(2), strExc(3))
            Call ClsPDGetCustomerNameAndAddress(stCuID, strExc(0), strExc(1), strExc(2), strExc(3))
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(30 + iIdx) & "=" & CNULL(ChgSQL(Trim(strExc(1))))
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(35 + iIdx) & "=" & CNULL(ChgSQL(Trim(strExc(2))))
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(40 + iIdx) & "=" & CNULL(ChgSQL(Trim(strExc(3))))
         Else
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(30 + iIdx) & "=NULL"
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(35 + iIdx) & "=NULL"
            stUpdateSQL2 = stUpdateSQL2 & ",PA" & Format(40 + iIdx) & "=NULL"
         End If
         '2018/7/18 END
      '2006/3/23 end
      End If
   Next
   
   'Add By Sindy 2016/4/29 加存代表人
   For i = 79 To 84
      pa(i) = txtCaseField(i - 40): stUpdateSQL2 = stUpdateSQL2 & ",pa" & i & "=" & CNULL(ChgSQL(pa(i)))
   Next
   For i = 109 To 132
      pa(i) = txtCaseField(i - 64): stUpdateSQL2 = stUpdateSQL2 & ",pa" & i & "=" & CNULL(ChgSQL(pa(i)))
   Next
   '2016/4/29 END
   
   'Added by Lydia 2018/09/11
   '電子送件有規費的一律設自動扣款(同內專) --敏莉
   stCP118 = txtCP118
   stCP152 = ""
   If txtCP118 = "Y" And Val(txtCP84) > 0 Then
      stCP118 = "A"
      stCP152 = Pub_FcpSetPayToday("2", Text9.Text, txtPayToday.Text)
   End If
   'end 2018/09/11
   
   'Modified by Lydia 2018/09/11 +CP118,CP152
   strSql = "UPDATE CASEPROGRESS SET CP27=" & TransDate(Text9, 2) & ", cp14=" & CNULL(Text10) & _
                    ", cp64=" & CNULL(ChgSQL(Text14)) & ", cp84=" & Format(Val(txtCP84.Text)) & _
                    ", CP16=NVL(CP16,0)-NVL(CP17,0)+" & Format(Val(txtCP84.Text)) & _
                    ", CP17=" & Format(Val(txtCP84.Text)) & ", CP18=NVL(CP18,0)" & _
                    ", cp110=" & CNULL(m_CP110) & ",CP22=NULL,CP118='" & stCP118 & "',CP152=" & CNULL(stCP152, True) & " " & stUpdateSQL1 & _
                    " where cp09='" & strReceiveNo & "'"
                    
   cnnConnection.Execute strSql
   
   'Add by Morgan 2008/3/18
   If m_PA143 <> pa(143) Then
      stUpdateSQL2 = stUpdateSQL2 & ",PA143='" & m_PA143 & "'"
   End If
   
   If stUpdateSQL2 <> "" Then
      '要去掉第一個字(",")
      strSql = "Update Patent Set " & Mid(stUpdateSQL2, 2) & _
                        " Where " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      cnnConnection.Execute strSql
   End If
    
   '若有輸入催審期限
   If Me.Text5.Text <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
   'intMax = objPublicData.GetNextProgressNo
   intMax = GetNextProgressNo
      'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
      strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
         "NP07,NP08,NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 催審 & "," & _
         PUB_GetWorkDay1(TransDate(Text5.Text, 2), True) & "," & TransDate(Text5.Text, 2) & ",'" & strUserNum & "'," & intMax & ")"
         
        cnnConnection.Execute strSql
   
      intMax = intMax + 1
   End If
   
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130 'Add by Morgan 2009/3/20
   
   'Add by Morgan 2007/1/30 讓與同時更新下一程序非專業部掌控之案件性質(未續辦)的智權人員
   pub_ChgSalesTargetIsNp pa(1), pa(2), pa(3), pa(4), PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
   
   'Modified by Morgan 2012/2/2 +pa75
   strExc(0) = "select pa26,pa27,pa28,pa29,pa30,pa75 from patent where " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Add by Morgan 2011/6/13
      '若有期限則於更新資料後清除原來的備註並加入新的備註
      If bolChkMemo416 = True Then
         'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
         'For intI = 0 To 4
         '   If Not IsNull(RsTemp(intI)) Then
         '      'Modified by Morgan 2012/2/2 +pa75
         '      'Modified by Morgan 2013/9/11 改抓設定檔
         '      'strNewMemo416 = PUB_Get416Memo(ChangeCustomerL(RsTemp(intI)), ChangeCustomerL("" & RsTemp("pa75")))
         '      strNewMemo416 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp("pa75")), ChangeCustomerL(RsTemp(intI)))
         '      If strNewMemo416 <> "" Then Exit For
         '   End If
         'Next
         strNewMemo416 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL("" & RsTemp.Fields("pa75")), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
         'end 2022/08/02
         
         If strNewMemo416 <> strOldMemo416 Then
            If strOldMemo416 <> "" Then 'Added by Lydia 2022/08/02
              strSql = "update nextprogress set np15=replace(np15,'" & ChgSQL(strOldMemo416) & "','')" & _
                 " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=416"
              cnnConnection.Execute strSql, intI
            End If 'Added by Lydia 2022/08/02
            If strNewMemo416 <> "" Then 'Added by Lydia 2022/08/02
               strSql = "update nextprogress set np15='" & ChgSQL(strNewMemo416) & "'||';'||np15" & _
                   " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=416"
               cnnConnection.Execute strSql, intI
            End If 'Added by Lydia 2022/08/02
         End If
         
      ElseIf bolChkMemo605 = True Then
         strExc(9) = PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1")
         'Modified by Morgan 2012/6/4 +pa26
         'Modified by Morgan 2013/9/11 改抓設定檔
         'strNewMemo605 = PUB_Get605Memo(strExc(9), RsTemp("pa26"), pa(1) & pa(2) & pa(3) & pa(4))
         'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
         'strNewMemo605 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp("pa26"))
         strNewMemo605 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"))
         
         If strNewMemo605 <> strOldMemo605 Then
            If strOldMemo605 <> "" Then 'Added by Lydia 2022/08/02
               strSql = "update nextprogress set np15=replace(np15,'" & ChgSQL(strOldMemo605) & "','')" & _
                   " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=605"
               cnnConnection.Execute strSql, intI
            End If 'Added by Lydia 2022/08/02
            If strNewMemo605 <> "" Then 'Added by Lydia 2022/08/02
              strSql = "update nextprogress set np15='" & ChgSQL(strNewMemo605) & "'||';'||np15" & _
                 " where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " and np06 is null and np07=605"
              cnnConnection.Execute strSql, intI
            End If
         End If
         
      End If
   End If
   'end 2011/6/13
   
 'Added by Lydia 2015/02/26 若已開請款單則換承辦人或核稿人時發Mail通知靜芳
   If m_CP60 > "X" Then
      'Modified by Lydia 2019/10/17 本所案號+"-"
      'PUB_PointReAssignInform Text1 & Text2 & Text3 & Text4, m_CP60, Text10.Tag, Text10.Text
      PUB_PointReAssignInform pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)), m_CP60, Text10.Tag, Text10.Text
   End If
   
   'Added by Lydia 2018/09/11 FCP電子送件若發文時若有規費，則自動產生行事曆。
   If txtCP118 = "Y" And Val(txtCP84) > 0 And stCP152 <> "" Then
       If Pub_AddReceiptCalendar1(pa(1), pa(2), pa(3), pa(4), m_CP10, stCP152) = True Then
       End If
   End If
   'end 2018/09/11
   
   cnnConnection.CommitTrans
   
   FormSave = True
   
CheckingErr:

   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      FormSave = False
   End If
    
End Function

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(8) = pa(5)
      Case "英"
         Label2(8) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label2(8) = pa(7)
   End Select
End Sub

'Add By Sindy 2016/4/28
Private Sub Combo2_Click(Index As Integer)

   Dim i As Integer, strTmp As String
   
   If Combo2(Index) = "" Then
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

Private Sub Form_Activate()
    '若有按下變更事項按鈕, 則重新讀取資料
    If m_blnClkChgEvnBtn = True Then
        ReadPatent
        Label2(0) = strReceiveNo
        m_blnClkChgEvnBtn = False
    End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   With frm060104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   ReDim pa(TF_PA)
   ReadPatent
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1300
   lstNameAgent.Width = 1300

   Label2(0) = strReceiveNo
   Combo1.ListIndex = 0
    m_blnClkChgEvnBtn = False
   SSTab1.Tab = 0 'Added by Morgan 2016/6/2 預設第一頁籤,否則User會誤以為沒有輸入受讓人重複輸到其他受讓人欄位
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_c = Nothing
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As Object
Dim i As Integer
   
   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            For i = 3 To 7
               If pa(i + 23) <> "" Then ChgType (i)
            Next
            Label2(8) = pa(5)
         End If
      Case "FG"
         If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            
         End If
   End Select
   'Modify by Morgan 2006/6/8 加cp93,cp94,cp95,cp96
   'Modified by Lydia 2015/02/26 +cp60
   'Modified by Lydia 2018/09/11 +cp118,cp82
   strExc(0) = "select cp13,st02,cp06,cp27,cp14,cp64,cp56,CP10, CP55, cp17,CP110,CP89,CP90,CP91,CP92,cp93,cp94,cp95,cp96,CP60,CP142,cp118,cp82,CP164" & _
      " from caseprogress,staff where cp09='" & strReceiveNo & "' and cp13=st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         m_CP110 = "" & .Fields("CP110")
         m_CP142 = "" & .Fields("CP142") 'Add By Sindy 2015/12/17
         m_CP164 = "" & .Fields("CP164") 'Add By Sindy 2021/4/20
         m_CP14 = "" & .Fields("CP14") 'Add by Sindy 2016/11/16
         m_CP17 = "" & .Fields("cp17")
         txtCP84.Tag = m_CP17
         txtCP84.Text = txtCP84.Tag
         
         If Not IsNull(.Fields(1)) Then Label2(1) = .Fields(1)
         If Not IsNull(.Fields(2)) Then Label2(2) = TransDate(.Fields(2), 1)
         If Not IsNull(.Fields(3)) Then
            Text9 = TransDate(.Fields(3), 1)
         Else
            Text9 = strSrvDate(2)
         End If
         If Not IsNull(.Fields(4)) Then Text10 = .Fields(4): ChgType (10)
         'Added by Lydia 2015/02/26
         Text10.Tag = Text10.Text
         If Not IsNull(.Fields("CP60")) Then
            m_CP60 = .Fields("CP60")
         Else
            m_CP60 = ""
         End If
         'end 2015/02/26
         If Not IsNull(.Fields(5)) Then Text14 = .Fields(5)
         m_CP10 = "" & .Fields(7).Value
         'Add By Sindy 2016/5/5 有申請書樣本，產生申請書按鈕才要亮起來
         'Modify By Sindy 2017/3/16 合併案和讓與同申請書樣本
         If m_CP10 = 讓與 Or m_CP10 = 合併 Then
            cmdOK(4).Enabled = True
         Else
            cmdOK(4).Enabled = False
         End If
         '2016/5/5 END
         
         If Not IsNull(.Fields("CP56")) Then txtAppNew(1) = .Fields("CP56"): ChgType 11
         If Not IsNull(.Fields("CP89")) Then txtAppNew(2) = .Fields("CP89"): ChgType 12
         If Not IsNull(.Fields("CP90")) Then txtAppNew(3) = .Fields("CP90"): ChgType 13
         If Not IsNull(.Fields("CP91")) Then txtAppNew(4) = .Fields("CP91"): ChgType 14
         If Not IsNull(.Fields("CP92")) Then txtAppNew(5) = .Fields("CP92"): ChgType 15
         
         'Add By Sindy 2016/4/29
         For i = 1 To 5 '讓與申請人代表
            Call SetCombo2(i)
         Next
         '2016/4/29 END
         
         'Add by Morgan 2006/6/8 讓與人
         m_Giver(1) = ChangeCustomerS("" & .Fields("cp55"))
         m_Giver(2) = ChangeCustomerS("" & .Fields("cp93"))
         m_Giver(3) = ChangeCustomerS("" & .Fields("cp94"))
         m_Giver(4) = ChangeCustomerS("" & .Fields("cp95"))
         m_Giver(5) = ChangeCustomerS("" & .Fields("cp96"))
         'Added by Lydia 2018/09/11
          m_CP118 = "" & .Fields("cp118") '電子送件
          If m_CP118 <> "" Then txtCP118.Text = "Y"
        
          m_CP82 = "" & .Fields("cp82") '發文時間
          'end 2018/09/11
      End If
   End With
End Sub

'Add By Sindy 2016/4/29
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
      Case 0 '發文日
         If Not ChkDate(Text9) Then
         ElseIf Val(Text9.Text) > PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1))) Then
            MsgBox "發文日大於系統日下一個工作日, 請重新輸入!!!", vbExclamation + vbOKOnly
         Else
            ChgType = True
         End If
      Case 555 '催審期限
         If ChkDate(Text5) Then
            ChgType = True
         End If
      Case 3, 4, 5, 6, 7
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(i + 23), strTempName) Then
         If ClsLawLawGetName(pa(i + 23), strTempName) Then
            Label2(i) = strTempName
            ChgType = True
         End If

      Case 10
         'ADD BY SONIA 2015/9/21 承辦人為外專程序時,改為操作人員
         Text10 = GetFCPUser(Text10)
         'END 2015/9/21
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(Text10, strTempName) Then
         If ClsPDGetStaff(Text10, strTempName) Then
            Label2(9) = strTempName
            ChgType = True
         End If
         
      Case 11, 12, 13, 14, 15
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.GetCusCAJnam(txtAppNew(i - 10).Text, strExc(1), strExc(2), strExc(3)) = True Then
         If ClsLawGetCusCAJnam(txtAppNew(i - 10).Text, strExc(1), strExc(2), strExc(3)) = True Then
            ChgType = True
            For j = 1 To 3
               txtAppName(3 * (i - 11) + j) = strExc(j)
            Next
         End If
   End Select
End Function

Private Sub Text10_GotFocus()
  TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then
      If Not ChgType(10) Then Cancel = True
   Else
      MsgBox "承辦人不可空白 !", vbCritical
      Cancel = True
   End If
End Sub

Private Sub Text14_GotFocus()
  TextInverse Text14
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> "" Then
      If Not ChgType(555) Then Cancel = True
   End If
   If Cancel = True Then Text5_GotFocus
End Sub


Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_LostFocus()
   ' 重新計算催審期限
   ReCaculateSpecDate
End Sub

' 重新計算催審期限
Private Sub ReCaculateSpecDate()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   
   Text5.Text = ""
   strSql = "SELECT CF05 FROM CASEFEE " & _
            "WHERE CF01='" & pa(1) & "' AND " & _
                  "CF02='" & pa(9) & "' AND " & _
                  "CF03='" & m_CP10 & "'"
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("CF05")) And rsTmp.Fields("CF05") <> 0 Then
         Text5.Text = TransDate(CompDate(2, Val(rsTmp.Fields("CF05")), TransDate(Text9, 2)), 1)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> "" Then
      If Not ChgType(0) Then
            Cancel = True
      'Added by Lydia 2018/09/11 當發文日有改時
      Else
            If Text9.Tag <> Text9 Then
                  Text9.Tag = Text9
                  txtPayToday.Text = Pub_FcpSetPayToday("1", Text9.Text, txtCP118.Text)
            End If
      'end 2018/09/11
      End If
   Else
      MsgBox "發文日不可空白 !", vbCritical
      Cancel = True
   End If
End Sub

Private Function TxtValidate() As Boolean

   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean

   TxtValidate = False

   If Me.Text9.Enabled = True Then
      Cancel = False
      Text9_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If

   If Me.Text5.Enabled = True Then
      Cancel = False
      Text5_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If

   If Me.Text10.Enabled = True Then
      Cancel = False
      Text10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If

   If txtCP84.Enabled = True Then
      Cancel = False
      txtCP84_Validate Cancel
      If Cancel = True Then
         txtCP84.SetFocus
         txtCP84_GotFocus
         Exit Function
      End If
   End If

   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
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
   
   'Add By Sindy 2015/12/17 檢查是否有指定送件日期,若有不可小於指定日期送件
   If m_CP142 <> "" Then
      'Modify By Sindy 2021/11/11 淑華說之後可以含當天發文
      'If m_CP142 >= strSrvDate(1) Then
      If m_CP142 > strSrvDate(1) Then
         'Add By Sindy 2021/4/20
         'Modify By Sindy 2021/10/20 + 3.之後
         If ((m_CP164 = "1" Or m_CP164 = "") And m_CP142 > strSrvDate(1)) Or _
            m_CP164 = "3" Then '1.當天 3.之後
         '2021/4/20 END
            MsgBox "有指定送件日期（" & ChangeWStringToTDateString(m_CP142) & "），不可提前送件!!!"
            Exit Function
         End If
      End If
   End If
   '2015/12/17 END
   
   'Added by Lydia 2018/09/11
   If txtCP118 = "Y" And Val(txtCP84) > 0 Then
      If txtPayToday = "" Then
         MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
         txtPayToday.SetFocus
         Exit Function
      End If
   End If
   'end 2018/09/11
   
   TxtValidate = True
End Function

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

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP84_Validate(Cancel As Boolean)
   '台灣
   If pa(9) = "000" Then
      If Val(txtCP84.Text) <> Val(m_CP17) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & m_CP17 & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            txtCP84_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub

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

'Added by Lydia 2018/09/11
Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
   CloseIme
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP118_Change()
    txtPayToday.Text = Pub_FcpSetPayToday("1", Text9.Text, txtCP118.Text)
End Sub

Private Sub txtPayToday_GotFocus()
   TextInverse txtPayToday
   CloseIme
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2018/09/11

'Add By Sindy 2022/5/17
Private Sub txtRecDate_GotFocus()
   TextInverse txtRecDate
End Sub
Private Sub txtRecDate_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtEmail_GotFocus()
   TextInverse txtEmail
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtRecDate_Validate(Cancel As Boolean)
   If txtRecDate.Tag <> txtRecDate.Text Then
      If txtRecDate = "Y" Then
         txtEmail = "Y"
      End If
   End If
   txtRecDate.Tag = txtRecDate.Text
End Sub
'2022/5/17 END
