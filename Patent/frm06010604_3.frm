VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010604_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "一般來函輸入"
   ClientHeight    =   5940
   ClientLeft      =   156
   ClientTop       =   960
   ClientWidth     =   9012
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9012
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   8412
      TabIndex        =   63
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6930
      TabIndex        =   61
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7512
      TabIndex        =   62
      Top             =   0
      Width           =   900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4680
      Left            =   90
      TabIndex        =   42
      Top             =   1230
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   8255
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "來函輸入1"
      TabPicture(0)   =   "frm06010604_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label22"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label24"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label26"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label28"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label29"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label3(7)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label43"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label37"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label9"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text29"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblDelivery"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "MSHFlexGrid1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text14(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text16"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text17"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text18"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text7"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Frame2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Frame1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text14(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text8"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text9"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text6"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text14(2)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtDelivery"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "來函輸入2"
      TabPicture(1)   =   "frm06010604_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text23"
      Tab(1).Control(1)=   "Text19"
      Tab(1).Control(2)=   "Text27(2)"
      Tab(1).Control(3)=   "Text27(1)"
      Tab(1).Control(4)=   "Text27(0)"
      Tab(1).Control(5)=   "Text26"
      Tab(1).Control(6)=   "Text20(0)"
      Tab(1).Control(7)=   "Text20(1)"
      Tab(1).Control(8)=   "Text20(2)"
      Tab(1).Control(9)=   "Text20(3)"
      Tab(1).Control(10)=   "Text20(4)"
      Tab(1).Control(11)=   "Text20(5)"
      Tab(1).Control(12)=   "Text31"
      Tab(1).Control(13)=   "Label6"
      Tab(1).Control(14)=   "Label30"
      Tab(1).Control(15)=   "Label31"
      Tab(1).Control(16)=   "Label32"
      Tab(1).Control(17)=   "Label33"
      Tab(1).Control(18)=   "Label34"
      Tab(1).Control(19)=   "Label35"
      Tab(1).Control(20)=   "Label36"
      Tab(1).Control(21)=   "Label46"
      Tab(1).Control(22)=   "Label45"
      Tab(1).Control(23)=   "Label44"
      Tab(1).Control(24)=   "Label42"
      Tab(1).Control(25)=   "Label41"
      Tab(1).Control(26)=   "Label40"
      Tab(1).Control(27)=   "Label39"
      Tab(1).Control(28)=   "Label38"
      Tab(1).ControlCount=   29
      TabCaption(2)   =   "被舉發"
      TabPicture(2)   =   "frm06010604_3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label47"
      Tab(2).Control(1)=   "Label15(0)"
      Tab(2).Control(2)=   "Label16(5)"
      Tab(2).Control(3)=   "chkItem(6)"
      Tab(2).Control(4)=   "chkItem(0)"
      Tab(2).Control(5)=   "chkItem(1)"
      Tab(2).Control(6)=   "chkItem(4)"
      Tab(2).Control(7)=   "chkItem(3)"
      Tab(2).Control(8)=   "chkItem(5)"
      Tab(2).Control(9)=   "chkItem(2)"
      Tab(2).Control(10)=   "txtItemCount"
      Tab(2).Control(11)=   "txtItemList"
      Tab(2).Control(12)=   "txtDay(0)"
      Tab(2).Control(13)=   "txtDay(1)"
      Tab(2).Control(14)=   "txtMonth(1)"
      Tab(2).Control(15)=   "txtYear(1)"
      Tab(2).Control(16)=   "txtYear(0)"
      Tab(2).Control(17)=   "txtMonth(0)"
      Tab(2).ControlCount=   18
      Begin VB.TextBox txtDelivery 
         Height          =   270
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   111
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1800
         Width           =   1152
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   5640
         MaxLength       =   30
         TabIndex        =   17
         Top             =   2412
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72570
         MaxLength       =   2
         TabIndex        =   100
         Top             =   1680
         Width           =   285
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -73290
         MaxLength       =   3
         TabIndex        =   99
         Top             =   1680
         Width           =   420
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70905
         MaxLength       =   3
         TabIndex        =   102
         Top             =   1680
         Width           =   420
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70230
         MaxLength       =   2
         TabIndex        =   103
         Top             =   1680
         Width           =   285
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -69645
         MaxLength       =   2
         TabIndex        =   104
         Top             =   1680
         Width           =   285
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72030
         MaxLength       =   2
         TabIndex        =   101
         Top             =   1680
         Width           =   285
      End
      Begin VB.TextBox txtItemList 
         Enabled         =   0   'False
         Height          =   450
         Left            =   -74460
         TabIndex        =   93
         Text            =   "第項"
         Top             =   1200
         Width           =   2580
      End
      Begin VB.TextBox txtItemCount 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -72075
         TabIndex        =   91
         Top             =   750
         Width           =   375
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷設計專利權"
         Enabled         =   0   'False
         Height          =   210
         Index           =   2
         Left            =   -71310
         TabIndex        =   94
         Top             =   780
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "專利權人所屬國家對中華民國申請專利不予受理者"
         Height          =   210
         Index           =   5
         Left            =   -71310
         TabIndex        =   97
         Top             =   1410
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "專利權人為非專利申請權人者"
         Height          =   210
         Index           =   3
         Left            =   -71310
         TabIndex        =   95
         Top             =   990
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "共有專利申請權非由全體共有人提出申請者"
         Height          =   210
         Index           =   4
         Left            =   -71310
         TabIndex        =   96
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "被請求撤銷部分之請求項："
         Height          =   210
         Index           =   1
         Left            =   -74730
         TabIndex        =   92
         Top             =   990
         Width           =   2625
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "被請求撤銷全部請求項：共計"
         Height          =   210
         Index           =   0
         Left            =   -74730
         TabIndex        =   90
         Top             =   780
         Width           =   2715
      End
      Begin VB.TextBox Text23 
         Height          =   270
         Left            =   -70095
         MaxLength       =   3
         TabIndex        =   24
         Top             =   930
         Width           =   684
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   972
         Width           =   7065
      End
      Begin VB.ComboBox Text8 
         Height          =   276
         ItemData        =   "frm06010604_3.frx":0054
         Left            =   5640
         List            =   "frm06010604_3.frx":0056
         TabIndex        =   1
         Text            =   "Text8"
         Top             =   672
         Width           =   1332
      End
      Begin VB.TextBox Text19 
         Height          =   270
         Left            =   -73320
         TabIndex        =   23
         Top             =   972
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   6870
         MaxLength       =   7
         TabIndex        =   13
         Top             =   1800
         Width           =   1272
      End
      Begin VB.Frame Frame1 
         Height          =   552
         Left            =   1440
         TabIndex        =   76
         Top             =   1212
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "文到當日"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "文到次日"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   4
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   552
         Left            =   4200
         TabIndex        =   75
         Top             =   1212
         Width           =   4332
         Begin VB.TextBox Text10 
            Height          =   270
            Left            =   780
            MaxLength       =   2
            TabIndex        =   6
            Top             =   200
            Width           =   375
         End
         Begin VB.TextBox Text12 
            Height          =   270
            Left            =   2880
            MaxLength       =   7
            TabIndex        =   10
            Top             =   200
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Height          =   270
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   8
            Top             =   200
            Width           =   375
         End
         Begin VB.OptionButton Option4 
            Caption         =   "文到           天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1370
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                        日"
            Height          =   225
            Index           =   2
            Left            =   2640
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "          月"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   7
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   0
         Top             =   672
         Width           =   855
      End
      Begin VB.TextBox Text27 
         Height          =   270
         Index           =   2
         Left            =   -73320
         MaxLength       =   1
         TabIndex        =   31
         Top             =   3072
         Width           =   375
      End
      Begin VB.TextBox Text27 
         Height          =   270
         Index           =   1
         Left            =   -73320
         MaxLength       =   1
         TabIndex        =   22
         Top             =   672
         Width           =   375
      End
      Begin VB.TextBox Text27 
         Height          =   270
         Index           =   0
         Left            =   -70440
         MaxLength       =   1
         TabIndex        =   21
         Top             =   372
         Width           =   375
      End
      Begin VB.TextBox Text26 
         Height          =   270
         Left            =   -73320
         MaxLength       =   7
         TabIndex        =   20
         Top             =   372
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Height          =   270
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   16
         Top             =   2412
         Width           =   495
      End
      Begin VB.TextBox Text17 
         Height          =   270
         Left            =   5640
         MaxLength       =   7
         TabIndex        =   15
         Top             =   2112
         Width           =   1272
      End
      Begin VB.TextBox Text16 
         Height          =   270
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   14
         Top             =   2112
         Width           =   1152
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   4080
         MaxLength       =   7
         TabIndex        =   12
         Top             =   1800
         Width           =   1152
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1572
         Left            =   1440
         TabIndex        =   18
         Top             =   2712
         Width           =   7092
         _ExtentX        =   12510
         _ExtentY        =   2773
         _Version        =   393216
         Cols            =   12
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
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷自「        年      月      日」至「        年      月      日」之專利權期間延長"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   -74730
         TabIndex        =   98
         Top             =   1710
         Width           =   7440
      End
      Begin VB.Label lblDelivery 
         AutoSize        =   -1  'True
         Caption         =   "送達日期:"
         Height          =   180
         Left            =   240
         TabIndex        =   110
         Top             =   384
         Width           =   768
      End
      Begin MSForms.TextBox Text29 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   4320
         Width           =   7095
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12515;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   285
         Index           =   0
         Left            =   -73185
         TabIndex        =   25
         Top             =   1275
         Width           =   6615
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "11668;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   285
         Index           =   1
         Left            =   -73200
         TabIndex        =   26
         Top             =   1575
         Width           =   6615
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "11668;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   285
         Index           =   2
         Left            =   -73200
         TabIndex        =   27
         Top             =   1875
         Width           =   6615
         VariousPropertyBits=   671105051
         MaxLength       =   100
         Size            =   "11668;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   285
         Index           =   3
         Left            =   -73320
         TabIndex        =   28
         Top             =   2175
         Width           =   6735
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "11880;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   285
         Index           =   4
         Left            =   -73320
         TabIndex        =   29
         Top             =   2475
         Width           =   6735
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "11880;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   285
         Index           =   5
         Left            =   -73320
         TabIndex        =   30
         Top             =   2775
         Width           =   6735
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "11880;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text31 
         Height          =   1005
         Left            =   -73320
         TabIndex        =   32
         Top             =   3375
         Width           =   6735
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "11880;1773"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         Caption         =   "約定期限:"
         Height          =   255
         Left            =   240
         TabIndex        =   109
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "法院案號:"
         Height          =   255
         Left            =   4560
         TabIndex        =   108
         Top             =   2412
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "項"
         Height          =   180
         Index           =   5
         Left            =   -71670
         TabIndex        =   107
         Top             =   795
         Width           =   180
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "被請求撤銷全部專利權"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   -71310
         TabIndex        =   106
         Top             =   540
         Width           =   1950
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "被請求撤銷發明(新型)專利權"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74730
         TabIndex        =   105
         Top             =   540
         Width           =   2490
      End
      Begin VB.Label Label6 
         Caption         =   "對造案件數代號:"
         Height          =   180
         Left            =   -71535
         TabIndex        =   89
         Top             =   972
         Width           =   1305
      End
      Begin VB.Label Label37 
         Caption         =   "本案期限:"
         Height          =   252
         Left            =   240
         TabIndex        =   85
         Top             =   2712
         Width           =   852
      End
      Begin VB.Label Label43 
         Caption         =   "進度備註:"
         Height          =   252
         Left            =   240
         TabIndex        =   84
         Top             =   4332
         Width           =   852
      End
      Begin VB.Label Label30 
         Caption         =   "對造號數:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   83
         Top             =   972
         Width           =   852
      End
      Begin VB.Label Label31 
         Caption         =   "對造案件名稱(中):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   82
         Top             =   1272
         Width           =   1452
      End
      Begin VB.Label Label32 
         Caption         =   "對造案件名稱(英):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   81
         Top             =   1572
         Width           =   1452
      End
      Begin VB.Label Label33 
         Caption         =   "對造案件名稱(外):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   80
         Top             =   1872
         Width           =   1452
      End
      Begin VB.Label Label34 
         Caption         =   "對造名稱(中):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   79
         Top             =   2172
         Width           =   1212
      End
      Begin VB.Label Label35 
         Caption         =   "對造名稱(英):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   78
         Top             =   2472
         Width           =   1212
      End
      Begin VB.Label Label36 
         Caption         =   "對造名稱(外):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   77
         Top             =   2772
         Width           =   1212
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Index           =   7
         Left            =   2685
         TabIndex        =   74
         Top             =   2115
         Width           =   1650
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2910;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   69
         Top             =   675
         Width           =   2010
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3545;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Index           =   3
         Left            =   7020
         TabIndex        =   68
         Top             =   672
         Width           =   1620
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label46 
         Caption         =   "案件備註:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   60
         Top             =   3372
         Width           =   852
      End
      Begin VB.Label Label45 
         Caption         =   "(Y:自動代繳)"
         Height          =   252
         Left            =   -72840
         TabIndex        =   59
         Top             =   3072
         Width           =   1092
      End
      Begin VB.Label Label44 
         Caption         =   "領證自動代繳:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   58
         Top             =   3072
         Width           =   1332
      End
      Begin VB.Label Label42 
         Caption         =   "(N:不請款)"
         Height          =   252
         Left            =   -72720
         TabIndex        =   57
         Top             =   672
         Width           =   852
      End
      Begin VB.Label Label41 
         Caption         =   "是否向客戶請款:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   56
         Top             =   672
         Width           =   1332
      End
      Begin VB.Label Label40 
         Caption         =   "(Y:閉卷)"
         Height          =   252
         Left            =   -69840
         TabIndex        =   55
         Top             =   372
         Width           =   852
      End
      Begin VB.Label Label39 
         Caption         =   "是否閉卷:"
         Height          =   252
         Left            =   -71520
         TabIndex        =   54
         Top             =   372
         Width           =   852
      End
      Begin VB.Label Label38 
         Caption         =   "專利權消滅日:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   53
         Top             =   372
         Width           =   1212
      End
      Begin VB.Label Label29 
         Caption         =   "(N:不算)"
         Height          =   252
         Left            =   2040
         TabIndex        =   52
         Top             =   2412
         Width           =   732
      End
      Begin VB.Label Label28 
         Caption         =   "是否算案件數:"
         Height          =   252
         Left            =   240
         TabIndex        =   51
         Top             =   2412
         Width           =   1332
      End
      Begin VB.Label Label26 
         Caption         =   "承辦人:"
         Height          =   252
         Left            =   240
         TabIndex        =   50
         Top             =   2112
         Width           =   972
      End
      Begin VB.Label Label25 
         Caption         =   "承辦期限:"
         Height          =   252
         Left            =   4560
         TabIndex        =   49
         Top             =   2112
         Width           =   972
      End
      Begin VB.Label Label23 
         Caption         =   "本所期限:"
         Height          =   255
         Left            =   3060
         TabIndex        =   48
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "法定期限:"
         Height          =   255
         Left            =   5880
         TabIndex        =   47
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "來函期限:"
         Height          =   252
         Left            =   240
         TabIndex        =   46
         Top             =   1404
         Width           =   852
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "下一程序:"
         Height          =   180
         Left            =   4560
         TabIndex        =   45
         Top             =   660
         Width           =   768
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "來函性質:"
         Height          =   180
         Left            =   240
         TabIndex        =   44
         Top             =   672
         Width           =   768
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "機關文號:"
         Height          =   180
         Left            =   240
         TabIndex        =   43
         Top             =   972
         Width           =   768
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   37
      Top             =   36
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   36
      Top             =   36
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   35
      Top             =   36
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   34
      Top             =   36
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4680
      TabIndex        =   33
      Top             =   36
      Width           =   1575
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   64
      Top             =   330
      Width           =   7785
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "13732;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(Y:閉卷)"
      Height          =   180
      Index           =   4
      Left            =   7920
      TabIndex        =   88
      Top             =   930
      Width           =   645
   End
   Begin VB.Label lblPA57 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   7200
      TabIndex        =   87
      Top             =   930
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否閉卷"
      Height          =   180
      Index           =   3
      Left            =   6240
      TabIndex        =   86
      Top             =   930
      Width           =   720
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   6
      Left            =   4680
      TabIndex        =   73
      Top             =   930
      Width           =   1350
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2381;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   72
      Top             =   930
      Width           =   1350
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2381;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日"
      Height          =   180
      Index           =   2
      Left            =   3600
      TabIndex        =   71
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "收文號"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   70
      Top             =   930
      Width           =   540
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   67
      Top             =   630
      Width           =   1350
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2381;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   66
      Top             =   630
      Width           =   1350
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2381;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   65
      Top             =   636
      Width           =   768
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   41
      Top             =   36
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3600
      TabIndex        =   40
      Top             =   36
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   39
      Top             =   336
      Width           =   768
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Left            =   3600
      TabIndex        =   38
      Top             =   636
      Width           =   588
   End
End
Attribute VB_Name = "frm06010604_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/22 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'Modified by Morgan 2021/8/12 智財法院-->智商法院
Option Explicit

Dim strReceiveNo As String, strTemp As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String, cp(1 To T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer, intLastRow As Integer
Public MPa9 As String

'Add By Cheng 2002/01/28
Dim m_strCP09ByCheng As String '總收文號

Dim m_928Upd As Boolean '是否更新重新委任准駁
Dim m_928CP09 As String '重新委任收文號
Dim m_CP16 As String '預設請款金額
Dim m_blnClosed As Boolean '是否閉卷'Add By Sindy 2012/3/5
Dim oChk As CheckBox 'Added by Morgan 2012/11/13
Dim m_strMemo As String 'C類來函接洽單備註 ADD BY SONIA 2014/5/28

'Added by Morgan 2015/5/20
Dim m_stUPA(4) As String '一案兩請新型案號
Dim m_NewCP09 As String '來函收文號
'Added by Morgan 2017/5/10 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_DocDate As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/5/10
Dim stCP133 As String 'Added by Morgan 2020/11/13
Dim m_pAgreeOnDate As String 'Modify By Sindy 2021/4/27
Dim m_1812CP07 As String 'Added by Lydia 2022/05/10 通知補充聽證資料之期限
Dim m_PA177 As String 'Added by Lydia 2023/07/28 FCP專利連結通知
'Added by Lydia 2023/09/25
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_bolReKeyInOK As Boolean '是否與2次確認期限一致
'end 2023/09/25
Dim m_strExtNote As String 'Added by Morgan 2024/5/21 特殊訊息(訴願案的智慧局答辯函)，要加在承辦人通知主旨和C類接洽單上
Public m_UpdCP09 As String 'Added by Morgan 2024/11/15 更新期限的收文號

'Added by Morgan 2012/11/13
'Modified by Morgan 2013/1/14 增加舉發事項
Private Sub chkItem_Click(Index As Integer)
   Dim ii As Integer
      
   If Me.ActiveControl <> chkItem(Index) Then Exit Sub
   
   txtItemCount.Enabled = False
   txtItemList.Enabled = False
   txtYear(0).Enabled = False
   txtYear(1).Enabled = False
   txtMonth(0).Enabled = False
   txtMonth(1).Enabled = False
   txtDay(0).Enabled = False
   txtDay(1).Enabled = False
   
   If Index = 0 Or Index = 1 Then
      If chkItem(Index).Value = vbChecked Then
         For Each oChk In chkItem
            If oChk.Index <> Index Then
               oChk.Value = vbUnchecked
            End If
         Next
         
         Select Case Index
         Case 0
            txtItemCount.Enabled = True
            txtItemCount.SetFocus
            
         Case 1
            txtItemList.Enabled = True
            txtItemList.SetFocus
            If Left(txtItemList, 1) = "第" Then
               txtItemList.SelStart = 1
               txtItemList.SelLength = 0
            End If
         End Select
      End If
   ElseIf Index = 6 Then
      If chkItem(Index).Value = vbChecked Then
         For Each oChk In chkItem
            If oChk.Index <> Index Then
               oChk.Value = vbUnchecked
            End If
         Next
         
         txtYear(0).Enabled = True
         txtYear(0).SetFocus
         txtYear(1).Enabled = True
         txtMonth(0).Enabled = True
         txtMonth(1).Enabled = True
         txtDay(0).Enabled = True
         txtDay(1).Enabled = True
      End If
   ElseIf chkItem(Index).Value = vbChecked Then
      chkItem(0).Value = vbUnchecked
      chkItem(1).Value = vbUnchecked
      chkItem(6).Value = vbUnchecked
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)

   Select Case Index
      Case 0
         'Modified by Morgan 2022/2/18 因偶爾會執行兩次,改寫成函式呼叫以便鎖住按鈕
         cmdOK(Index).Enabled = False
         If Process() = False Then
            cmdOK(Index).Enabled = True
         End If
         'end 2022/2/18
      Case 1
         frm06010604_2.Show
         Unload Me
      Case 2
         Unload frm06010604_2
         Unload frm06010604_1
         Unload Me
   End Select
End Sub

Private Sub StartLetter(ByVal ET01 As String, ET02 As String, ByVal ET03 As String)
   Dim strTxt() As String, i As Integer
   Dim dblExRate As Double
   
   EndLetter ET01, ET02, ET03, strUserNum
   
   i = 0
   If m_stUPA(1) <> "" Then
      strExc(0) = "select pa11,pa77,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo from patent where pa01='" & m_stUPA(1) & "' and pa02='" & m_stUPA(2) & "' and pa03='" & m_stUPA(3) & "' and pa04='" & m_stUPA(4) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','新型案申請號','" & RsTemp("pa11") & "')"
            
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','新型案彼所案號','" & RsTemp("pa77") & "')"
         
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','新型案本所案號','" & RsTemp("CaseNo") & "')"
         
         dblExRate = PUB_GetUSXRate
         
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','美金費用','" & Format(Fix(2500 / dblExRate), "##0") & "')"
         
         'Add By Sindy 2021/8/3
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','約定期限','" & DBDATE(Text14(2)) & "')"
         '2021/8/3 END
      End If
   End If
   
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Function FormSave() As Boolean
   Dim intMax As Long, intStep As Integer, strTxt(1 To 10) As String, strTmp As String, i As Integer
   'edit by nickc 2007/02/02
   'Dim Ncp(1 To T_CP) As String
   Dim Ncp() As String
   'add by nickc 2007/02/02
   ReDim Ncp(1 To TF_CP) As String
   Dim strSql As String
   Dim strNP22 As String
   'Add By Cheng 2002/01/29
   Dim BlnCheck As Boolean '判斷是否有勾選本案期限
   Dim strDate1 As String '本所期限
   Dim strDate2 As String '法定期限
   Dim strCP20 As String, strCP16 As String
   Dim bolReKeyInCase As Boolean 'Added by Lydia 2023/09/25
   Dim bolAddNP As Boolean 'Added by Morgan 2023/11/29
   
   m_928Upd = PUB_928Check(pa, m_928CP09) 'Add by Morgan 2007/7/18
   bolReKeyInCase = False 'Added by Lydia 2023/09/25
   
   FormSave = True
   
On Error GoTo CheckingErr

   cnnConnection.BeginTrans

   'Add by Morgan 2007/7/18
   If m_928Upd = True And m_928CP09 <> "" Then
      PUB_928Update pa, m_928CP09
   End If
   'end 2007/7/18
   
   strNP22 = Empty
   'edit by nickc 2007/02/02 不用 dll 了
   'intMax = objPublicData.GetNextProgressNo
   intMax = GetNextProgressNo
   intStep = 1
   
   'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：收到學名藥廠P4通知1922,來函有專利號:(1)程序輸入C類來函「P4通知1922」(承辦期限3個工作天, 法定期限45天)
   If pa(1) = "FCP" And Text7 = "1922" And m_PA177 = "Y" Then
      strExc(1) = CompDate(2, 45, DBDATE(Label3(6)))
      strExc(2) = CompWorkDay(4, strExc(1), 1)
      Text14(1) = TransDate(strExc(1), 1)
      Text17 = TransDate(strExc(2), 1)
   End If
   'FG案:收到食藥署轉交第三人通知1923,承辦期限7個工作天
   If pa(1) = "FG" And Text7 = "1923" Then
      strExc(2) = CompWorkDay(8, DBDATE(Label3(6)))
      Text17 = TransDate(strExc(2), 1)
   End If
   'end 2023/07/28
   '1
      Ncp(1) = cp(1)
      Ncp(2) = cp(2)
      Ncp(3) = cp(3)
      Ncp(4) = cp(4)
      Ncp(5) = Label3(6)
      Ncp(6) = Text14(0)
      Ncp(7) = Text14(1)
      Ncp(8) = Text9
      'Modify by Morgan 2011/2/24 修正百年收文號問題
      'Ncp(9) = "C" & Left(strSrvDate(2), 2)
      Ncp(9) = "C" & CompAutoNumberYear(GetTaiwanThisYear)
      Ncp(10) = Text7
      'Ncp(12) = cp(12) 'Removed by Morgan 2012/7/24
'      Ncp(13) = cp(13)
        'Modify By Cheng 2003/04/07
        '智權人員存國家檔FCP承辦智權人員
      Ncp(13) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
      Ncp(12) = GetSalesArea(Ncp(13)) 'Added by Morgan 2012/7/24
      Ncp(14) = Text16
'2009/4/1 MODIFY BY SONIA 靜芳說全部掛承辦期限FCP-026581
'      If Text8 <> "" Then
''         Ncp(14) = Text16
'         Ncp(48) = Text17
'      Else
''         Ncp(14) = strUserNum
'      End If
      Ncp(48) = Text17
'2009/4/1 END
      
      'Modified by Lydia 2024/05/28 改成模組
      ''Added by Lydia 2022/05/03 FCP-062174審定前不收費控制: (補上)判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
      'If pa(16) = "" And InStr("FCP062174000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      '      Text27(1) = "N":  Ncp(16) = "":   Ncp(17) = "":   Ncp(18) = ""
      ''FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
      'ElseIf pa(16) <> "1" And InStr("FCP067004000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      If PUB_GetCP20forSpec(pa(1), pa(2), pa(3), pa(4), pa(16)) = "N" Then
      'end by Lydia 2022/05/03
            Text27(1) = "N":  Ncp(16) = "":   Ncp(17) = "":   Ncp(18) = ""
      End If
      'end by Lydia 2022/05/03
      
      'Added by Morgan 2025/3/4
      'Y2099001Murgitroyd+Meta集團(X80668000、 X80669000、X80670000)案件，1002核駁、1202審查意見通知函、1227最後通知預設不請款並備註"簡單報告"
      'Modify By Sindy 2025/6/19 +1221通知申復
      If pa(1) = "FCP" And (Text7 = "1202" Or Text7 = "1227" Or Text7 = "1221") And pa(75) = "Y2099001" And InStr("X80668000,X80669000,X80670000", pa(26)) > 0 Then
         Text29 = "簡單報告;" & Text29
      End If
      'end 2025/3/4
      
      'Modify by Morgan 2007/7/24 改為輸 N 與欄位一致
      'If Text27(1) = "Y" Then
      '   Ncp(20) = ""
      'Else
      '   Ncp(20) = "N"
      'End If
      Ncp(20) = Text27(1)
      If Ncp(20) = "" Then
         Ncp(16) = Val(m_CP16)
         Ncp(17) = 0
         Ncp(18) = Val(m_CP16) / 1000
      End If
      'end 2007/7/24
      
      
      '2010/11/17 modify by sonia 取消撤銷原處分准/駁欄,此處一定准,駁改在核駁輸
      'If Text7 = 撤銷原處分 Then Ncp(24) = Text13
      If Text7 = 撤銷原處分 Then Ncp(24) = "1"
      '2010/11/17 end
      If Text7 = 專利權消滅 Then Ncp(25) = Text26
      
      ' 承辦期限
      Ncp(48) = Empty
      '92.3.23 ADD BY SONIA
      If Text7 = 撤銷原處分 Or Text7 = 通知智慧局答辯函 Then
'         Ncp(14) = Text16
         Ncp(48) = Text17
         '2013//19 add by sonia
         If Me.Text7.Text = "1506" And Label3(1) = "行政訴訟" Then
            Ncp(35) = Text6
         End If
         '2013/8/19 end
      End If
      '92.3.23 END
      Ncp(26) = Text18
      'If Text8 = "" Then Ncp(27) = strSrvDate(2)
      Ncp(32) = "N"
      
      'Modify by Morgan 2008/2/25 加對照案件數代號
      'Ncp(36) = Text19
      Ncp(36) = Text19.Text & Text23.Text
      'end 2008/2/25
      
      For i = 0 To 5
         ' 91.01.22 modify by louis
         'Ncp(i + 37) = Text20(i)
         Ncp(i + 37) = ChgSQL(Text20(i))
      Next
      Ncp(43) = cp(9)
      
      'Added by Morgan 2012/11/13
      'Modified by Morgan 2013/1/14 增加舉發事項
      If SSTab1.TabVisible(2) = True Then
         For Each oChk In chkItem
            If oChk.Value = vbChecked Then
               If oChk.Index = 0 Then
                  Text29.Text = oChk.Caption & txtItemCount & "項;" & Text29
               ElseIf oChk.Index = 1 Then
                  Text29.Text = oChk.Caption & txtItemList & ";" & Text29
               ElseIf oChk.Index = 6 Then
                  Text29.Text = "請求撤銷自「" & txtYear(0) & "年" & txtMonth(0) & "月" & txtDay(0) & "日」至「" & txtYear(1) & "年" & txtMonth(1) & "月" & txtDay(1) & "日」之專利權期間延長;" & Text29.Text
               Else
                  Text29.Text = oChk.Caption & ";" & Text29
               End If
            End If
         Next
      End If
      'end 2012/11/13
      
      '2005/12/26 MODIFY BY SONIA
      'Ncp(64) = Text31
      Ncp(64) = Text29
      
      ' 發文日
      Ncp(27) = Empty
'CANCEL BY SONIA 2014/5/13 下面有寫
'      Select Case Text7
'         ' 來函性質為通知文件, 通知閱卷, 通知變更, 延長審查時間, 通知領証時應自動上發文日
'         'Case "1003", "1402", "1403", "1501", "1601", "1901", "1902", "1904", "1905":
'         '   Ncp(27) = strSrvDate(2)
'         ' 通知智慧局答辯函, 對方撤回時不需要發文日
'         'Case "1506", "1808":
'         '   Ncp(27) = Empty
'         'Case Else
'         '   If IsEmptyText(Text8) = True Then
'         '      Ncp(27) = strSrvDate(2)
'         '   End If
'         ' modify by sonia 90.7.20
'         '下列來函性質不自動上發文日, 由工程師發文
'         Case 通知修正, 通知申復, 通知補充說明, 通知改請發明, 通知改請新型, 通知改請設計
'            Ncp(27) = Empty
'         Case 通知改請追加, 通知改請聯合, 通知改請獨立, 通知分割, 通知面詢, 撤銷原處分
'            Ncp(27) = Empty
'         Case 通知參加訴願, 通知參加訴訟, 通知智慧局答辯函, 被異議理由, 被舉發理由
'            Ncp(27) = Empty
'         Case 發回補理由, 發回補答辯, 對方補充說明, 對方撤回, 所外鑑定報告結果, "1507"
'            Ncp(27) = Empty
'         Case Else
'            Ncp(27) = strSrvDate(2)
'      End Select
'END 2014/5/13
      'Add By Cheng 2002/05/31
      '若案件性質為下列者, 則發文日預設為 NULL
        'Modify By Cheng 2002/12/19
'      If (Me.Text7.Text >= "1201" And Me.Text7.Text <= "1203") Or _
'         Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or (Me.Text7.Text >= "1301" And Me.Text7.Text <= "1307") Or _
'         Me.Text7.Text = "1401" Or Me.Text7.Text = "1502" Or (Me.Text7.Text >= "1504" And Me.Text7.Text <= "1507") Or _
'         Me.Text7.Text = "1801" Or Me.Text7.Text = "1802" Or (Me.Text7.Text >= "1805" And Me.Text7.Text = "1808") Or _
'         Me.Text7.Text = "1903" Then
      'MODIFY BY SONIA 2014/5/13 +1227最後通知FCP-035287
      'Modified by Morgan 2016/5/19 +1232通知擇一申復--何淑華 Ex.FCP-49198
      'Modify By Sindy 2016/5/31 + 通知補文件1003,相關總收文號為201,235,209,210 不上發文日
      'Modified by Morgan 2024/4/26 +對方補充答辯狀 1809,經濟部答辯函 1508 也不上發文日
      'Modified by Morgan 2024/5/16 +1815第三方意見
      'Modify By Sindy 2025/6/19 +1221通知申復
      If Val(Me.Text7.Text) = 1001 Or Val(Me.Text7.Text) = 1002 Or (Val(Me.Text7.Text) >= 1201 And Val(Me.Text7.Text) <= 1203) Or _
         (Val(Me.Text7.Text) >= 1210 And Val(Me.Text7.Text) <= 1212) Or (Val(Me.Text7.Text) >= 1301 And Val(Me.Text7.Text) <= 1307) Or _
         Val(Me.Text7.Text) = 1401 Or Val(Me.Text7.Text) = 1502 Or (Val(Me.Text7.Text) >= 1504 And Val(Me.Text7.Text) <= 1508) Or _
         Val(Me.Text7.Text) = 1801 Or Val(Me.Text7.Text) = 1802 Or (Val(Me.Text7.Text) >= 1805 And Val(Me.Text7.Text) <= 1809) Or _
         Val(Me.Text7.Text) = 1903 Or Val(Me.Text7.Text) = 1227 Or Val(Me.Text7.Text) = 1232 Or Val(Me.Text7.Text) = 1815 Or _
         (Val(Me.Text7.Text) = 1003 And (cp(10) = "201" Or cp(10) = "235" Or cp(10) = "209" Or cp(10) = "210")) Or _
         Val(Me.Text7.Text) = 1221 Then
         Ncp(27) = Empty
        'Add By Cheng 2002/12/19
        '其餘案件性質發文日為系統日
      'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：收到學名藥廠P4通知1922,來函有專利號
      ElseIf pa(1) = "FCP" And Text7 = "1922" And m_PA177 = "Y" Then
         Ncp(27) = Empty   '「P4通知」發文時,自動收文告代=>不要上發文日
      'FG案:收到食藥署轉交第三人通知1923,承辦期限7個工作天
      ElseIf pa(1) = "FG" And Text7 = "1923" Then
         Ncp(27) = Empty
      'end 2023/07/28
      Else
         'Modify By Sindy 2022/5/9 修改智慧局來函kEY"其他來函"(電子公文或紙本來函)，若承辦人掛工程師，請不要自動上發文日
         If GetSalesArea(Text16) = "F21" And Val(Me.Text7.Text) = 1902 Then
            Ncp(27) = Empty
         Else
         '2022/5/9 END
            Ncp(27) = strSrvDate(2)
         End If
      End If
      
'Remove by Morgan 2007/7/24 改抓 CPM 設定
'      '92.10.19 ADD BY SONIA 預設服務費4000
'      If (Val(Me.Text7.Text) >= 1201 And Val(Me.Text7.Text) <= 1203) Or _
'         (Val(Me.Text7.Text) >= 1210 And Val(Me.Text7.Text) <= 1212) Or _
'         (Val(Me.Text7.Text) >= 1301 And Val(Me.Text7.Text) <= 1307) Or _
'         (Val(Me.Text7.Text) >= 1502 And Val(Me.Text7.Text) <= 1503) Or _
'         (Val(Me.Text7.Text) >= 1506 And Val(Me.Text7.Text) <= 1508) Or _
'         (Val(Me.Text7.Text) >= 1801 And Val(Me.Text7.Text) <= 1802) Or _
'         (Val(Me.Text7.Text) >= 1805 And Val(Me.Text7.Text) <= 1807) Then
'         Ncp(16) = 4000
'         Ncp(17) = 0
'         Ncp(18) = 4
'      Else
'         Ncp(16) = Empty
'         Ncp(17) = Empty
'         Ncp(18) = Empty
'      End If
'end 2007/7/14

'2009/4/1 MODIFY BY SONIA 取消限制條件靜芳說全部掛承辦期限FCP-026581
'      Select Case Text7
'         ' 通知閱卷, 通知變更, 延期受理
'         Case "1402", "1403", "1004", "1901", "1404":
'            Ncp(48) = Empty
'         Case Else:
'            If IsEmptyText(Text8) = False Then
'               Ncp(48) = Text17
'            End If
'      End Select
      Ncp(48) = Text17
'2009/4/1 end

      Ncp(133) = stCP133 'Added by Morgan 2020/11/13
      Ncp(134) = Val(Text11) 'Added by Morgan 2020/5/15
      
      'ADD BY SONIA 2014/5/28 Intersil及其子公司的案件在C類接洽單加印
      m_strMemo = ""
      'Remove by Morgan 2017/3/9 改從備註維護功能自行設定(與其他的內容合併)--敏莉 Ex.FCP-49174
      'Select Case Left(pa(26) & "000", 8)
      '   Case "X6217700", "X5272200", "X5422700", "X5819500", "X6380100", "X6554500", "X6036001", "X4899100", "X4899101"
      '      m_strMemo = "若有報價請一併CC給Intersil ！"
      'End Select
      'end 2017/3/9
      'END 2014/5/9
   
      'Add By Cheng 2002/01/28
      m_strCP09ByCheng = Empty
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.SaveNewCaseProgressDatabase("C", Ncp, intWhere, m_strCP09ByCheng) Then
      If Not ClsPDSaveNewCaseProgressDatabase("C", Ncp, intWhere, m_strCP09ByCheng) Then 'Memo by Lydia 2023/07/26 新增資料至CaseProgress基本檔Ncp(9)
         '911106 nick transation
         cnnConnection.RollbackTrans
         FormSave = False
         Exit Function
      End If
      
   '2
   'Add By Sindy 2012/3/5 原基本檔未閉卷時,才要更新 +And m_blnClosed = False
   If Text27(0) = "Y" And m_blnClosed = False Then
      'Modify by Morgan 2007/2/1 專利存是否存在不必上'N'
      'strTxt(intStep) = "UPDATE PATENT SET PA57='Y',PA17='N' WHERE " & ChgPatent(pA(1) & pA(2) & pA(3) & pA(4))
      'Add By Sindy 2012/3/5 +PA58,PA59及Update servicepractice
      'strTxt(intStep) = "UPDATE PATENT SET PA57='Y' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      If pa(1) = "FCP" Then
         strTxt(intStep) = "UPDATE PATENT SET PA57='Y',PA58=" & strSrvDate(1) & ",PA59='99' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      Else
         strTxt(intStep) = "UPDATE servicepractice SET sp15='Y',sp16=" & strSrvDate(1) & ",sp17='99' WHERE " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
      End If
      '2012/3/5 End
      'End 2007/2/1
      '911106 nick transation
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   'Add By Sindy 2012/3/5
   ElseIf Text27(0) = "" Then
      If pa(1) = "FCP" Then
         strTxt(intStep) = "UPDATE PATENT SET PA57=null,PA58=null,PA59=null WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      Else
         strTxt(intStep) = "UPDATE servicepractice SET sp15=null,sp16=null,sp17=null WHERE " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
      End If
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   '2012/3/5 End
   
   'Added by Lydia 2025/08/19 輸入C類來函時，去檢查上一道承辦人掛工程師，是否為未請款，若是，則發Mail通知工程師；
   If pa(1) = "FCP" And Text16 <> "" Then
      If PUB_ChkFCPtoCP14CP60(pa(1), pa(2), pa(3), pa(4), Text7, Ncp(9), Text16) = True Then
      End If
   End If
   'end 2025/08/19
   
   '3
   '911028 nick 重新抓
   'edit by nickc 2007/02/02 不用 dll 了
   'intMax = objPublicData.GetNextProgressNo
   intMax = GetNextProgressNo
   If Text7 = 所外鑑定報告結果 Then
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07='" & 催審 & "'"
        '911106 nick transation
        cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1' WHERE CP09='" & strReceiveNo & "'"
        '911106 nick transation
        cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   '4
   ElseIf Text7 = 撤銷原處分 Then
      strTmp = CompDate(1, 3, TransDate(Label3(6).Caption, 2)) '函收文日加3個月
      
      'strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
      '   "NP07,NP08,NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
      '   "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 改變原處分 & "," & _
      '   strTmp & "," & strTmp & ",'" & strUserNum & "'," & intMax & ")"
      ' 90.06.22 modify by Louis 一般來函的進度備註同時存入下一程序檔
      'Modify By Cheng 2002/06/03
      '新增至下一程序檔的智權人員, 原以點選收文資料的智權人員, 現改為抓專利基本檔, 若有代理人, 以代理人之國籍抓國家檔的"CFP承辦智權人員", 若無代理人, 則以申請人1之國籍抓國家檔的"FCP承辦智權人員"
'      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
'         "NP07,NP08,NP09,NP10,NP15,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
'         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 改變原處分 & "," & _
'         strTmp & "," & strTmp & ",'" & strUserNum & "'," & CNULL(Text29) & "," & intMax & ")"
        '911106 nick 重新抓
        'edit by nickc 2007/02/02 不用 dll 了
        'intMax = objPublicData.GetNextProgressNo
        intMax = GetNextProgressNo
        'Modify By Cheng 2003/04/07
        '智權人員存國家檔FCP承辦智權人員
        'Modify by Morgan 2011/10/12 智權人員改放程序,因為是要催智慧局(和內專一樣)
        '改變原處分(1503) => 重為處分(1503)
        'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
        strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
            "NP07,NP08,NP09,NP10,NP15,NP22) VALUES ('" & Ncp(9) & "','" & pa(1) & _
            "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 改變原處分 & "," & _
            PUB_GetWorkDay1(strTmp, True) & "," & strTmp & ",'" & strUserNum & "'," & CNULL(Text29) & "," & intMax & ")"
        '911106 nick transation
        cnnConnection.Execute strTxt(intStep)
        intMax = intMax + 1
        intStep = intStep + 1
   End If
   
   'Add by Morgan 2004/11/10 若下一程序有催審則上不續辦
   If Text7 = 撤銷原處分 Or Text7 = 通知面詢 Or Text7 = 對方撤回 Then
      strSql = "Update NextProgress Set NP06='Y' Where NP01='" & cp(9) & "' and NP06 is null and NP07='411'"
      cnnConnection.Execute strSql
   End If
   '2004/11/10 end
   '2010/11/17 ADD BY SONIA 原點選進度上核准
   If Text7 = 撤銷原處分 Then
      strSql = "Update CASEProgress Set CP24='1',CP25=" & DBDATE(Label3(6)) & " Where CP09='" & cp(9) & "' and CP24 is null AND CP25 IS NULL"
      cnnConnection.Execute strSql
   End If
   '2010/11/17 END
   
   '5若有輸入下一程序
   If Text8 <> "" Then
      'Add By Sindy 2016/4/18
      'Modified by Morgan 2023/11/29 沒勾選下一程序時要新增NP
      'If Me.Text7.Text = 通知補文件 And InStr(NewCasePtyList, cp(10)) > 0 And Text8 = 補文件 Then
      bolAddNP = True
      If Me.Text7.Text = 通知補文件 And InStr(NewCasePtyList, cp(10)) > 0 And Text8 = 補文件 Then
         With MSHFlexGrid1
         For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "v" Then
               bolAddNP = False
               Exit For
            End If
         Next
         End With
      'end 2023/11/29
      
         'A.依點選資料去更新np的期限,不再新增np資料
         '後面程式已經會更新勾選的期限
'         With MSHFlexGrid1
'            For i = 1 To .Rows - 1
'               If .TextMatrix(i, 0) = "v" Then
'                  strSql = "UPDATE NEXTPROGRESS SET NP08=" & DBDATE(Me.Text14(0).Text) & ",NP09=" & DBDATE(Me.Text14(1).Text) & " WHERE NP22=" & .TextMatrix(i, 8) & " and np01='" & .TextMatrix(i, 9) & "'"
'                  cnnConnection.Execute strSql
'               End If
'            Next
'         End With
         'B.進度檔若有201,209,210,235未發文未取消收文的資料時,同時更新該筆期限,若該筆已有期限且不相同時則先詢問是否更新
         strExc(0) = "SELECT cp09,cp06,cp07,decode('" & pa(9) & "','000',cpm03,cpm04) CSName FROM CASEPROGRESS,casepropertymap WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND " & "CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                     " AND CP10 in('201','209','210','235') AND CP27 IS NULL AND CP57 IS NULL" & _
                     " AND cp01=cpm01(+) AND cp10=cpm02(+)" & _
                     " ORDER BY CP09 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            strTmp = ""
            If Val("" & RsTemp.Fields("cp06")) > 0 And Val("" & RsTemp.Fields("cp06")) <> DBDATE(Me.Text14(0).Text) Then
               strTmp = "本所期限=" & RsTemp.Fields("cp06")
            End If
            If Val("" & RsTemp.Fields("cp07")) > 0 And Val("" & RsTemp.Fields("cp07")) <> DBDATE(Me.Text14(1).Text) Then
               strTmp = strTmp & "法定期限=" & RsTemp.Fields("cp07")
            End If
            If strTmp <> "" Then
               If MsgBox("此案" & RsTemp.Fields("CSName") & "已有期限(" & strTmp & ")，是否要更新為此通知補文件的新期限？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                  strTmp = "" '代表要更新
               End If
            End If
            If strTmp = "" Then '代表要更新
               strSql = "Update CaseProgress Set CP06=" & DBDATE(Me.Text14(0).Text) & ",CP07=" & DBDATE(Me.Text14(1).Text) & " Where CP09='" & RsTemp.Fields("cp09") & "'"
               cnnConnection.Execute strSql
            End If
         End If
      'Modified by Morgan 2023/11/29
      'Else
      End If
      If bolAddNP = True Then
      'end 2023/11/29
      '2016/4/18 END
         
         If Text20(3) <> "" Then
            strTmp = Text20(3)
         ElseIf Text20(4) <> "" Then
            strTmp = Text20(4)
         ElseIf Text20(5) <> "" Then
            strTmp = Text20(5)
         End If
      'Modify By Cheng 2002/06/03
'      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
'         "NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) VALUES ('" & Ncp(9) & "','" & pa(1) & _
'         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & Text8 & "," & _
'         TransDate(Text14(0), 2) & "," & TransDate(Text14(1), 2) & "," & CNULL(cp(13)) & _
'         "," & CNULL(Text9) & "," & CNULL(strTmp) & "," & CNULL(Text29) & "," & intMax & ")"
      'Modify By Cheng 2002/09/25
'      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
'         "NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) VALUES ('" & Ncp(9) & "','" & pa(1) & _
'         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & Text8 & "," & _
'         TransDate(Text14(0), 2) & "," & TransDate(Text14(1), 2) & ",'" & PUB_GetSaleNo(pa(1), pa(2), pa(3), pa(4)) & _
'         "'," & CNULL(Text9) & "," & CNULL(strTmp) & "," & CNULL(Text29) & "," & intMax & ")"
         '911106 nick 重新抓
         'edit by nickc 2007/02/02 不用 dll 了
         'intMax = objPublicData.GetNextProgressNo
         intMax = GetNextProgressNo
         'Modify By Cheng 2003/04/07
         '智權人員存國家檔FCP承辦智權人員
         'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(TransDate(Text14(2), 2)):約定期限
         strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
            "NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22,NP23) VALUES ('" & Ncp(9) & "','" & pa(1) & _
            "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & Text8 & "," & _
            TransDate(Text14(0), 2) & "," & TransDate(Text14(1), 2) & ",'" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & _
            "'," & CNULL(Text9) & "," & CNULL(ChgSQL(strTmp)) & "," & CNULL(Text29) & "," & intMax & "," & CNULL(TransDate(Text14(2), 2)) & ")"
         '911106 nick transation
         cnnConnection.Execute strTxt(intStep)
         
         ' 將下一程序資料的序號存起來
         strNP22 = CStr(intMax)
         intMax = intMax + 1
         intStep = intStep + 1
         With MSHFlexGrid1
            For i = 1 To .Rows - 1
               If .TextMatrix(i, 0) = "v" Then
                  'Add By Cheng 2003/01/17
                  '若來函性質不為通知補文件, 延期受理
                  If Me.Text7.Text <> 通知補文件 And Me.Text7.Text <> 延期受理 Then
                     'Modify By Cheng 2003/01/17
                     '序號已向後移一欄
      '               strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP22=" & .TextMatrix(i, 7)
                     'Modify by Morgan 2006/1/24 加NP01
                     strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP22=" & .TextMatrix(i, 8) & " and np01='" & .TextMatrix(i, 9) & "'"
                     '911106 nick transation
                     cnnConnection.Execute strTxt(intStep)
                     intStep = intStep + 1
                  End If
               End If
            Next
         End With
      End If
   End If
   '91.12.1 CANCEL BY SONIA
   '6
   'If Text8 = "" And Text14(0) <> "" And Text14(1) <> "" Then
   '   strTxt(intStep) = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Text14(0), 2) & _
   '      ",CP07=" & TransDate(Text14(1), 2) & " WHERE CP09='" & strReceiveNo & "'"
   '     '911106 nick transation
   '     cnnConnection.Execute strTxt(intStep)
   '   intStep = intStep + 1
   'End If
   '91.12.1 END
   'Add By Cheng 2001/12/31
   '若來函性質為爭議程序(18XX), 則更新專利基本檔的是否有爭議欄(PA19)為"Y"
   If Left(Me.Text7.Text, 2) = "18" Then
        strTxt(intStep) = "UPDATE PATENT SET PA19='Y' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        '911106 nick transation
        cnnConnection.Execute strTxt(intStep)
        intStep = intStep + 1
   End If
   '911106 nick transation
   'FormSave = objLawDll.ExecSQL(intStep - 1, strTxt)
   ' 列印接洽結案單
   'Modify By Cheng 2001/12/20
'   If IsEmptyText(strNP22) = False Then
'      g_PrtForm001.PrintForm strNP22, pa(1), pa(2), pa(3), pa(4)
'   End If
   '911106 nick 移到最下面
   '8
   'If Text7 = 被異議理由 Then
   '   strExc(0) = "SELECT MAX(CP05),CP09 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND " & _
   '      "CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='" & 領證及繳年費 & "' AND CP27 IS NULL AND CP57 IS NULL GROUP BY CP09"
   '   intI = 1
   '   Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   '   ' 90.06.26 modify by louis 無資料時不做
   '   If rsTemp.RecordCount > 0 Then
   '      rsTemp.MoveFirst
   '      If rsTemp.Fields(0) > 0 Then
   '         frm06010604_31.Tag = "6" & rsTemp.Fields(1)
   '         frm06010604_31.Show vbModal
   '      End If
   '   End If
   'End If
   '
   'Add By Cheng 2002/01/29
   If FormSave = True Then
      ' 90.06.26 modify by louis, 來函性質為專用權消滅1604時, 更新基本檔專用權是否存在為N
      If Text7 = "1604" Then
         strSql = "UPDATE PATENT SET PA17 = 'N' " & _
                  "WHERE PA01 = '" & pa(1) & "' AND " & _
                        "PA02 = '" & pa(2) & "' AND " & _
                        "PA03 = '" & pa(3) & "' AND " & _
                        "PA04 = '" & pa(4) & "' "
         cnnConnection.Execute strSql
      End If
      'Modify By Cheng 2003/01/17
      'Add By Cheng 2002/01/29
      '若來函性質為通知補文件(1003), 延期受理(1004), 若未勾選本案期限, 則以相關總收文號去更新案件進度檔的本所期限及法定期限
      '若來函性質為通知補文件(1003), 延期受理(1004), 若有勾選本案期限, 則更新下一程序檔的本所期限及法定期限
      If Me.Text7.Text = 通知補文件 Or Me.Text7.Text = 延期受理 Then
         BlnCheck = False
         With Me.MSHFlexGrid1
            For i = 1 To .Rows - 1
               If LCase("" & .TextMatrix(i, 0)) = "v" Then
                  BlnCheck = True
                    'Modify By Cheng 2003/01/17
                    '修改本所期限及法定期限
'                  '本所期限
'                  strDate1 = IIf(Len(Trim(.TextMatrix(i, 2))) > 0, _
'                              IIf(Len(Trim(.TextMatrix(i, 2)) <> 8), Val(Replace(.TextMatrix(i, 2), "/", "")) + 19110000, Replace(.TextMatrix(i, 2), "/", "")), _
'                              "")
'                  '法定期限
'                  strDate2 = IIf(Len(Trim(.TextMatrix(i, 3))) > 0, _
'                              IIf(Len(Trim(.TextMatrix(i, 3)) <> 8), Val(Replace(.TextMatrix(i, 3), "/", "")) + 19110000, Replace(.TextMatrix(i, 3), "/", "")), _
'                              "")
                  '本所期限
                  strDate1 = DBDATE(Me.Text14(0).Text)
                  '法定期限
                  strDate2 = DBDATE(Me.Text14(1).Text)
                    'Modify By Cheng 2003/01/17
                    '序號已向後移一欄
'                  strSQL = "UPDATE NEXTPROGRESS SET NP08='" & strDate1 & "' " & _
'                           " And NP09 = '" & strDate2 & "' " & _
'                           " WHERE NP22=" & .TextMatrix(i, 7)
                  'Modify by Morgan 2006/1/24 加NP01
                  'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(DBDATE(Me.Text14(2).Text)):約定期限
                  strSql = "UPDATE NEXTPROGRESS SET NP08='" & strDate1 & "' " & _
                           ",NP09 = '" & strDate2 & "' " & _
                           ",NP23=" & CNULL(DBDATE(Me.Text14(2).Text)) & _
                           " WHERE NP22=" & .TextMatrix(i, 8) & " and np01='" & .TextMatrix(i, 9) & "'"
                  cnnConnection.Execute strSql, intI
               End If
            Next i
         End With
         'Modified by Morgan 2024/10/21 若延期的相關收文號為AB類表示已收文，都要再更新進度檔期限(Ex:FCP-070987已收文但NP也有補文件期限)--敏莉
         'If BlnCheck = False Then
         If BlnCheck = False Or (Text7.Text = 延期受理 And cp(43) < "C") Then
         'end 2024/10/21
            'Modify By Cheng 2003/01/17
            '修改本所期限, 法定期限
'            '本所期限
'            strDate1 = IIf(Len(Trim(Me.Text14(0).Text)) > 0, _
'                        IIf(Len(Trim(Me.Text14(0).Text) <> 8), Val(Replace(Me.Text14(0).Text, "/", "")) + 19110000, Replace(Me.Text14(0).Text, "/", "")), _
'                        "")
'            '法定期限
'            strDate2 = IIf(Len(Trim(Me.Text14(1).Text)) > 0, _
'                        IIf(Len(Trim(Me.Text14(1).Text) <> 8), Val(Replace(Me.Text14(1).Text, "/", "")) + 19110000, Replace(Me.Text14(1).Text, "/", "")), _
'                        "")
            '本所期限
            strDate1 = DBDATE(Me.Text14(0).Text)
            '法定期限
            strDate2 = DBDATE(Me.Text14(1).Text)
            'Modified by Morgan 2024/10/21 +cp27 is null
            strSql = "UPDATE CaseProgress SET CP06 = " & strDate1 & _
                     " ,CP07 = " & strDate2 & _
                     " WHERE CP09 = '" & "" & cp(43) & "' and cp27 is null"
            cnnConnection.Execute strSql, intI
         End If
      End If
      
      'Added by Morgan 2024/11/15
      If m_UpdCP09 <> "" Then
         strDate1 = DBDATE(Me.Text14(0).Text) '本所期限
         strDate2 = DBDATE(Me.Text14(1).Text) '法定期限
         strSql = "UPDATE CaseProgress SET CP06 = " & strDate1 & ",CP07 = " & strDate2 & " WHERE CP09 = '" & m_UpdCP09 & "' and cp27 is null"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql, intI
      End If
      'end 2024/11/15
   End If
   
'Add by Morgan 2005/1/31 通知審查中(1905)&延長審查時間(1501)的催審下一程序期限改為來函收文日加3個月。--靜芳,秀玲
   If Me.Text7.Text = 通知審查中 Then
      strExc(1) = CompDate(1, 3, TransDate(Label3(6), 2))
      'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
      strSql = "Update NextProgress Set NP08=" & PUB_GetWorkDay1(strExc(1), True) & ",NP09=" & strExc(1) & " WHERE NP01=(SELECT CP43 FROM CASEPROGRESS WHERE CP09='" & Label3(5) & "') AND NP07='" & 催審 & "' AND NP06 IS NULL"
      cnnConnection.Execute strSql
   ElseIf Me.Text7.Text = 延長審查時間 Then
      strExc(1) = CompDate(1, 3, TransDate(Label3(6), 2))
      'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
      strSql = "Update NextProgress Set NP08=" & PUB_GetWorkDay1(strExc(1), True) & ",NP09=" & strExc(1) & " WHERE NP01='" & Label3(5) & "' AND NP07='" & 催審 & "' AND NP06 IS NULL"
      cnnConnection.Execute strSql
      
   'Add by Morgan 2010/8/10
   '審查意見通知函1202+18個月更新催審期限
   'Modified by Morgan 2012/12/27 +最後通知1227
   'Modified by Lydia 2020/03/06 +被舉發理由1802
   'Modified by Lydia 2021/05/18 +通知補正1201
   'Modified by Lydia 2021/08/25 國外部凡是C類工程師的來函(排除核准1001、核發1008，另外非核准和一般來函性質1204,1217,1913,1603,1604)，有設核駁及審查意見通知函備註皆要帶備註到接洽單
   'ElseIf (Text7 = "1202" Or Text7 = "1227" Or Text7 = "1802" Or Text7 = "1201") Then
   'Modified by Lydia 2021/08/31 +判斷來函承辦人為工程師 And PUB_GetST03(Text16) = "F21" ; ex.發生了FCP065502通知即將公開1207直接收文告代
   'Modified by Lydia 2021/09/03 因為只需在C類接洽單列印,這裡反而不用改
   'ElseIf InStr("1001,1008,1204,1217,1913,1603,1604", Text7) = 0 And PUB_GetST03(Text16) = "F21" Then
   '   If (Text7 = "1202" Or Text7 = "1227" Or Text7 = "1802" Or Text7 = "1201") Then  'Added by Lydia 2021/08/25
   'Modify By Sindy 2025/6/19 +1221通知申復
   ElseIf (Text7 = "1202" Or Text7 = "1227" Or Text7 = "1802" Or Text7 = "1201" Or Text7 = "1221") Then
         strExc(1) = CompDate(1, 18, TransDate(Label3(6), 2))
         'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
         strSql = "Update NextProgress Set NP08=" & PUB_GetWorkDay1(strExc(1), True) & ",NP09=" & strExc(1) & " WHERE NP02='" & cp(1) & "' and NP03='" & cp(2) & "'" & _
            " and NP04='" & cp(3) & "' and NP05='" & cp(4) & "' AND NP07='" & 催審 & "' AND NP06 IS NULL" & _
            " and exists(select * from caseprogress where cp09=np01 and instr('101,102,103,104,105,301,302,303,304,305,306,307,107',cp10)>0)"
         cnnConnection.Execute strSql, intI
   '   End If 'Added by Lydia 2021/08/25 'Remove by Lydia 2021/09/03
      
'Modified by Morgan 2013/9/18 改呼叫共用函數
'      '2012/9/18 ADD BY SONIA 審查意見通知1202且為尼康客戶或Y51508案件要內部收文901,承辦期限為系統日起14天,不必抓工作天,不請款--陳毓芳
'      'Modified by Morgan 2012/10/19 +Y52003--陳毓芳
'      If Left(pa(26), 6) = "X56040" Or Left(pa(26), 6) = "X48340" Or Left(pa(26), 6) = "X45149" Or Left(pa(26), 6) = "X60049" Or Left(pa(75), 6) = "Y51508" Or Left(pa(75), 6) = "Y52003" Then
'         strExc(1) = AutoNo("B", 6)
'         strExc(5) = Val(CompDate(2, 14, strSrvDate(1)))
'         strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp43,cp48)" & _
'            " values('" & Ncp(1) & "','" & Ncp(2) & "','" & Ncp(3) & "','" & Ncp(4) & "'," & DBDATE(Ncp(5)) & ",'" & strExc(1) & "','901','90','" & Ncp(12) & "','" & Ncp(13) & "','" & Ncp(14) & "'" & _
'            ",'" & Ncp(9) & "'," & CNULL(strExc(5), True) & ")"
'         cnnConnection.Execute strSql, intI
'      End If
'      '2012/9/18 END
'
'      '2012/11/9 ADD BY SONIA 審查意見通知1202且為Y20065案件要內部收文901,承辦期限為系統日起10天(代理人要求是主管機關發文日15天),不必抓工作天,不請款--陳毓芳
'      '2012/11/15 MODIFY BY SONIA 加入 Y27766
'      If Left(pa(75), 6) = "Y20065" Or Left(pa(75), 6) = "Y27766" Then
'         strExc(1) = AutoNo("B", 6)
'         strExc(5) = Val(CompDate(2, 10, strSrvDate(1)))
'         strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp43,cp48)" & _
'            " values('" & Ncp(1) & "','" & Ncp(2) & "','" & Ncp(3) & "','" & Ncp(4) & "'," & DBDATE(Ncp(5)) & ",'" & strExc(1) & "','901','90','" & Ncp(12) & "','" & Ncp(13) & "','" & Ncp(14) & "'" & _
'            ",'" & Ncp(9) & "'," & CNULL(strExc(5), True) & ")"
'         cnnConnection.Execute strSql, intI
'      End If
'      '2012/11/9 END
'
'      '2012/10/19 ADD BY SONIA Y53309審查意見通知1202或核駁要內部收文901,承辦期限為系統日起7天(日曆天)--吳若芬
'      '2013/1/24 MODIFY BY SONIA 加 Y51542
'      'Modified by Morgan 2013/3/5 取消 Y51542 --吳彩菱
'      'Modified by Morgan 2013/8/28 ,加 Y34210 + X51446 --邱子瑜,Y51542 --吳彩菱
'      'Modified by Morgan 2013/8/30 ,+ Y47453 & X55778 --羅惠蓮
'      'Modified by Morgan 2013/9/6 + Y20065 --邱子瑜
'      If Left(pa(75), 6) = "Y53309" Or Left(pa(75), 6) = "Y51542" Or Left(pa(75), 6) = "Y20065" Or _
'         (Left(pa(75), 6) = "Y34210" And Left(pa(26), 6) = "X51446") Or _
'         (Left(pa(75), 6) = "Y47453" And Left(pa(26), 6) = "X55778") Then
'
'         strExc(1) = AutoNo("B", 6)
'         strExc(2) = "901"
'
'         'Added by Morgan 2013/8/28
'         'Y51542 改收其他翻譯 --吳彩菱
'         If Left(pa(75), 6) = "Y51542" Then
'            strExc(2) = "927"
'         End If
'         'Y34210 + X51446 14天 --邱子瑜
'         If Left(pa(75), 6) = "Y34210" And Left(pa(26), 6) = "X51446" Then
'            strExc(5) = Val(CompDate(2, 14, strSrvDate(1)))
'         'Added by Morgan 2013/9/6
'         'Y20065 15天 --邱子瑜
'         ElseIf Left(pa(75), 6) = "Y20065" Then
'               strExc(5) = Val(CompDate(2, 15, strSrvDate(1)))
'         Else
'         'end 2013/8/28
'            strExc(5) = Val(CompDate(2, 7, strSrvDate(1)))
'         End If 'Added by Morgan 2013/8/28
'Add by Lydia 2014/12/3 核駁及審查意見通知函備註
     ' If PUB_ChkAutoRec(pa(1), pa(75), pa(26), , strexc(2), strexc(5), , , pa(27), pa(28), pa(29), pa(30)) = True Then
       Dim sMemo As String
       If pa(1) = "FCP" Then  'add by sonia 2024/11/21 FG案不用
         'Remove by Lydia 2021/11/05
         'strExc(2) = "": strExc(5) = ""
         'strExc(7) = "": strExc(3) = "": strExc(4) = "": strExc(10) = ""
         'If Not IsNull(pa(27)) Then strExc(7) = ChangeCustomerL(pa(27))
         'If Not IsNull(pa(28)) Then strExc(3) = ChangeCustomerL(pa(28))
         'If Not IsNull(pa(29)) Then strExc(4) = ChangeCustomerL(pa(29))
         'If Not IsNull(pa(30)) Then strExc(10) = ChangeCustomerL(pa(30))
         'end 2021/11/05
         'Modified by Morgan 2020/11/12 +Ncp(133)
         'Modified by Lydia 2021/11/05 分別傳回B類收文(承辦期限、所限)和C類來函(承辦期限和指定送件日期)
         'sMemo = PUB_GetIncomMemoNew(pa(1) & pa(2) & pa(3) & pa(4), pa(1), strExc(2), ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)), , strExc(5), Ncp(133) _
                     , strExc(7), strExc(3), strExc(4), strExc(10))
         'strExc(7) = "": strExc(3) = "": strExc(4) = "": strExc(10) = ""
         Dim stBCP10 As String, stBCP48   As String, stBCP06 As String, stCCP48 As String, stCCP142 As String
         'sMemo = PUB_GetIncomMemoNew(pa(1) & pa(2) & pa(3) & pa(4), pa(1), stBCP10, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30)), _
                        "", Ncp(133), stBCP48, stBCP06, stCCP48, stCCP142)
         sMemo = PUB_GetIncomMemoNew(pa(1) & pa(2) & pa(3) & pa(4), pa(1), ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30)), _
                        "", Ncp(133), Text7.Text, stCCP48, stCCP142, stBCP10, stBCP48, stBCP06)
                        
         'Added by Lydia 2021/11/05 更新C類來函的承辦期限和指定送件日期，一併更新指定送件日期之前CP164=2
         If stCCP48 <> "" Then
             'Modified by Lydia 2021/11/16 加註cp64
             strSql = "Update CaseProgress set cp48=" & stCCP48 & ", cp141='3', cp142=" & stCCP142 & ", cp164='2' " & _
                         ", cp64='客戶指定" & ChangeWStringToTDateString(stCCP142) & "之前送件;'||cp64 where cp09='" & Ncp(9) & "' "
             cnnConnection.Execute strSql, intI
         End If
         'end 2021/11/05
         
         'Added by Lydia 2025/02/05 輸入中間程序來函時自動產生行事曆
         If PUB_AddSCforIncomMemo(pa(1), pa(2), pa(3), pa(4), Ncp(9), Text7, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30))) = False Then
             GoTo CheckingErr
         End If
         'end 2025/02/05
       End If   'add by sonia 2024/11/21 FG案不用
        
      'Modified by Lydia 2021/11/05 PUB_GetIncomMemoNew已有另外抓B類收文設定
      'If Len(sMemo) > 0 Then
      '   If strExc(2) = "" Then strExc(2) = "901"
      If stBCP10 <> "" Then
'end 'Add by Lydia 2014/12/3
         strExc(1) = AutoNo("B", 6)
'end 2013/9/18
         
         'Add By Sindy 2021/6/17 非智慧局期限，要掛本所期限
         'Remove by Lydia 2021/11/05 改從PUB_GetIncomMemoNew取得
         'Call GetPrjState6HM(pa(1), strExc(2), "cpm34", strExc(0))
         'strExc(6) = "" '本所期限
         ''110/7/21 淑華改, 本所期限=承辦期限; 因為這是算客戶指定的期限(為應該出給客戶的期限)
'         'If Val(strExc(5)) > 0 And strExc(0) = "N" And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
'         '   strExc(6) = PUB_GetFCPOurDeadline(DBDATE(strExc(5)), , , , "N")
         'If Val(strExc(5)) > 0 And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         '   'strExc(6) = PUB_GetFCPOurDeadline(DBDATE(strExc(5)), , , , "N")
         '   strExc(6) = DBDATE(strExc(5))
         'End If
         ''2021/6/17 END
         'end 2021/11/05
         
         'Modified by Lydia 2024/05/28 改成模組
         ''Modified by Morgan 2019/8/8 +CP20(抓設定)
         ''Modified by Lydia 2022/01/05 改抓變數 strExc(2)=> stBCP10
         ''Added by Lydia 2022/05/03 FCP-062174審定前不收費控制: (補上)判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
         'If pa(16) = "" And InStr("FCP062174000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
         '    strCP20 = "N": strCP16 = ""
         ''FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
         'ElseIf pa(16) <> "1" And InStr("FCP067004000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
         If PUB_GetCP20forSpec(pa(1), pa(2), pa(3), pa(4), pa(16)) = "N" Then
         'end 2024/05/28
             strCP20 = "N": strCP16 = ""
         Else
              strCP20 = PUB_GetCP20(pa(1), stBCP10, strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
         End If 'added by Lydia 2022/05/03
         'Modify By Sindy 2021/6/17 + ,cp06
         'Modified by Lydia 2021/11/05 改變數strExc(2)=>stBCP10, strExc(5)=> stBCP48, strExc(6)=> stBCP06
         'Modified by Lydia 2022/01/05 +CP16
         strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp43,cp48,cp06,cp16)" & _
            " values('" & Ncp(1) & "','" & Ncp(2) & "','" & Ncp(3) & "','" & Ncp(4) & "'," & DBDATE(Ncp(5)) & ",'" & strExc(1) & "','" & stBCP10 & "','90','" & Ncp(12) & "','" & Ncp(13) & "','" & Ncp(14) & "'" & _
            ",'" & strCP20 & "','" & Ncp(9) & "'," & CNULL(stBCP48, True) & "," & CNULL(stBCP06, True) & "," & CNULL(strCP16, True) & ")"
         cnnConnection.Execute strSql, intI
      End If
      '2012/10/19 END
   
   End If
'2005/1/31 end

   'Added by Morgan 2012/7/24
   '輸入通知依職權修正時再由電腦系統自動產生修正之內部收文(下一程序〝修正〞上Y)並自動掛上10個工作天之承辦期限--靜芳101/6/25請作單
   If Text7 = "1225" Then
      strSql = "update nextprogress set np06='Y' where np01='" & Ncp(9) & "' and np07='204' and np06 is null"
      cnnConnection.Execute strSql, intI
      
      strExc(1) = AutoNo("B", 6)
      'Modified by Lydia 2024/05/28 改成模組
      ''Added by Lydia 2022/05/03 FCP-062174審定前不收費控制: (補上)判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
      'If pa(16) = "" And InStr("FCP062174000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      '      strExc(6) = "N": strExc(2) = "null": strExc(3) = "null": strExc(4) = "null"
      ''FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
      'ElseIf pa(16) <> "1" And InStr("FCP067004000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      If PUB_GetCP20forSpec(pa(1), pa(2), pa(3), pa(4), pa(16)) = "N" Then
      'end 2024/05/28
            strExc(6) = "N": strExc(2) = "null": strExc(3) = "null": strExc(4) = "null"
      Else
         strExc(6) = PUB_GetCP20(pa(1), "204")
      'end 2022/05/03
         strExc(2) = GetFCPFee("FCP", "204")
         strExc(3) = GetPatentOfficialFee("FCP", "204", Ncp(7), pa(8), pa(9), pa(16))
         strExc(4) = (Val(strExc(2)) - Val(strExc(3))) / 1000
      End If 'added by Lydia 2022/05/03
      strExc(5) = Pub_GetHandleDay("FCP", "000", "204")
      'Modified by Lydia 2022/05/05 +cp20=strExc(6)
      strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14,cp16,cp17,cp18,cp20 ,cp43,cp48)" & _
         " values('" & Ncp(1) & "','" & Ncp(2) & "','" & Ncp(3) & "','" & Ncp(4) & "'," & DBDATE(Ncp(5)) & "," & CNULL(DBDATE(Ncp(6)), True) & _
         "," & CNULL(DBDATE(Ncp(7)), True) & ",'" & Ncp(8) & "','" & strExc(1) & "','204','90','" & Ncp(12) & "','" & Ncp(13) & "','" & Ncp(14) & "'" & _
         "," & strExc(2) & "," & strExc(3) & "," & strExc(4) & ", '" & strExc(6) & "' ,'" & Ncp(9) & "'," & CNULL(strExc(5), True) & ")"
      cnnConnection.Execute strSql, intI
   End If
   'end 2012/7/24
   
   'Addded by Lydia 2022/02/18  輸入1401智慧局通知面詢時，系統自動收文408面詢至進度檔(帶入期限-承辦、所限、法限)，預設為要請款
   If Text7 = "1401" Then
      strSql = "update nextprogress set np06='Y' where np01='" & Ncp(9) & "' and np07='408' and np06 is null"
      cnnConnection.Execute strSql, intI
      
      strExc(1) = AutoNo("B", 6)
      'Modified by Lydia 2024/05/28 改成模組
      ''Added by Lydia 2022/05/03 FCP-062174審定前不收費控制: (補上)判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
      'If pa(16) = "" And InStr("FCP062174000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      '      strExc(6) = "N": strExc(2) = "null": strExc(3) = "null": strExc(4) = "null"
      ''FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
      'ElseIf pa(16) <> "1" And InStr("FCP067004000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      If PUB_GetCP20forSpec(pa(1), pa(2), pa(3), pa(4), pa(16)) = "N" Then
      'end 2024/05/28
         strExc(6) = "N": strExc(2) = "null": strExc(3) = "null": strExc(4) = "null"
      Else
         strExc(6) = PUB_GetCP20(pa(1), "408")
      'end 2022/05/03
         strExc(2) = GetFCPFee("FCP", "408")
         strExc(3) = GetPatentOfficialFee("FCP", "408", Ncp(7), pa(8), pa(9), pa(16))
         strExc(4) = (Val(strExc(2)) - Val(strExc(3))) / 1000
         strExc(5) = Pub_GetHandleDay("FCP", "000", "408")
      End If 'Added by Lydia 2022/05/03
      'Modified by Lydia 2022/03/02 所限=法限, Ncp(5)=>Ncp(6)
      'Modified by Lydia 2022/03/22 debug : FCP-65683的收文日和所限放錯位置
      'strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14,cp16,cp17,cp18,cp43,cp48)" & _
         " values('" & Ncp(1) & "','" & Ncp(2) & "','" & Ncp(3) & "','" & Ncp(4) & "'," & DBDATE(Ncp(6)) & "," & CNULL(DBDATE(Ncp(6)), True) & _
         "," & CNULL(DBDATE(Ncp(7)), True) & ",'" & Ncp(8) & "','" & strExc(1) & "','408','90','" & Ncp(12) & "','" & Ncp(13) & "','" & Ncp(14) & "'" & _
         "," & strExc(2) & "," & strExc(3) & "," & strExc(4) & ",'" & Ncp(9) & "'," & CNULL(strExc(5), True) & ")"
      'Modified by Lydia 2022/05/05 +cp20=strExc(6)
      'Modified by Lydia 2022/06/02 增加備註,cp64
      'Modified by Lydia 2025/11/12 debug DBDATE(Ncp(7))=>DBDATE(Ncp(6))
      strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14,cp16,cp17,cp18,cp20,cp43,cp48,cp64)" & _
         " values('" & Ncp(1) & "','" & Ncp(2) & "','" & Ncp(3) & "','" & Ncp(4) & "'," & DBDATE(Ncp(5)) & "," & CNULL(DBDATE(Ncp(6)), True) & _
         "," & CNULL(DBDATE(Ncp(7)), True) & ",'" & Ncp(8) & "','" & strExc(1) & "','408','90','" & Ncp(12) & "','" & Ncp(13) & "','" & Ncp(14) & "'" & _
         "," & strExc(2) & "," & strExc(3) & "," & strExc(4) & ", '" & strExc(6) & "','" & Ncp(9) & "'," & CNULL(strExc(5), True) & ",'需繳納規費及準備PA給工程師')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2022/02/18
   
   'Added by Lydia 2022/05/10 通知補充聽證資料之期限 ：產生內部收文”202補文件”，法限:掛程序輸入的期限，本所:法限-2工作天，承辦期限: 本所-2工作天，承辦人掛工程師，進度檔備註:補呈聽證資料
   If m_1812CP07 <> "" Then
      'Modified by Lydia 2024/05/28 改成模組
      ''FCP-062174審定前不收費控制: (補上)判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
      'If pa(16) = "" And InStr("FCP062174000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      '      strExc(6) = "N": strExc(2) = "null": strExc(3) = "null": strExc(4) = "null"
      ''FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
      'ElseIf pa(16) <> "1" And InStr("FCP067004000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      If PUB_GetCP20forSpec(pa(1), pa(2), pa(3), pa(4), pa(16)) = "N" Then
      'end 2024/05/28
         strExc(6) = "N": strExc(2) = "null": strExc(3) = "null": strExc(4) = "null"
      Else
         strExc(6) = PUB_GetCP20(pa(1), "202")
         strExc(2) = GetFCPFee("FCP", "202")
         strExc(3) = GetPatentOfficialFee("FCP", "202", Ncp(7), pa(8), pa(9), pa(16))
         strExc(4) = (Val(strExc(2)) - Val(strExc(3))) / 1000
         strExc(5) = Pub_GetHandleDay("FCP", "000", "202")
      End If
      strExc(1) = AutoNo("B", 6)
      If m_1812CP07 <= strSrvDate(1) Then
        m_1812CP07 = strSrvDate(1): strExc(7) = strSrvDate(1): strExc(8) = strSrvDate(1)
      Else
         strExc(7) = CompWorkDay(3, m_1812CP07, 1)
         strExc(8) = CompWorkDay(5, m_1812CP07, 1)
         If strExc(7) <= strSrvDate(1) Then strExc(7) = strSrvDate(1)
         If strExc(8) <= strSrvDate(1) Then strExc(8) = strSrvDate(1)
      End If
      strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14,cp16,cp17,cp18,cp20,cp43,cp48,cp64)" & _
         " values('" & Ncp(1) & "','" & Ncp(2) & "','" & Ncp(3) & "','" & Ncp(4) & "'," & DBDATE(Ncp(5)) & "," & CNULL(strExc(7), True) & _
         "," & CNULL(DBDATE(m_1812CP07), True) & ",'" & Ncp(8) & "','" & strExc(1) & "','202','90','" & Ncp(12) & "','" & Ncp(13) & "','" & Ncp(14) & "'" & _
         "," & CNULL(strExc(2), True) & "," & CNULL(strExc(3), True) & "," & CNULL(strExc(4), True) & ", '" & strExc(6) & "','" & Ncp(9) & "'," & CNULL(strExc(8), True) & ",'補呈聽證資料')"
      cnnConnection.Execute strSql, intI
   End If
   
   'Added by Morgan 2017/5/10 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, Ncp(9), pa(1), pa(2), pa(3), pa(4), Ncp(10), m_strExtNote
   'Added by Morgan 2021/6/11 紙本公文--何淑華
   Else
      PUB_FCPOAInform Ncp(9), pa(1), pa(2), pa(3), pa(4), Ncp(10), m_strExtNote
   End If
   'end 2017/5/10
   
   'Added by Lydia 2023/09/25
   If m_strIR01 <> "" Then
      'F2外專不做2次確認
      If Text14(1) <> "" And Left(Pub_StrUserSt03, 2) <> "F2" Then
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm06010604_1", Ncp(9), m_bolReKeyInOK
         bolReKeyInCase = True
      Else
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm06010604_1", IIf(Pub_StrUserSt03 = "F22", Ncp(9), ""), , , False
      End If
   End If
   'end 2023/09/25
   
   cnnConnection.CommitTrans
   '911106 nick 從上面移下來
   '8
   If Text7 = 被異議理由 Then
      strExc(0) = "SELECT MAX(CP05),CP09 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND " & _
         "CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='" & 領證及繳年費 & "' AND CP27 IS NULL AND CP57 IS NULL GROUP BY CP09"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      ' 90.06.26 modify by louis 無資料時不做
      If RsTemp.RecordCount > 0 Then
         RsTemp.MoveFirst
         If RsTemp.Fields(0) > 0 Then
            frm06010604_31.Tag = "6" & RsTemp.Fields(1)
            frm06010604_31.Show vbModal
         End If
      End If
   End If
   
'911106 nick transation
   Exit Function
CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False

End Function

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()
Dim ret As Long
Dim strTmp As String

'   prevWndProc = GetWindowLong(Text8.hwnd, GWL_WNDPROC)
'   ret = SetWindowLong(Text8.hwnd, GWL_WNDPROC, AddressOf WndProc)
   MoveFormToCenter Me
   
   'EnableTextBox Text13, False  '2010/11/17 cancel by sonia
   
   intWhere = 國外_FC
   With frm06010604_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      ReadPatent
   End With
   Combo1.ListIndex = 0
   
   'Added by Lydia 2023/09/25
   m_strIR01 = frm06010604_2.m_strIR01
   m_strIR02 = frm06010604_2.m_strIR02
   m_strIR03 = frm06010604_2.m_strIR03
   m_strIR04 = frm06010604_2.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      lblDelivery.Visible = True: txtDelivery.Visible = True
   Else
      lblDelivery.Visible = False: txtDelivery.Visible = False
   End If
   'end 2023/09/25
   
   '2013/8/19 add by sonia 內專需求
   Label8.Visible = False
   Text6.Visible = False
   Text6.Locked = True
   '2013/8/19 end
   
   'Modify By Cheng 2002/05/31
'   Text6 = strSrvDate(2)
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   Text9.Text = "（" & strTmp & "）智專一（二）字第號"
   
   'Added by Morgan 2017/5/10 電子公文
   If m_DocNo <> "" Then
      If m_DocWord <> "" Then
         Text9 = m_DocWord & "字第" & m_DocNo & "號"
      ElseIf m_DocNo <> "" Then
         Text9 = Replace(Text9, "第號", "第" & m_DocNo & "號")
      End If
      '期限
      If m_DeadLine <> "" Then
         Option1(1).Value = True
         If Len(m_DeadLine) >= 7 Then
            Option4(2).Value = True
            Text12 = m_DeadLine
            Text12_Validate False
         ElseIf Right(m_DeadLine, 1) = "日" Then
            Option4(0).Value = True
            Text10 = Val(m_DeadLine)
            Text10_Validate False
         ElseIf Right(m_DeadLine, 1) = "月" Then
            Option4(1).Value = True
            Text11 = Val(m_DeadLine)
            Text11_Validate False
         End If
      End If
   End If
   'end 2017/5/10
   
   Check908 pa 'Add by Morgan 2009/10/1
   
   SSTab1.Tab = 0
   
   'Add By Sindy 2021/5/7
   If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
      Label9.Visible = True
      Text14(2).Visible = True
   Else
      Label9.Visible = False
      Text14(2).Visible = False
   End If
   '2021/5/7 END
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, i As Integer
Dim rsTemp1 As New ADODB.Recordset, bolTmp As Boolean
'Dim Cancel As Boolean 'Add By Sindy 2016/8/17 'TEST
   
   'Modify by Morgan 2006/1/3 改寫法
   Text8.Clear
   Text8.AddItem "分割", 0: Text8.ItemData(0) = 分割
   Text8.AddItem "改請獨立", 0: Text8.ItemData(0) = 改請獨立
   Text8.AddItem "改請聯合", 0: Text8.ItemData(0) = 改請聯合
   Text8.AddItem "改請追加", 0: Text8.ItemData(0) = 改請追加
   Text8.AddItem "改請設計", 0: Text8.ItemData(0) = 改請設計
   Text8.AddItem "改請新型", 0: Text8.ItemData(0) = 改請新型
   Text8.AddItem "改請發明", 0: Text8.ItemData(0) = 改請發明
   Text8.AddItem "退費", 0: Text8.ItemData(0) = 退費
   Text8.AddItem "變更", 0: Text8.ItemData(0) = 變更
   Text8.AddItem "舉發答辯", 0: Text8.ItemData(0) = 舉發答辯
   Text8.AddItem "異議答辯", 0: Text8.ItemData(0) = 異議答辯
   Text8.AddItem "領證及繳年費", 0: Text8.ItemData(0) = 領證及繳年費
   Text8.AddItem "申復", 0: Text8.ItemData(0) = 申復
   Text8.AddItem "補充說明", 0: Text8.ItemData(0) = 補充說明
   Text8.AddItem "修正", 0: Text8.ItemData(0) = 修正
   Text8.AddItem "補文件", 0: Text8.ItemData(0) = 補文件
   Text8.AddItem "訴願", 0: Text8.ItemData(0) = 訴願 '--靜芳
   '2006/1/3 end
   
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   Label3(6) = frm06010604_1.Text5
   Label3(5) = strReceiveNo
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   ' 90.06.26 modify by louis 是否閉卷
   lblPA57 = Empty
   m_PA177 = "" 'Added by Lydia 2023/07/28
   If pa(1) = "FCP" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         MPa9 = pa(9)
         Label3(2) = pa(10)
         Text1 = pa(11)
         Text31 = pa(91)
         Text27(0) = pa(57)
         'Add By Sindy 2012/3/5
         If pa(57) = "Y" Then
            m_blnClosed = True
         Else
            m_blnClosed = False
         End If
         '2012/3/5 End
         If pa(71) = "" Then
            If pa(75) = "" Then
               strExc(0) = "SELECT CU75 FROM CUSTOMER WHERE " & ChgCustomer(pa(26))
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 And Not IsNull(RsTemp.Fields(0)) Then Text27(2) = RsTemp.Fields(0)
            Else
               strExc(0) = "SELECT FA42 FROM FAGENT WHERE " & ChgFagent(pa(75))
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 And Not IsNull(RsTemp.Fields(0)) Then Text27(2) = RsTemp.Fields(0)
            End If
         Else
            Text27(2) = pa(71)
         End If
         ' 90.06.26 modify by louis 是否閉卷
         lblPA57 = pa(57)
         m_PA177 = pa(177) 'Added by Lydia 2023/07/28 FCP專利連結通知
      End If
   Else
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadServicePracticeDatabase(pA(), intWhere) Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         
         Label3(2) = pa(10)
         Text1 = pa(11)
         Text31 = pa(18)
         Text27(0) = pa(15)
         'Add By Sindy 2012/3/5
         If pa(15) = "Y" Then
            m_blnClosed = True
         Else
            m_blnClosed = False
         End If
         '2012/3/5 End
      End If
   End If
   ' 90.06.26 modify by louis, 下一程序名稱帶出來
   'strExc(0) = "SELECT '',CPM03," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP13," & _
   '   "NP14," & SQLDate("NP11") & ",NP22 FROM NEXTPROGRESS,CASEPROPERTYMAP " & _
   '   "WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
   '   " AND (NP06<>'Y' OR NP06 IS NULL) AND NP01=CPM01(+) AND NP07=CPM02(+)"
    'Modify By Cheng 2003/01/17
    '在相關人後加備註欄
'   strExc(0) = "SELECT '',CPM03," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP13," & _
'      "NP14," & SQLDate("NP11") & ",NP22 FROM NEXTPROGRESS,CASEPROPERTYMAP " & _
'      "WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'      " AND (NP06<>'Y' OR NP06 IS NULL) AND NP02=CPM01(+) AND NP07=CPM02(+)"
   'Modify by Morgan 2006/1/24 加NP01
   'Mofieid by Morgan 2025/3/4 +NP07
   strExc(0) = "SELECT '',CPM03," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP13," & _
      "NP14,NP15," & SQLDate("NP11") & ",NP22,NP01,NP07 FROM NEXTPROGRESS,CASEPROPERTYMAP " & _
      "WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND (NP06<>'Y' OR NP06 IS NULL) AND NP02=CPM01(+) AND NP07=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   cp(9) = strReceiveNo
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.ReadCaseProgressDatabase(cp(), intWhere) Then
   If ClsPDReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(10) <> "" Then
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(cp(1), cp(10), strExc(0), BolTmp) Then Label3(1) = strExc(0)
         If ClsPDGetCaseProperty(cp(1), cp(10), strExc(0), bolTmp) Then Label3(1) = strExc(0)
      End If
      ' 90.06.26 modify by louis 讀檔時不須帶入承辦人, 改在輸入完來函性質後才去帶出承辦人
      'If Left(cp(10), 1) = "1" Then
      '   strExc(0) = "SELECT CP14 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10=" & 翻譯
      '   intI = 1
      '   Set rsTemp1 = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      '   If intI = 1 Then
      '      If Not IsNull(rsTemp1.Fields(0)) Then Text16 = rsTemp1.Fields(0): ChgType 16
      '   End If
      'Else
      '   If cp(14) <> "" Then Text16 = cp(14): ChgType 16
      'End If
      
      ' 90.10.09 modify by louis
      'If cp(10) = 撤銷原處分 Then
      '   Label17.Visible = True
      '   Text13.Visible = True
      '   Label16.Visible = True
      '   Label17.Visible = True
      'Else
      '   Label17.Visible = False
      '   Text13.Visible = False
      '   Label16.Visible = False
      '   Label17.Visible = False
      'End If
      
      ' 90.06.26 modify by louis 進度備註放錯位置且不帶出來
      'Text31 = cp(64)
   End If
   
   'TEST
'   'Add By Sindy 2016/8/17 若為新申請案之第一次通知補文件，則法定期限預設為指定日期且為案件之申請日+4個月
'   strExc(0) = "SELECT c2.cp09 FROM CASEPROGRESS c1,CASEPROGRESS c2" & _
'         " WHERE c1.cp01='" & pa(1) & "'" & _
'         " and c1.cp02='" & pa(2) & "'" & _
'         " and c1.cp03='" & pa(3) & "'" & _
'         " and c1.cp04='" & pa(4) & "'" & _
'         " and c1.cp10='1003' and c1.cp43=c2.cp09(+)" & _
'         " and instr('" & NewCasePtyList & "',c2.cp10)>0" & _
'         " order by SQLDatet2(c2.CP05) asc,c2.CP66 asc,c2.CP67 asc,c2.CP09 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      If cp(9) = RsTemp.Fields(0) Then
'         If Val(pa(10)) > 0 Then
'            Option4(2).Value = True
'            '指定日期且為案件之申請日+4個月
'            Text12 = TransDate(CompDate(1, 4, TransDate(pa(10), 2)), 1)
'            Cancel = False
'            Call Text12_Validate(Cancel)
'            If Cancel = True Then
'               Text12.SetFocus
'               Exit Sub
'            End If
'         End If
'      End If
'   End If
'   '2016/8/17 END

   If m_DocNo <> "" Then stCP133 = PUB_GetEDocDate(m_DocNo) 'Added by Morgan 2020/11/13 官方發文日
End Sub

Private Function ChgType(i As Integer, Optional SstrKind As Integer) As Boolean
 Dim strTempName As String, bolExcept As Boolean, stCPM01 As String
   ChgType = False
   Select Case i
      Case 7
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Text7.Text, strTempName, False) Then
         If ClsPDGetCaseProperty(pa(1), Text7.Text, strTempName, False) Then
            Label3(4) = strTempName
            
            'Added by Morgan 2017/6/8 下一程序都要清除
            Text8 = ""
            Label3(3) = ""
            'end 2017/6/8
            
            If Me.m_DeadLine = "" Then 'Added by Morgan 2017/5/10 電子公文
               
               Text14(0) = ""
               Text14(1) = ""
               'Text8 = "" 'Removed by Morgan 2017/6/8 移到上面
               'Label3(3) = "" 'Removed by Morgan 2017/6/8 移到上面
               stCPM01 = pa(1)
               
               'Added by Morgan 2013/10/9 台灣新型審查意見通知函來函期限預設文到次日起1個月--102/9/5 靜芳請作單
               'Modified by Morgan 2013/12/19 +1201,1202 --靜芳
               If pa(1) = "FCP" And (Text7.Text = "1221" Or Text7.Text = "1201" Or Text7.Text = "1202") And pa(8) = "2" Then
                  Option1(1).Value = True
                  Option4(1).Value = True
                  Text11 = 1
                  GetTime
               Else
               'end 2013/10/9
               
                  'Add by Morgan 2008/1/11
                  '若來函為1202(審查意見通知)且申請人均為本國人時法限=收文日+60天,所限=法限-4天
                  bolExcept = False
                  'Modified by Morgan 2012/12/27 +最後通知1227
                  If pa(1) = "FCP" And (Text7.Text = "1202" Or Text7.Text = "1227") Then
                     If PUB_ExistForeigner(pa(1) & pa(2) & pa(3) & pa(4)) = False Then
                        'Modify by Morgan 2008/9/3 改抓P案設定
                        'Option1(1).Value = True
                        'Text10 = 60
                        'Text14(1) = TransDate(CompDate(2, Text10, TransDate(Label3(6), 2)), 1)
                        'Text14(0) = TransDate(CompDate(2, -4, TransDate(Text14(1), 2)), 1)
                        'bolExcept = True
                        stCPM01 = "P"
                     End If
                  End If
                  
                  'Modify by Morgan 2008/9/3
                  'If bolExcept = False Then
                     strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & stCPM01 & "' AND CPM02='" & Text7.Text & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     With RsTemp
                        If intI = 1 Then
                           If Not IsNull(.Fields(1)) Then
                              Option4(0).Value = True
                              Text10 = .Fields(1)
                           ElseIf Not IsNull(.Fields(2)) Then
                              Option4(1).Value = True
                              Text11 = .Fields(2)
                           Else
                              Option4(0).Value = True
                              Text10 = ""
                              Text11 = ""
                           End If
                           If Not IsNull(.Fields(0)) Then
                              If .Fields(0) = "1" Then
                                 Option1(0).Value = True
                              Else
                                 Option1(1).Value = True
                              End If
                           End If
                           GetTime
                        End If
                     End With
                  'End If
               End If 'Added by Morgan 2013/10/9
            
            End If 'Added by Morgan 2017/5/10
            
            '承辦期限
            Text17 = ""
            'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
            'strExc(0) = "SELECT CF04 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & Text7.Text & "'"
            'intI = 1
            'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            'If intI = 1 Then
            '   With RsTemp
            '      If Not IsNull(.Fields(0)) Then
            '         Text17 = TransDate(CompWorkDay(Val(.Fields(0)), TransDate(Label3(6).Caption, 2)), 1)
            '      End If
            '   End With
            'End If
            Text17 = TransDate(Pub_GetHandleDay(pa(1), pa(9), Text7.Text, TransDate(Label3(6).Caption, 2)), 1)
            'End 2007/10/11
            'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：收到學名藥廠P4通知1922,來函有專利號:(1)程序輸入C類來函「P4通知1922」(承辦期限3個工作天, 法定期限45天)
            If pa(1) = "FCP" And Text7 = "1922" And m_PA177 = "Y" Then
               strExc(1) = CompDate(2, 45, DBDATE(Label3(6)))
               strExc(2) = CompWorkDay(4, strExc(1), 1)
               'Text14(1) = TransDate(strExc(1), 1) '到存檔直接寫入
               Text17 = TransDate(strExc(2), 1)
            End If
            'FG案:收到食藥署轉交第三人通知1923,承辦期限7個工作天
            If pa(1) = "FG" And Text7 = "1923" Then
               strExc(2) = CompWorkDay(8, DBDATE(Label3(6)))
               Text17 = TransDate(strExc(2), 1)
            End If
            'end 2023/07/28
            
            '下一程序 modify by sonia
            strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & Text7.Text & "'"
            ' 90.07.06 modify by louis (以系統類別+案件性質取案件國家收費表的下一救濟程序)
            'strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & cp(10) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With RsTemp
                  If Not IsNull(.Fields(0)) Then Text8 = .Fields(0): ChgType 8, Val(Text8)
               End With
            End If
            
            ChgType = True
         Else
            Label3(4) = ""
         End If
         Text8.Tag = Text8 'Added by Morgan 2012/11/5
         
      Case 8
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Format(SstrKind), strTempName, False) Then
         If ClsPDGetCaseProperty(pa(1), Format(SstrKind), strTempName, False) Then
            Label3(3) = strTempName
            ChgType = True
         Else
            Label3(3) = ""
         End If
      Case 16
        'Modify By Cheng 2003/04/08
        '若有輸入承辦人
        If Me.Text16.Text <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetStaff(Text16.Text, strTempName) Then
            If ClsPDGetStaff(Text16.Text, strTempName) Then
               Label3(7) = strTempName
               ChgType = True
            Else
               Label3(7) = ""
            End If
        '若未輸入承辦人
        Else
            Label3(7) = ""
            ChgType = True
        End If
   End Select
End Function

'Added by Morgan 2015/4/20
'例外承辦期限控制
Private Sub SetException()
   'Modified by Lydia 2016/08/15 改成共用模組
'   '先正達OA承辦期限設7個工作天,若下一程序為 804,501-509時設2個工作天(24Hr)
'   If InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left(pa(75) & "000", 8)) > 0 Then
'      If Text8 = "804" Or Text8 >= "501" And Text8 <= "509" Then
'         Text17 = TransDate(CompWorkDay(2, TransDate(Label3(6).Caption, 2), 0), 1)
'      ElseIf (Text7 = "1202" Or Text7 = "1227") Then
'         Text17 = TransDate(CompWorkDay(7, TransDate(Label3(6).Caption, 2), 0), 1)
'      End If
'   'Added by Morgan 2015/7/3 --吳彩菱
'   'Y51753+X45149010 承辦天數:14 起算日期:系統日
'   ElseIf Left(pa(75) & "000", 8) = "Y5175300" And Left(pa(26) & "000", 8) = "X4514901" Then
'      If (Text7 = "1202" Or Text7 = "1227") Then
'         Text17 = TransDate(CompDate(2, 14, strSrvDate(1)), 1)
'      End If
'   End If
   'Modify By Sindy 2017/5/9 + , TransDate(Text14(0), 2)
   'Call Pub_SetExceptCP48(pa(75), pa(26), Text7.Text, TransDate(Label3(6).Caption, 2), Text17, Text8.Text)
   'Modified by Morgan 2020/11/13 +stCP133
   Call Pub_SetExceptCP48(pa(75), pa(26), Text7.Text, TransDate(Label3(6).Caption, 2), Text17, Text8.Text, TransDate(Text14(0), 2), , stCP133)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim ret As Long
   'If prevWndProc <> 0 Then
   '   ret = SetWindowLong(Text8.hwnd, GWL_WNDPROC, prevWndProc)
   '   prevWndProc = 0
   'End If
   PUB_SendMailCache 'Added by Morgan 2021/6/11
   Set frm06010604_3 = Nothing
End Sub

Private Sub Text10_GotFocus()
   InverseTextBox Text10
End Sub

Private Sub Text11_GotFocus()
   InverseTextBox Text11
End Sub

Private Sub Text12_GotFocus()
   InverseTextBox Text12
End Sub

'Private Sub itmNext_Click(Index As Integer)
' Dim i As Integer
'   Select Case Index
'      Case 0
'         i = 補文件
'      Case 1
'         i = 修正
'      Case 2
'         i = 補充說明
'      Case 3
'         i = 申復
'      Case 4
'         i = 領證及繳年費
'      Case 5
'         i = 異議答辯
'      Case 6
'         i = 舉發答辯
'      Case 7
'         i = 變更
'      Case 8
'         i = 退費
'      Case 9
'         i = 改請發明
'      Case 10
'         i = 改請新型
'      Case 11
'         i = 改請設計
'      Case 12
'         i = 改請追加
'      Case 13
'         i = 改請聯合
'      Case 14
'         i = 改請獨立
'      Case 15
'         i = 分割
'   End Select
'   ChgType 8, i
'   Text8.Text = i
'End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   Dim iDays As Integer 'Added by Morgan 2019/7/11
   
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
      MsgBox "來函期限不可空白 !", vbCritical
      Cancel = True
   Else
      If ChkDate(Text12) Then
         If Val(Text12) < Val(strSrvDate(2)) Then
            MsgBox "來函期限不可小於系統日 !", vbCritical
            Cancel = True
         Else
            Text14(1) = Text12
            
            'Modified by Morgan 2014/11/20 外專改回舊規則
            ''Added by Morgan 2014/10/9
            'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
            '   Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
            'Else
            ''end 2014/10/9
            

            'Modified by Morgan 2019/7/11
            'Text14(0) = TransDate(CompDate(2, -2, TransDate(Text14(1), 2)), 1)
            iDays = 2
            'end 2019/7/11
            
               'Add By Cheng 2002/12/16
               '若前畫面點選的案件性質為101~107, 301~302, 來函性質為通知補文件(1003), 則本所期限 = 法定期限 - 4 天
               If ((Val(cp(10)) >= 101 And Val(cp(10)) <= 107) Or (Val(cp(10)) >= 301 And Val(cp(10)) <= 302)) And Me.Text7.Text = "1003" Then
                  'Modified by Morgan 2019/7/11
                  'Text14(0) = TransDate(CompDate(2, -4, TransDate(Text14(1), 2)), 1)
                  iDays = 4
                  'end 2019/7/11
               End If
               '92.3.11 add by sonia 法定期限>系統日+1月則本所期限=法定期限-4天
               If CompDate(1, 1, GetTodayDate) < TransDate(Text12, 2) Then
               'If DateDiff("m", GetTaiwanTodayDate, TransDate(Text12, 1)) > 2 Then
                   'Modified by Morgan 2019/7/11
                   'Text14(0) = TransDate(CompDate(2, -4, TransDate(Text14(1), 2)), 1)
                   iDays = 4
                  'end 2019/7/11
               End If
               '92.3.11 end
               
            'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
            If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
               'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
               'Modify By Sindy 2025/2/12 傳入本所案號,C
               Text14(0) = TransDate(PUB_GetFCPOurDeadline(Text14(1), iDays, , m_pAgreeOnDate, , , pa(1), pa(2), pa(3), pa(4), "C"), 1)
               Text14(2) = TransDate(m_pAgreeOnDate, 1) 'Add By Sindy 2021/5/7
            Else
               Text14(0) = TransDate(CompDate(2, -1 * iDays, TransDate(Text14(1), 2)), 1)
            End If
            'Added by Morgan 2019/7/11
               
            'End If 'Added by Morgan 2014/10/9
            'end 2014/11/20
            
            'Added by Morgan 2022/9/20 若所限或約定小於系統日時設定為系統日 Ex:FCP-063165 通知面詢法限為當天
            If Text14(0).Enabled = False Then
               If Text14(0) < strSrvDate(2) Then
                  Text14(0) = strSrvDate(2)
               End If
            End If
            If Text14(2).Enabled = False Then
               If Text14(2) < strSrvDate(2) Then
                  Text14(2) = strSrvDate(2)
               End If
            End If
            'end 2022/9/20
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub GetTime()
 Dim i As Integer
   If Option4(0).Value = True Then
      If Val(Text10) > 0 Then
         Text14(1) = TransDate(CompDate(2, Val(Text10), TransDate(Label3(6), 2)), 1)
         If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
         'Modify by Morgan 2008/9/3 改60天以上
         'If Text10 = "60" Or Text10 = "90" Then
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      End If
   ElseIf Option4(1).Value = True Then
      If Val(Text11) > 0 Then
         Text14(1) = TransDate(CompDate(1, Val(Text11), TransDate(Label3(6), 2)), 1)
         If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
         ''Modify by Morgan 2008/9/3 改2個月以上
         'If Text11 = "2" Then
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
   End If
   If Text14(1) <> "" Then
      'Modified by Morgan 2014/11/20 外專改回舊規則
      ''Added by Morgan 2014/10/9
      'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
      '   Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
      'Else
      ''end 2014/10/9
      
      'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
      If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
         'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
         'Modify By Sindy 2025/2/12 傳入本所案號,C
         Text14(0) = TransDate(PUB_GetFCPOurDeadline(Text14(1), i, , m_pAgreeOnDate, , , pa(1), pa(2), pa(3), pa(4), "C"), 1)
         Text14(2) = TransDate(m_pAgreeOnDate, 1) 'Add By Sindy 2021/5/7
      Else
      'end 2019/7/11
         
         Text14(0) = TransDate(CompDate(2, i, TransDate(Text14(1), 2)), 1)
         
      End If 'Added by Morgan 2019/7/11
      
      'End If 'Added by Morgan 2014/10/9
      'end 2014/11/20
   End If
   
End Sub

'2010/11/17 cancel by sonia
'Private Sub Text13_GotFocus()
'   InverseTextBox Text13
'End Sub
'
'Private Sub Text13_KeyPress(KeyAscii As Integer)
'   If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
'      KeyAscii = 0
'      Beep
'   End If
'End Sub
'2010/11/17 end

Private Sub Text14_GotFocus(Index As Integer)
   InverseTextBox Text14(Index)
End Sub

Private Sub Text14_Validate(Index As Integer, Cancel As Boolean)
   If Text14(Index) <> "" Then
      If Not ChkDate(Text14(Index)) Then
         Cancel = True
      Else
         'Add By Cheng 2002/10/15
         '若有輸入本所期限時, 不可小於系統日
         If Index = 0 Then
            If Len(Me.Text14(0).Text) = 8 Then
               If Val(Me.Text14(0).Text) < strSrvDate(1) Then
                  MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            ElseIf Len(Me.Text14(0).Text) = 7 Or Len(Me.Text14(0).Text) = 6 Then
               If Val(Me.Text14(0).Text) + 19110000 < strSrvDate(1) Then
                  MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                  Cancel = True
               End If
            End If
         End If
         
         If Index = 1 Then
            If Not ChkRange(Text14(0), Text14(1), "本所期限、法定期限") Then
               Cancel = True
            Else
               'edit by nickc 2007/02/05 不用 dll 了
               'If objLawDll.ChkMRec(TransDate(Label3(6).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
               If ClsLawChkMRec(TransDate(Label3(6).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
                  If Text14(0) <> TransDate(strExc(1), 1) Then
                     If MsgBox("與櫃台之來函收文記錄本所期限 ( " & TransDate(strExc(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                        frm06010604_1.Show
                        Unload frm06010604_2
                        Unload Me
                     Else
                        Text14(0) = ""
                        Text14(1) = ""
                     End If
                  ElseIf Text14(1) <> TransDate(strExc(2), 1) Then
                     If MsgBox("與櫃台之來函收文記錄法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                        frm06010604_1.Show
                        Unload frm06010604_2
                        Unload Me
                     Else
                        Text14(0) = ""
                        Text14(1) = ""
                     End If
                  End If
               'Added by Morgan 2017/5/10 電子公文
               ElseIf m_DocNo <> "" Then
                  If m_DeadLine <> "" Then
                     If Len(m_DeadLine) >= 7 Then
                        strExc(2) = m_DeadLine
                     ElseIf Right(m_DeadLine, 1) = "日" Then
                        strExc(2) = CompDate(2, Val(m_DeadLine), Label3(6))
                     ElseIf Right(m_DeadLine, 1) = "月" Then
                        strExc(2) = CompDate(1, Val(m_DeadLine), Label3(6))
                     End If
                     If Text14(1) <> TransDate(strExc(2), 1) Then
                        If MsgBox("與電子公文之法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                           Exit Sub
                        End If
                     End If
                  End If
               'end 2017/5/10
               Else
                  If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Cancel = True
               End If
            End If
         End If
      End If
   End If
   If Cancel = True Then TextInverse Text14(Index)
End Sub

Private Sub Text16_GotFocus()
   InverseTextBox Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTmp As String
   
   Cancel = False
   Label3(7) = Empty
   If IsEmptyText(Text16) = False Then
      strTemp = Empty
      strTemp = GetStaffName(Text16)
      Label3(7) = strTemp
      If IsEmptyText(strTemp) Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的承辦人"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         Text16_GotFocus
      End If
   End If
End Sub

Private Sub Text17_GotFocus()
   InverseTextBox Text17
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   If Text17 <> "" Then
      If ChkWorkDay(TransDate(Text17, 2)) Then
         'Modify by Morgan 2010/8/12 百年蟲
         'If Len(Me.Text14(0).Text) > 0 And Len(Me.Text17.Text) > 0 And Text17 > Text14(0) Then
         If Len(Me.Text14(0).Text) > 0 And Len(Me.Text17.Text) > 0 And Val(Text17) > Val(Text14(0)) Then
            MsgBox "承辦期限不可大於本所期限，請重新輸入 !", vbCritical
            Cancel = True
         End If
      Else
         MsgBox "承辦期限不正確，請重新輸入 !", vbCritical
         Cancel = True
      End If
      
   'Remove by Morgan 2011/3/21 找不到要檢查的需求先取消,否則無法作業
'   Else
'      If Text8 <> "" Then
'         MsgBox "有下一程序且有定義工作天數時不可空白 !", vbCritical
'         Cancel = True
'      End If
   End If

   If Cancel = True Then TextInverse Text17
End Sub

Private Sub Text18_GotFocus()
   InverseTextBox Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text19_GotFocus()
   InverseTextBox Text19
End Sub

Private Sub Text20_GotFocus(Index As Integer)
   InverseTextBox Text20(Index)
End Sub

Private Sub Text20_LostFocus(Index As Integer)
   If Text19 <> "" Then
      Select Case Index
         Case 2
            If Text20(0) = "" And Text20(1) = "" And Text20(2) = "" Then
               MsgBox "對造案件名稱不可同時空白 !", vbCritical
               Text20(0).SetFocus
            End If
         Case 5
            If Text20(3) = "" And Text20(4) = "" And Text20(5) = "" Then
               MsgBox "對造名稱不可同時空白 !", vbCritical
               Text20(3).SetFocus
            End If
      End Select
   End If
End Sub
'Add by Morgan 2008/2/25
Private Sub Text23_GotFocus()
   Dim ii As Integer
   Select Case Me.Text7.Text
   Case "1802"
      ii = InStr(Me.Text23.Text, "N")
      Me.Text23.SelStart = ii
      Me.Text23.SelLength = 0
   
   Case "1405", "1810"
      ii = InStr(Me.Text23.Text, "e")
      Me.Text23.SelStart = ii
      Me.Text23.SelLength = 0
   
   Case Else
      TextInverse Me.Text23
   End Select
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
   Select Case Me.Text7.Text
   '受理技術報告須輸 eXX
   Case "1405"
   Case Else
      KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text23_LostFocus()
   If Me.Text7.Text = "1405" Or Me.Text7.Text = "1810" Then
      strExc(0) = "SELECT CP36 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 in ('1405','1810') AND CP36='" & Text19.Text & Text23.Text & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "對造案件數代號不可重覆 !", vbCritical
         Text23_GotFocus
         Text23.SetFocus
      End If
   End If
End Sub
'end 2008/2/25

Private Sub Text26_GotFocus()
   InverseTextBox Text26
End Sub

Private Sub Text26_Validate(Cancel As Boolean)
   If Text26 = "" Then
      If Text7 = 專利權消滅 Then
         MsgBox "來函性質為專利權消滅時，不可空白 !", vbCritical
         Cancel = True
      End If
   Else
      If Not ChkDate(Text26) Then
         MsgBox "日期不正確，請重新輸入 !", vbCritical
         Cancel = True
      End If
   End If
End Sub

Private Sub Text27_GotFocus(Index As Integer)
   InverseTextBox Text27(Index)
End Sub

Private Sub Text27_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
      'Modify by Morgan 2007/7/24 是否向客戶請款改輸N
   If Index = 1 Then
      If KeyAscii <> Asc("N") And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   Else
   'End 2007/7/24
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
   If Index = 0 Then
      If Text7 = 專利權消滅 Then
         If KeyAscii <> 89 Then
            MsgBox "來函性質為專利權消滅時，必須為 Y !", vbCritical
            KeyAscii = 89
         End If
      End If
   End If
End Sub

Private Sub Text31_GotFocus()
   InverseTextBox Text31
End Sub
'Modify By Cheng 2002/05/31
'Private Sub Text6_GotFocus()
'   InverseTextBox Text6
'End Sub
'
'Private Sub Text6_Validate(Cancel As Boolean)
'   If Text6 = "" Then
'      MsgBox "一般來函日期不可空白 !", vbCritical
'      Cancel = True
'   Else
'      If ChkDate(Text6) Then
'         If Val(Text6) > Val(strSrvDate(2)) Then
'            MsgBox "一般來函日期不可大於系統日 !", vbCritical
'            Cancel = True
'         End If
'      Else
'         Cancel = True
'      End If
'   End If
'   If Cancel = True Then TextInverse Text6
'End Sub

'2013/8/19 add by sonia
Private Sub Text6_GotFocus()
Dim intPos As Integer
   
   If Me.Text7.Text = "1506" And Label3(1) = "行政訴訟" Then
      With Me.Text6
         If Len("" & .Text) > 0 Then
            intPos = InStr("" & .Text, "第")
            If intPos > 0 Then
               .SelStart = intPos
               .SelLength = 0
            End If
         End If
      End With
   End If
End Sub
'2013/8/19 end

Private Sub Text6_Validate(Cancel As Boolean)
   If Not CheckLengthIsOK(Text6, Text6.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub Text7_Change()
   'Add by Morgan 2007/7/24
   If Len(Text7) = 4 Then
      'Modified by Lydia 2024/05/28 改成模組
      ''Added by Lydia 2022/05/03 FCP-062174審定前不收費控制: (補上)判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
      'If pa(16) = "" And InStr("FCP062174000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      '      Text27(1) = "N": m_CP16 = ""
      ''FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
      'ElseIf pa(16) <> "1" And InStr("FCP067004000", pa(1) & pa(2) & pa(3) & pa(4)) > 0 Then
      If PUB_GetCP20forSpec(pa(1), pa(2), pa(3), pa(4), pa(16)) = "N" Then
      'end 2024/05/28
            Text27(1) = "N": m_CP16 = ""
      Else
      'end 2022/05/03
         'Modify by Morgan 2008/3/27 +pa75
         'Modify by Morgan 2008/4/10 +本所案號
         Text27(1) = PUB_GetCP20(Text2, Text7, m_CP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
      End If 'added by Lydia 2022/05/03
      SetException 'Added by Morgan 2015/4/20
      
      'Added by Morgan 2025/3/4 改案件性質要清除勾選，因要控制延期受理不可以點實審及優先權
      strExc(0) = ""
      For intI = 1 To MSHFlexGrid1.Rows - 1
         If MSHFlexGrid1.TextMatrix(intI, 0) <> "" Then
            MSHFlexGrid1.TextMatrix(intI, 0) = ""
            strExc(0) = "Y"
         End If
      Next
      If strExc(0) = "Y" Then
         MsgBox "來函性質有變動，本案期限勾選已清除，若有需要，請重新勾選！", vbInformation
      End If
      'end 2025/3/4
   End If
   
   'Added by Morgan 2012/11/13
   If pa(9) = 台灣國家代號 And Text7.Text = "1802" Then
      SSTab1.TabVisible(2) = True
      If pa(8) = "3" Then
         chkItem(0).Enabled = False
         chkItem(1).Enabled = False
         chkItem(2).Enabled = True
      Else
         chkItem(2).Enabled = False
      End If
   Else
      SSTab1.TabVisible(2) = False
   End If
   'end 2012/11/13
   
End Sub

Private Sub Text7_GotFocus()
   InverseTextBox Text7
End Sub

Private Sub Text7_LostFocus()
   Dim rsTmp As ADODB.Recordset
   Dim rsTmp1 As ADODB.Recordset
   Dim rsTmp2 As ADODB.Recordset
       
    'Add By Cheng 2003/05/14
    '若未輸入來函性質則不檢查
    If Me.Text7.Text = "" Then Exit Sub
   If ChgType(7) = False Then
      Me.SSTab1.Tab = 0
      Me.Text7.SetFocus
      Text7_GotFocus
   End If
   
   ' 90.06.26 modify by louis
   
'Remove by Morgan 2007/7/24 改抓CPM設定值
'   Select Case Text7
'      Case "1604":
'         'Remove by Morgan 2007/2/1 不會再輸1604
'         'Text27(0) = "Y"
'         'End 2007/2/1
'         Text27(1) = Empty
'      Case "1605":
'         Text27(1) = Empty
'      Case "1903":
'         Text27(1) = "Y"
'      Case Else:
'         Text27(1) = Empty
'   End Select
'end 2007/7/24

    'Modify By Cheng 2002/12/11
'   'Add By Cheng 2002/05/31
'   '以下案件性質資料, 畫面上預設之承辦人抓該案號案件進度檔案件性質為"翻譯"(201)或"檢視中說"(209)或"製作中說"(210)且收文日最小+總收文號最小+"A"類總收文號,
'   '再以該最小總收文號抓承辦人繪圖人員案件進度檔的"核稿人"(EP04),若無資料或有資料而無核稿人時,則抓案件進度檔的承辦人(若有承辦人資料帶出姓名)
'   If (Me.Text7.Text >= "1201" And Me.Text7.Text <= "1203") Or _
'      Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or (Me.Text7.Text >= "1301" And Me.Text7.Text <= "1307") Or _
'      Me.Text7.Text = "1401" Or Me.Text7.Text = "1502" Or (Me.Text7.Text >= "1504" And Me.Text7.Text <= "1507") Or _
'      Me.Text7.Text = "1801" Or Me.Text7.Text = "1802" Or (Me.Text7.Text >= "1805" And Me.Text7.Text = "1808") Or _
'      Me.Text7.Text = "1903" Then
'
'      Set rsTmp2 = New ADODB.Recordset
'      strSQL = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND (CP10='201' OR CP10='209' OR CP10='210') AND SUBSTR(CP09,1,1)='A' " & _
'               " ORDER BY CP05,CP09 "
'      rsTmp2.CursorLocation = adUseClient
'      rsTmp2.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp2.RecordCount > 0 Then
'         rsTmp2.MoveFirst
'         Set rsTmp1 = New ADODB.Recordset
'         strSQL = "SELECT EP04 FROM ENGINEERPROGRESS WHERE EP02='" & rsTmp2.Fields(0).Value & "'"
'         rsTmp1.CursorLocation = adUseClient
'         rsTmp1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsTmp1.RecordCount > 0 Then
'            If Len("" & rsTmp1.Fields(0).Value) > 0 Then
'               Text16 = "" & rsTmp1.Fields(0).Value
'               Label3(7) = GetStaffName(Text16)
'            Else
'               GoTo GetPromoterNO
'            End If
'         Else
'GetPromoterNO:
'            Set rsTmp = New ADODB.Recordset
'            strSQL = "SELECT CP14 FROM CASEPROGRESS " & _
'                     "WHERE CP09 = '" & rsTmp2.Fields(0).Value & "'"
'            rsTmp.CursorLocation = adUseClient
'            rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsTmp.RecordCount > 0 Then
'               rsTmp.MoveFirst
'               If IsNull(rsTmp.Fields("CP14")) = False Then
'                  Text16 = rsTmp.Fields("CP14")
'                  Label3(7) = GetStaffName(Text16)
'               End If
'            End If
'            rsTmp.Close
'            Set rsTmp = Nothing
'         End If
'         If rsTmp1.State <> adStateClosed Then rsTmp1.Close
'         Set rsTmp1 = Nothing
'      End If
'      rsTmp2.Close
'      Set rsTmp2 = Nothing
'
'   Else
'      ' 90.06.26 modify by louis
'      Select Case Text7
'         ' 以下案件性質的資料其承辦人為同本所案號但案件性質為翻議(201)的那一筆承辦人
'         Case "1002", "1201", "1202", "1203", "1301", "1302", "1303", "1304", "1305", "1306", "1307", _
'              "1401", "1502", "1504", "1505", "1506", "1801", "1802", "1805", "1806", "1807", "1808", "1903":
'               Set rsTmp = New ADODB.Recordset
'               strSQL = "SELECT * FROM CASEPROGRESS " & _
'                        "WHERE CP01 = '" & pa(1) & "' AND " & _
'                              "CP02 = '" & pa(2) & "' AND " & _
'                              "CP03 = '" & pa(3) & "' AND " & _
'                              "CP04 = '" & pa(4) & "' AND " & _
'                              "CP10 = '" & "201" & "' "
'               rsTmp.CursorLocation = adUseClient
'               rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'               If rsTmp.RecordCount > 0 Then
'                  rsTmp.MoveFirst
'                  If IsNull(rsTmp.Fields("CP14")) = False Then
'                     Text16 = rsTmp.Fields("CP14")
'                  End If
'               End If
'               rsTmp.Close
'               Set rsTmp = Nothing
'         Case Else
'      End Select
'   End If
    'Modify By Cheng 2003/04/08
''92.3.23 MODIFY BY SONIA 取消1502,1506
'    If (Val(Me.Text7.Text) >= 1201 And Val(Me.Text7.Text) <= 1203) Or _
'        (Val(Me.Text7.Text) = 1210 And Val(Me.Text7.Text) = 1212) Or (Val(Me.Text7.Text) >= 1301 And Val(Me.Text7.Text) <= 1307) Or _
'        Val(Me.Text7.Text) = 1401 Or (Val(Me.Text7.Text) >= 1504 And Val(Me.Text7.Text) <= 1505) Or Val(Me.Text7.Text) = 1507 Or _
'        Val(Me.Text7.Text) = 1801 Or Val(Me.Text7.Text) = 1802 Or (Val(Me.Text7.Text) >= 1805 And Val(Me.Text7.Text) <= 1808) Or _
'        Val(Me.Text7.Text) = 1903 Then
'        'Modify By Cheng 2002/12/18
'        '不預設承辦人
''        Me.Text16.Text = PUB_GetSpecCP14(pa(1) & pa(2) & pa(3) & pa(4))
''        If Me.Text16.Text <> "" Then ChgType 16
'        Me.Text16.Text = ""
'        Me.Label3(7).Caption = ""
'    Else
'        Text16 = cp(14): ChgType 16
'    End If
   
   'Add By Sindy 2016/5/31 預設承辦人
   'Mark by Lydia 2023/12/19 併入共用模組
'   If cp(10) = "201" Then '核稿人/承辦人
'      strExc(0) = "select cp14,ep04 from caseprogress,engineerprogress" & _
'                  " where cp09='" & cp(9) & "'" & _
'                  " and cp09=ep02(+)"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If "" & RsTemp.Fields("ep04") <> "" Then
'            Text16.Text = RsTemp.Fields("ep04")
'         Else
'            Text16.Text = RsTemp.Fields("cp14")
'         End If
'      End If
'   ElseIf cp(10) = "235" Or cp(10) = "209" Or cp(10) = "210" Then '承辦人
'      strExc(0) = "select cp14 from caseprogress" & _
'                  " where cp09='" & cp(9) & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Text16.Text = RsTemp.Fields("cp14")
'      End If
'   Else
   'end 2023/12/19
      'Modify by Morgan 2004/4/20
      '若帶出人員屬 F51 或 F52 則取消預設
      'Text16 = PUB_GetFCPPromoterNo(cp(9), Me.Text7.Text, cp(14)): ChgType 16
      Text16 = PUB_GetFCPPromoterNo(cp(9), Me.Text7.Text, cp(14))
      If Text16 <> "" Then
         Dim stArea As String
         stArea = GetStaffDepartment(Text16.Text)
         If stArea = "F51" Or stArea = "F52" Then
            Text16.Text = ""
'         Else
'            ChgType 16
         End If
      End If
      'Modify end
   'End If 'Mark by Lydia 2023/12/19
   ChgType 16
   '2016/5/31 END
   
   '2006/4/7 ADD BY SONIA
   If Text7 = 延期受理 Then
      Text8.Enabled = False
   Else
      Text8.Enabled = True
   End If
   '2006/4/7 END
   
   'Add by Morgan 2008/2/25
   '若案件性質為被異議(1801), 被舉發(1802), 對造號數預設為申請案號
   Select Case Me.Text7.Text
   Case "1802"
       If Me.Text19.Text = "" Then Me.Text19.Text = Me.Text1.Text
       If Me.Text23.Text = "" Then Me.Text23.Text = "N"
   
   '若案件性質為受理技術報告(1405), 對造號數預設為申請案號
   Case "1405", "1810"
       If Me.Text19.Text = "" Then Me.Text19.Text = Me.Text1.Text
       If Me.Text23.Text = "" Then Me.Text23.Text = "e"
       
   Case Else
       Me.Text23.Text = ""
       Me.Text19.Text = "" 'Add by Morgan 2011/6/16
   End Select
   'end 2008/2/25
   
   '2013/8/19 add by sonia 內專需求
   Label8.Visible = False
   Text6.Visible = False
   Text6.Locked = True
   If Text7 = "1506" And Label3(1) = "行政訴訟" Then
      Label8.Visible = True
      Text6.Visible = True
      Text6.Locked = False
      Text6 = Val(Left(DBDATE(cp(27)), 4) - 1911) & "年度行專訴字第號"
   End If
   '2013/8/19 end
   
   'Added by Lydia 2023/09/25
   If cp(1) = "FCP" And m_strIR01 <> "" And txtDelivery.Visible = True Then
      If Len(Text7) = 4 Then
         If InStr("1210,1211,1812", Text7) > 0 Then '1210通知準備程序，1211通知言詞辯論，1812通知聽證=>不用輸
            txtDelivery.Enabled = False
            txtDelivery = ""
         Else '游標設定在"送達日期"之欄位，供程序人員輸入
            txtDelivery.Enabled = True
            txtDelivery.SetFocus
            txtDelivery_GotFocus
         End If
      End If
   End If
   'end 2023/09/25
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
Dim strTempName As String
        
    'Add By Cheng 2003/01/17
    '若未輸入來函性質, 則不檢查
    If Me.Text7.Text = "" Then Exit Sub
   If Text7 = "" Then
      MsgBox "來函性質不可空白 !", vbCritical
      Cancel = True
   Else
      If Len(Text7) <> 4 Then
         MsgBox "來函性質錯誤，請重新輸入 !", vbCritical
         Cancel = True
      Else
         If Text7 = 核准 Or Text7 = 核駁 Or Text7 = 改變原處分 Then
            MsgBox "來函性質不可為核准或核駁或改變原處分!", vbCritical
            Cancel = True
         Else
            'Modify By Cheng 2002/06/03
'            If ChgType(7) = False Then Cancel = True
            'Add By Cheng 2002/06/03
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCaseProperty(pA(1), Text7.Text, strTempName, False) = False Then
            If ClsPDGetCaseProperty(pa(1), Text7.Text, strTempName, False) = False Then
               Label3(4).Caption = ""
               Cancel = True
            End If
            '2010/11/17 add by sonia
            If Text7.Text = "1502" And cp(10) <> "501" And cp(10) <> "503" And cp(10) <> "507" Then
               MsgBox "來函為撤銷原處分, 請點選 訴願或行政訴訟或行政訴訟上訴！", vbCritical + vbOKOnly, MsgText(9001)
               Cancel = True
            End If
            '2010/11/17 end
         End If
      End If
      'Modify By Cheng 2002/06/03
'      ' 90.06.26 modify by louis
'      Select Case Text7
'         Case "1604":
'            Text27(0) = "Y"
'         Case Else:
'            Text27(0) = Empty
'      End Select
'      'Add By Cheng 2002/05/31
'      '以下案件性質資料, 畫面上預設之承辦人抓該案號案件進度檔案件性質為"翻譯"(201)或"檢視中說"(209)或"製作中說"(210)且收文日最小+總收文號最小+"A"類總收文號,
'      '再以該最小總收文號抓承辦人繪圖人員案件進度檔的"核稿人"(EP04),若無資料或有資料而無核稿人時,則抓案件進度檔的承辦人(若有承辦人資料帶出姓名)
'      If (Me.Text7.Text >= "1201" And Me.Text7.Text <= "1203") Or _
'         Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or (Me.Text7.Text >= "1301" And Me.Text7.Text <= "1307") Or _
'         Me.Text7.Text = "1401" Or Me.Text7.Text = "1502" Or (Me.Text7.Text >= "1504" And Me.Text7.Text <= "1507") Or _
'         Me.Text7.Text = "1801" Or Me.Text7.Text = "1802" Or (Me.Text7.Text >= "1805" And Me.Text7.Text = "1808") Or _
'         Me.Text7.Text = "1903" Then
'
'         Set rsTmp2 = New ADODB.Recordset
'         strSQL = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND (CP10='201' OR CP10='209' OR CP10='210') AND SUBSTR(CP09,1,1)='A' " & _
'                  " ORDER BY CP05,CP09 "
'         rsTmp2.CursorLocation = adUseClient
'         rsTmp2.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsTmp2.RecordCount > 0 Then
'            rsTmp2.MoveFirst
'            Set rsTmp1 = New ADODB.Recordset
'            strSQL = "SELECT EP04 FROM ENGINEERPROGRESS WHERE EP02='" & rsTmp2.Fields(0).Value & "'"
'            rsTmp1.CursorLocation = adUseClient
'            rsTmp1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsTmp1.RecordCount > 0 Then
'               If Len("" & rsTmp1.Fields(0).Value) > 0 Then
'                  Text16 = "" & rsTmp1.Fields(0).Value
'                  Label3(7) = GetStaffName(Text16)
'               Else
'                  GoTo GetPromoterNO
'               End If
'            Else
'GetPromoterNO:
'               Set rsTmp = New ADODB.Recordset
'               strSQL = "SELECT CP14 FROM CASEPROGRESS " & _
'                        "WHERE CP09 = '" & rsTmp2.Fields(0).Value & "'"
'               rsTmp.CursorLocation = adUseClient
'               rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'               If rsTmp.RecordCount > 0 Then
'                  rsTmp.MoveFirst
'                  If IsNull(rsTmp.Fields("CP14")) = False Then
'                     Text16 = rsTmp.Fields("CP14")
'                     Label3(7) = GetStaffName(Text16)
'                  End If
'               End If
'               rsTmp.Close
'               Set rsTmp = Nothing
'            End If
'            If rsTmp1.State <> adStateClosed Then rsTmp1.Close
'            Set rsTmp1 = Nothing
'         End If
'         rsTmp2.Close
'         Set rsTmp2 = Nothing
'
'      Else
'         ' 90.06.26 modify by louis
'         Select Case Text7
'            ' 以下案件性質的資料其承辦人為同本所案號但案件性質為翻議(201)的那一筆承辦人
'            Case "1002", "1201", "1202", "1203", "1301", "1302", "1303", "1304", "1305", "1306", "1307", _
'                 "1401", "1502", "1504", "1505", "1506", "1801", "1802", "1805", "1806", "1807", "1808", "1903":
'                  Set rsTmp = New ADODB.Recordset
'                  strSQL = "SELECT * FROM CASEPROGRESS " & _
'                           "WHERE CP01 = '" & pa(1) & "' AND " & _
'                                 "CP02 = '" & pa(2) & "' AND " & _
'                                 "CP03 = '" & pa(3) & "' AND " & _
'                                 "CP04 = '" & pa(4) & "' AND " & _
'                                 "CP10 = '" & "201" & "' "
'                  rsTmp.CursorLocation = adUseClient
'                  rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                  If rsTmp.RecordCount > 0 Then
'                     rsTmp.MoveFirst
'                     If IsNull(rsTmp.Fields("CP14")) = False Then
'                        Text16 = rsTmp.Fields("CP14")
'                     End If
'                  End If
'                  rsTmp.Close
'                  Set rsTmp = Nothing
'            Case Else
'         End Select
'      End If
   End If
'2010/11/17 cancel by sonia
'   If Cancel = False Then
'      If Text7 = "1502" Then
'         EnableTextBox Text13, True
'      Else
'         EnableTextBox Text13, False
'      End If
'   End If
'2010/11/17 end

   If Cancel = True Then
      TextInverse Text7
   End If
End Sub

'Added by Morgan 2015/4/20
Private Sub Text8_Change()
   If Len(Text8) = 3 Then SetException
End Sub

Private Sub Text8_Click()
   If Text8.ListIndex >= 0 Then
      Text8.Text = Text8.ItemData(Text8.ListIndex)
      ChgType 8, Text8.Text
   End If
   
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Val(Text8.Text) = 0 Then
      Text8.Text = ""
   Else
      Text8.Text = Val(Text8)
      Cancel = Not ChgType(8, Text8.Text)
   End If
End Sub

Private Sub Text8_GotFocus()
   Text8.SelStart = 0
   Text8.SelLength = Len(Text8.Text)
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1500: .Text = "下一程序"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
        'Add By Cheng 2003/01/17
        '加備註
      .col = 6: .ColWidth(6) = 1500: .Text = "備註"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 1500: .Text = "解除期限日期"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 0 'NP22
      'Add by Morgan 2006/1/24
      .col = 9: .ColWidth(9) = 0 '總收文號
      'Added by Morgan 2025/3/4
      .col = 10: .ColWidth(10) = 0 '案件性質
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
Dim i As Integer
   
   'Remove by Morgan 2007/10/29 取消
   ''Add by Morgan 2007/10/24 延期受理不可以點
   'If Text7 = "1004" Then
   '   MsgBox "【延期受理】不可點選下一程序！"
   '   Exit Sub
   'End If
   
   'Added by Morgan 2025/3/4 延期受理不可以點實審及優先權
   If Text7 = "1004" Then
      strExc(1) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 10)
      strExc(2) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 6)
      strExc(3) = ""
      If strExc(1) = "416" Then
         strExc(3) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1)
      ElseIf strExc(1) = "202" And InStr(strExc(2), "優先權證明") > 0 Then
         strExc(3) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 1) & "-優先權證明"
      End If
      If strExc(3) <> "" Then
         MsgBox "【延期受理】不可點選【" & strExc(3) & "】！", vbExclamation
         Exit Sub
      End If
   End If
   'end 2025/3/4
   
   GridClick MSHFlexGrid1, intLastRow, 0, 1
End Sub

Private Sub Text9_GotFocus()
'   InverseTextBox Text9
    'edit by nickc 2007/07/11 切換輸入法改用API
   'Text9.IMEMode = 1
   OpenIme
'  TextInverse Text9
Dim intPos As Integer
'Modify By Cheng 2002/04/22
'當來函性質為"1601"或"1604"時, 將游標設定在機關文號欄的"第"的後面, 其餘則放在"專"的後面
With Me.Text9
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, IIf(Me.Text7.Text = "1601" Or Me.Text7.Text = "1604", "第", "專"))
      If intPos > 0 Then
         .SelStart = intPos
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub Text9_LostFocus()
    'edit by nickc 2007/07/11 切換輸入法改用API
   'Text9.IMEMode = 2
   CloseIme
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim bFind As String
TxtValidate = False

'92.3.11 Add By SONIA
' 當來函性質為延期受理時, 未收文期限至少要選取一筆
'Modify by Morgan 2004/11/3 cp43 = "C類收文號" 的也要提示
'If Me.Text7.Text = 延期受理 And cp(43) = "" Then
If Me.Text7.Text = 延期受理 And (cp(43) = "" Or cp(43) > "C") Then
   bFind = False
   For ii = 1 To Me.MSHFlexGrid1.Rows - 1
      If Me.MSHFlexGrid1.TextMatrix(ii, 0) = "v" Then
         bFind = True
         Exit For
      End If
   Next ii
   If bFind = False Then
      If PUB_GetCPAftExt(pa(9), cp(9), m_UpdCP09) = False Then 'Added by Morgan 2025/9/2 若有抓到已收文的收文號則直接更新進度
         'Modify by Morgan 2004/11/3 改訊息
         'MsgBox "請先選取未收文期限的資料來做延期受理的處理!!!", vbExclamation + vbOKOnly
         MsgBox "請於延期程序之相關總收文號欄位補輸原下一程序之A類總收文號!!!", vbExclamation + vbOKOnly
         Cancel = True
         Exit Function
      End If
   End If
End If

'Added by Morgan 2018/8/28
'延期受理函及通知補文件函管控必須要有期限--葉敏莉
'通知補文件目前有預設期限應可不控管
If Text14(1) = "" And (Me.Text7.Text = 延期受理 Or Me.Text7.Text = 通知補文件) Then
   MsgBox "請輸入期限後再按確定！", vbExclamation, Label3(4).Caption & "期限檢查"
   Exit Function
End If
'end 2018/8/28

'Added by Morgan 2019/9/5 --何淑華 (108.9.2與智慧局客服電話確認,新型案申請人為外國人發審查意見通知函應為二個月期限)
'新型案輸入審查意見通知函若來函期限不是 2 個月,請彈訊息
'非新型案輸入審查意見通知函若來函期限不是 3 個月,請彈訊息
'Modified by Morgan 2020/8/19 +1232 --淑華
'modify by sonia 2024/11/21 加入FCP條件，因為FG不用
If pa(1) = "FCP" And (Text7 = "1202" Or Text7 = "1232") Then
   If pa(8) = "2" And Not (Option4(1).Value = True And Val(Text11) = 2) Then
      MsgBox "來函期限不是 2 個月，請與智慧局確認！"
   ElseIf pa(8) <> "2" And Not (Option4(1).Value = True And Val(Text11) = 3) Then
      MsgBox "來函期限不是 3 個月，請與智慧局確認！"
   End If
End If
'end 2019/9/5

'Modify By Cheng 2002/05/31
'If Me.Text6.Enabled = True Then
'   Cancel = False
'   Text6_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If

If Me.Text7.Enabled = True Then
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2016/4/18
If Me.Text7.Text = 通知補文件 And InStr(NewCasePtyList, cp(10)) > 0 And Text8 = 補文件 Then
   bFind = False
   For ii = 1 To Me.MSHFlexGrid1.Rows - 1
      If Me.MSHFlexGrid1.TextMatrix(ii, 0) = "v" Then
         If InStr(Me.MSHFlexGrid1.TextMatrix(ii, 6), "優先權證明文件") > 0 Then
            MsgBox "不可點選備註裡有「優先權證明文件」字樣的期限!!!", vbExclamation + vbOKOnly
            Cancel = True
            Exit Function
         End If
         bFind = True
         'Exit For
      End If
   Next ii
   
   'Removed by Morgan 2023/11/29 取消--何淑華 Ex:FCP-070671
   'If bFind = False Then
   '   MsgBox "請於本案期限中點選一筆相關期限!!!", vbExclamation + vbOKOnly
   '   Cancel = True
   '   Exit Function
   'End If
   'end 2023/11/29
   
End If
'2016/4/18 END

If Me.Text8.Enabled = True Then
   Cancel = False
   Text8_Validate Cancel
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

If Me.Text11.Enabled = True Then
   Cancel = False
   Text11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text12.Enabled = True Then
   Cancel = False
   Text12_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Modify By Cheng 2002/11/18
For Each objTxt In Text14
'   If objTxt.Enabled = True Then
      Cancel = False
'      Text14_Validate objTxt.Index, Cancel
        If Text14(objTxt.Index) <> "" Then
           If Not ChkDate(Text14(objTxt.Index)) Then
              Cancel = True
           Else
              'Add By Cheng 2002/10/15
              '若有輸入本所期限時, 不可小於系統日
              If objTxt.Index = 0 Then
                 If Len(Me.Text14(0).Text) = 8 Then
                    If Val(Me.Text14(0).Text) < strSrvDate(1) Then
                       MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                       Cancel = True
                    End If
                 ElseIf Len(Me.Text14(0).Text) = 7 Or Len(Me.Text14(0).Text) = 6 Then
                    If Val(Me.Text14(0).Text) + 19110000 < strSrvDate(1) Then
                       MsgBox "本所期限不可小於系統日!!!", vbExclamation + vbOKOnly
                       Cancel = True
                    End If
                 End If
              End If
              
              If objTxt.Index = 1 Then
                 If Not ChkRange(Text14(0), Text14(1), "本所期限、法定期限") Then
                    Cancel = True
                 Else
                    'edit by nickc 2007/02/05 不用 dll 了
                    'If objLawDll.ChkMRec(TransDate(Label3(6).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
                    If ClsLawChkMRec(TransDate(Label3(6).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
                       If Text14(0) <> TransDate(strExc(1), 1) Then
                          If MsgBox("與櫃台之來函收文記錄本所期限 ( " & TransDate(strExc(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                                'Modify By Cheng 2002/12/18
                                '按下否則不繼續作業
'                             frm06010604_1.Show
'                             Unload frm06010604_2
'                             Unload Me
                                Cancel = True
                          Else
                                'Modify By Cheng 2002/12/18
                                '按下是則繼續作業
'                             Text14(0) = ""
'                             Text14(1) = ""
                          End If
                       ElseIf Text14(1) <> TransDate(strExc(2), 1) Then
                          If MsgBox("與櫃台之來函收文記錄法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                                'Modify By Cheng 2002/12/18
                                '按下否則不繼續作業
'                             frm06010604_1.Show
'                             Unload frm06010604_2
'                             Unload Me
                                Cancel = True
                          Else
                                'Modify By Cheng 2002/12/18
                                '按下是則繼續作業
'                             Text14(0) = ""
'                             Text14(1) = ""
                          End If
                       End If
                     'Added by Morgan 2017/5/10 電子公文
                     ElseIf m_DocNo <> "" Then
                        If m_DeadLine <> "" Then
                           If Len(m_DeadLine) >= 7 Then
                              strExc(2) = m_DeadLine
                           ElseIf Right(m_DeadLine, 1) = "日" Then
                              strExc(2) = CompDate(2, Val(m_DeadLine), Label3(6))
                           ElseIf Right(m_DeadLine, 1) = "月" Then
                              strExc(2) = CompDate(1, Val(m_DeadLine), Label3(6))
                           End If
                           If Text14(1) <> TransDate(strExc(2), 1) Then
                              If MsgBox("與電子公文之法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                                 Cancel = True
                              End If
                           End If
                        End If
                     'end 2017/5/10
                    Else
                       If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Cancel = True
                    End If
                 End If
              End If
           End If
        End If
      If Cancel = True Then
         Exit Function
      End If
'   End If
Next

If Me.Text16.Enabled = True Then
   Cancel = False
   Text16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text17.Enabled = True Then
   Cancel = False
   Text17_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.Text26.Enabled = True Then
   Cancel = False
   Text26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add by Morgan 2004/5/18
'檢查機關文號
If Me.Text9.Enabled = True Then
   Cancel = False
   Text9_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add by Morgan 2010/2/8
'來函性質為1801或1802時控制1.對造號數不可空白 2.對造名稱(中,英,日)不可全部空白
If Text7 = "1801" Or Text7 = "1802" Then
   If RTrim(Text19) = "" Then
      SSTab1.Tab = 1
      MsgBox "對造號數不可空白！", vbExclamation
      Text19.SetFocus
      Exit Function
      
   ElseIf RTrim(Text23) = "" Then
      SSTab1.Tab = 1
      MsgBox "對造案件數代號不可空白！", vbExclamation
      Text23.SetFocus
      Exit Function
      
   ElseIf RTrim(Text20(3) & Text20(4) & Text20(5)) = "" Then
      SSTab1.Tab = 1
      MsgBox "對造名稱不可空白！", vbExclamation
      Text20(3).SetFocus
      Exit Function
   Else
      PUB_ChkCustNameExist Text20(3), Text20(4), Text20(5)
   End If
End If

'Added by Morgan 2012/11/5
If Text8 = "" And Text8.Tag <> "" Then
   If MsgBox("是否確定不需管制下一程序？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Text8.SetFocus
      Exit Function
   End If
End If
'end 2012/11/5

'Added by Morgan 2012/11/13
'102新法被舉發要勾撤銷被請求項目
'Modified by Morgan 2013/1/14 增加舉發事項
If SSTab1.TabVisible(2) = True And Val(strSrvDate(1)) > 20130000 Then
   Cancel = True
   For Each oChk In chkItem
      If oChk.Value = vbChecked Then
         If oChk.Index = 0 Then
            If txtItemCount = "" Then
               SSTab1.Tab = 2
               MsgBox "請輸入項數", vbExclamation, "舉發聲明"
               If txtItemCount.Enabled Then txtItemCount.SetFocus
               Exit Function
            End If
         ElseIf oChk.Index = 1 Then
            If txtItemList = "第項" Then
               SSTab1.Tab = 2
               MsgBox "請輸入項次", vbExclamation, "舉發聲明"
               If txtItemList.Enabled Then txtItemList.SetFocus
               Exit Function
            ElseIf PUB_ChkItemList(txtItemList) = False Then
               SSTab1.Tab = 2
               MsgBox "撤銷部分請求項格式錯誤！", vbExclamation, "舉發聲明"
               If txtItemList.Enabled Then txtItemList.SetFocus
               Exit Function
            End If
         ElseIf oChk.Index = 6 Then
            For intI = 0 To 1
               If txtYear(intI) = "" Then
                  SSTab1.Tab = 2
                  MsgBox "請輸入年度!", vbExclamation, "舉發聲明"
                  txtYear(intI).SetFocus
                  Exit Function
               End If
               If txtMonth(intI) = "" Then
                  SSTab1.Tab = 2
                  MsgBox "請輸入月份!", vbExclamation, "舉發聲明"
                  txtMonth(intI).SetFocus
                  Exit Function
               End If
               If txtDay(intI) = "" Then
                  SSTab1.Tab = 2
                  MsgBox "請輸入日期!", vbExclamation, "舉發聲明"
                  txtDay(intI).SetFocus
                  Exit Function
               End If
               If Not IsDate((Val(txtYear(intI)) + 1911) & "/" & txtMonth(intI) & "/" & txtDay(intI)) Then
                  SSTab1.Tab = 2
                  MsgBox "日期錯誤，請重新輸入！", vbExclamation, "舉發聲明"
                  txtYear(intI).SetFocus
                  Exit Function
               End If
            Next
            If CDate((Val(txtYear(0)) + 1911) & "/" & txtMonth(0) & "/" & txtDay(0)) > CDate((Val(txtYear(1)) + 1911) & "/" & txtMonth(1) & "/" & txtDay(1)) Then
               SSTab1.Tab = 2
               MsgBox "起日不可晚於迄日，請重新輸入！", vbExclamation, "舉發聲明"
               txtYear(0).SetFocus
               Exit Function
            End If
            
         End If
         Cancel = False
         Exit For
      End If
   Next
   If Cancel = True Then
      SSTab1.Tab = 2
      MsgBox "請選擇被請求撤銷項目！", vbExclamation, "舉發聲明"
      Exit Function
   End If
End If
'end 2012/11/13

'Added by Morgan 2015/10/14
If Text14(1) <> "" Then
   If DBDATE(Text14(1)) > CompDate(1, 6, Label3(6)) Then
      MsgBox "法定期限大於來函收文日6個月!!", vbCritical
      Exit Function
   End If
End If
'end 2015/10/14

'Added by Lydia 2023/09/25 若為來函期限2次確認退回時需檢查法限是否一致
If m_strIR01 <> "" Then
   If PUB_ChkReKeyInOk(m_strIR01, m_strIR02, m_strIR03, m_strIR04, Text14(1).Text, m_bolReKeyInOK) = False Then
      Text14(1).SetFocus
      Exit Function
   End If
   If cp(1) = "FCP" And txtDelivery.Visible = True And txtDelivery.Enabled = True Then
      If Trim(txtDelivery) = "" Then
         MsgBox "送達日期不可空白！", vbExclamation
         txtDelivery.SetFocus
         txtDelivery_GotFocus
         Exit Function
      Else
         If TransDate(Label3(6), 2) <> TransDate(txtDelivery, 2) Then
            MsgBox "送達日期與來函收文日不一致，請確認！", vbExclamation
            txtDelivery.SetFocus
            txtDelivery_GotFocus
            Exit Function
         End If
      End If
   End If
End If
'end 2023/09/25

If Check1004 = False Then Exit Function 'Added by Morgan 2015/10/22

'Added by Morgan 2024/5/21 --Sharon
'訴願案設定(排除日本部)
'"智慧局答辯函"的相關收文號為"501訴願"， 設定本所期限為5個工作天，承辦期限往前-2天，在接洽單上備註"請務必在5個工作天內送件"
'"206 補充說明"的相關收文號為"501訴願"， 設定本所期限為3個工作天，承辦期限往前-2天，於分案時，自動發一封email給承辦人，主旨: "此為訴願案補充說明，請務必在3個工作天內送件"
'期限計算都不含當日
m_strExtNote = ""
If Text7 = "1506" And cp(10) = "501" And pa(150) <> "3" Then
   m_strExtNote = "請務必在5個工作天內送件!!"
   strExc(1) = CompDate(2, 1, strSrvDate(1))  '先抓次日以免是非工作日上班
   strExc(2) = TransDate(CompWorkDay(5, strExc(1)), 1) '本所期限=收文日+5個工作天
   strExc(3) = TransDate(CompWorkDay(3, strExc(1)), 1) '承辦期限=本所期限-2個工作天=收文日+3個工作天
   If Text14(0) <> strExc(2) Or Text17 <> strExc(3) Then
      Text14(0) = strExc(2)
      Text17 = strExc(3)
      MsgBox "【1506 智慧局答辯函】的相關收文號為【501訴願】，本所期限已設定為收文日+5個工作天(不含收文日)，承辦期限已設定為本所期限-2個工作天！", vbExclamation
   End If
End If
'end 2024/5/21

TxtValidate = True
End Function

Private Sub Text9_Validate(Cancel As Boolean)
   If CheckLengthIsOK(Text9, Text9.MaxLength) = False Then
      Cancel = True
      Text9.SetFocus
      Text9_GotFocus
   End If
End Sub

Private Sub txtItemCount_GotFocus()
   TextInverse txtItemCount
   CloseIme
End Sub

Private Sub txtItemCount_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDay_GotFocus(Index As Integer)
   TextInverse txtDay(Index)
   CloseIme
End Sub

Private Sub txtDay_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub txtMonth_GotFocus(Index As Integer)
   TextInverse txtMonth(Index)
   CloseIme
End Sub

Private Sub txtMonth_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii > Asc("9") Or KeyAscii < Asc("0")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtYear_GotFocus(Index As Integer)
   TextInverse txtYear(Index)
   CloseIme
End Sub

Private Sub txtYear_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub
'Added by Morgan 2015/10/21
'檢查申復延期受理期限
Private Function Check1004() As Boolean
   If Text7 = "1004" Then
      'Added by Morgan 2019/9/5--Sharon
      'Modified by Morgan 2020/8/19 +1232(239) --淑華
      'Modified by Morgan 2025/9/2 不必再限制案件性質 --Winfrey/Sharon
      'Modified by Morgan 2025/10/17 NP補文件的延期相關收文號可能是申請程序(A類),+判斷CP30
      'If cp(43) > "C" Then
      If cp(43) > "C" Or cp(30) <> "" Then
      'end 2025/10/17
         'strExc(0) = "select np09 from caseprogress a,nextprogress where cp09='" & cp(43) & "' and cp10 in ('1202','1232') and np01(+)=cp09 and np07 in ('204','205','239')"
         strExc(0) = "select np09 from caseprogress a,nextprogress where cp09='" & cp(43) & "' and np01(+)=cp09 and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04"
      Else
         'strExc(0) = "select a.cp07 np09 from caseprogress a,caseprogress b where a.cp09='" & cp(43) & "' and a.cp10 in ('204','205','239') and b.cp09(+)=a.cp43 and b.cp10 in ('1202','1232')"
         strExc(0) = "select a.cp07 np09 from caseprogress a where a.cp09='" & cp(43) & "'"
      End If
      
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If DBDATE(Text14(1)) <> "" & RsTemp("np09") Then
            MsgBox "來函期限與原期限【" & TransDate(RsTemp("np09"), 1) & "】不同，請確認！", vbExclamation
         End If
      End If
      'end 2019/9/5
      
      'Modified by Morgan 2020/8/19 +239 --淑華
      strExc(0) = " select np08, np09, 1 srt from caseprogress a, nextprogress where a.cp09='" & cp(9) & "' and a.cp10='404' and np01(+)=a.cp43 and np07 in ('205','239')" & _
         " union all select b.cp06, b.cp07, 2 from caseprogress a, caseprogress b where a.cp09='" & cp(9) & "' and a.cp10='404' and b.cp09(+)=a.cp43 and b.cp10 in ('205','239')" & _
         " order by srt"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '法限相同時所限設為原所限
         If DBDATE(Text14(1)) = RsTemp("np09") Then
            
            'Modified by Morgan 2019/7/15 取消,外專台灣案所限改以工作天計算--David確認
            'strExc(1) = TransDate(RsTemp("np08"), 1)
            'If strExc(1) <> Text14(0) Then
            '   MsgBox "本次來函為申復的延期受理，本所期限將更新為延期發文時的期限 " & strExc(1) & " !!", vbInformation, "申復延期受理函檢查"
            '   Text14(0) = strExc(1)
            'End If
            'end 2019/7/15
            
         '超過延期發文法限+10天
         ElseIf DBDATE(Text14(1)) > CompDate(2, 10, RsTemp("np09")) Then
            If MsgBox("法定期限超過申復延期發文時的期限+10天!!是否確定要繼續?", vbYesNo + vbExclamation + vbDefaultButton2, "申復延期受理函檢查") = vbNo Then
               Exit Function
            End If
            
         End If
      End If
   End If
   Check1004 = True
End Function

'Added by Morgan 2022/2/18 從cmdOK_Click抽出
Private Function Process() As Boolean
   Process = False
   
   'Add By Cheng 2003/01/17
   Dim ii As Integer '回圈序號
   
   ' 90.07.31 modify by louis (來函性質不可為空白)
   If IsEmptyText(Text7) = True Then
      MsgBox "來函性質不可為空白時 !", vbCritical
      Text7.SetFocus
      Exit Function
   End If
   If Text8 <> "" And (Text14(0) = "" Or Text14(1) = "") Then
      ' 90.07.31 modify by louis (更新顯示訊息及Focus)
      MsgBox "下一程序不為空白時，本所期限與法定期限不可空白 !", vbCritical
      If IsEmptyText(Text14(0)) = True Then
         'Modify By Cheng 2002/05/31
   '               Text14(0).SetFocus
         Me.Text8.SetFocus
      Else
         'Modify By Cheng 2002/05/31
   '               Text14(1).SetFocus
         Me.Text8.SetFocus
      End If
      Exit Function
   End If
   
   'Added by Morgan 2022/4/28
   m_UpdCP09 = ""
   If Text14(1) <> "" And Text8 = "" Then
      'Added by Morgan 2024/11/15 沒有點選下一程序期限才要彈訊息--Winfrey
      intI = 0
      For ii = 1 To Me.MSHFlexGrid1.Rows - 1
          If MSHFlexGrid1.TextMatrix(ii, 0) <> "" Then
            intI = 1
            Exit For
          End If
      Next ii
      If intI = 0 Then
      'end 2024/11/15
         If MsgBox("本案有官方來函期限，是否管制下一程序？", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            'Modified by Morgan 2022/5/4 延期受理不可輸入下一程序(考慮案件性質會輸錯,仍要User確認)
            'Text8.SetFocus
            If Text8.Enabled = True Then Text8.SetFocus
            'end 2022/5/4
            MsgBox "請點選下一程序期限或新增下一程序欄位性質！", vbInformation 'Added by Morgan 2024/11/15 --Winfrey
            Exit Function 'Added by Morgan 2024/10/21 選是應該要回原畫面
         
         'Added by Morgan 2024/11/15 --Winfrey
         '若非1004則 +再彈訊息詢
         ElseIf Text7 <> "1004" Then
            If MsgBox("期限是否已管制於進度檔?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
               '帶出進度檔性質供點選(一併更新期限)
               Set frm060101_2.fmParent = Me
               frm060101_2.Show vbModal
               If m_UpdCP09 = "" Then
                  MsgBox "未點選進度檔！", vbCritical
                  Exit Function
               End If
            ElseIf MsgBox("確定不需管制期限？", vbQuestion + vbOKCancel + vbDefaultButton2) = vbCancel Then
               Exit Function
            End If
         'end 2024/11/15
         End If
         
      End If 'Added by Morgan 2024/11/15
   End If
   'end 2022/4/28
   
   'Add By Cheng 2002/03/11
   If Me.Text14(0).Text <> "" Then
      If Len(Me.Text14(0).Text) = 8 Then
         If Val(Me.Text14(0).Text) < strSrvDate(1) Then
            MsgBox "本所期限不可小於系統日期!!!", vbExclamation
              'Modify By Cheng 2002/12/19
              If Me.Text14(0).Enabled Then
                  Me.Text14(0).SetFocus
                  Me.Text14(0).SelStart = 0
                  Me.Text14(0).SelLength = Len(Me.Text14(0).Text)
                  Exit Function
              End If
         End If
      Else
         If Val(Me.Text14(0).Text) + 19110000 < ServerDate Then
            MsgBox "本所期限不可小於系統日期!!!", vbExclamation
              'Modify By Cheng 2002/12/19
              If Me.Text14(0).Enabled Then
                  Me.Text14(0).SetFocus
                  Me.Text14(0).SelStart = 0
                  Me.Text14(0).SelLength = Len(Me.Text14(0).Text)
                  Exit Function
              End If
         End If
      End If
   End If
   'Add By Cheng 2002/05/06
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.Text14(0).Text) > 0 And Len(Me.Text17.Text) > 0 Then
      If Val(Me.Text14(0).Text) < Val(Me.Text17.Text) Then
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         Exit Function
      End If
   End If
   'Add By Cheng 2003/01/17
   '若輸入的來函性質為通知補文件或延期受理時
   If Me.Text7.Text = 通知補文件 Or Me.Text7.Text = 延期受理 Then
      For ii = 1 To Me.MSHFlexGrid1.Rows - 1
          '若有勾選本案期限
          If Me.MSHFlexGrid1.TextMatrix(ii, 0) <> "" Then
              If Me.Text14(0).Text = "" Or Me.Text14(1).Text = "" Then
                  MsgBox "本所期限及法定期限不可空白!!!", vbExclamation + vbOKOnly
                  Exit Function
              End If
          End If
      Next ii
   End If
   'Add by Morgan 2006/2/27 爭議案不可收文1605(通知年費逾期)
   If Text7.Text = "1605" And pa(23) <> "1" Then
      MsgBox "非申請案不可收文1605(通知年費逾期)！", vbExclamation
      Exit Function
   End If
   '2006/2/27 end
   
   'Add by Morgan 2008/2/25
   '若來函性質為被舉發(1802),第三人提起技術報告(1810),受理技術報告申請(1405) 則要檢查對造號數及對造案件數代號
   If Me.Text7.Text = "1802" Or Me.Text7.Text = "1810" Or Me.Text7.Text = "1405" Then
      strExc(0) = ""
      If Len(Trim("" & Me.Text19.Text)) <= 0 Then
         strExc(0) = "1"
      ElseIf Len(Trim("" & Me.Text23.Text)) <= 1 Then
         strExc(0) = "2"
      End If
      If strExc(0) <> "" Then
         MsgBox "請輸入本案件的對造資料 !", vbCritical
         Me.SSTab1.Tab = 1
         If strExc(0) = "1" Then
            Me.Text19.SetFocus
         ElseIf strExc(0) = "2" Then
            Me.Text23.SetFocus
            Text23_GotFocus
         End If
         Exit Function
      End If
   End If
   'end 2008/2/25
   'Add by Amy 2022/09/30 cp36放寬至200 檢查大小(存對造號數&對造案件數代號)
   If CheckLengthIsOK(Text19.Text & Text23.Text, 200, False) = False Then
        MsgBox "對造號數+對造案件數代號 " & vbCrLf & _
        MsgText(9205) & "200" & MsgText(9206) & "!", vbExclamation + vbOKOnly
        Me.Text19.SetFocus
        Exit Function
   End If
   'end 2022/09/30
   
   '2013/8/19  add by sonia
   If Me.Text7.Text = "1506" And Label3(1) = "行政訴訟" Then
     If Me.Text6.Text = Val(Left(DBDATE(cp(27)), 4) - 1911) & "年度行專訴字第號" Then
        MsgBox "行政訴訟之智慧局答辯函, 請輸入法院案號, 以便將來智商法院來函可查詢!!!", vbExclamation + vbOKOnly
        SSTab1.Tab = 0 'Added by Morgan 2023/5/31
        Me.Text6.SetFocus
        Exit Function
     End If
   End If
   '2013/8/19 end
   
   'Added by Morgan 2015/5/20
   If Text7.Text = "1232" Then
      'Added by Morgan 2016/5/19 Ex FCP-49368(櫃台來函本所案號輸錯,應為FCP-49198)
      If pa(8) <> "1" Then
         MsgBox "發明案才可輸入 " & Label3(4) & "！", vbCritical
         Exit Function
      End If
      'end 2016/5/19
      If PUB_IsDualApply(pa, m_stUPA) = False Then
         MsgBox "一案兩請未建立關聯案！", vbCritical
         Exit Function
      End If
   End If
   'end 2015/5/20
   
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Function
   
   'Add by Sindy 2021/11/22 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Function
   End If
   
   'Added by Lydia 2022/05/10 通知補充聽證資料之期限
   m_1812CP07 = ""
   If pa(1) = "FCP" And Text7.Text = "1812" Then
      If MsgBox("來函內文中，是否有通知補充聽證資料之期限？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
JumpReInput:
         strExc(2) = ""
         strExc(2) = UCase(InputBox("請在下方輸入通知補充聽證資料之期限，或者是空白表示沒有期限：", "通知補充聽證資料之期限", strSrvDate(2)))
         If strExc(2) <> "" Then
         '檢查日期
             If ChkDate(strExc(2)) = False Then
                   GoTo JumpReInput
             End If
             If strExc(2) < strSrvDate(2) Then
                If MsgBox("輸入期限" & strExc(2) & "小於系統日期，是否正確？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
                   GoTo JumpReInput
                End If
             End If
         End If
         If strExc(2) <> "" Then
            m_1812CP07 = DBDATE(strExc(2))
         End If
      End If
   End If
   'end 2022/05/10
   
   'Added by Morgan 2023/12/28
   '發明申請(101)的通知即將公開(1207)檢查是否有補文件未收文或未發文 --敏莉
   If Text7 = "1207" And cp(10) = "101" Then
      PUB_ChkUnAddDoc cp(9)
   End If
   'end 2023/12/28
   
   'Added by Lydia 2023/09/25
   If m_strIR01 <> "" Then
      '下載信件檔
      If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", IIf(pa(9) <> 台灣國家代號, "PAT", "RX"), , True) = False Then
         Exit Function
      End If
   End If
   'end 2023/09/25
   
   'Add By Sindy 2025/1/6
   If Me.Text7.Text = 通知補文件 And Trim(lblPA57.Caption) = "Y" Then
      If MsgBox("本案已上閉卷，是否要管制此期限？" & vbCrLf & vbCrLf & _
         "【是】(來函輸入2)畫面上(是否閉卷)Ｙ清空。系統將會" & vbCrLf & _
         "      清除基本檔的（是否閉卷、閉卷日期、閉卷原因）以利管控期限" & vbCrLf & _
         "      待補完文件後，請自行恢復閉卷" & vbCrLf & vbCrLf & _
         "【否】不管制期限", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
         Text27(0) = ""
      Else
         Text27(0) = "Y"
      End If
   End If
   '2025/1/6 END
   
   'Add by Morgan 2004/7/28
   '加漏斗
   Screen.MousePointer = vbHourglass
   If FormSave = False Then
      Screen.MousePointer = vbDefault
      MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Function
   End If
   
   'Added by Morgan 2017/12/4 FCP57047 收到智慧局未出接洽單之來函: 歸卷、一般來函(延期受理、通知補文件)、通知實審日、通知公開 提醒 "收到智慧局來函需退承辦報告客戶" -- Sharon
   If pa(1) = "FCP" And pa(2) = "057047" And (Text7 = "1003" Or Text7 = "1004" Or Text7 = "1207") Then
      MsgBox "本案收到智慧局來函需退承辦報告客戶！", vbInformation
   End If
   'end 2017/12/4
   
   'Add by Morgan 2008/8/21
   If Text7 = "1207" And Left(pa(75), 6) = "Y20304" Then
      MsgBox "本案為 ASAHI 的案件，來函要寄客戶！", vbInformation
   End If
   'END 2008/8/21
   
   
   'Added by Morgan 2012/11/5
   If Text8 = "" And Left(pa(75), 6) = "Y53309" Then
      MsgBox "本案需調卷轉承辦組報告並寄代！", vbInformation
   End If
   'end 2012/11/5
   
   'add by sonia 2016/6/17 Y34232+X48637的1207通知即將公開 (YASUTOMI+大日本印刷)也要提醒
   'modify by sonia 2017/8/14 再加Y54116+X48637的1207通知即將公開也要提醒
   'If pa(75) = "Y34232" And (Left(pa(26), 8) = "X48637" Or Left(pa(27), 8) = "X48637" Or Left(pa(28), 8) = "X48637" Or Left(pa(29), 8) = "X48637" Or Left(pa(30), 8) = "X48637") And Text7 = "1207" Then
   'modify by sonia 2017/10/16 再加Y47649+X48637,Y48651+X48637 也要提醒
   'modify by sonia 2017/12/19 取消Y48651+X48637
   'Modified by Morgan 2018/5/22 取消 Y54116+X48637 -- 郭怡瑩,吳若芬
   If (pa(75) = "Y34232" Or pa(75) = "Y47649") And (Left(pa(26), 8) = "X48637" Or Left(pa(27), 8) = "X48637" Or Left(pa(28), 8) = "X48637" Or Left(pa(29), 8) = "X48637" Or Left(pa(30), 8) = "X48637") And Text7 = "1207" Then
      MsgBox "請調卷退承辦智權同仁通知即將公開函！", vbInformation
   End If
   'end 2016/6/17
   
   'Added by Morgan 2022/4/12 +Y55105 --蘇暐嵐
   '智慧局所有來函彈跳提醒(除C類來函承辦人為工程師及已閉卷的案子)
   If pa(75) = "Y55105" And pa(57) = "" Then
      If PUB_GetST03(Text16) <> "F21" Then
         MsgBox "收到智慧局來函,請通知承辦寄代！", vbInformation
      End If
   End If
   'end 2022/4/12
            
   Screen.MousePointer = vbDefault
   'Modify By Cheng 2002/12/09
   '若新增至案件進度檔的C類資料, 若案件性質為1002,1201~1203,1210~1212,1301~1307,1401,1502,1504~1507,
   '1801,1802,1805~1808,1903, 則列印C類接洽記錄單
   '         'Add By Cheng 2002/01/28
   '         '若新增案件進度檔的來函性質為 1201-1203,1210,1211,1301-1307,1401,1502,1504-1507,1801,1802,1805-1808,1903 則列印C類接洽記錄單
   '         If (Text7 >= "1201" Or Text7 <= "1203") Or _
   '            Text7 = "1210" Or Text7 = "1211" Or (Text7 >= "1301" Or Text7 <= "1307") Or Text7 = "1401" Or Text7 = "1502" Or (Text7 >= "1504" Or Text7 <= "1507") Or _
   '            Text7 = "1801" Or Text7 = "1802" Or (Text7 >= "1805" Or Text7 <= "1808") Or Text7 = "1903" Then
   'MODIFY BY SONIA 2014/5/15 +1227最後通知
   'Modified by Lydia 2016/02/05 +1232通知擇一申復
   'Modify By Sindy 2016/5/31 + 通知補文件1003,相關總收文號為201,235,209,210要加印C類接洽記錄單
   'Modify By Sindy 2025/6/19 +1221通知申復
   If Val(Me.Text7.Text) = 1002 Or (Val(Text7.Text) >= 1201 And Val(Text7.Text) <= 1203) Or _
      (Val(Text7.Text) >= 1210 And Val(Text7.Text) <= 1212) Or (Val(Text7.Text) >= 1301 And Val(Text7.Text) <= 1307) Or _
      Val(Text7.Text) = 1401 Or Val(Text7.Text) = 1502 Or (Val(Text7.Text) >= 1504 And Val(Text7.Text) <= 1507) Or _
      Val(Text7.Text) = 1801 Or Val(Text7.Text) = 1802 Or (Val(Text7.Text) >= 1805 And Val(Text7.Text) <= 1808) Or _
      Val(Text7.Text) = 1903 Or Val(Text7.Text) = 1227 Or Val(Text7.Text) = 1232 Or _
      (Val(Me.Text7.Text) = 1003 And (cp(10) = "201" Or cp(10) = "235" Or cp(10) = "209" Or cp(10) = "210")) Or _
      Val(Text7.Text) = 1221 Then
      
      If m_strExtNote <> "" Then m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & m_strExtNote 'Added by Morgan 2024/5/22
      '列印C類接洽記錄單
      'Modify By Cheng 2003/03/04
      '未閉卷的才印
      'g_PrtForm001.PrintCForm m_strCP09ByCheng
      'Modified by Lydia 2018/12/17 FCP案C類接洽單同時列印並且上傳到卷宗區
      'If PUB_CaseClosed_1(pa(1), pa(2), pa(3), pa(4)) = False Then g_PrtForm001.PrintCForm m_strCP09ByCheng, m_strMemo
      'Modified by Lydia 2019/03/18 改成開啟Word
      'If PUB_CaseClosed_1(pa(1), pa(2), pa(3), pa(4)) = False Then g_PrtForm001.PrintCForm m_strCP09ByCheng, m_strMemo, , True
      If PUB_CaseClosed_1(pa(1), pa(2), pa(3), pa(4)) = False Then g_PrtForm001.PrintCFormNew m_strCP09ByCheng, m_strMemo, , True
   End If
   
   'Added by Morgan 2015/5/20
   If Text7.Text = "1232" Then
      StartLetter "17", m_strCP09ByCheng, "01"
      NowPrint m_strCP09ByCheng, "17", "01", False, strUserNum
   End If
   'end 2015/5/20
   
   'Add By Sindy 2023/12/4
   If (Label3(1) = "準備程序" Or Label3(1) = "言詞辯論") And pa(9) = "000" Then
      '1210=通知準備程序 1211=通知言詞辯論
      If Text7 = "1210" Or Text7 = "1211" Then
         Dim m_CP14 As String, m_CP13 As String
         Dim m_StrTo As String, m_StrSub As String, m_StrCont As String
         m_CP14 = ""
         strSql = " select CASEPROGRESS.*,NVL(ST04,' ') as ST04,cp13 from CASEPROGRESS,STAFF where CP14 = ST01(+) and CP09='" & strReceiveNo & "'"
         '請同時發給m_CP14(承辦人=>專利工程師), 但先檢查m_CP14 若為離職人員則改發 工程師主管
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount > 0 Then
           m_CP13 = adoRecordset.Fields("cp13")
           If adoRecordset.Fields("ST04") = "2" Then '離職人員代理
               m_CP14 = PUB_GetFCPEngSup(adoRecordset.Fields("cp14"), True) 'FCP工程師主管
           Else
               m_CP14 = CheckStr(adoRecordset.Fields("cp14"))
           End If
         End If
         'Modify By Sindy 2023/12/8 法律所調整內專行政訴訟開庭通知之系統通知信也請一併轉陳亮之; 商標一併調整
         'Modified by Lydia 2024/10/30 串法律所案源資料，抓出法律所案件有承辦人且收文日最大的進度，抓承辦人及所有出庭律師。
         'm_StrTo = Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & m_CP13 & IIf(Trim(m_CP14) <> "", ";" & m_CP14, "")
         m_StrTo = PUB_GetLosCL02list(Text2, Text3, Text4, Text5)
         m_StrTo = IIf(m_StrTo <> "", m_StrTo & ";", "") & Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & m_CP13 & IIf(Trim(m_CP14) <> "", ";" & m_CP14, "")
         'end 2024/10/30
         
         m_StrSub = "開庭通知--來函案件：" & Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5
         m_StrCont = "本所案號：" & Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5 & vbCrLf & _
                     "案件名稱：" & Me.Combo1.Text & vbCrLf & _
                     "案件性質：" & Label3(4) & vbCrLf & _
                     "申請人　：" & GetCustomerName(pa(26)) & vbCrLf & _
                     "承辦人　：" & GetStaffName(m_CP14) & vbCrLf & _
                     "智權人員　：" & GetStaffName(m_CP13) & vbCrLf & _
                     "法定期限：" & DBYEAR(Text14(1).Text) - 1911 & " 年 " & DBMONTH(Text14(1).Text) & " 月 " & DBDAY(Text14(1).Text) & " 日 " & vbCrLf & _
                     "時間地點：" & Text29 & vbCrLf & _
                     "法院案號：" & Text6
         PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
      End If
   End If
   '2023/12/4 END
               
   'Modified by Morgan 2017/6/6 調整順序
   Unload frm06010604_2
   Unload Me
   
   'Added by Lydia 2023/09/25
   If Me.m_strIR01 <> "" Then
      Unload frm06010604_1

      Forms(0).Tmpfrm04010519.GoNext
      Set Forms(0).Tmpfrm04010519 = Nothing
   'end 2023/09/25
   'Modified by Morgan 2017/5/10 電子公文
   'frm06010604_1.Show
   'frm06010604_1.Clear
   'Modified by Lydia 2023/09/25 +Else
   ElseIf m_DocNo <> "" Then
      Unload frm06010604_1
      frm060119.GoNext
   Else
      frm06010604_1.Show
      frm06010604_1.Clear
   End If
   'end 2017/5/10
   'end 2017/6/6
   
   Process = True
End Function

'Added by Lydia 2023/09/25
Private Sub txtDelivery_GotFocus()
   TextInverse txtDelivery
End Sub

Private Sub txtDelivery_Validate(Cancel As Boolean)
   If Trim(txtDelivery) <> "" Then
      If Not ChkDate(txtDelivery) Then
         txtDelivery.SetFocus
         txtDelivery_GotFocus
         Cancel = True
      End If
   End If
End Sub


