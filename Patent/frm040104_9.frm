VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_9 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文-異議/舉發"
   ClientHeight    =   6360
   ClientLeft      =   1212
   ClientTop       =   948
   ClientWidth     =   8064
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8064
   Begin VB.TextBox txtCP118 
      Height          =   270
      Left            =   7008
      MaxLength       =   1
      TabIndex        =   101
      Top             =   1632
      Width           =   255
   End
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   3975
      MaxLength       =   4
      TabIndex        =   18
      Top             =   5667
      Width           =   540
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   45
      TabIndex        =   71
      Top             =   2760
      Width           =   7935
      _ExtentX        =   13991
      _ExtentY        =   4466
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "對造(中)"
      TabPicture(0)   =   "frm040104_9.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label32"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label29"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text5(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text5(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "對造(英)"
      TabPicture(1)   =   "frm040104_9.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "Text5(6)"
      Tab(1).Control(3)=   "Text5(9)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "對造(日)"
      TabPicture(2)   =   "frm040104_9.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "Text5(7)"
      Tab(2).Control(3)=   "Text5(10)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "舉發聲明"
      TabPicture(3)   =   "frm040104_9.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label13"
      Tab(3).Control(1)=   "Label15"
      Tab(3).Control(2)=   "Label16(5)"
      Tab(3).Control(3)=   "Label17"
      Tab(3).Control(4)=   "Label19"
      Tab(3).Control(5)=   "chkItem(6)"
      Tab(3).Control(6)=   "chkItem(0)"
      Tab(3).Control(7)=   "chkItem(1)"
      Tab(3).Control(8)=   "chkItem(4)"
      Tab(3).Control(9)=   "chkItem(3)"
      Tab(3).Control(10)=   "chkItem(5)"
      Tab(3).Control(11)=   "chkItem(2)"
      Tab(3).Control(12)=   "txtItemCount"
      Tab(3).Control(13)=   "txtItemList"
      Tab(3).Control(14)=   "txtMonth(0)"
      Tab(3).Control(15)=   "txtYear(0)"
      Tab(3).Control(16)=   "txtYear(1)"
      Tab(3).Control(17)=   "txtMonth(1)"
      Tab(3).Control(18)=   "txtDay(1)"
      Tab(3).Control(19)=   "txtDay(0)"
      Tab(3).Control(20)=   "chkItem(7)"
      Tab(3).Control(21)=   "chkItem(8)"
      Tab(3).ControlCount=   22
      Begin VB.CheckBox chkItem 
         Caption         =   "同一人就相同創作，於同日分別申請發明專利及新型專利，其發明專利審定前，新型專利權已當然消滅或撤銷確定者"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   -71715
         TabIndex        =   99
         Top             =   990
         Width           =   4605
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "同一人於同日就相同創作分別申請發明及新型專利，已於申請時分別聲明，而其發明及新型專利權同時並存者"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -71715
         TabIndex        =   98
         Top             =   630
         Width           =   4605
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72165
         MaxLength       =   2
         TabIndex        =   92
         Top             =   2190
         Width           =   285
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -69780
         MaxLength       =   2
         TabIndex        =   95
         Top             =   2190
         Width           =   285
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70365
         MaxLength       =   2
         TabIndex        =   94
         Top             =   2190
         Width           =   285
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -71040
         MaxLength       =   3
         TabIndex        =   93
         Top             =   2190
         Width           =   420
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -73425
         MaxLength       =   3
         TabIndex        =   90
         Top             =   2190
         Width           =   420
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72705
         MaxLength       =   2
         TabIndex        =   91
         Top             =   2190
         Width           =   285
      End
      Begin VB.TextBox txtItemList 
         Enabled         =   0   'False
         Height          =   420
         Left            =   -74595
         TabIndex        =   83
         Text            =   "第項"
         Top             =   1320
         Width           =   2715
      End
      Begin VB.TextBox txtItemCount 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -72390
         TabIndex        =   81
         Top             =   870
         Width           =   375
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷設計專利權"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   -74865
         TabIndex        =   87
         Top             =   1980
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "專利權人所屬國家對中華民國申請專利不予受理者"
         Height          =   210
         Index           =   5
         Left            =   -71715
         TabIndex        =   86
         Top             =   1770
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "專利權人為非專利申請權人者"
         Height          =   210
         Index           =   3
         Left            =   -71715
         TabIndex        =   85
         Top             =   1350
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "共有專利申請權非由全體共有人提出申請者"
         Height          =   210
         Index           =   4
         Left            =   -71715
         TabIndex        =   84
         Top             =   1560
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷部分之請求項："
         Height          =   210
         Index           =   1
         Left            =   -74865
         TabIndex        =   82
         Top             =   1110
         Width           =   2400
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷全部請求項：共計"
         Height          =   210
         Index           =   0
         Left            =   -74865
         TabIndex        =   80
         Top             =   900
         Width           =   2535
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
         Left            =   -74865
         TabIndex        =   89
         Top             =   2220
         Width           =   7440
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   10
         Left            =   -73350
         TabIndex        =   12
         Top             =   660
         Width           =   5535
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "9763;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   7
         Left            =   -73350
         TabIndex        =   13
         Top             =   960
         Width           =   5535
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "9763;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   9
         Left            =   -73350
         TabIndex        =   10
         Top             =   660
         Width           =   5535
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "9763;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   6
         Left            =   -73350
         TabIndex        =   11
         Top             =   960
         Width           =   5535
         VariousPropertyBits=   671107099
         MaxLength       =   250
         Size            =   "9763;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   8
         Left            =   1650
         TabIndex        =   8
         Top             =   660
         Width           =   5535
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "9763;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   5
         Left            =   1650
         TabIndex        =   9
         Top             =   960
         Width           =   5535
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "9763;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "請求撤銷全部或部分請求項"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74865
         TabIndex        =   97
         Top             =   630
         Width           =   2160
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "( 例如：第 1,3,5-12 項 )"
         Height          =   180
         Left            =   -74595
         TabIndex        =   96
         Top             =   1740
         Width           =   1800
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "項"
         Height          =   180
         Index           =   5
         Left            =   -71985
         TabIndex        =   88
         Top             =   915
         Width           =   180
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "請求撤銷全部專利權"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -71715
         TabIndex        =   79
         Top             =   390
         Width           =   1620
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "請求撤銷發明(新型)專利權"
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
         Left            =   -74865
         TabIndex        =   78
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱(日):"
         Height          =   180
         Left            =   -74910
         TabIndex        =   77
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(日):"
         Height          =   180
         Left            =   -74910
         TabIndex        =   76
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱(英):"
         Height          =   180
         Left            =   -74910
         TabIndex        =   75
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(英):"
         Height          =   180
         Left            =   -74910
         TabIndex        =   74
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱(中):"
         Height          =   180
         Left            =   90
         TabIndex        =   73
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(中):"
         Height          =   180
         Left            =   90
         TabIndex        =   72
         Top             =   660
         Width           =   1065
      End
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   270
      Left            =   6675
      MaxLength       =   8
      TabIndex        =   16
      Top             =   5340
      Width           =   975
   End
   Begin VB.TextBox txtCP120 
      Height          =   270
      Left            =   1485
      MaxLength       =   1
      TabIndex        =   17
      Top             =   5670
      Width           =   255
   End
   Begin VB.TextBox txtCP84 
      Height          =   285
      Left            =   6615
      TabIndex        =   7
      Top             =   2430
      Width           =   1092
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Index           =   4
      Left            =   4605
      MaxLength       =   9
      TabIndex        =   4
      Top             =   2160
      Width           =   1092
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Index           =   3
      Left            =   4605
      MaxLength       =   9
      TabIndex        =   3
      Top             =   1890
      Width           =   1092
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Index           =   2
      Left            =   960
      MaxLength       =   9
      TabIndex        =   2
      Top             =   2160
      Width           =   1092
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Index           =   1
      Left            =   960
      MaxLength       =   9
      TabIndex        =   1
      Top             =   1890
      Width           =   1092
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Index           =   0
      Left            =   960
      MaxLength       =   9
      TabIndex        =   0
      Top             =   1620
      Width           =   1092
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   960
      TabIndex        =   15
      Top             =   5340
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   5256
      MaxLength       =   1
      TabIndex        =   19
      Top             =   5670
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   32
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   31
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   30
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   29
      Top             =   720
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm040104_9.frx":0070
      Left            =   960
      List            =   "frm040104_9.frx":007D
      Style           =   2  '單純下拉式
      TabIndex        =   25
      Top             =   1230
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Index           =   3
      Left            =   4032
      TabIndex        =   20
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5256
      TabIndex        =   22
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6084
      TabIndex        =   24
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label lblCP118 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:         (Y:是)"
      Height          =   180
      Left            =   5832
      TabIndex        =   102
      Top             =   1668
      Width           =   1992
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   756
      Left            =   6492
      TabIndex        =   21
      Top             =   5616
      Width           =   1500
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;1333"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   300
      Index           =   12
      Left            =   960
      TabIndex        =   5
      Top             =   2445
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   375
      Index           =   11
      Left            =   960
      TabIndex        =   23
      Top             =   5970
      Width           =   5535
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "9763;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   300
      Index           =   4
      Left            =   3975
      TabIndex        =   6
      Top             =   2445
      Width           =   1515
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "2672;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   300
      Index           =   0
      Left            =   4590
      TabIndex        =   14
      Top             =   1620
      Width           =   1110
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1958;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   18
      Left            =   3090
      TabIndex        =   100
      Top             =   5715
      Width           =   765
   End
   Begin VB.Label lblCaseFee 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7665
      TabIndex        =   69
      Tag             =   "Y"
      Top             =   5310
      Width           =   255
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   5715
      TabIndex        =   68
      Top             =   5385
      Width           =   765
   End
   Begin VB.Label lblCP120 
      AutoSize        =   -1  'True
      Caption         =   "說明書要電子檔:          (Y:是)"
      Height          =   180
      Left            =   90
      TabIndex        =   67
      Top             =   5715
      Width           =   2220
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人:"
      Height          =   180
      Left            =   5532
      TabIndex        =   66
      Top             =   5712
      Width           =   948
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   5715
      TabIndex        =   65
      Top             =   2490
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   15
      Left            =   5805
      TabIndex        =   64
      Top             =   2205
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3757;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Index           =   4
      Left            =   3765
      TabIndex        =   63
      Top             =   2205
      Width           =   675
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   13
      Left            =   5805
      TabIndex        =   62
      Top             =   1935
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3757;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Index           =   3
      Left            =   3765
      TabIndex        =   61
      Top             =   1935
      Width           =   675
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   4
      Left            =   2160
      TabIndex        =   60
      Top             =   2205
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2805;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   59
      Top             =   2205
      Width           =   675
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   2160
      TabIndex        =   58
      Top             =   1935
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2805;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   57
      Top             =   1935
      Width           =   675
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   7200
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   7200
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   56
      Top             =   5400
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   14
      Left            =   2475
      TabIndex        =   55
      Top             =   5400
      Width           =   3180
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5609;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "預估結果:           (1.准 2.駁)"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   54
      Top             =   2490
      Width           =   2655
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   4020
      TabIndex        =   53
      Top             =   510
      Width           =   1185
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2090;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家"
      Height          =   180
      Index           =   1
      Left            =   3180
      TabIndex        =   52
      Top             =   510
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   120
      TabIndex        =   51
      Top             =   510
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   50
      Top             =   765
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   120
      TabIndex        =   49
      Top             =   1020
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   5340
      TabIndex        =   48
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   3180
      TabIndex        =   47
      Top             =   1020
      Width           =   765
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   5340
      TabIndex        =   46
      Top             =   1020
      Width           =   585
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   45
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   1665
      Width           =   675
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   0
      Left            =   5340
      TabIndex        =   43
      Top             =   765
      Width           =   765
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "業務區:"
      Height          =   180
      Left            =   3180
      TabIndex        =   42
      Top             =   765
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   960
      TabIndex        =   41
      Top             =   510
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3387;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   6180
      TabIndex        =   40
      Top             =   1020
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2302;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   960
      TabIndex        =   39
      Top             =   1020
      Width           =   1920
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3387;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   5
      Left            =   6180
      TabIndex        =   38
      Top             =   510
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2328;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   6
      Left            =   4020
      TabIndex        =   37
      Top             =   1020
      Width           =   1185
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2090;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   7
      Left            =   1680
      TabIndex        =   36
      Top             =   1290
      Width           =   5820
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "10266;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   8
      Left            =   2160
      TabIndex        =   35
      Top             =   1665
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2805;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   6180
      TabIndex        =   34
      Top             =   765
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2302;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   4020
      TabIndex        =   33
      Top             =   765
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2117;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   5970
      Width           =   765
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "對造號數:"
      Height          =   180
      Left            =   3090
      TabIndex        =   27
      Top             =   2490
      Width           =   765
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Index           =   0
      Left            =   3780
      TabIndex        =   26
      Top             =   1650
      Width           =   585
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   7710
      TabIndex        =   70
      Top             =   5355
      Width           =   255
   End
End
Attribute VB_Name = "frm040104_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Text5,lstNameAgent,Label2)
'Memo By Morgan 2012/12/13 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'整理 by Morgan 2008/2/22
Option Explicit

'Modify by Morgan 2005/7/15 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String
Dim cp() As String

Dim intWhere As Integer
Dim m_strCust1 As String '申請人1
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/23 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim oChk As CheckBox 'Added by Morgan 2012/10/5
Dim m_bolFMP As Boolean 'Added by Lydia 2023/06/20 是否為FMP案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/20 是否為寰華

Private Function Process(Index As Integer) As Boolean
   Dim strTmp As String
   '檢查輸入資料的完整性
   If CheckDataIntegrity = False Then Exit Function
   'Add By Cheng 2002/05/22
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Function
   
   'Add by Morgan 2009/3/23 設定是否算發文室案件
   If pa(9) = 台灣國家代號 Then
      'Add by Morgan 2009/4/28
      If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text5(0)) = False Then
         Exit Function
      End If
      If m_CP123s = "Y" Then
      'end 2009/4/28
         'modify by sonia 2014/6/23 加傳發文規費, P-108903
         If ModifyDispatch(cp(9), m_CP09s, m_CP123s, txtCP84, Text5(0)) = False Then
             Exit Function
         End If
      End If
      'Add by Amy 2014/10/14 P台灣案發文控制
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        If pa(1) = "P" And cp(9) < "C" Then
            If cp(9) < "B" Then
                'A類一定要有接洽單才可發文
                'Modify by Amy 2014/11/27 取消ChkOneDayHasCP27判斷,接洽單改檢查,因考慮可能同時發文其他案件性質情形
                'If PUB_CheckPDF2(cp(9), 0, True, strExc(0)) = False And ChkOneDayHasCP27(pa(1), pa(2), pa(3), pa(4), cp(5) + 19110000) = False Then
                If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
                    Exit Function
                End If
            End If
            'AB類申請書確認檢查,符合條件才可發文
            'Modified by Morgan 2015/3/17
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" Then
               If PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            'end 2015/3/17
                  MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
                  Exit Function
               End If 'Added by Morgan 2015/3/17
            End If
        End If
      End If
      'end 2014/10/14

   Else
   
      'Added by Morgan 2016/6/29 非臺灣案電子化
      If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
         If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
             If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
                 Exit Function
             End If
         End If
      End If
      'end 2016/6/29
   
      'Add by Morgan 2009/11/11
      If PUB_ChkFileNP(cp(9), "'997','998'") Then
         MsgBox "下一程序已有一般提申或收達期限，若為重新發文時需要先刪除後才可作業！"
         Exit Function
      End If
   End If
   
   If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Function
   Process = True
   
   'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
   PUB_CheckEMail Combo2
   PUB_CheckEMail pa(75), pa(144)
   If pa(145) <> "" Then
      PUB_CheckEMail pa(75), pa(145)
   End If
   'end 2008/2/20
   
   'Add by Morgan 2007/6/14
   If pa(9) = "000" Then
      PUB_ReAsignInform pa(1), pa(2), pa(3), pa(4), cp(9)
   End If
   
   '2012/7/23 add by sonia
   '台灣案發文規費與收文規費不符時,mail給智權人員
   'Modified by Morgan 2013/1/14 舉發案規費鎖定但也要發Mail
   'If txtCP84.Enabled = True And pa(9) = "000" And Val(Me.txtCP84.Text) <> Val(cp(17)) Then
   If pa(9) = "000" And Val(Me.txtCP84.Text) <> Val(cp(17)) Then
      '2013/7/2 modify by sonia 改用共用module
      PUB_ChkOfficialFee cp(9), Me.txtCP84.Text
   End If
   '2012/7/23 end

   
   If pa(9) = 台灣國家代號 Then '通知函
      strTmp = "00"
      'Added by Morgan 2024/11/7 --玲玲
      If PUB_CheckCuNation(pa(26), pa(1), pa(2), pa(3), pa(4)) = "1" Then   '大-->台
         If cp(10) = "803" Then '舉發
            strTmp = "01"
         End If
      End If
      'end 2024/11/7
   'Added by Morgan 2022/3/18
   ElseIf pa(9) = "020" Then
      strTmp = "02"
   'end 2022/3/18
   Else
      strTmp = "01"
   End If
   EndLetter "02", cp(9), strTmp, strUserNum
   'Modify by Amy 2014/08/25 +傳strLetterRecNo
   NowPrint cp(9), "02", strTmp, False, strUserNum, 0, , , , , , , , , , , , cp(9)
   
   'Added by Lydia 2024/03/06 外專機械設計組人員異動調整程式：內專協辦工程師完成送件之後，需通知外專工程師進行請款
   'Move by Lydia 2024/03/12 改使用Outlook草稿，從FormSave移出
   'Mark by Lydia 2024/04/09 FMP案不用通知--- Phoebe
   'If m_bolFMP = True And cp(1) = "P" And Mid(cp(14), 4, 1) = "9" Then
   '   Call Pub_SetEngMail(cp(9))
   'End If
   ''end 2024/03/06
   'end 2024/04/09
   
   'Add By Cheng 2002/04/30
   '若有未發文資料顯示警告
   PUB_GetCPunIssueDatas "" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)

End Function

'Added by Morgan 2012/10/5
'Modified by Morgan 2013/1/14 增加舉發事項及規費計算
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
   SetOfficialFee
End Sub

'Added by Morgan 2013/1/14
Private Sub SetOfficialFee()
   Dim bolSet As Boolean
   For Each oChk In chkItem
      If oChk.Value = vbChecked Then
         '請求撤銷發明(新型)專利權 (每件新台幣5,000元，且每一請求項加收新台幣800元)
         If oChk.Index = 0 Then
            txtCP84 = 5000 + Val(txtItemCount) * 800
         ElseIf oChk.Index = 1 Then
            txtCP84 = 5000 + PUB_GetItemCount(txtItemList) * 800
            
         '請求撤銷自「 年  月  日」至「 年  月  日」之專利權期間延長(每件新台幣10,000元)
         ElseIf oChk.Index = 6 Then
            txtCP84 = 10000
            
         '請求撤銷全部專利權 (設計專利，每件新台幣8,000元；發明專利，每件新台幣10,000元；新型專利，每件新台幣9,000元)
         Else
            If pa(8) = "1" Then
               txtCP84 = 10000
            ElseIf pa(8) = "2" Then
               txtCP84 = 9000
            ElseIf pa(8) = "3" Then
               txtCP84 = 8000
            End If
         End If
         bolSet = True
         Exit For
      End If
   Next
   If bolSet = False Then txtCP84 = 0
End Sub

Private Sub cmdok_Click(Index As Integer)
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0
         'Modify by Morgan 2010/2/10 改呼叫函數方式以便鎖定按鍵
         cmdOK(Index).Enabled = False
         If Not Process(Index) Then
            cmdOK(Index).Enabled = True
         Else
            'Add By Sindy 2013/5/20
            If frm040104_1.bolIsEMPFlow = True Then
               Unload frm040104_1
               frm090202_4.Show
               frm090202_4.QueryData
            Else
            '2013/5/20 End
               frm040104_1.Show
               ' 90.07.11 modify by louis (回第一個畫面清除)
               frm040104_1.Clear
            End If
            Unload Me
         End If
      Case 1
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            frm040104_1.Show
         End If
         Unload Me
      Case 3
         Where1103ComeFrom Me, pa(1), pa(2), pa(3), pa(4)
   End Select
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
End Sub

Private Function FormSave() As Boolean
Dim intStep As Integer, iMax As Long, strTxt(1 To 25) As String, ii As Integer
Dim strTmp1(1 To 3) As String
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
FormSave = True
cnnConnection.BeginTrans

   'Added by Morgan 2013/6/7 自 lstNameAgent_Validate 移來,否則若觸發 Form_Activate 事件會跑 ReadPatent 導致 cp(110) 被清除
   cp(110) = ""
   If lstNameAgent.Visible = True Then
      For ii = 0 To lstNameAgent.ListCount - 1
         If lstNameAgent.Selected(ii) = True Then
            'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
            'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
            'Modified by Morgan 2021/12/15f Forms2.0 改用模組
            'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
            cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         End If
      Next
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   End If
   'end 2013/6/7
   

   intStep = 1
   '910703 Sieg 402
    '申請人1
    If Text7(0).Text <> "" Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCustomerNameAndAddress(Text7(0).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
        If ClsPDGetCustomerNameAndAddress(Text7(0).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
            '修改申請人時
            If InStr(ChangeCustomerL(pa(26)), ChangeCustomerL(Text7(0).Text)) = 0 Then
                If cp(60) <> "" Then
                    strExc(1) = pa(1)
                    strExc(2) = pa(2)
                    strExc(3) = pa(3)
                    strExc(4) = pa(4)
                    strExc(5) = cp(60)
                    strExc(6) = Text7(0)
                    strExc(7) = strExc(0)
                    'edit by nickc 2007/02/05 不用 dll 了
                    'If Not objLawDll.UpdAcc0k0(strExc(), True, True) Then
                    If Not ClsLawUpdAcc0k0(strExc(), True, True) Then
                        Text7(0).SetFocus
                        Text7_GotFocus 0
                        GoTo ErrorHandler
                    End If
                 End If
                strTxt(intStep) = "UPDATE PATENT SET PA26=" & CNULL(ChangeCustomerL(Text7(0))) & _
                                        ",PA31=" & CNULL(strTmp1(1)) & ",PA36=" & CNULL(strTmp1(2)) & _
                                        ",PA41=" & CNULL(strTmp1(3)) & ",PA79=NULL,PA80=NULL,PA81=NULL" & _
                                        ",PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                cnnConnection.Execute strTxt(intStep)
                intStep = intStep + 1
            End If
        End If
    Else
        strTxt(intStep) = "UPDATE PATENT SET PA26=NULL,PA31=NULL,PA36=NULL,PA41=NULL," & _
                                "PA79=NULL,PA80=NULL,PA81=NULL,PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        cnnConnection.Execute strTxt(intStep)
        intStep = intStep + 1
    End If
    'Add By Cheng 2003/08/18
    '申請人2
    If Text7(1).Text <> "" Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCustomerNameAndAddress(Text7(1).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
        If ClsPDGetCustomerNameAndAddress(Text7(1).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
            '修改申請人時
            If InStr(ChangeCustomerL(pa(27)), ChangeCustomerL(Me.Text7(1).Text)) = 0 Then
                strTxt(intStep) = "UPDATE PATENT SET PA27=" & CNULL(ChangeCustomerL(Text7(1))) & _
                                    ",PA32=" & CNULL(strTmp1(1)) & ",PA37=" & CNULL(strTmp1(2)) & _
                                    ",PA42=" & CNULL(strTmp1(3)) & ",PA79=NULL,PA80=NULL,PA81=NULL" & _
                                    ",PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                cnnConnection.Execute strTxt(intStep)
                intStep = intStep + 1
            End If
        End If
    Else
        strTxt(intStep) = "UPDATE PATENT SET PA27=NULL,PA32=NULL,PA37=NULL,PA42=NULL," & _
                                "PA79=NULL,PA80=NULL,PA81=NULL,PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(intStep)
        intStep = intStep + 1
    End If
    '申請人3
    If Text7(2).Text <> "" Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCustomerNameAndAddress(Text7(2).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
        If ClsPDGetCustomerNameAndAddress(Text7(2).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
            '修改申請人時
            If InStr(ChangeCustomerL(pa(28)), ChangeCustomerL(Me.Text7(2).Text)) = 0 Then
                strTxt(intStep) = "UPDATE PATENT SET PA28=" & CNULL(ChangeCustomerL(Text7(2))) & _
                                    ",PA33=" & CNULL(strTmp1(1)) & ",PA38=" & CNULL(strTmp1(2)) & _
                                    ",PA43=" & CNULL(strTmp1(3)) & ",PA79=NULL,PA80=NULL,PA81=NULL" & _
                                    ",PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                cnnConnection.Execute strTxt(intStep)
                intStep = intStep + 1
            End If
        End If
    Else
        strTxt(intStep) = "UPDATE PATENT SET PA28=NULL,PA33=NULL,PA38=NULL,PA43=NULL," & _
                                "PA79=NULL,PA80=NULL,PA81=NULL,PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(intStep)
        intStep = intStep + 1
    End If
    '申請人4
    If Text7(3).Text <> "" Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCustomerNameAndAddress(Text7(3).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
        If ClsPDGetCustomerNameAndAddress(Text7(3).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
            '修改申請人時
            If InStr(ChangeCustomerL(pa(29)), ChangeCustomerL(Me.Text7(3).Text)) = 0 Then
                strTxt(intStep) = "UPDATE PATENT SET PA29=" & CNULL(ChangeCustomerL(Text7(3))) & _
                                    ",PA34=" & CNULL(strTmp1(1)) & ",PA39=" & CNULL(strTmp1(2)) & _
                                    ",PA44=" & CNULL(strTmp1(3)) & ",PA79=NULL,PA80=NULL,PA81=NULL" & _
                                    ",PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                cnnConnection.Execute strTxt(intStep)
                intStep = intStep + 1
            End If
        End If
    Else
        strTxt(intStep) = "UPDATE PATENT SET PA29=NULL,PA34=NULL,PA39=NULL,PA44=NULL," & _
                                "PA79=NULL,PA80=NULL,PA81=NULL,PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(intStep)
        intStep = intStep + 1
    End If
    '申請人5
    If Text7(4).Text <> "" Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCustomerNameAndAddress(Text7(4).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
        If ClsPDGetCustomerNameAndAddress(Text7(4).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
            '修改申請人時
            If InStr(ChangeCustomerL(pa(30)), ChangeCustomerL(Me.Text7(4).Text)) = 0 Then
                strTxt(intStep) = "UPDATE PATENT SET PA30=" & CNULL(ChangeCustomerL(Text7(4))) & _
                                    ",PA35=" & CNULL(strTmp1(1)) & ",PA40=" & CNULL(strTmp1(2)) & _
                                    ",PA45=" & CNULL(strTmp1(3)) & ",PA79=NULL,PA80=NULL,PA81=NULL" & _
                                    ",PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                cnnConnection.Execute strTxt(intStep)
                intStep = intStep + 1
            End If
        End If
    Else
        strTxt(intStep) = "UPDATE PATENT SET PA30=NULL,PA35=NULL,PA40=NULL,PA45=NULL," & _
                                "PA79=NULL,PA80=NULL,PA81=NULL,PA82=NULL,PA83=NULL,PA84=NULL WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(intStep)
        intStep = intStep + 1
    End If
   
   If Combo2 <> "" Then
      'Modify by Morgan 2008/2/22
      'cp(44) = ChangeCustomerL(Combo2)
      intI = InStr(Combo2, "-")
      If intI > 0 Then
         cp(44) = Left(Combo2, intI - 1)
         cp(116) = Mid(Combo2, intI + 1)
      Else
         cp(44) = Combo2
         cp(116) = ""
      End If
      cp(44) = ChangeCustomerL(cp(44))
      'end 2008/2/22
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.GetCaseThatCode(pA) Then pA(45) = ""
      If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   Else
      cp(44) = ""
      cp(116) = ""
      cp(45) = ""
   End If
   
   'Added by Morgan 2012/10/5
   'Modified by Morgan 2013/1/14 增加舉發事項及規費自動計算
   If SSTab1.TabVisible(3) = True Then
      For Each oChk In chkItem
         If oChk.Value = vbChecked Then
            If oChk.Index = 0 Then
               Text5(11).Text = oChk.Caption & txtItemCount & "項;" & Text5(11)
            ElseIf oChk.Index = 1 Then
               Text5(11).Text = oChk.Caption & txtItemList & ";" & Text5(11)
            ElseIf oChk.Index = 6 Then
               Text5(11).Text = "請求撤銷自「" & txtYear(0) & "年" & txtMonth(0) & "月" & txtDay(0) & "日」至「" & txtYear(1) & "年" & txtMonth(1) & "月" & txtDay(1) & "日」之專利權期間延長;" & Text5(11)
            Else
               Text5(11).Text = oChk.Caption & ";" & Text5(11)
            End If
         End If
      Next
   End If
   'end 2012/10/5

   'Modify by morgan 2004/8/11 加 cp84
   'Modify by Morgan 2005/7/15 加 cp110
   'Modify by Morgan 2008/11/10 +cp120
   'Modify by Morgan 2010/2/9 +CP38,CP39,CP41,CP42
   'Modified by Lydia 2021/05/25 +CP113工作時數
   'Modified by Lydia 2023/06/20 +CP14
   'Modified by Morgan 2024/1/22 +CP118
   strTxt(intStep) = "UPDATE CASEPROGRESS SET CP22=" & CNULL(Text10) & ",CP23=" & CNULL(Text5(12)) & _
      ",cp27=" & TransDate(Text5(0), 2) & ",CP14=" & CNULL(cp(14)) & _
      ",CP44=" & CNULL(cp(44)) & ",CP116=" & CNULL(pa(116)) & ",CP45=" & CNULL(pa(45)) & _
      ",cp36=" & CNULL(ChgSQL(Text5(4))) & ",cp37=" & CNULL(ChgSQL(Text5(5))) & _
      ",cp38=" & CNULL(ChgSQL(Text5(6))) & ",cp39=" & CNULL(ChgSQL(Text5(7))) & _
      ",cp40=" & CNULL(ChgSQL(Text5(8))) & ",cp41=" & CNULL(ChgSQL(Text5(9))) & _
      ",cp42=" & CNULL(ChgSQL(Text5(10))) & ",cp64=" & CNULL(ChgSQL(Text5(11))) & _
      ",cp84=" & Format(Val(txtCP84.Text)) & ",cp110=" & CNULL(cp(110)) & _
      ",cp120=" & CNULL(txtCP120) & " ,cp113=" & CNULL(txtCP113, True) & _
      ",cp118='" & txtCP118 & "'" & _
      " where cp09='" & cp(9) & "'"
      
    'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(intStep)
    
   intStep = intStep + 1
   iMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了 iMax = objPublicData.GetNextProgressNo
   
'Removed by Morgan 2014/12/2 取消(原異議用)--陳玲玲
'   If Text5(1) <> "" Then
'      '重抓智權人員
'      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07," & _
'         "NP08,NP09,NP10,NP22,NP15) VALUES ('" & cp(9) & "','" & pa(1) & "','" & _
'         pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & Text5(1) & "'," & _
'         TransDate(Text5(2), 2) & "," & TransDate(Text5(3), 2) & "," & CNULL(PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))) & "," & iMax & "," & CNULL(Me.Text5(11).Text) & ")"
'        cnnConnection.Execute strTxt(intStep)
'       iMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了 iMax = objPublicData.GetNextProgressNo
'      intStep = intStep + 1
'   End If
'end 2014/12/2
   
   strExc(0) = "SELECT CF23 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & cp(10) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(0)) Or RsTemp.Fields(0) <> 0 Then
            '若本所期限非工作天則抓最近的工作天
         strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP22) VALUES ('" & cp(9) & "','" & pa(1) & "','" & pa(2) & _
            "','" & pa(3) & "','" & pa(4) & "'," & 收達 & "," & _
            PUB_GetWorkDay1(CompDate(2, RsTemp.Fields(0), TransDate(Text5(0), 2)), True) & "," & _
            CompDate(2, RsTemp.Fields(0), TransDate(Text5(0), 2)) & ",'" & _
            strUserNum & "'," & iMax & ")"
        cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
      End If
   End If
   
   '將"對造案件中英文名稱"更新至專利基本檔的"案件中英日名稱"
   'Modify by Morgan 2010/2/9 +PA06,PA07
   strTxt(intStep) = "Update Patent Set PA05='" & ChgSQL(Text5(5)) & "',PA06='" & ChgSQL(Text5(6)) & "'" & _
      ",PA07='" & ChgSQL(Text5(7)) & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   '案件性質為異議(801), 舉發(803) ,若有輸入對照號數則更新申請案號
   If cp(10) = "801" Or cp(10) = "803" Then
       If Me.Text5(4).Text <> "" Then
           strTxt(intStep) = "Update Patent Set PA11='" & Me.Text5(4).Text & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
           cnnConnection.Execute strTxt(intStep)
           intStep = intStep + 1
       End If
   End If
   
   'Add by Morgan 2009/3/23
   If pa(9) = 台灣國家代號 Then
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130
      'Modify by Amy 2014/09/09 for 台灣案電子化
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
         If cp(9) < "C" Then
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , , pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10))
            'Modify by Amy 2015/02/13 此固定出客戶通知函,故不考慮未出客戶函 修改判斷條件
            'PUB_AddLetterProgress cp(9), 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
              '1.　電子送件有規費的有收據；無規費的無回執
              '2.非電子送件要計件的有回執；不計件的無回執
             'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'             If cp(118) = "Y" Then
'                If Val(txtCP84) > 0 Then
'                    PUB_AddLetterProgress cp(9), 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
'                Else
'                    PUB_AddLetterProgress cp(9), 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
'                End If
'             Else
                If Left(m_CP123s, 1) = "Y" Then
                    PUB_AddLetterProgress cp(9), 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
                Else
                    PUB_AddLetterProgress cp(9), 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
                End If
'             End If
'             'end 2015/02/13
              'end 2015/03/06
         End If
      End If
      'end 2014/09/09
      'Add by Amy 2015/02/13 更新收據/回執設定
      'Modify by Amy 2015/03/06 +發文日參數
      PUB_UpdateLP19 cp(1), cp(2), cp(3), cp(4), m_CP09s, m_CP123s, Text5(0)
   
   'Added by Morgan 2016/5/26 非臺灣案電子化
   ElseIf Left(Pub_StrUserSt03, 1) <> "F" Then
      '客戶通知函
      If 內專全面電子化啟用日 <= Val(strSrvDate(1)) Then
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , pa(9), pa(1), pa(2), pa(3), pa(4))
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10), pa(9), , , IIf(Left(cp(12), 1) = "F", True, False))
         PUB_AddLetterProgress cp(9), 0, True, strExc(1), False, pa(26), cp(10), pa(75)
      End If
   'end 2016/5/26
   End If

   'Add by Morgan 2009/8/17
   If txtChkRltDate <> "" Then
      PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43)
   End If
   
   'Add by Morgan 2009/11/11 +收達期限管控
   If pa(9) <> "000" Then PUB_SetArriveDate cp(9)
   'end 2009/11/11
   
   cnnConnection.CommitTrans
   Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    FormSave = False
End Function

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(7) = pa(5)
      Case "英"
         Label2(7) = pa(6)
      Case "日"
         Label2(7) = pa(7)
   End Select
End Sub

Private Sub Combo2_Click()
   Combo2_Validate False
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo2.Text = "" Then
      If pa(9) <> 台灣國家代號 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
      
   ElseIf Not ChgType(12) Then
      Cancel = True
      
   Else
      strNo = Combo2.Text
      
      'Add by Morgan 2008/2/22 加聯絡人判斷
      iPos = InStr(strNo, "-")
      If iPos > 0 Then
         strNo = Left(strNo, iPos - 1)
      End If
      'end 2008/2/22
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If PUB_CheckStatus(strNo) = False Then Cancel = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   'Modified by Lydia 2023/06/20
   'ReDim cp(1 To TF_CP)
   ReDim cp(TF_CP)
   With frm040104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      cp(9) = .Tag
   End With
   
   'Add by Morgan 2005/7/15
   ReDim pa(1 To TF_PA)
   ReadPatent
   
   SSTab1.Tab = 0 'Added by Lydia 2021/05/25
   
   cp(110) = "" '要清空,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致 Added by Morgan 2012/9/7
   
   'Add by Morgan 2005/7/15
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text10.Visible = False
   lstNameAgent.Clear
   If pa(9) = "000" Then
      'Modified by Morgan 2021/12/15 +傳入bForm2=True
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10), True 'Modified by Morgan 2013/2/19 +傳cp(10)
      'end 2021/12/15
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/14 END
   
   Label2(0) = cp(9)
   Combo1.ListIndex = 0
   
   'Add by Morgan 2008/11/10 預設說明書是否要電子檔
   If pa(9) = "000" Then 'Added by Morgan 2024/1/24 因增加大陸案也會設電子送件
      txtCP120 = PUB_GetDefautCP120(cp(10), pa(9))
      
   'Added by Morgan 2024/1/24
   Else
      txtCP120 = ""
      txtCP120.Enabled = False
   End If

   'Add by Morgan 2009/11/11
   If pa(9) <> "000" Then
      If PUB_ChkFileNP(cp(9), "'997','998'") Then MsgBox "下一程序已有一般提申或收達期限，若為重新發文時需要先刪除後才可作業！"
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2024/03/06
   
   'Set frm040104_9 = Nothing 'Removed by Morgan 2021/12/15 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As TextBox, i As Integer
Dim strTempName As String '客戶名稱
Dim m_Fee As String         '銷帳服務費 2012/8/1 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/1 add by sonia

   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   cp(1) = pa(1)
   cp(2) = pa(2)
   cp(3) = pa(3)
   cp(4) = pa(4)
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      If pa(9) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strExc(0)) Then Label2(12) = strExc(0)
         If ClsPDGetNation(pa(9), strExc(0)) Then Label2(12) = strExc(0)
      End If
      Label2(3) = pa(22)
      ChgType (999) 'pa(8)
      Label2(7) = pa(5)
      If pa(26) <> "" Then
         Text7(0) = pa(26)
'         ChgType (11) ' Label2(8)
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GETCUSTOMER(Text7(0).Text, strTempName) Then
        If ClsPDGetCustomer(Text7(0).Text, strTempName) Then
            Label2(8) = strTempName
        Else
            Label2(8) = ""
        End If
      End If
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.Text7(0).Text
        'Add By Cheng 2003/08/18
        If pa(27) <> "" Then
            Text7(1) = pa(27)
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GETCUSTOMER(Text7(1).Text, strTempName) Then
            If ClsPDGetCustomer(Text7(1).Text, strTempName) Then
                Label2(2) = strTempName
            Else
                Label2(2) = ""
            End If
        End If
        If pa(28) <> "" Then
            Text7(2) = pa(28)
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GETCUSTOMER(Text7(2).Text, strTempName) Then
            If ClsPDGetCustomer(Text7(2).Text, strTempName) Then
                Label2(4) = strTempName
            Else
                Label2(4) = ""
            End If
        End If
        If pa(29) <> "" Then
            Text7(3) = pa(29)
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GETCUSTOMER(Text7(3).Text, strTempName) Then
            If ClsPDGetCustomer(Text7(3).Text, strTempName) Then
                Label2(13) = strTempName
            Else
                Label2(13) = ""
            End If
        End If
        If pa(30) <> "" Then
            Text7(4) = pa(30)
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GETCUSTOMER(Text7(4).Text, strTempName) Then
            If ClsPDGetCustomer(Text7(4).Text, strTempName) Then
                Label2(15) = strTempName
            Else
                Label2(15) = ""
            End If
        End If
   End If
   
   
   'Mark by Lydia 2023/06/20 改成在下方使用模組
'   If pa(9) = 台灣國家代號 Then
'      strExc(1) = "CPM03,"
'   Else
'      strExc(1) = "CPM04,"
'   End If
'   'Modify by Morgan 2004/8/11 加 cp17
'   '2012/8/1 MODIFY BY SONIA 加 cp77
'   'Modify by Amy 2014/10/14 +CP14,CP05
'   'Modified by Lydia 2021/05/25 +CP113
'   strExc(0) = "select " & strExc(1) & " staff.st02 as st1,a0902,staff1.st02 as st2," & _
'      "cp27,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp64,CP13,CP22,CP23,CP10,CP60, CP18, cp17,CP110,CP77,CP14,CP05,CP113" & _
'      " from caseprogress,casepropertymap,staff,staff staff1,acc090 " & _
'      " where cp09='" & cp(9) & "' and " & _
'      "cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
'      "cp13=staff1.st01(+) and cp12=a0901(+)"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'   With RsTemp
'      If intI = 1 Then
'         cp(110) = "" & .Fields("CP110")
'         'Add by Morgan 2004/8/11
'         cp(17) = "" & .Fields("cp17")
'         cp(5) = "" & .Fields("cp05") 'Add by Amy 2014/10/14
'         cp(14) = "" & .Fields("cp14") 'Add by Amy 2014/10/14
'
'         '2012/8/1 add by sonia 若有銷帳則要扣除銷帳規費
'         If Val("" & .Fields("cp77")) > 0 Then
'            If GetCP77Detail(cp(9), m_Fee, m_Official) = True Then
'               cp(17) = cp(17) - m_Official
'            End If
'         End If
'         '2012/8/1 end
'
'         If Not IsNull(.Fields(0)) Then Label2(5) = .Fields(0)
'         If Not IsNull(.Fields(1)) Then Label2(1) = .Fields(1)
'         If Not IsNull(.Fields(2)) Then Label2(10) = .Fields(2)
'         If Not IsNull(.Fields(3)) Then Label2(9) = .Fields(3)
'         If Not IsNull(.Fields(4)) Then
'            Text5(0) = TransDate(.Fields(4), 1)
'         Else
'            Text5(0) = strSrvDate(2)
'         End If
'
'         Text5(4).Text = "" & .Fields("cp36")
'
'         'Modify by Moragn 2005/8/31 預設本所案件名稱
'         'Text5(5).Text = "" & .Fields("cp37")
'         'Modify by Morgan 2010/2/9 +英日文欄位改不預設
'         'If "" & .Fields("cp37") = "" Then
'         '   Text5(5).Text = "" & .Fields("cp37")
'         'Else
'         '   Text5(5).Text = pa(5)
'         'End If
'         Text5(5).Text = "" & .Fields("cp37")
'         Text5(6).Text = "" & .Fields("cp38")
'         Text5(7).Text = "" & .Fields("cp39")
'         Text5(9).Text = "" & .Fields("cp41")
'         Text5(10).Text = "" & .Fields("cp42")
'         'end 2010/2/9
'         Text5(8).Text = "" & .Fields("cp40")
'         Text5(11).Text = "" & .Fields("cp64") 'Added by Morgan 2013/1/14 原來漏了
'         cp(13) = "" & .Fields("cp13") 'Added by Morgan 2013/1/14 原來漏了
'
'         If Not IsNull(.Fields(14)) Then Text10 = .Fields(14)
'         If Not IsNull(.Fields(15)) Then Text5(12) = .Fields(15)
'         If Not IsNull(.Fields(16)) Then cp(10) = .Fields(16)
'         If Not IsNull(.Fields("CP60")) Then
'            cp(60) = .Fields("CP60")
'         Else
'            cp(60) = ""
'         End If
'         'Modified by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設 => cp(9), pa(9), pa(26)
'         'AddAgent Combo2, pa
'         AddAgent Combo2, pa, , , , cp(9), pa(9), pa(26)
'
'        'Added by Lydia 2021/05/25
'        cp(113) = "": txtCP113 = ""
'        cp(113) = "" & .Fields("cp113")
'        txtCP113 = cp(113)
'        'end 2021/05/25
'      End If
'   End With
   'Added by Lydia 2023/06/20 改成共用模組
   If PUB_ReadCaseProgressDatabase(cp(), intWhere, False) Then
      '判斷FCP案,寰華案
      If Left(cp(12), 1) = "F" And pa(9) <> "000" Then
         m_bolFMP = True
      Else
         m_bolFMP = False
      End If
      m_bolFMP2 = False
      If m_bolFMP = True Then
         m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
         '寰華案:承辦人為外專程序時,改為操作人員
         If m_bolFMP2 = True Then
            cp(14) = GetFCPUser(cp(14))
         End If
      End If
      If cp(10) <> "" Then
         If ClsPDGetCaseProperty(cp(1), cp(10), strExc(0), IIf(pa(9) = 台灣國家代號, False, True)) Then
            Label2(5) = strExc(0)
         End If
      End If
      If cp(14) <> "" Then
         If ClsPDGetStaff(cp(14), strExc(0)) Then Label2(1) = strExc(0)
      End If
      If cp(12) <> "" Then
         If ClsPDGetStaffDeptName(cp(12), strExc(0)) Then Label2(10) = strExc(0)
      End If
      If cp(13) <> "" Then
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(9) = strExc(0)
      End If
      If cp(27) = "" Then
         Text5(0) = strSrvDate(2)
      Else
         Text5(0) = TransDate(cp(27), 1)
      End If
      '2012/8/1 add by sonia 若有銷帳則要扣除銷帳規費
      If Val(cp(77)) > 0 Then
         If GetCP77Detail(cp(9), m_Fee, m_Official) = True Then
            cp(17) = cp(17) - m_Official
         End If
      End If
      '2012/8/1 end
      Text5(4).Text = "" & cp(36) '對造號數
      Text5(5).Text = "" & cp(37)
      Text5(6).Text = "" & cp(38)
      Text5(7).Text = "" & cp(39)
      Text5(9).Text = "" & cp(41)
      Text5(10).Text = "" & cp(42)
      Text5(8).Text = "" & cp(40)
      Text5(11).Text = "" & cp(64)
      Text10 = "" & cp(22)  '是否出名
      Text5(12) = "" & cp(23) '預估結果
      'Modified by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設 => cp(9), pa(9), pa(26)
      AddAgent Combo2, pa, , , , cp(9), pa(9), pa(26)
      txtCP113 = cp(113) '工作時數
   End If
   'end 2023/06/20
   
   'Add by Morgan 2004/9/8 檢查是否有延期，若有則規費預設0
   If PUB_ChkDelay(cp(9)) = True Then cp(17) = "0"
   
   'Add by Morgan 2004/8/11
   txtCP84.Tag = cp(17)
   
   'Add by Morgan 2009/8/17
   If Text5(0) <> "" Then
      PUB_SetChkResultDate pa(1), pa(9), cp(10), Text5(0), txtChkRltDate, cp, pa(8)
      Text5(0).Tag = Text5(0)
   End If
   
   'Added by Morgan 2012/10/5
   'Modified by Morgan 2013/1/14 增加舉發事項及規費自動計算
   If pa(9) = 台灣國家代號 And cp(10) = "803" Then
      txtCP84.Enabled = False
      SSTab1.TabVisible(3) = True
      If pa(8) = "3" Then
         chkItem(0).Enabled = False
         chkItem(1).Enabled = False
         chkItem(2).Enabled = True
      Else
         chkItem(2).Enabled = False
      End If
   Else
      SSTab1.TabVisible(3) = False
   End If
   
   'Added by Morgan 2024/1/22 大陸案預設電子送件--郭
   txtCP118 = cp(118)
   If pa(9) = "020" Then
      lblCP118.Visible = True
      txtCP118.Visible = True
      'Modified by Morgan 2024/1/25 有設定大陸P案要公文正本者預設紙本送件--郭
      'Removed by Morgan 2024/1/30 改分案預設--郭
      'If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
      '   txtCP118 = ""
      'Else
      '   txtCP118 = "Y"
      'End If
   '台灣案維持只有紙本送件(不顯示)
   Else
      lblCP118.Visible = False
      txtCP118.Visible = False
   End If
   'end 2024/1/22
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String, bolTmp As Boolean
   ChgType = False
   Select Case i
      Case 0
         '2011/12/8 MODIFY BY SONIA 發文日可輸系統日的下一個工作日
         'If Not ChkDate(Text5(0)) Or Val(Text5(0)) > Val(strSrvDate(2)) Then
         '   MsgBox "發文日期不正確或發文日大於系統日，請重新輸入 !", vbCritical
         If Not ChkDate(Text5(0)) Or DBDATE(Val(Text5(0))) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
            MsgBox "發文日期不正確或發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
         '2011/12/8 END
         Else
            ChgType = True
'Removed by Morgan 2014/12/31 取消(原異議用)--陳玲玲
'            If Text5(0) <> "" Then
'               strExc(0) = TransDate(Text5(0).Text, 2)
'               'edit by nickc 2007/02/05 不用 dll 了
'               'If objLawDll.GetCaseFeeDelay(pa(1), pa(9), cp(10), strExc) Then
'               If ClsLawGetCaseFeeDelay(pa(1), pa(9), cp(10), strExc) Then
'                  Text5(3) = TransDate(strExc(1), 1)
'                  Text5(2) = TransDate(strExc(2), 1)
'                    'Add By Cheng 2003/12/08
'                    '本所期限若非工作天則抓最近工作天
'                    'Me.Text5(2).Text = TransDate(PUB_GetWorkDay1(Me.Text5(2).Text, True), 1)
'               End If
'            End If
'end 2014/12/31

         End If
      Case 1
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Text5(i), strTempName, BolTmp) Then
         If ClsPDGetCaseProperty(pa(1), Text5(i), strTempName, bolTmp) Then
            Label2(11) = strTempName
            ChgType = True
         Else
            Label2(11) = ""
         End If
      Case 12 '代理人
         strExc(1) = Combo2.Text
         'Add by Morgan 2008/2/22 加判斷是否為聯絡人
         If InStr(strExc(1), "-") > 0 Then
            If ClsPDGetContact(strExc(1), strTempName) Then
               Combo2 = strExc(1)
               Label2(14) = strTempName
               ChgType = True
            End If

         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
         ElseIf PUB_GetAgentName(pa(1), strExc(1), strTempName) = True Then
            Combo2.Text = strExc(1)
            Label2(14).Caption = strTempName
            ChgType = True
         Else
            Label2(14).Caption = ""
         End If
         
      Case 999
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTempName, BolTmp, pA(9)) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, pa(8), strTempName, bolTmp, pa(9)) = 1 Then
            Label2(6) = strTempName
         End If
   End Select
End Function

Private Sub Text10_GotFocus()
  TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

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

Private Sub txtItemCount_Validate(Cancel As Boolean)
   SetOfficialFee
End Sub

Private Sub txtItemList_GotFocus()
   CloseIme
End Sub

Private Sub txtItemList_Validate(Cancel As Boolean)
   SetOfficialFee
End Sub

Private Sub Text5_GotFocus(Index As Integer)
    TextInverse Text5(Index)
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    Case 12 '預估結果
        KeyAscii = UpperCase(KeyAscii)
        If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End Select
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If Text5(Index) <> "" Then
            If ChgType(0) = False Then
               Cancel = True
            Else
               'Add by Morgan 2009/8/17
               If Text5(0).Tag <> Text5(0) Then
                  PUB_SetChkResultDate pa(1), pa(9), cp(10), Text5(0), txtChkRltDate, cp, pa(8)
                  Text5(0).Tag = Text5(0)
               End If
            End If
         Else
            MsgBox "發文日不可空白 !", vbCritical
            Cancel = True
         End If
         
'Removed by Morgan 2014/12/2 取消(原異議用)--陳玲玲
'      Case 1 '下一程序
'         If Text5(Index) <> "" Then
'            'Add/Modify By Cheng 2002/01/04
'            If Len(Me.Text5(Index).Text) <> 3 Then
'               MsgBox "下一程序欄位值必須為三碼 !", vbCritical
'               Cancel = True
'            Else
'               Select Case Text5(Index)
'                  Case "202", "206"
'                     If Not ChgType(1) Then Cancel = True: TextInverse Text5(Index)
'                  Case Else
'                     MsgBox "下一程序只能為補充說明或補文件 !", vbCritical
'                     Cancel = True
'               End Select
'            End If
'         Else
'            Label2(11) = ""
'         End If
'      Case 2, 3 '下一程序本所期限, 下一程序法定期限
'         If Text5(1) = "" Then
'            If Text5(Index) <> "" Then
'               MsgBox "下一程序為空白時，此欄必須為空白 !", vbCritical
'               Cancel = True
'            End If
'         Else
'            If Text5(Index) = "" Then
'               MsgBox "下一程序為空白時，此欄不得為空白 !", vbCritical
'               Cancel = True
'            Else
'               If Not ChkDate(Text5(Index)) Then
'                  MsgBox "日期不正確，請重新輸入 !", vbCritical
'                  Cancel = True
'               Else
'                  If Index = 1 Then
'                     If ChkRange(Text5(2), Text5(3), "日期") = False Then
'                        Text5(2).SetFocus
'                        Exit Sub
'                     End If
'                  End If
'               End If
'                'Add By Cheng 2003/12/08
'                '若本所期限非工作天則直接調整至最近的工作天
'                If Index = 2 And Cancel = False Then
'                    Me.Text5(2).Text = TransDate(PUB_GetWorkDay1(Me.Text5(2).Text, True), 1)
'                End If
'                'End
'            End If
'         End If
'end 2014/12/2

      Case 4
         If Text5(Index) = "" Then
            MsgBox "對照號數不可空白 !", vbCritical
            Cancel = True
         End If
      Case 12 '預估准駁
         'Add/Modify By Cheng 2002/06/24
         If GetCF15("" & pa(1), "" & pa(9), "" & cp(10)) Then
            If Text5(12) = "" Then
               MsgBox "預估准駁不可空白 !"
               Cancel = True
            End If
         End If
   End Select
   If Cancel = True Then TextInverse Text5(Index)
End Sub

Private Sub Text7_GotFocus(Index As Integer)
    TextInverse Text7(Index)
End Sub

Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
Dim strTempName As String '客戶名稱
    
    Cancel = False
    Select Case Index
    Case 0 '申請人1
        If Me.Text7(Index).Text = "" Then
            Me.Label2(8).Caption = ""
        Else
            If Me.Text7(Index).Text <> m_strCust1 Then
                If Not PUB_EditCustOk(Me.Label2(0).Caption, pa(1), pa(2), pa(3), pa(4)) Then
                    Cancel = True
                End If
            End If
            If Cancel = False Then
                '910703 Sieg 402
                'edit by nickc 2007/02/02 不用 dll 了
                'If objPublicData.GETCUSTOMER(Text7(Index).Text, strTempName) Then
                If ClsPDGetCustomer(Text7(Index).Text, strTempName) Then
                    Label2(8) = strTempName
                Else
                    Label2(8) = ""
                    Cancel = True
                End If
            End If
        End If
    Case 1 '申請人2
        If Me.Text7(Index).Text = "" Then
            Me.Label2(2).Caption = ""
        Else
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GETCUSTOMER(Text7(Index).Text, strTempName) Then
            If ClsPDGetCustomer(Text7(Index).Text, strTempName) Then
                Label2(2) = strTempName
            Else
                Label2(2) = ""
                Cancel = True
            End If
        End If
    Case 2 '申請人3
        If Me.Text7(Index).Text = "" Then
            Me.Label2(4).Caption = ""
        Else
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GETCUSTOMER(Text7(Index).Text, strTempName) Then
            If ClsPDGetCustomer(Text7(Index).Text, strTempName) Then
                Label2(4) = strTempName
            Else
                Label2(4) = ""
                Cancel = True
            End If
        End If
    Case 3 '申請人4
        If Me.Text7(Index).Text = "" Then
            Me.Label2(13).Caption = ""
        Else
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GETCUSTOMER(Text7(Index).Text, strTempName) Then
            If ClsPDGetCustomer(Text7(Index).Text, strTempName) Then
                Label2(13) = strTempName
            Else
                Label2(13) = ""
                Cancel = True
            End If
        End If
    Case 4 '申請人5
        If Me.Text7(Index).Text = "" Then
            Me.Label2(15).Caption = ""
        Else
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GETCUSTOMER(Text7(Index).Text, strTempName) Then
            If ClsPDGetCustomer(Text7(Index).Text, strTempName) Then
                Label2(15) = strTempName
            Else
                Label2(15) = ""
                Cancel = True
            End If
        End If
    End Select
    
   'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
   If Cancel = False Then
      If PUB_CheckStatus(Text7(Index).Text) = False Then Cancel = True
   End If
    
    If Cancel Then TextInverse Text7(Index)
   
End Sub


Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Cheng 2002/03/08
Private Function CheckDataIntegrity() As Boolean
Dim Cancel As Boolean
   'add by nickc 2008/05/01
   If IsDebt(pa(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        GoTo IntegrityOrNot
   End If
'Add By Cheng 2003/08/18
If Me.Text7(0).Text = "" Then
    MsgBox "申請人1不可空白 !", vbCritical
    Text7(0).SetFocus
    Text7_GotFocus 0
    GoTo IntegrityOrNot
End If
If Text5(8) = "" Then
    SSTab1.Tab = 0
    MsgBox "對造名稱不可空白 !", vbCritical
    Text5(8).SetFocus
    Text5_GotFocus 8
    GoTo IntegrityOrNot
End If
If Text5(5) = "" Then
   SSTab1.Tab = 0
    MsgBox "對造案件名稱不可空白 !", vbCritical
    Text5(5).SetFocus
    Text5_GotFocus 5
    GoTo IntegrityOrNot
End If
'檢查代理人欄位
Cancel = False
Combo2_Validate Cancel
If Cancel = True Then
    Me.Combo2.SetFocus
    GoTo IntegrityOrNot
End If
CheckDataIntegrity = True
Exit Function

IntegrityOrNot:
CheckDataIntegrity = False
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/15 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/15
   
For Each objTxt In Text5
   If objTxt.Enabled = True Then
      Cancel = False
      Text5_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text5(objTxt.Index).SetFocus
         Text5_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next
If Text5(8) = "" Then
   MsgBox "對造名稱不可空白 !", vbCritical
   Text5(8).SetFocus
   Exit Function
End If
If Text5(5) = "" Then
   MsgBox "對造案件名稱不可空白 !", vbCritical
   Text5(5).SetFocus
   Exit Function
End If

'Add by Morgan 2004/8/11
If txtCP84.Enabled = True Then
   Cancel = False
   txtCP84_Validate Cancel
   If Cancel = True Then
      txtCP84.SetFocus
      txtCP84_GotFocus
      Exit Function
   End If
End If

'Add by Morgan 2004/9/14
If Combo2.Enabled = True Then
   Cancel = False
   Combo2_Validate Cancel
   If Cancel = True Then
      Combo2.SetFocus
      Exit Function
   End If
End If

   'Add by Morgan 2005/7/15
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   '2005/7/14 END
   
'Added by Morgan 2012/10/5
'102新法舉發要勾選聲明
'Modified by Morgan 2013/1/14 增加舉發事項及規費計算
If SSTab1.TabVisible(3) = True And Val(strSrvDate(1)) > 20130000 Then
   Cancel = True
   For Each oChk In chkItem
      If oChk.Value = vbChecked Then
         If oChk.Index = 0 Then
            If txtItemCount = "" Then
               SSTab1.Tab = 3
               MsgBox "請輸入項數", vbExclamation, "舉發聲明"
               If txtItemCount.Enabled Then txtItemCount.SetFocus
               Exit Function
            End If
         ElseIf oChk.Index = 1 Then
            If txtItemList = "第項" Then
               SSTab1.Tab = 3
               MsgBox "請輸入項次", vbExclamation, "舉發聲明"
               If txtItemList.Enabled Then txtItemList.SetFocus
               Exit Function
            ElseIf PUB_ChkItemList(txtItemList) = False Then
               SSTab1.Tab = 3
               MsgBox "撤銷部分請求項格式錯誤！", vbExclamation, "舉發聲明"
               If txtItemList.Enabled Then txtItemList.SetFocus
               Exit Function
            End If
         ElseIf oChk.Index = 6 Then
            For intI = 0 To 1
               If txtYear(intI) = "" Then
                  SSTab1.Tab = 3
                  MsgBox "請輸入年度!", vbExclamation, "舉發聲明"
                  txtYear(intI).SetFocus
                  Exit Function
               End If
               If txtMonth(intI) = "" Then
                  SSTab1.Tab = 3
                  MsgBox "請輸入月份!", vbExclamation, "舉發聲明"
                  txtMonth(intI).SetFocus
                  Exit Function
               End If
               If txtDay(intI) = "" Then
                  SSTab1.Tab = 3
                  MsgBox "請輸入日期!", vbExclamation, "舉發聲明"
                  txtDay(intI).SetFocus
                  Exit Function
               End If
               If Not IsDate((Val(txtYear(intI)) + 1911) & "/" & txtMonth(intI) & "/" & txtDay(intI)) Then
                  SSTab1.Tab = 3
                  MsgBox "日期錯誤，請重新輸入！", vbExclamation, "舉發聲明"
                  txtYear(intI).SetFocus
                  Exit Function
               End If
            Next
            If CDate((Val(txtYear(0)) + 1911) & "/" & txtMonth(0) & "/" & txtDay(0)) > CDate((Val(txtYear(1)) + 1911) & "/" & txtMonth(1) & "/" & txtDay(1)) Then
               SSTab1.Tab = 3
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
      SSTab1.Tab = 3
      MsgBox "請選擇舉發聲明項目！", vbExclamation, "舉發聲明"
      Exit Function
   End If
   
   txtCP84_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'end 2012/10/5
   
   'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
   If Pub_ChkACS112isNull(pa(1), pa(2), pa(3), pa(4), txtCP113) = True Then
         txtCP113.SetFocus
         txtCP113_GotFocus
         Exit Function
   End If
   'end 2021/05/25
   
'Added by Morgan 2024/1/22
If pa(9) = "020" And txtCP118 = "" Then
   If MsgBox("請確認本案是否為紙本送件？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      txtCP118.SetFocus
      Exit Function
   End If
End If
'end 2024/1/22

TxtValidate = True
End Function

'Add By Cheng 2002/06/24
Private Function GetCF15(strCF01 As String, strCF02 As String, strCF03 As String) As Boolean
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF15 = False
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
StrSQLa = "Select CF15 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   If IsNull(rsA.Fields(0).Value) = False Then
      GetCF15 = False
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

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


Private Sub txtCP120_GotFocus()
   TextInverse txtCP120
End Sub

Private Sub txtCP120_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add by Morgan 2004/8/11
Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_Validate(Cancel As Boolean)
   '台灣
   If pa(9) = "000" Then
      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            If txtCP84.Enabled = True Then txtCP84_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub

'Add by Morgan 2005/7/15
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   cp(110) = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/15f Forms2.0 改用模組
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   If bolCheck = True Then
      Text10 = ""
   Else
      Text10 = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub
'Add by Morgan 2009/8/17
Private Sub lblCaseFee_Click()
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = pa(9)
   frm12040102_2.txtCF(3) = cp(10)
   frm12040102_2.Show vbModal
   If Val(Text5(0)) > 0 Then
      PUB_SetChkResultDate pa(1), pa(9), cp(10), Text5(0), txtChkRltDate, cp, pa(8)
   End If
End Sub

Private Sub lblCaseFee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblCaseFee, lblCaseFees
End Sub

Private Sub lblCaseFee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblCaseFee, lblCaseFees
End Sub

Private Sub txtChkRltDate_Validate(Cancel As Boolean)
   If txtChkRltDate <> "" Then
      If ChkDate(txtChkRltDate) = False Then
         Cancel = True
      End If
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


