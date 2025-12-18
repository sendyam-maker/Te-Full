VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_9 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-異議/舉發"
   ClientHeight    =   6456
   ClientLeft      =   -2148
   ClientTop       =   3360
   ClientWidth     =   7884
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6456
   ScaleWidth      =   7884
   Begin TabDlg.SSTab SSTab1 
      Height          =   1995
      Left            =   60
      TabIndex        =   51
      Top             =   4410
      Width           =   7710
      _ExtentX        =   13610
      _ExtentY        =   3514
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "對造(中)"
      TabPicture(0)   =   "frm060104_9.frx":0000
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
      TabPicture(1)   =   "frm060104_9.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text5(9)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text5(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label30"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label33"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "對造(日)"
      TabPicture(2)   =   "frm060104_9.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text5(10)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text5(7)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label31"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label34"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "舉發聲明"
      TabPicture(3)   =   "frm060104_9.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtItemCount"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "txtMonth(0)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtYear(0)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtYear(1)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtMonth(1)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtDay(1)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtDay(0)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "chkItem(0)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "chkItem(1)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "chkItem(4)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "chkItem(3)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "chkItem(5)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "chkItem(2)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "txtItemList"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "chkItem(6)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Label17"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Label13"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Label15"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Label16(5)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).ControlCount=   19
      Begin VB.TextBox txtItemCount 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -72435
         TabIndex        =   65
         Top             =   630
         Width           =   375
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72750
         MaxLength       =   2
         TabIndex        =   70
         Top             =   1590
         Width           =   285
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -73470
         MaxLength       =   3
         TabIndex        =   69
         Top             =   1590
         Width           =   420
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -71085
         MaxLength       =   3
         TabIndex        =   72
         Top             =   1590
         Width           =   420
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70410
         MaxLength       =   2
         TabIndex        =   73
         Top             =   1590
         Width           =   285
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -69825
         MaxLength       =   2
         TabIndex        =   74
         Top             =   1590
         Width           =   285
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72210
         MaxLength       =   2
         TabIndex        =   71
         Top             =   1590
         Width           =   285
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷全部請求項：共計"
         Height          =   210
         Index           =   0
         Left            =   -74910
         TabIndex        =   78
         Top             =   660
         Width           =   2505
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷部分之請求項："
         Height          =   210
         Index           =   1
         Left            =   -74910
         TabIndex        =   77
         Top             =   870
         Width           =   2400
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "共有專利申請權非由全體共有人提出申請者"
         Height          =   210
         Index           =   4
         Left            =   -71670
         TabIndex        =   76
         Top             =   1080
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "專利權人為非專利申請權人者"
         Height          =   210
         Index           =   3
         Left            =   -71670
         TabIndex        =   75
         Top             =   870
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "專利權人所屬國家對中華民國申請專利不予受理者"
         Height          =   210
         Index           =   5
         Left            =   -71670
         TabIndex        =   67
         Top             =   1290
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷設計專利權"
         Enabled         =   0   'False
         Height          =   210
         Index           =   2
         Left            =   -71670
         TabIndex        =   66
         Top             =   660
         Width           =   4335
      End
      Begin VB.TextBox txtItemList 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -74640
         TabIndex        =   64
         Text            =   "第項"
         Top             =   1080
         Width           =   2580
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
         Left            =   -74910
         TabIndex        =   68
         Top             =   1620
         Width           =   7440
      End
      Begin MSForms.TextBox Text5 
         Height          =   288
         Index           =   10
         Left            =   -73164
         TabIndex        =   61
         Top             =   540
         Width           =   5628
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "9927;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   288
         Index           =   7
         Left            =   -73164
         TabIndex        =   60
         Top             =   840
         Width           =   5676
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "10012;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   288
         Index           =   9
         Left            =   -73164
         TabIndex        =   57
         Top             =   528
         Width           =   5700
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "10054;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   288
         Index           =   6
         Left            =   -73164
         TabIndex        =   56
         Top             =   840
         Width           =   5724
         VariousPropertyBits=   671105051
         MaxLength       =   250
         Size            =   "10096;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   288
         Index           =   8
         Left            =   1836
         TabIndex        =   53
         Top             =   528
         Width           =   5748
         VariousPropertyBits=   671105051
         MaxLength       =   600
         Size            =   "10139;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   288
         Index           =   5
         Left            =   1836
         TabIndex        =   52
         Top             =   840
         Width           =   5748
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "10139;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "( 例如：第 1,3,5-12 項 )"
         Height          =   180
         Left            =   -74640
         TabIndex        =   82
         Top             =   1380
         Width           =   1800
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
         Left            =   -74910
         TabIndex        =   81
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "請求撤銷全部專利權"
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
         Left            =   -71670
         TabIndex        =   80
         Top             =   420
         Width           =   1755
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "項"
         Height          =   180
         Index           =   5
         Left            =   -72030
         TabIndex        =   79
         Top             =   675
         Width           =   180
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱(日):"
         Height          =   180
         Left            =   -74730
         TabIndex        =   63
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(日):"
         Height          =   180
         Left            =   -74730
         TabIndex        =   62
         Top             =   540
         Width           =   1065
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱(英):"
         Height          =   180
         Left            =   -74730
         TabIndex        =   59
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(英):"
         Height          =   180
         Left            =   -74730
         TabIndex        =   58
         Top             =   540
         Width           =   1065
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱(中):"
         Height          =   180
         Left            =   270
         TabIndex        =   55
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱(中):"
         Height          =   180
         Left            =   270
         TabIndex        =   54
         Top             =   540
         Width           =   1065
      End
   End
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   3225
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2580
      Width           =   600
   End
   Begin VB.TextBox txtCP84 
      Height          =   288
      Left            =   3615
      TabIndex        =   7
      Top             =   3180
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   23
      Top             =   900
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   22
      Top             =   900
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   21
      Top             =   900
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   20
      Top             =   900
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060104_9.frx":0070
      Left            =   1020
      List            =   "frm060104_9.frx":007D
      Style           =   2  '單純下拉式
      TabIndex        =   13
      Top             =   1830
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "補件期限(&D)"
      Height          =   400
      Index           =   3
      Left            =   4110
      TabIndex        =   10
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5340
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6165
      TabIndex        =   12
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   12
      Left            =   1560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2880
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   915
      Index           =   11
      Left            =   975
      TabIndex        =   9
      Top             =   3480
      Width           =   4560
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "8043;1614"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   4
      Left            =   975
      TabIndex        =   6
      Top             =   3180
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   3
      Left            =   6780
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2880
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   2
      Left            =   4170
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2880
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   5070
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2580
      Width           =   615
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1085;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   975
      TabIndex        =   0
      Top             =   2580
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "下一程序約定期限:"
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   83
      Top             =   2925
      Width           =   1485
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   6300
      TabIndex        =   8
      Top             =   3180
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   12
      Left            =   2370
      TabIndex        =   50
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   5370
      TabIndex        =   49
      Top             =   3195
      Width           =   900
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   2790
      TabIndex        =   48
      Top             =   3195
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   60
      X2              =   7700
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   7700
      Y1              =   2580
      Y2              =   2580
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   11
      Left            =   5730
      TabIndex        =   47
      Top             =   2580
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
      Left            =   180
      TabIndex        =   46
      Top             =   570
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   45
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發證日:"
      Height          =   180
      Index           =   0
      Left            =   3180
      TabIndex        =   44
      Top             =   900
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   180
      TabIndex        =   43
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號:"
      Height          =   180
      Left            =   3180
      TabIndex        =   42
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   180
      TabIndex        =   41
      Top             =   1530
      Width           =   765
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   3180
      TabIndex        =   40
      Top             =   1530
      Width           =   765
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   3180
      TabIndex        =   39
      Top             =   570
      Width           =   585
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   38
      Top             =   1830
      Width           =   765
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人:"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   37
      Top             =   2190
      Width           =   585
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   4620
      TabIndex        =   36
      Top             =   2190
      Width           =   765
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "業務區:"
      Height          =   180
      Left            =   3180
      TabIndex        =   35
      Top             =   2190
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1020
      TabIndex        =   34
      Top             =   570
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3492;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   4020
      TabIndex        =   33
      Top             =   570
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   4020
      TabIndex        =   32
      Top             =   900
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   1020
      TabIndex        =   31
      Top             =   1200
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   4020
      TabIndex        =   30
      Top             =   1230
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   1020
      TabIndex        =   29
      Top             =   1530
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   6
      Left            =   4020
      TabIndex        =   28
      Top             =   1530
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   27
      Top             =   1860
      Width           =   6060
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "10689;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   8
      Left            =   1020
      TabIndex        =   26
      Top             =   2190
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3598;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   9
      Left            =   5460
      TabIndex        =   25
      Top             =   2190
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2752;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   10
      Left            =   3840
      TabIndex        =   24
      Top             =   2190
      Width           =   750
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1323;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   180
      TabIndex        =   19
      Top             =   3450
      Width           =   765
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "對造號數:"
      Height          =   180
      Left            =   180
      TabIndex        =   18
      Top             =   3195
      Width           =   765
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "下一程序法定期限:"
      Height          =   180
      Left            =   5250
      TabIndex        =   17
      Top             =   2925
      Width           =   1485
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "下一程序本所期限:"
      Height          =   180
      Index           =   0
      Left            =   2640
      TabIndex        =   16
      Top             =   2925
      Width           =   1485
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "下一程序:"
      Height          =   180
      Left            =   4245
      TabIndex        =   15
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   2640
      Width           =   585
   End
End
Attribute VB_Name = "frm060104_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/17 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'2006/3/22 整理
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/8/4 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, m_CP110 As String, m_AgentName As String

Dim intWhere As Integer
Dim m_CP10 As String
'Add by Morgan 2004/8/11
Dim m_CP17 As String '收文規費
Dim bolDelay As Boolean 'Add by Morgan 2004/9/8 是否延期過
Dim strDelayCP09 As String 'Added by Morgan 2011/11/11 延期收文號
Dim m_CP14 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/20 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim oChk As CheckBox 'Added by Morgan 2012/10/5
Dim m_CP142 As String 'Add By Sindy 2015/12/17
Dim m_CP164 As String 'Add By Sindy 2021/4/20


'Added by Morgan 2012/10/29
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

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         'Add By Cheng 2002/05/21
         If CheckDataValid = False Then Exit Sub
               
         'Add by Morgan 2009/4/28
         If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text5(0)) = False Then
            Exit Sub
         End If
         If m_CP123s = "Y" Then
         'end 2009/4/28
            'Add by Morgan 2009/3/20 設定是否算發文室案件
            'modify by sonia 2014/6/23 加傳發文規費, P-108903
            If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, txtCP84, Text5(0)) = False Then
                Exit Sub
            End If
            'end 2009/3/20
         End If
         
         'Add by Sindy 2021/11/17 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         If FormSave = False Then
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
            Exit Sub
         Else
         
            'Add by Morgan 2008/2/20 檢查代理人Email
            PUB_CheckEMail pa(75), pa(144)
            If pa(145) <> "" Then
               PUB_CheckEMail pa(75), pa(145)
            End If
            'end 2008/2/20

            If pa(1) = "FCP" Then
'Modified by Morgan 2020/3/3 改呼叫共用
'               'Add By Sindy 2016/7/7 + 代理人為Y4829203Hewlett-Packard Company Intellectual Property Administration
'               '承辦人為工程師(ST03 IN ('F21','F51','F52))時,於存檔後彈訊息
'               If ChangeCustomerL(pa(75)) = "Y48292030" And _
'                  (PUB_GetST03(m_CP14) = "F21" Or PUB_GetST03(m_CP14) = "F51" Or PUB_GetST03(m_CP14) = "F52") Then
'                  'Add By Sindy 2016/7/18
'                  strExc(0) = "select cp130 from caseprogress where cp09='" & strReceiveNo & "'"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     If "" & RsTemp.Fields(0) <> "" Then
'                  '2016/7/18 END
'                        MsgBox "請優先請款並且在提申當天上傳報告!!"
'                     End If
'                  End If
'               End If
'               '2016/7/7 END
'
'               'Add By Sindy 2016/10/17 凡代理人Y33844   KLARQUIST SPARKMAN, LLP的案件，
'               '若是工程師中間程序(例: 申復、再審、訴願、補充說明、...)發文時，
'               '彈訊息"請在送件後3天內並且要當月優先請款"，請排除901.告代、902.回代、1202.審查意見、1002.核駁.....。
'               If (PUB_GetST03(m_CP14) = "F21" Or PUB_GetST03(m_CP14) = "F51" Or PUB_GetST03(m_CP14) = "F52") And _
'                  Not (m_CP10 = "901" And m_CP10 = "902" And m_CP10 = "1202" And m_CP10 = "1002") Then
'                  strExc(0) = "select cp130 from caseprogress where cp09='" & strReceiveNo & "'"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     If "" & RsTemp.Fields(0) <> "" Then
'                        If ChangeCustomerL(pa(75)) = "Y33844000" Then
'                           MsgBox "請在送件後3天內並且要當月優先請款!!"
'                        'Add By Sindy 2016/10/20
'                        ElseIf ChangeCustomerL(pa(75)) = "Y51982000" Then
'                           '中間程序須優先請款(於法限前優先請款完成)
'                           MsgBox "申復/再審/修正等送智慧局的案件，於收到指示後7天內送程序送件同時請款(由Wilson指示備註)!!"
'                        '2016/10/20 END
'                        'Add By Sindy 2016/10/24
'                        ElseIf ChangeCustomerL(pa(75)) = "Y20272000" Then
'                           MsgBox "中間程序送件當日簡單報告!!"
'                        '2016/10/24 END
'                        'Add By Sindy 2016/11/16
'                        ElseIf ChangeCustomerL(pa(75)) = "Y34440B30" Then
'                           MsgBox "請當日優先請款報告!!"
'                        '2016/11/16 END
'                        End If
'                     End If
'                  End If
'               End If
'               '2016/10/17 END
               PUB_FCPAlert strReceiveNo
'end 2020/3/3
            End If

            'Add By Sindy 2023/11/9
            If frm060104_1.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2023/11/9 End
            
            'Added by Lydia 2024/03/06 外專機械設計組人員異動調整程式：內專協辦工程師完成送件之後，需通知外專工程師進行請款
            'Move by Lydia 2024/03/12 改使用Outlook草稿，從FormSave移出
            'Mark by Lydia 2024/04/18 FCP案直接併入frm060104_k的Outlook，所以也不用---Sharon
            'If pa(1) = "FCP" And Mid(m_CP14, 4, 1) = "9" Then
            '   Call Pub_SetEngMail(strReceiveNo)
            'End If
            ''end 2024/03/06
            'end 2024/04/18
            
            'Add By Cheng 2002/04/30
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
            
            Unload Me
         End If
      Case 1
         frm060104_1.Show
         Unload Me
      Case 3
         Select Case Label2(5)
            Case "異議"
               Me.Tag = "1"
            Case "舉發"
               Me.Tag = "2"
         End Select
         frm060104_d.Show
         ' 90.06.27 modify by louis
         'Modify By Sindy 2021/8/17 + , Text5(12)
         frm060104_d.SetData Text5(2), Text5(3), Text5(12)
         Me.Hide
   End Select
End Sub

Private Function FormSave() As Boolean
Dim intStep As Integer
Dim m_NP09 As String  '2006/3/22 ADD BY SONIA
Dim lngFee As Long 'Added by Morgan 2011/11/11
 
   '911105 nick transation
   FormSave = True
    On Error GoTo CheckingErr
cnnConnection.BeginTrans

   intStep = 1
   
   'Added by Morgan 2012/10/5
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
   
   'Modify by morgan 2005/8/4 加 cp110
   'Modify by Morgan 2007/7/19 加 cp113
   'MODIFY BY SONIA 2015/9/21 加 cp14
   strExc(1) = "UPDATE CASEPROGRESS SET cp27=" & TransDate(Text5(0), 2) & ",cp14=" & CNULL(m_CP14) & "," & _
      "cp36=" & CNULL(ChgSQL(Text5(4))) & ",cp37=" & CNULL(ChgSQL(Text5(5))) & "," & _
      "cp38=" & CNULL(ChgSQL(Text5(6))) & ",cp39=" & CNULL(ChgSQL(Text5(7))) & "," & _
      "cp40=" & CNULL(ChgSQL(Text5(8))) & ",cp41=" & CNULL(ChgSQL(Text5(9))) & "," & _
      "cp42=" & CNULL(ChgSQL(Text5(10))) & ",cp64=" & CNULL(ChgSQL(Text5(11))) & _
      ",cp84=" & Format(Val(txtCP84.Text)) & IIf(bolDelay, "", ", CP16=NVL(CP16,0)-NVL(CP17,0)+" & Format(Val(txtCP84.Text)) & _
      ", CP17=" & Format(Val(txtCP84.Text))) & ", CP18=NVL(CP18,0),cp110=" & CNULL(m_CP110) & ",CP22=NULL" & _
      ", CP113=" & CNULL(txtCP113.Text, True) & _
      " where cp09='" & strReceiveNo & "'"
   cnnConnection.Execute strExc(1)
    
   'Added by Morgan 2011/11/11
   '更新收文規費為延期發文規費(含補收款)
   If bolDelay = True And strDelayCP09 <> "" Then
      lngFee = PUB_GetDelayPayFee(strDelayCP09)
      If lngFee > 0 Then
         strSql = "update caseprogress set cp16=NVL(cp16,0)-NVL(cp17,0)+" & lngFee & ",cp17=" & lngFee & " where cp09='" & strReceiveNo & "'"
         cnnConnection.Execute strSql, intI
      End If
   End If
   
   If Text5(1) <> "" Then
      'Modify By Cheng 2002/07/04
'      strExc(2) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07," & _
'         "NP08,NP09,NP10,NP22,NP15) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & _
'         pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & Text5(1) & "'," & _
'         TransDate(Text5(2), 2) & "," & TransDate(Text5(3), 2) & "," & cnull(chgsql(TEXT6)) & "," & objPublicData.GetNextProgressNo & "," & CNULL(Me.Text5(11).Text) & ")"
      'edit by nickc 2007/02/02 不用 dll 了
      'strExc(2) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07," & _
         "NP08,NP09,NP10,NP22,NP15) VALUES ('" & strReceiveNo & "','" & pA(1) & "','" & _
         pA(2) & "','" & pA(3) & "','" & pA(4) & "','" & Text5(1) & "'," & _
         TransDate(Text5(2), 2) & "," & TransDate(Text5(3), 2) & "," & CNULL(PUB_GetFCPSalesNo(pA(1), pA(2), pA(3), pA(4))) & "," & objPublicData.GetNextProgressNo & "," & CNULL(Me.Text5(11).Text) & ")"
      'Modify By Sindy 2021/8/17 + NP23
      strExc(2) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07," & _
         "NP08,NP09,NP10,NP22,NP15,NP23) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & _
         pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & Text5(1) & "'," & _
         TransDate(Text5(2), 2) & "," & TransDate(Text5(3), 2) & "," & CNULL(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))) & "," & GetNextProgressNo & "," & CNULL(Me.Text5(11).Text) & "," & TransDate(Text5(12), 2) & ")"
      cnnConnection.Execute strExc(2)
    
      intStep = 2
   End If
   'Add By Cheng 2002/07/04
   intStep = intStep + 1
   strExc(intStep) = "Update Patent Set PA05='" & Me.Text5(5).Text & "' ,PA06=" & CNULL(ChgSQL(Text5(6))) & ",PA07='" & Me.Text5(7).Text & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   cnnConnection.Execute strExc(intStep)
   'Add By Cheng 2003/10/28
   '案件性質為異議(801), 舉發(803) ,若有輸入對照號數則更新申請案號
   If m_CP10 = "801" Or m_CP10 = "803" Then
       If Me.Text5(4).Text <> "" Then
           intStep = intStep + 1
           strExc(intStep) = "Update Patent Set PA11='" & Me.Text5(4).Text & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
           cnnConnection.Execute strExc(intStep)
       End If
   End If
   
   '2006/3/22 ADD BY SONIA 掛催審期限
   strExc(0) = "SELECT CF05 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND " & _
      "CF02='" & pa(9) & "' AND CF03='" & m_CP10 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(0)) And RsTemp.Fields(0) <> 0 Then
         m_NP09 = PUB_DBDATE(CompDate(2, Val(RsTemp.Fields(0)), TransDate(Text5(0), 2)))
         intStep = intStep + 1
         'edit by nickc 2007/02/02 不用 dll 了
         'strExc(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
            "NP07,NP08,NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pA(1) & _
            "','" & pA(2) & "','" & pA(3) & "','" & pA(4) & "'," & 催審 & "," & _
            m_NP09 & "," & m_NP09 & ",'" & strUserNum & "'," & objPublicData.GetNextProgressNo & ")"
         'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
         strExc(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
            "NP07,NP08,NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
            "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 催審 & "," & _
            PUB_GetWorkDay1(m_NP09, True) & "," & m_NP09 & ",'" & strUserNum & "'," & GetNextProgressNo & ")"
         cnnConnection.Execute strExc(intStep)
      End If
   End If
   '2006/3/22 END
   
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130 'Add by Morgan 2009/3/20
   
   cnnConnection.CommitTrans
   Exit Function

CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False
   
End Function

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(7) = pa(5)
      Case "英"
         Label2(7) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label2(7) = pa(7)
   End Select
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
   'Add by Morgan 2005/8/4
   ReDim pa(TF_PA)
   ReadPatent
   'Add by Morgan 2005/8/4
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1300
   lstNameAgent.Width = 1300

   Label2(0) = strReceiveNo
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2024/03/06
   
   Set frm060104_9 = Nothing
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As Object, i As Integer
Dim strTempName As String

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
            Label2(2) = pa(21)
            Label2(3) = pa(22)
            Label2(4) = pa(77)
            ChgType (999) 'pa(8)
            Label2(7) = pa(5)
            If pa(26) <> "" Then ChgType (11) ' Label2(8)
         End If
      Case "FG"
      
   End Select
   'Modify by Morgan 2004/8/12 Add cp17
   strExc(0) = "select cpm03,staff.st02 as st1,a0902,staff1.st02 as st2," & _
      "cp27,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp64,CP13,CP10, cp17,CP110,cp113,cp14,CP142,CP164 from caseprogress,casepropertymap," & _
      "staff,staff staff1,acc090 where cp09='" & strReceiveNo & "' and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+) and cp12=a0901(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP14 = "" & .Fields("CP14") 'Add by Moragn 2007/8/7
      m_CP142 = "" & .Fields("CP142") 'Add by Sindy 2015/12/17
      m_CP164 = "" & .Fields("CP164") 'Add By Sindy 2021/4/20
      m_CP110 = "" & .Fields("CP110")
      'Add by Morgan 2004/8/12
      m_CP17 = "" & .Fields("cp17")
      
      'Add by Morgan 2004/9/8 檢查是否有延期，若有則規費預設0
      bolDelay = PUB_ChkDelay(strReceiveNo, strDelayCP09)
      If bolDelay = True Then m_CP17 = "0"
      '2004/9/8 end
      
      Text5(11).Text = "" & .Fields("cp64") 'Added by Morgan 2013/1/14 原來漏了
      
      txtCP84.Tag = m_CP17
      txtCP84.Text = txtCP84.Tag
      If Not IsNull(.Fields(0)) Then Label2(5) = .Fields(0)
      'MODIFY BY SONIA 2015/9/21 承辦人為外專程序時,改為操作人員
      'If Not IsNull(.Fields(1)) Then Label2(1) = .Fields(1)
      m_CP14 = GetFCPUser(m_CP14)
      If ClsPDGetStaff(m_CP14, strTempName) Then Label2(1) = strTempName
      'END 2015/9/21
      If Not IsNull(.Fields(2)) Then Label2(10) = .Fields(2)
      If Not IsNull(.Fields(3)) Then Label2(9) = .Fields(3)
      If Not IsNull(.Fields(4)) Then
         Text5(0) = TransDate(.Fields(4), 1)
      Else
         Text5(0) = strSrvDate(2)
      End If
      For i = 5 To 12
         If Not IsNull(.Fields(i)) Then Text5(i - 1) = .Fields(i)
      Next
      'If Not IsNull(.Fields(13)) Then Text6 = .Fields(13) 'Removed by Morgan 2013/1/14 沒用了
      'Add By Cheng 2002/07/17
      m_CP10 = ""
      If Not IsNull(.Fields("CP10")) Then m_CP10 = .Fields("CP10")
      txtCP113 = "" & .Fields("cp113")
   End If
   End With
   
   '92.7.15 cancel by sonia
   'If IsEmptyText(Text5(0)) = False Then
   '   CaculateNP08NP09
   'End If
   '92.7.15 end
   
   'Added by Morgan 2012/10/29
   'Modified by Morgan 2013/1/14 增加舉發事項及規費自動計算
   If pa(9) = 台灣國家代號 And m_CP10 = "803" Then
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
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 0 '發文日
'         If ChkDate(Text5(0)) Or Val(Text5(0).Text) > Val(strSrvDate(2)) Then
'            ChgType = True
'         End If
         If Not ChkDate(Text5(0)) Then
         ElseIf Val(Text5(0).Text) > PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1))) Then
            MsgBox "發文日大於系統日下一個工作日, 請重新輸入!!!", vbExclamation + vbOKOnly
         Else
            ChgType = True
         End If
      
      Case 11
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(26), strTempName) Then
         If ClsLawLawGetName(pa(26), strTempName) Then
            Label2(8) = strTempName
            ChgType = True
         Else
            Label2(8) = ""
         End If
      Case 999
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTempName, , 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, pa(8), strTempName, , 台灣國家代號) = 1 Then
            Label2(6) = strTempName
         End If
   End Select
End Function

Private Sub txtItemList_GotFocus()
   CloseIme
End Sub

Private Sub txtItemList_Validate(Cancel As Boolean)
   SetOfficialFee
End Sub

Private Sub Text5_GotFocus(Index As Integer)
TextInverse Text5(Index)
End Sub

Private Sub Text5_LostFocus(Index As Integer)
   Select Case Index
      Case 3
         If Text5(2) <> "" Then
            If ChkRange(Text5(2), Text5(3), "日期") = False Then
               Text5(2).SetFocus
            End If
            'Add By Sindy 2021/8/17
            If ChkRange(Text5(12), Text5(3), "日期") = False Then
               Text5(12).SetFocus
            End If
            '2021/8/17 END
         End If
      Case 7
         If Text5(5) = "" And Text5(6) = "" And Text5(7) = "" Then
            MsgBox "對造案件名稱不可同時空白 !", vbCritical
            Text5(5).SetFocus
         End If
      Case 10
         If Text5(8) = "" And Text5(9) = "" And Text5(10) = "" Then
            MsgBox "對造名稱不可同時空白 !"
            Text5(8).SetFocus
         End If
   End Select
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0 '發文日
         If Text5(Index) <> "" Then
            If ChgType(0) = False Then
               Cancel = True
            Else
               'CaculateNP08NP09    '92.7.15 cancel by sonia
            End If
         Else
            MsgBox "發文日不可空白 !", vbCritical
            Cancel = True
         End If
      Case 1 '下一程序
         If Text5(Index) <> "" Then
            
            'Add By Cheng 2002/01/04
            If Len(Me.Text5(Index).Text) <> 3 Then
               MsgBox "下一程序欄位值必須為三碼 !", vbCritical
               Cancel = True
            End If
            '92.7.15 add by sonia
            CaculateNP08NP09
            '92.7.15 end
            Select Case Text5(Index)
               Case "202"
                  Label2(11) = "補文件"
               Case "206"
                  Label2(11) = "補充說明"
               Case Else
                  MsgBox "下一程序只能為補充說明或補文件 !", vbCritical
                  Cancel = True
            End Select
         End If
      Case 2, 3
         If Text5(1) = "" Then
            If Text5(Index) <> "" Then
               MsgBox "下一程序為空白時，此欄必須為空白 !", vbCritical
               Cancel = True
            End If
         Else
            If Text5(Index) = "" Then
               MsgBox "下一程序為空白時，此欄不得為空白 !", vbCritical
               Cancel = True
            Else
               If Not ChkDate(Text5(Index)) Then
                  Cancel = True
               End If
            End If
         End If
      Case 4
         If Text5(Index) = "" Then
            MsgBox "對造號數不可空白 !", vbCritical
            Cancel = True
         End If
   End Select
   If Cancel Then TextInverse Text5(Index)
End Sub

' 計算本所期限及法定期限
Private Sub CaculateNP08NP09()
   If IsEmptyText(Text5(0)) = False Then
      strExc(0) = TransDate(Text5(0).Text, 2)
      'edit by nickc 2007/02/05 不用 dll 了
      'If objLawDll.GetCaseFeeDelay(pa(1), pa(9), m_CP10, strExc) Then
      If ClsLawGetCaseFeeDelay(pa(1), pa(9), m_CP10, strExc) Then
         Text5(3) = TransDate(strExc(1), 1)
         Text5(2) = TransDate(strExc(2), 1)
         Text5(12) = TransDate(strExc(3), 1) 'Add By Sindy 2021/6/17 約定期限
      End If
   End If
End Sub

'Add By Cheng 2002/05/21
Private Function CheckDataValid() As Boolean
Dim ii As Integer
'Add by Morgan 2004/8/12
Dim Cancel As Boolean

   CheckDataValid = False
   '檢查發文日
   If Text5(0) <> "" Then
      If ChgType(0) = False Then
         Me.Text5(0).SetFocus
         Text5_GotFocus 0
         Exit Function
      End If
   Else
      MsgBox "發文日不可空白 !", vbCritical
      Me.Text5(0).SetFocus
      Text5_GotFocus 0
      Exit Function
   End If
   '檢查下一程序
   If Text5(1) <> "" Then
      'Add By Cheng 2002/01/04
      If Len(Me.Text5(1).Text) <> 3 Then
         MsgBox "下一程序欄位值必須為三碼 !", vbCritical
         Me.Text5(1).SetFocus
         Text5_GotFocus 1
         Exit Function
      End If
      Select Case Text5(1)
         Case "202"
            Label2(11) = "補文件"
         Case "206"
            Label2(11) = "補充說明"
         Case Else
            MsgBox "下一程序只能為補充說明或補文件 !", vbCritical
            Me.Text5(1).SetFocus
            Text5_GotFocus 1
            Exit Function
      End Select
   End If
   '檢查下一程序之
   For ii = 2 To 3
      If Text5(1) = "" Then
         If Text5(ii) <> "" Then
            MsgBox "下一程序為空白時，此欄必須為空白 !", vbCritical
            Me.Text5(ii).SetFocus
            Text5_GotFocus ii
            Exit Function
         End If
      Else
         If Text5(ii) = "" Then
            MsgBox "下一程序為空白時，此欄不得為空白 !", vbCritical
            Me.Text5(ii).SetFocus
            Text5_GotFocus ii
            Exit Function
         Else
            If Not ChkDate(Text5(ii)) Then
               Me.Text5(ii).SetFocus
               Text5_GotFocus ii
               Exit Function
            Else
               If ii = 3 Then
                  If ChkRange(Text5(2), Text5(3), "日期") = False Then
                     Me.Text5(2).SetFocus
                     Text5_GotFocus 2
                     Exit Function
                  End If
                  'Add By Sindy 2021/8/17
                  If ChkRange(Text5(12), Text5(3), "日期") = False Then
                     Me.Text5(12).SetFocus
                     Text5_GotFocus 12
                     Exit Function
                  End If
                  '2021/8/17 END
               End If
            End If
         End If
      End If
   Next ii
   '檢查對造號數
   If Text5(4) = "" Then
      MsgBox "對造號數不可空白 !", vbCritical
      Me.Text5(4).SetFocus
      Text5_GotFocus 4
      Exit Function
   End If
   '檢查對造案件名稱
   If Text5(5) = "" And Text5(6) = "" And Text5(7) = "" Then
      MsgBox "對造案件名稱不可同時空白 !", vbCritical
      Text5(5).SetFocus
      Text5_GotFocus 5
      Exit Function
   End If
   '檢查對告名稱
   If Text5(8) = "" And Text5(9) = "" And Text5(10) = "" Then
      MsgBox "對造名稱不可同時空白 !"
      Text5(8).SetFocus
      Text5_GotFocus 8
      Exit Function
   End If
   
   'Add by Morgan 2004/8/12
   If txtCP84.Enabled = True Then
      Cancel = False
      txtCP84_Validate Cancel
      If Cancel = True Then
         txtCP84.SetFocus
         txtCP84_GotFocus
         Exit Function
      End If
   End If
   
   'Add by Morgan 2005/8/4
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   'Add by Morgan 2007/7/20
   txtCP113_Validate Cancel
   If Cancel = True Then Exit Function
   'end 2007/7/20
      
   'Added by Morgan 2012/10/29
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
   End If
   'end 2012/10/29
   
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
   
   CheckDataValid = True
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

Private Sub txtItemCount_Validate(Cancel As Boolean)
   SetOfficialFee
End Sub

Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

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
   Cancel = Not PUB_CheckCP113(txtCP113, pa(1), m_CP10, m_CP14)
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
'Add by Morgan 2005/8/4
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

