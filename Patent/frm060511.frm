VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060511 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專案件清單Excel"
   ClientHeight    =   8736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8928
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8736
   ScaleWidth      =   8928
   Begin VB.CommandButton Cmd2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "其他特定清單"
      Height          =   360
      Left            =   6864
      Style           =   1  '圖片外觀
      TabIndex        =   93
      Top             =   8328
      Visible         =   0   'False
      Width           =   1668
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修正清單編號"
      Height          =   564
      Left            =   9240
      TabIndex        =   89
      Top             =   168
      Width           =   1212
   End
   Begin VB.ComboBox Combo2 
      Height          =   276
      ItemData        =   "frm060511.frx":0000
      Left            =   3864
      List            =   "frm060511.frx":0002
      TabIndex        =   2
      Top             =   456
      Width           =   2340
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm060511.frx":0004
      Left            =   4440
      List            =   "frm060511.frx":0006
      TabIndex        =   32
      Top             =   5400
      Width           =   1380
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除"
      Height          =   324
      Left            =   72
      TabIndex        =   80
      Top             =   72
      Width           =   950
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&Q)"
      Height          =   324
      Left            =   7032
      TabIndex        =   79
      Top             =   72
      Width           =   852
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&E)"
      Height          =   324
      Left            =   7992
      TabIndex        =   72
      Top             =   72
      Width           =   852
   End
   Begin VB.CheckBox Check2 
      Caption         =   "排除年費不續辦"
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   1
      Left            =   1560
      TabIndex        =   31
      Top             =   5400
      Width           =   1620
   End
   Begin VB.CheckBox Check2 
      Caption         =   "只印未核准"
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   3
      Left            =   6984
      TabIndex        =   36
      Top             =   5808
      Width           =   1400
   End
   Begin VB.CheckBox Check2 
      Caption         =   "只印已核准"
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   2
      Left            =   5472
      TabIndex        =   35
      Top             =   5808
      Width           =   1400
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生EXCEL(&E)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      Style           =   1  '圖片外觀
      TabIndex        =   42
      Top             =   8328
      Width           =   2100
   End
   Begin VB.Frame Frame4 
      Caption         =   "案件清單顯示ITEM"
      Height          =   1940
      Left            =   24
      TabIndex        =   55
      Top             =   6312
      Width           =   8724
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   6504
         MaxLength       =   10
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   168
         Width           =   1212
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "查詢"
         Height          =   300
         Left            =   7848
         TabIndex        =   39
         Top             =   168
         Width           =   636
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   936
         TabIndex        =   65
         Top             =   1344
         Width           =   280
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   936
         TabIndex        =   64
         Top             =   984
         Width           =   280
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除 ->"
         Height          =   285
         Index           =   2
         Left            =   4344
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   516
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<- 新增"
         Height          =   285
         Index           =   2
         Left            =   4344
         TabIndex        =   40
         Top             =   216
         Width           =   735
      End
      Begin VB.CommandButton CmcClear 
         Caption         =   "清除"
         Height          =   300
         Index           =   2
         Left            =   70
         TabIndex        =   56
         Top             =   936
         Width           =   684
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "輸入代號/名稱："
         Height          =   228
         Index           =   9
         Left            =   5208
         TabIndex        =   67
         Top             =   216
         Width           =   1308
      End
      Begin VB.Label Lbl1 
         Caption         =   "排列由上而下 "
         Height          =   228
         Index           =   7
         Left            =   96
         TabIndex        =   63
         Top             =   672
         Width           =   1092
      End
      Begin VB.Label lblCnt 
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
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
         Height          =   228
         Index           =   2
         Left            =   672
         TabIndex        =   62
         Top             =   360
         Width           =   468
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "數量："
         ForeColor       =   &H00FF0000&
         Height          =   228
         Index           =   5
         Left            =   96
         TabIndex        =   61
         Top             =   360
         Width           =   636
      End
      Begin MSForms.ListBox ListBox1 
         Height          =   330
         Index           =   3
         Left            =   5232
         TabIndex        =   59
         Top             =   552
         Width           =   3300
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "5821;582"
         MatchEntry      =   0
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox ListBox1 
         Height          =   330
         Index           =   2
         Left            =   1224
         TabIndex        =   57
         Top             =   216
         Width           =   3060
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "5397;582"
         MatchEntry      =   0
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "含閉卷／銷卷"
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   0
      Left            =   96
      TabIndex        =   30
      Top             =   5400
      Width           =   1400
   End
   Begin VB.Frame Frame3 
      Caption         =   "清單選項"
      Height          =   1764
      Left            =   24
      TabIndex        =   50
      Top             =   3480
      Width           =   8724
      Begin VB.CheckBox Check1 
         Caption         =   "行事曆(獨立工作表，含本所案號、事由、管制日期)"
         Height          =   180
         Index           =   5
         Left            =   2808
         TabIndex        =   16
         Top             =   168
         Width           =   4400
      End
      Begin VB.CheckBox Check1 
         Caption         =   "未付款帳單(請款日期在"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   17
         Top             =   480
         Width           =   2268
      End
      Begin VB.CheckBox Check1 
         Caption         =   "未請款(發文日在"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   1680
         TabIndex        =   23
         Top             =   840
         Width           =   1644
      End
      Begin VB.CheckBox Check1 
         Caption         =   "未發文"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   21
         Top             =   840
         Width           =   1092
      End
      Begin VB.CheckBox Check1 
         Caption         =   "案件清單"
         Height          =   204
         Index           =   0
         Left            =   210
         TabIndex        =   15
         Top             =   216
         Width           =   1212
      End
      Begin VB.CheckBox Check1 
         Caption         =   "法限落在"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   216
         TabIndex        =   25
         Top             =   1128
         Width           =   1032
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   38
         Left            =   6744
         TabIndex        =   20
         Top             =   408
         Width           =   372
         VariousPropertyBits=   679495707
         MaxLength       =   1
         Size            =   "656;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   252
         Index           =   7
         Left            =   3432
         TabIndex        =   92
         Top             =   432
         Width           =   156
         ForeColor       =   16711680
         Caption         =   "~"
         Size            =   "275;444"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   252
         Index           =   6
         Left            =   4560
         TabIndex        =   91
         Top             =   456
         Width           =   3540
         ForeColor       =   0
         Caption         =   ")  只列出有未付款之案件：　　   (Y:是)"
         Size            =   "6244;444"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   37
         Left            =   3600
         TabIndex        =   19
         Top             =   420
         Width           =   900
         VariousPropertyBits=   679495707
         MaxLength       =   8
         Size            =   "1587;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   36
         Left            =   2520
         TabIndex        =   18
         Top             =   420
         Width           =   900
         VariousPropertyBits=   679495707
         MaxLength       =   8
         Size            =   "1587;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   3
         X1              =   120
         X2              =   8592
         Y1              =   744
         Y2              =   744
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   31
         Left            =   6504
         TabIndex        =   28
         Top             =   1068
         Width           =   372
         VariousPropertyBits=   679495707
         MaxLength       =   1
         Size            =   "656;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   252
         Index           =   4
         Left            =   5040
         TabIndex        =   81
         Top             =   1128
         Width           =   3516
         ForeColor       =   16711935
         Caption         =   "是否預估請款金額          (Y:是，含C類來函)"
         Size            =   "6202;444"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   252
         Index           =   3
         Left            =   5424
         TabIndex        =   75
         Top             =   840
         Width           =   1236
         ForeColor       =   0
         Caption         =   "前但尚未請款)"
         Size            =   "2180;444"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   252
         Index           =   2
         Left            =   4296
         TabIndex        =   74
         Top             =   816
         Width           =   156
         ForeColor       =   0
         Caption         =   "~"
         Size            =   "275;444"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   15
         Left            =   3336
         TabIndex        =   22
         Top             =   792
         Width           =   900
         VariousPropertyBits=   679495707
         MaxLength       =   8
         Size            =   "1587;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   16
         Left            =   4464
         TabIndex        =   24
         Top             =   792
         Width           =   900
         VariousPropertyBits=   679495707
         MaxLength       =   8
         Size            =   "1587;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   252
         Index           =   1
         Left            =   3288
         TabIndex        =   71
         Top             =   1128
         Width           =   1236
         ForeColor       =   0
         Caption         =   "前但尚未收文"
         Size            =   "2180;444"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label1 
         Height          =   252
         Index           =   0
         Left            =   2160
         TabIndex        =   70
         Top             =   1080
         Width           =   156
         ForeColor       =   16711680
         Caption         =   "~"
         Size            =   "275;444"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   19
         Left            =   2328
         TabIndex        =   27
         Top             =   1080
         Width           =   900
         VariousPropertyBits=   679495707
         MaxLength       =   8
         Size            =   "1587;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   18
         Left            =   1248
         TabIndex        =   26
         Top             =   1080
         Width           =   900
         VariousPropertyBits=   679495707
         MaxLength       =   8
         Size            =   "1587;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDB 
         Height          =   300
         Index           =   20
         Left            =   4224
         TabIndex        =   29
         Top             =   1416
         Width           =   372
         VariousPropertyBits=   679495707
         MaxLength       =   1
         Size            =   "656;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Lbl1 
         Caption         =   "未發文/未請款/未收文是否匯出同一工作表(Sheet)：        (Y: 是，最後面欄位固定為狀態)"
         ForeColor       =   &H00FF00FF&
         Height          =   228
         Index           =   4
         Left            =   216
         TabIndex        =   60
         Top             =   1464
         Width           =   7044
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "代理人編號"
      Height          =   880
      Left            =   24
      TabIndex        =   48
      Top             =   1224
      Width           =   8724
      Begin VB.CommandButton CmcClear 
         Caption         =   "清除"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   220
         Width           =   684
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   4510
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   120
         Width           =   1000
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<- 新增"
         Height          =   285
         Index           =   0
         Left            =   3700
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除 ->"
         Height          =   285
         Index           =   0
         Left            =   3700
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Lbl1 
         Caption         =   "數量："
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   6
         Left            =   120
         TabIndex        =   86
         Top             =   600
         Width           =   564
      End
      Begin VB.Label lblCnt 
         Caption         =   "0"
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
         Height          =   228
         Index           =   0
         Left            =   720
         TabIndex        =   84
         Top             =   600
         Width           =   204
      End
      Begin MSForms.ListBox ListBox1 
         Height          =   336
         Index           =   0
         Left            =   1032
         TabIndex        =   53
         Top             =   108
         Width           =   2600
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "4593;593"
         MatchEntry      =   0
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   276
         Index           =   1
         Left            =   5568
         TabIndex        =   49
         Top             =   144
         Width           =   3000
         BackColor       =   16777215
         Size            =   "5292;487"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "申請人編號"
      Height          =   880
      Left            =   24
      TabIndex        =   46
      Top             =   2184
      Width           =   8724
      Begin VB.CheckBox Check3 
         Caption         =   "不同申請人分成不同Sheet"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   6288
         TabIndex        =   12
         Top             =   504
         Width           =   2268
      End
      Begin VB.CheckBox Check3 
         Caption         =   "限第一申請人"
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   0
         Left            =   4632
         TabIndex        =   11
         Top             =   504
         Width           =   1476
      End
      Begin VB.CommandButton CmcClear 
         Caption         =   "清除"
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   216
         Width           =   684
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   4510
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   144
         Width           =   1000
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除 ->"
         Height          =   285
         Index           =   1
         Left            =   3700
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<- 新增"
         Height          =   285
         Index           =   1
         Left            =   3700
         TabIndex        =   9
         Top             =   144
         Width           =   735
      End
      Begin VB.Label Lbl1 
         Caption         =   "數量："
         ForeColor       =   &H00000000&
         Height          =   228
         Index           =   11
         Left            =   96
         TabIndex        =   87
         Top             =   576
         Width           =   564
      End
      Begin VB.Label lblCnt 
         Caption         =   "0"
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
         Height          =   228
         Index           =   1
         Left            =   720
         TabIndex        =   85
         Top             =   576
         Width           =   204
      End
      Begin MSForms.ListBox ListBox1 
         Height          =   330
         Index           =   1
         Left            =   1030
         TabIndex        =   51
         Top             =   110
         Width           =   2600
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "4586;582"
         MatchEntry      =   0
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   276
         Index           =   2
         Left            =   5568
         TabIndex        =   47
         Top             =   144
         Width           =   3000
         BackColor       =   16777215
         Size            =   "5292;487"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.TextBox txtDB 
      Height          =   300
      Index           =   40
      Left            =   7800
      TabIndex        =   14
      Top             =   3096
      Width           =   372
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "656;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl1 
      Caption         =   "只有個案設定：         (Y:是)"
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   14
      Left            =   6552
      TabIndex        =   94
      Top             =   3144
      Width           =   2292
   End
   Begin MSForms.TextBox txtDB 
      Height          =   320
      Index           =   35
      Left            =   960
      TabIndex        =   3
      Top             =   816
      Width           =   5004
      VariousPropertyBits=   679495707
      MaxLength       =   200
      Size            =   "8826;564"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl1 
      Caption         =   "清單備註："
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   13
      Left            =   48
      TabIndex        =   90
      Top             =   870
      Width           =   924
   End
   Begin VB.Label Lbl1 
      Caption         =   "顯示日期格式："
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   12
      Left            =   2592
      TabIndex        =   88
      Top             =   516
      Width           =   1284
   End
   Begin MSForms.Label Label1 
      Height          =   252
      Index           =   5
      Left            =   6984
      TabIndex        =   83
      Top             =   5400
      Width           =   156
      ForeColor       =   0
      Caption         =   "~"
      Size            =   "275;444"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDB 
      Height          =   324
      Index           =   34
      Left            =   7128
      TabIndex        =   34
      Top             =   5376
      Width           =   1020
      VariousPropertyBits=   679495707
      MaxLength       =   8
      Size            =   "1799;564"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDB 
      Height          =   324
      Index           =   33
      Left            =   5904
      TabIndex        =   33
      Top             =   5376
      Width           =   1020
      VariousPropertyBits=   679495707
      MaxLength       =   8
      Size            =   "1799;564"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl1 
      Caption         =   "日期範圍："
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   10
      Left            =   3456
      TabIndex        =   82
      Top             =   5424
      Width           =   924
   End
   Begin MSForms.TextBox txtDB 
      Height          =   336
      Index           =   25
      Left            =   1680
      TabIndex        =   78
      Top             =   8340
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   30
      Size            =   "2117;593"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   372
      Index           =   4
      Left            =   1104
      TabIndex        =   77
      Top             =   24
      Width           =   4104
      Size            =   "7239;656"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl1 
      Caption         =   "排序欄位(從上到下位置ex.1,3,5)："
      ForeColor       =   &H00FF0000&
      Height          =   372
      Index           =   1
      Left            =   72
      TabIndex        =   76
      Top             =   8328
      Width           =   1572
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   3
      Left            =   2448
      TabIndex        =   73
      Top             =   3108
      Width           =   4032
      BackColor       =   16777215
      Size            =   "7112;487"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDB 
      Height          =   336
      Index           =   26
      Left            =   5376
      TabIndex        =   69
      Top             =   8304
      Visible         =   0   'False
      Width           =   2268
      VariousPropertyBits=   679495707
      Size            =   "4000;593"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDB 
      Height          =   336
      Index           =   7
      Left            =   9576
      TabIndex        =   68
      Top             =   2136
      Visible         =   0   'False
      Width           =   3132
      VariousPropertyBits=   679495707
      Size            =   "5524;593"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDB 
      Height          =   504
      Index           =   11
      Left            =   1200
      TabIndex        =   37
      Top             =   5760
      Width           =   4068
      VariousPropertyBits=   -1467987941
      MaxLength       =   200
      Size            =   "7175;889"
      Value           =   "(輸入姓名不分大小寫也不用加上稱謂，用空白或;區隔KeyWord，程式以模糊比對進行尋找)"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl1 
      Caption         =   "指定聯絡人："
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   8
      Left            =   72
      TabIndex        =   66
      Top             =   5784
      Width           =   1092
   End
   Begin MSForms.TextBox txtDB 
      Height          =   320
      Index           =   23
      Left            =   7008
      TabIndex        =   4
      Top             =   816
      Width           =   372
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "656;564"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl1 
      Caption         =   "欄位抬頭：　　  (1.英文  2.中文)"
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   3
      Left            =   6096
      TabIndex        =   58
      Top             =   864
      Width           =   2676
   End
   Begin MSForms.TextBox txtDB 
      Height          =   336
      Index           =   10
      Left            =   1392
      TabIndex        =   13
      Top             =   3072
      Width           =   1020
      VariousPropertyBits=   679495707
      MaxLength       =   9
      Size            =   "1799;593"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl1 
      Caption         =   "固定請款對象："
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   2
      Left            =   48
      TabIndex        =   45
      Top             =   3120
      Width           =   1308
   End
   Begin MSForms.Label lblFM2 
      Height          =   276
      Index           =   0
      Left            =   1656
      TabIndex        =   44
      Top             =   480
      Width           =   852
      BackColor       =   16777215
      Size            =   "1503;487"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDB 
      Height          =   336
      Index           =   6
      Left            =   9696
      TabIndex        =   43
      Top             =   936
      Visible         =   0   'False
      Width           =   3180
      VariousPropertyBits=   679495707
      Size            =   "5609;593"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDB 
      Height          =   320
      Index           =   5
      Left            =   960
      TabIndex        =   1
      Top             =   456
      Width           =   660
      VariousPropertyBits=   679495707
      MaxLength       =   6
      Size            =   "1164;564"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl1 
      Caption         =   "需求人員："
      ForeColor       =   &H00FF0000&
      Height          =   228
      Index           =   0
      Left            =   48
      TabIndex        =   0
      Top             =   504
      Width           =   924
   End
End
Attribute VB_Name = "frm060511"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRead As New ADODB.Recordset
Dim rsQD As New ADODB.Recordset
Dim strConSql(0 To 5) As String  '條件清單語法
Dim strTitle As String  '欄位抬頭
Dim maxField As Integer 'FCPEListRec: Table欄位數
Dim maxSeq As String '執行序號
Dim strNowNo As String '現在流水號
Dim strPreNo As String '匯入流水號
Dim strFER05 As String, strFER06 As String, strFER07 As String, strFER10 As String '匯入: 需求人員, 代理人, 申請人, 固定請款對象
Dim strFileName As String '檔案名稱
Dim intR As Integer, intQ As Integer, strQuery As String

Dim oObj As Object
Dim intL(0 To 2) As Integer 'ListBox1對應的txtDB
Private Const contSpec As String = "(輸入姓名不分大小寫也不用加上稱謂，用空白或;區隔KeyWord，程式以模糊比對進行尋找)"

Private Sub ClearForm(Optional ByVal bolReset As Boolean)
   
   For Each oObj In txtDB
      oObj.Text = ""
      oObj.Tag = ""
   Next
   txtDB(11).Text = contSpec
   For Each oObj In Text1
      oObj.Text = ""
   Next
   
   For Each oObj In lblFM2
      oObj.Caption = ""
   Next
   For Each oObj In LblCnt
      oObj.Caption = ""
   Next
   For Each oObj In Check1
      oObj.Value = 0
   Next
   For Each oObj In Check2
      oObj.Value = 0
   Next
   For Each oObj In Check3
      oObj.Value = 0
   Next
   
   For Each oObj In ListBox1
      oObj.Clear
   Next
 
   Combo1.Text = ""
   Combo2.Text = ""
   strPreNo = "": strFER05 = "": strFER06 = "": strFER07 = "": strFER10 = ""
   Call doQueryItem  '清單全部列出
End Sub

Public Function QueryData(Optional ByVal PKno As String) As Boolean  '匯入記錄
Dim strCon As String
    
   If PKno <> "" Then  '從查詢畫面傳入流水號
      strCon = " AND FER01=" & CNULL(PKno)
   Else
      '需求人員
      If txtDB(5) <> "" Then
         strCon = strCon & " AND FER05=" & CNULL(txtDB(5))
      End If
      '代理人編號
      If txtDB(6) <> "" Then
         strCon = strCon & " AND INSERT(FER06," & CNULL(txtDB(6)) & ") > 0"
      ElseIf Text1(0) <> "" And lblFM2(1).Caption <> "" Then
         strCon = strCon & " AND INSERT(FER06," & CNULL(Text1(0)) & ") > 0"
      End If
      '客戶編號
      If txtDB(7) <> "" Then
         strCon = strCon & " AND INSERT(FER07," & CNULL(txtDB(7)) & ") > 0"
      ElseIf Text1(1) <> "" And lblFM2(2).Caption <> "" Then
         strCon = strCon & " AND INSERT(FER07," & CNULL(Text1(1)) & ") > 0"
      End If
      '固定請款對象
      If txtDB(10) <> "" And lblFM2(3).Caption <> "" Then
         strCon = strCon & " AND INSERT(FER10," & CNULL(txtDB(10)) & ") > 0"
      End If
   End If
   
   strQuery = " select a.*,b.st02 as fer05n,getfagentnamelist(fer06) as fer06n" & _
              " ,getcustomernamelist(fer07) as fer07n" & _
              " ,decode(substr(fer10,1,1),'Y',getfagentnamelist(fer10),'X',getcustomernamelist(fer10),null) as fer10n" & _
              " ,getferlist(fer26,'1') as fer26n From fcpelistrec a, staff b where fer05=st01(+)" & strCon
   strQuery = strQuery & " order by fer03 desc, fer04 desc "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      QueryData = True
      rsQD.MoveFirst
      Call SetCtrlData(rsQD)
      'Added by Lydia 2025/01/17
      If Trim("" & rsQD.Fields("FER39")) <> "" Then
         MsgBox "此為固定清單記錄，請改用其他特定清單功能！", vbInformation
      End If
      'end 2025/01/17
   End If
   
End Function

Private Sub Check1_Click(Index As Integer)
   If Check1(Index).Value = 1 Then
      Select Case Index
         Case 0 '案件清單
            If Check1(1).Value = 1 Or Check1(2).Value = 1 Or Check1(3).Value = 1 Or Check1(4).Value = 1 Then
                MsgBox "不可同時執行【未發文／未請款／未收文／未付款】的清單！", vbExclamation
                Check1(1).Value = False
                Check1(2).Value = False
                Check1(3).Value = False
                Check1(4).Value = False
                txtDB(15) = "": txtDB(16) = ""
                txtDB(18) = "": txtDB(19) = "": txtDB(20) = ""
            End If
         Case 1, 2, 3 '未發文/未請款/未收文
            If Check1(0).Value = 1 Or Check1(4).Value = 1 Then
                MsgBox "不可同時執行【案件清單／未付款】的清單！", vbExclamation
                Check1(0).Value = False
                Check1(4).Value = False
            End If
         Case 4 '未付款
            If Check1(0).Value = 1 Or Check1(2).Value = 1 Or Check1(3).Value = 1 Then
                MsgBox "不可同時執行【案件清單／未發文／未請款／未收文】的清單！", vbExclamation
                Check1(0).Value = False
                Check1(1).Value = False
                Check1(2).Value = False
                Check1(3).Value = False
                txtDB(15) = "": txtDB(16) = ""
                txtDB(18) = "": txtDB(19) = "": txtDB(20) = ""
            End If
      End Select
   End If
End Sub

Private Sub Check2_Click(Index As Integer)
   If Check2(Index).Value = 1 Then
      Select Case Index
         Case 2 '只印已核准
            If Check2(3).Value = 1 Then
               MsgBox "不可同時勾選【只印未核准】！", vbExclamation
               Check2(3).Value = 0
            End If
         Case 3 '只印未核准
            If Check2(2).Value = 1 Then
               MsgBox "不可同時勾選【只印已核准】！", vbExclamation
               Check2(2).Value = 0
            End If
      End Select
   End If
End Sub

Private Sub CmcClear_Click(Index As Integer)
   
   If ListBox1(Index).ListCount > 0 Then
      strExc(1) = ""
      Select Case Index
         Case 0: strExc(1) = "代理人編號"
         Case 1: strExc(1) = "申請人編號"
         Case 2: strExc(1) = "案件清單顯示ITEM"
      End Select
      If MsgBox("是否要清除已輸入的" & strExc(1) & "？", vbInformation + vbYesNo + vbSystemModal + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   ListBox1(Index).Clear
   LblCnt(Index) = ""
   txtDB(intL(Index)) = ""
End Sub

Private Sub cmd2_Click()

  Call frm060511_2.SetParent(Me)
  frm060511_2.Show
  Me.Hide
End Sub

Private Sub cmdAdd_Click(Index As Integer)
   
   Call AddlstNo(Index)
   LblCnt(Index) = ListBox1(Index).ListCount

End Sub

Private Sub cmdClear_Click()
   Call ClearForm
End Sub

Private Sub cmdDown_Click()
Dim intP As Integer, ii As Integer
Dim strTempB As String, tmpArr2 As Variant

   If ListBox1(2).ListCount > 0 Then
      For ii = 0 To ListBox1(2).ListCount - 1
         If ListBox1(2).Selected(ii) = True Then
            strTempB = strTempB & ";" & ListBox1(2).List(ii)
         End If
      Next ii
      If strTempB <> "" Then
         strTempB = Mid(strTempB, 2)
         tmpArr2 = Empty
         tmpArr2 = Split(strTempB, ";")
         For ii = 0 To UBound(tmpArr2)
            If Trim(tmpArr2(ii)) <> "" Then
               '考慮複數筆的移動後,原本記錄的位置有變=>全部改成用欄位值判斷現在位置
               intP = GetNowPos(2, Trim(tmpArr2(ii)))
               If intP = ListBox1(2).ListCount - 1 Then
                  MsgBox "欄位已在最後一筆!" & vbCrLf & tmpArr2(ii), vbInformation
                  ListBox1(2).Selected(intP) = True
               ElseIf intP >= 0 Then
                  ListBox1(2).RemoveItem intP
                  ListBox1(2).AddItem tmpArr2(ii), intP + 1
                  ListBox1(2).Selected(intP + 1) = True
               End If
            End If
         Next ii
         txtDB(intL(2)).Text = ComposeList(2)
      End If
   End If
End Sub

'取得指定欄位在ListBox的位置
Private Function GetNowPos(ByVal p_idx As Integer, ByVal pPosNo As String) As Integer
Dim intP As Integer, ii As Integer

   intP = -1
   If p_idx = 2 Then
      intQ = 4
   Else
      intQ = 10
   End If
   If ListBox1(p_idx).ListCount > 0 Then
      For ii = 0 To ListBox1(p_idx).ListCount - 1
         If Trim(Left(ListBox1(p_idx).List(ii), intQ)) = Trim(Left(pPosNo, intQ)) Then
            intP = ii
            Exit For
         End If
      Next
   End If
   If intP >= 0 Then
      GetNowPos = intP
   End If
End Function

Private Sub SetCtrlData(ByRef rsOLD As ADODB.Recordset)
Dim intX As Integer, intP As Integer, intB As Integer
Dim tmpArr As Variant

   If "" & rsOLD.Fields("FER01") & rsOLD.Fields("FER02") <> "" Then
      Call ClearForm
      Call cmdItem_Click
      lblFM2(4).Caption = "匯入記錄：" & rsOLD.Fields("FER01") & String(4, " ") & "需求人員：" & rsOLD.Fields("FER05") & "  " & rsOLD.Fields("FER05N") & String(4, " ") & "Create Date：" & ChangeWStringToTDateString("" & rsOLD.Fields("FER03")) & "  " & Format("" & rsOLD.Fields("FER04"), "00:00:00")
      strPreNo = "" & rsOLD.Fields("FER01") '匯入編號
      strFER05 = "" & rsOLD.Fields("FER05")
      strFER06 = "" & rsOLD.Fields("FER06")
      strFER07 = "" & rsOLD.Fields("FER07")
      strFER10 = "" & rsOLD.Fields("FER10")
      For intP = 5 To maxField   'FCPEListRec: Table欄位數
         Select Case intP
            Case 5, 10 '需求人員,
               If "" & rsOLD.Fields("FER" & Format(intP, "00")) <> "" Then
                  txtDB(intP).Text = "" & rsOLD.Fields("FER" & Format(intP, "00"))
                  If intP = 5 Then
                     lblFM2(0).Caption = "" & rsOLD.Fields("FER" & Format(intP, "00") & "N")
                  ElseIf intP = 10 Then
                     lblFM2(3).Caption = "" & rsOLD.Fields("FER" & Format(intP, "00") & "N")
                  End If
                  txtDB(intP).Tag = txtDB(intP).Text
               End If
               
            Case 6, 7, 26  '代理人,申請人,顯示ITEM
               If "" & rsOLD.Fields("FER" & Format(intP, "00")) <> "" Then
                  txtDB(intP).Text = "" & rsOLD.Fields("FER" & Format(intP, "00"))
                  txtDB(intP).Tag = txtDB(intP).Text
                  intB = 0
                  tmpArr = Empty
                  tmpArr = Split("" & rsOLD.Fields("FER" & Format(intP, "00") & "N"), ";")
                  For intX = UBound(tmpArr) To 0 Step -1
                     If Trim(tmpArr(intX)) <> "" And InStr("," & tmpArr(intX), "(無)") = 0 Then
                        If intP = 6 Then  '代理人
                           ListBox1(0).AddItem Trim(tmpArr(intX)), 0
                           intB = intB + 1
                        ElseIf intP = 7 Then '申請人
                           ListBox1(1).AddItem Trim(tmpArr(intX)), 0
                           intB = intB + 1
                        ElseIf intP = 26 Then '顯示ITEM
                           ListBox1(2).AddItem Trim(tmpArr(intX)), 0
                           intB = intB + 1
                        End If
                     End If
                  Next intX
                  If intP = 6 Then
                     LblCnt(0) = intB
                     ListBox1(0).ListIndex = 0
                  ElseIf intP = 7 Then
                     LblCnt(1) = intB
                     ListBox1(1).ListIndex = 0
                  ElseIf intP = 26 Then
                     LblCnt(2) = intB
                     ListBox1(2).ListIndex = 0
                  End If
                End If
            Case 15, 16, 18, 19, 20, 23, 25, 31, 33, 34, 35, 36, 37, 38, 40
               txtDB(intP).Text = "" & rsOLD.Fields("FER" & Format(intP, "00"))
               txtDB(intP).Tag = txtDB(intP).Text
            Case 8, 9
               If "" & rsOLD.Fields("FER" & Format(intP, "00")) = "Y" Then
                  If intP = 8 Then  '申請人條件1:限基本檔第一申請人(PA26)
                     Check3(0).Value = 1
                  ElseIf intP = 9 Then '申請人條件2:不同申請人分成不同Sheet
                     Check3(1).Value = 1
                  End If
               End If
            Case 11   '指定聯絡人: 去掉固定提示
               If "" & rsOLD.Fields("FER" & Format(intP, "00")) <> "" Then
                  txtDB(intP).Text = "" & rsOLD.Fields("FER" & Format(intP, "00"))
                  txtDB(intP).Tag = txtDB(intP).Text
               End If
            Case 12, 13, 14, 17, 21, 22
               If "" & rsOLD.Fields("FER" & Format(intP, "00")) = "Y" Then
                  If intP = 12 Then   '案件清單
                     Check1(0).Value = 1
                  ElseIf intP = 13 Then '未發文清單
                     Check1(1).Value = 1
                  ElseIf intP = 14 Then  '未請款清單
                     Check1(2).Value = 1
                  ElseIf intP = 17 Then  '未收文清單(下一程序)
                     Check1(3).Value = 1
                  ElseIf intP = 21 Then  '未付款之案件
                     Check1(4).Value = 1
                  ElseIf intP = 22 Then  '行事曆清單
                     Check1(5).Value = 1
                  End If
               End If
            Case 24
               '顯示日期格式: 1. 西元年(YYYY/MM/DD), 2. 民國年, 3. 西元年(DD.MM.YYYY)
               If "" & rsOLD.Fields("FER" & Format(intP, "00")) <> "" Then
                  Combo2.ListIndex = Val("" & rsOLD.Fields("FER" & Format(intP, "00"))) - 1
               Else
                  Combo2.Text = ""
               End If
            Case 27, 28, 29, 30
               If "" & rsOLD.Fields("FER" & Format(intP, "00")) = "Y" Then
                  If intP = 27 Then   '含閉卷／銷卷
                     Check2(0).Value = 1
                  ElseIf intP = 28 Then '排除年費不續辦
                     Check2(1).Value = 1
                  ElseIf intP = 29 Then '只印已核准
                     Check2(2).Value = 1
                  ElseIf intP = 30 Then '只印未核准
                     Check2(3).Value = 1
                  End If
               End If
            Case 32 '日期範圍種類32
               If "" & rsOLD.Fields("FER" & Format(intP, "00")) <> "" Then
                  Combo1.ListIndex = Val("" & rsOLD.Fields("FER" & Format(intP, "00"))) - 1
               Else
                  Combo1.Text = ""
               End If
            Case Else
         End Select
      Next intP
      MsgBox "匯入記錄：" & rsOLD.Fields("FER01") & vbCrLf & vbCrLf & "      完畢！", vbInformation, "匯入記錄"
   End If
End Sub

Private Sub cmdExcel_Click()
Dim strField As String, strValue As String
Dim intP As Integer

   If TxtValidate = False Then Exit Sub
   
   Call Pub_ChkExcelPath
   strNowNo = ""
   strFileName = ""
   strTitle = ""
   strExc(1) = ""
   If txtDB(6) <> "" Then
      strExc(1) = strExc(1) & "+" & Mid(txtDB(6), 1, 9)
   End If
   If txtDB(7) <> "" Then
      strExc(1) = strExc(1) & "+" & Mid(txtDB(7), 1, 9)
   End If
   If txtDB(10) <> "" Then
      strExc(1) = strExc(1) & "+固定請款對象" & Mid(txtDB(10), 1, 9)
   End If
   strFileName = strExcelPath & strSrvDate(1) & "_" & Mid(strExc(1), 2)
   '加上備註
   If txtDB(35) <> "" Then
      strExc(1) = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(txtDB(35), "/", "／"), "\", "＼"), ":", "："), "*", "＊"), "?", "？"), "<", "＜"), ">", "＞"), "|", "｜"), Chr(34), "")
      strFileName = strFileName & "_" & strExc(1)
   End If
   strFileName = strFileName & MsgText(43)
   If PUB_ChkFileOpening(strFileName) = True Then
      Exit Sub
   End If
   If Dir(strFileName) <> "" Then
      Kill strFileName
   End If
   
   cmdExcel.Enabled = False
   Screen.MousePointer = vbHourglass
   If ProcExcel = True Then

On Error GoTo ErrHandle
      '因為有可能無資料，仍要寫記錄
      If Trim(txtDB(5) & txtDB(6) & txtDB(7) & txtDB(10)) <> "" Then
         '新增記錄
         For intP = 5 To maxField   'FCPEListRec: Table欄位數
            Select Case intP
               '直接帶入畫面的值
               Case 5, 6, 7, 10, 15, 16, 18, 19, 20, 23, 25, 26, 31, 33, 34, 35, 36, 37, 38, 40
                  strField = strField & ", FER" & Format(intP, "00")
                  strValue = strValue & ", " & CNULL(ChgSQL(Trim(txtDB(intP))))
               Case 8
                  '申請人條件1:限基本檔第一申請人(PA26)
                  If Check3(0).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 9
                  '申請人條件2:不同申請人分成不同Sheet
                  If Check3(1).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 11   '指定聯絡人: 去掉固定提示
                  strField = strField & ", FER" & Format(intP, "00")
                  strValue = strValue & ", " & CNULL(ChgSQL(Replace(txtDB(intP), contSpec, "")))
               Case 12
                  '案件清單
                  If Check1(0).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 13
                  '未發文清單
                  If Check1(1).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 14
                  '未請款清單
                  If Check1(2).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 17
                  '未收文清單(下一程序)
                  If Check1(3).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 21
                  '未付款之案件
                  If Check1(4).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 22
                  '行事曆清單
                  If Check1(5).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 24
                  '顯示日期格式: 1. 西元年(YYYY/MM/DD), 2. 民國年, 3. 西元年(DD.MM.YYYY)
                  If Trim(Combo2) <> "" Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", " & CNULL(Left(Combo2, 1))
                  End If
               Case 27
                  '含閉卷／銷卷
                  If Check2(0).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 28
                  '排除年費不續辦
                  If Check2(1).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 29
                  '只印已核准
                  If Check2(2).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 30
                  '只印未核准
                  If Check2(3).Value = 1 Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", 'Y'"
                  End If
               Case 32
                  '日期範圍種類
                  If Trim(Combo1) <> "" Then
                     strField = strField & ", FER" & Format(intP, "00")
                     strValue = strValue & ", " & CNULL(Left(Combo1, 1))
                  End If
               Case Else
            End Select
         Next intP
                  
         'Modified by Lydia 2025/01/17 改成模組
         'strSql = "select decode(mno,null,to_char(to_number(substr(to_char(sysdate,'yyyymmdd'),1,4))-1911)||'0001', to_number(mno)+1) mnowno from (" & _
         '         " select max(fer01) mno from fcpelistrec where fer03>=substr(to_char(sysdate,'yyyymmdd'),1,4)||'0100') "
         'intI = 1
         'strNowNo = ""
         ''PKEY：民國年+流水號4碼
         'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         'If intI = 1 Then
         '   strNowNo = RsTemp.Fields("mnowno")
         'End If
         strNowNo = GetNowFER01
         'end 2025/01/17
         If strNowNo <> "" Then
            intP = -1
            '流水號+Created ID, Date, Time
            strField = ", FER01, FER02, FER03, FER04" & strField
            strValue = ", " & CNULL(strNowNo) & ", " & CNULL(strUserNum) & ", to_char(sysdate,'yyyymmdd'), to_char(sysdate,'hh24miss')" & strValue
            strSql = "Insert Into FCPEListRec(" & Mid(strField, 2) & ") Values (" & Mid(strValue, 2) & ") "
            cnnConnection.Execute strSql, intP
         End If
      End If  'If Trim(txtDB(5) & txtDB(6) & txtDB(7) & txtDB(10)) <> "" Then
   End If
   
   If strFileName <> "" Then
      MsgBox "Excel檔案產生完成！檔案位置：" & strExcelPathN
   End If
   cmdExcel.Enabled = True
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , IIf(intP = -1, "新增記錄失敗", "")
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub cmdExit_Click()
   
   Unload Me
End Sub

Private Sub cmdItem_Click()
   Call doQueryItem
End Sub

Private Sub doQueryItem()
   
   strSql = "SELECT FEI01||FEI02 AS K01, FEI03 AS K02 FROM FCPELISTITEM WHERE 1=1"
   If Trim(Text1(2)) <> "" Then
      strSql = strSql & " AND INSTR(UPPER(FEI01||FEI02||FEI03),UPPER('" & ChgSQL(Trim(Text1(2))) & "')) > 0"
   End If
   strSql = strSql & " ORDER BY 1 desc"
   
   Me.Enabled = False
   intI = 0
   ListBox1(3).Clear
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         ListBox1(3).AddItem RsTemp.Fields("K01") & " " & RsTemp.Fields("K02"), 0
         RsTemp.MoveNext
      Loop
      ListBox1(3).ListIndex = 0
   End If
   Me.Enabled = True
End Sub

Private Sub cmdRemove_Click(Index As Integer)
   
   Call RemovelstNo(Index)
   
   LblCnt(Index) = ListBox1(Index).ListCount

End Sub

Private Sub cmdSearch_Click()

  Call frm060511_1.SetParent(Me, txtDB(5), txtDB(6), txtDB(7), txtDB(10))
  frm060511_1.Show
  Me.Hide
End Sub

Private Sub cmdUp_Click()
Dim intP As Integer, ii As Integer
Dim strTempB As String, tmpArr2 As Variant

   If ListBox1(2).ListCount > 0 Then
      For ii = 0 To ListBox1(2).ListCount - 1
         If ListBox1(2).Selected(ii) = True Then
            strTempB = strTempB & ";" & ListBox1(2).List(ii)
         End If
      Next ii
      If strTempB <> "" Then
         strTempB = Mid(strTempB, 2)
         tmpArr2 = Empty
         tmpArr2 = Split(strTempB, ";")
         For ii = 0 To UBound(tmpArr2)
            If Trim(tmpArr2(ii)) <> "" Then
               '考慮複數筆的移動後,原本記錄的位置有變=>全部改成用欄位值判斷現在位置
               intP = GetNowPos(2, Trim(tmpArr2(ii)))
               If intP = 0 Then
                  MsgBox "欄位已在第一筆!" & vbCrLf & tmpArr2(ii), vbInformation
                   ListBox1(2).Selected(intP) = True
               ElseIf intP > 0 Then
                  ListBox1(2).RemoveItem intP
                  ListBox1(2).AddItem tmpArr2(ii), intP - 1
                  ListBox1(2).Selected(intP - 1) = True
               End If
            End If
         Next ii
         txtDB(intL(2)).Text = ComposeList(2)
      End If
   End If
End Sub

Private Function ComposeList(p_idx As Integer) As String
Dim ii As Integer, strTempB As String

   If ListBox1(p_idx).ListCount = 0 Then
      txtDB(intL(p_idx)) = ""
   Else
      For ii = 0 To ListBox1(p_idx).ListCount - 1
         strTempB = strTempB & ";" & IIf(p_idx = 2, Trim(Left(ListBox1(p_idx).List(ii), 4)), Trim(Left(ListBox1(p_idx).List(ii), 10)))
      Next ii
      If strTempB <> "" Then
         strTempB = Mid(strTempB, 2)
      End If
      ComposeList = strTempB
   End If
   
End Function

Private Sub Combo1_Validate(Cancel As Boolean)

   If Combo1.Text <> "" Then
     If Val(Trim(Left(Combo1.Text, 1))) = 0 Or Val(Trim(Left(Combo1.Text, 1))) > Combo1.ListCount Then
        MsgBox "日期範圍請選擇！", vbExclamation, "檢核資料"
        Cancel = True
        Combo1.SetFocus
        Screen.MousePointer = Default
     Else
        Combo1.ListIndex = Val(Trim(Left(Combo1.Text, 1))) - 1
     End If
   End If
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   If Val(Trim(Left(Combo2.Text, 1))) = 0 Or Val(Trim(Left(Combo2.Text, 1))) > Combo2.ListCount Then
      MsgBox "請選擇顯示日期格式！" & vbCrLf & "1. 西元年(YYYY/MM/DD)" & vbCrLf & "2. 民國年" & vbCrLf & "3. 西元年(DD.MM.YYYY)", vbExclamation, "檢核資料"
      Cancel = True
      Combo2.SetFocus
      Screen.MousePointer = Default
   Else
      Combo2.ListIndex = Val(Trim(Left(Combo2.Text, 1))) - 1
   End If
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   
   Call ClearForm(True)
   Combo1.Clear
   Combo1.AddItem "1. 提申日期", 0
   Combo1.AddItem "2. 公告日期", 1
   Combo2.Clear
   Combo2.AddItem "1. 西元年(YYYY/MM/DD)", 0
   Combo2.AddItem "2. 民國年", 1
   Combo2.AddItem "3. 西元年(DD.MM.YYYY)", 2
   Combo2.ListIndex = 0
   
   ListBox1(0).Height = 720
   ListBox1(1).Height = 720
   ListBox1(2).Height = 1680
   ListBox1(3).Height = 1320
   
   intL(0) = 6 '代理人編號
   intL(1) = 7 '申請人編號
   intL(2) = 26 '案件清單顯示ITEM
   
   maxField = 37
   strSql = "select * from FCPEListRec where RowNum<2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      maxField = RsTemp.Fields.Count
   End If
   
   If Pub_StrUserSt03 <> "M51" Then
      cmd2.Visible = False
   Else
      cmd2.Visible = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsRead = Nothing
   Set rsQD = Nothing
   
   Set frm060511 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTempA As String
   Select Case Index
      Case 0  '代理人編號
         If Text1(Index).Tag <> Text1(Index).Text Then
            If Text1(Index).Text = "" Then
               lblFM2(1).Caption = ""
            Else
               Text1(Index) = ChangeCustomerL(Text1(Index))
               strTempA = GetFAgentName(Text1(Index))
               If strTempA <> "" Then
                  lblFM2(1).Caption = strTempA
               Else
                  lblFM2(1).Caption = ""
                  MsgBox "資料庫無資料 !", vbInformation
                  GoTo EXITSUB
               End If
            End If
         End If
         Text1(Index).Tag = Text1(Index).Text
      Case 1  '申請人編號
         If Text1(Index).Tag <> Text1(Index).Text Then
            If Text1(Index).Text = "" Then
               lblFM2(2).Caption = ""
            Else
               Text1(Index) = ChangeCustomerL(Text1(Index))
               strTempA = GetCustomerName(Text1(Index), "1")
               If strTempA <> "" Then
                  lblFM2(2).Caption = strTempA
               Else
                  lblFM2(2).Caption = ""
                  MsgBox "資料庫無資料 !", vbInformation
                  GoTo EXITSUB
               End If
            End If
         End If
         Text1(Index).Tag = Text1(Index).Text

   End Select
   Cancel = False
   Exit Sub
   
EXITSUB:
   Cancel = True
   Text1(Index).SetFocus
   Text1_GotFocus Index
End Sub

Private Sub txtDB_Change(Index As Integer)
   If Index = 11 Then
      PUB_RefreshText txtDB(Index)
   End If
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
End Sub

Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index <> 11 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtDB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 11 Then
      If Button = 2 Then Forms(0).PopupMenu2 txtDB(Index)
   End If
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
Dim strTempA As String
Dim tmpArr As Variant, intX As Integer

   If txtDB(Index).Text = "" Then Exit Sub
   
   Select Case Index
      Case 5  '需求人員
         If txtDB(Index).Tag <> txtDB(Index).Text Then
            strTempA = GetStaffName(txtDB(Index).Text, True)
            If strTempA = "" Then
               MsgBox "請輸入正確的員工編號!"
               GoTo EXITSUB
            Else
               lblFM2(0).Caption = strTempA
            End If
         End If
      Case 10  '固定請款對象
         If txtDB(Index).Tag <> txtDB(Index).Text Then
            If txtDB(Index).Text = "" Then
               lblFM2(3).Caption = ""
            Else
               txtDB(Index) = ChangeCustomerL(txtDB(Index))
               strTempA = ""
               Select Case Left(txtDB(Index), 1)
                  Case "Y": Call ClsPDGetAgent(txtDB(Index), strTempA)
                  Case "X": Call ClsPDGetCustomer(txtDB(Index), strTempA)
                  Case Else
                     MsgBox "請輸入代理人或申請人編號！"
               End Select
               If strTempA = "" Then
                  GoTo EXITSUB
               Else
                  lblFM2(3).Caption = strTempA
               End If
            End If
         End If
      Case 15, 16, 18, 19, 33, 34, 36, 37 '未請款之發文日範圍(15,16)、未收文之法限期間(18,19)、日期範圍(33,34)、未付款之請款期間(36,37)
         If CheckIsDate(txtDB(Index)) = False Then
            GoTo EXITSUB
         End If
      Case 20, 31, 38, 40 '20.未發文/未請款/未收文是否匯出同一工作表(Sheet)、31是否預估請款金額(Y:是，含C類來函), 38只列出有未付款之案件,40固定請款對象只有個案設定
         If txtDB(Index) <> "Y" Then
            MsgBox "請輸入Y or 空白!"
            GoTo EXITSUB
         End If
      Case 23 '欄位抬頭
         If txtDB(Index) <> "1" And txtDB(Index) <> "2" Then
            MsgBox "請輸入1 or 2 !"
            GoTo EXITSUB
         End If
      Case 25 '排序欄位
         If txtDB(Index) <> "" Then
            txtDB(Index) = Replace(txtDB(Index), " ", "")
            tmpArr = Split(txtDB(Index), ",")
            For intX = 0 To UBound(tmpArr)
               If Trim(tmpArr(intX)) <> "" Then
                  If Val(tmpArr(intX)) = 0 Or Val(tmpArr(intX)) > (Val(LblCnt(2)) + IIf(txtDB(20) = "Y", 1, 0)) Then
                     MsgBox "請輸入正確欄位的位置：" & tmpArr(intX)
                     GoTo EXITSUB
                  End If
               End If
            Next intX
         End If
   End Select
   
   If InStr("15,16,18,19,33,34,36,37", Format(Index, "00")) = 0 Then
      txtDB(Index).Tag = txtDB(Index).Text
   End If
   Exit Sub
   
EXITSUB:
   Cancel = True
   txtDB_GotFocus Index
   txtDB(Index).SetFocus
   Screen.MousePointer = vbDefault
End Sub

Private Sub AddlstNo(p_idx As Integer)
Dim intX As Integer, bFound As Boolean, ii As Integer
Dim strTempA As String, strNewItem As String, tmpArr1 As Variant

   bFound = True
   If p_idx = 0 Or p_idx = 1 Then
      If Trim(Text1(p_idx)) = "" Then
         Exit Sub
      Else
         Call Text1_Validate(p_idx, bFound)
         If bFound = True Then
            Exit Sub
         Else
            bFound = False
            strTempA = Trim(Text1(p_idx)) & " " & IIf(p_idx = 0, lblFM2(1).Caption, lblFM2(2).Caption)
         End If
      End If
   ElseIf p_idx = 2 Then
      If ListBox1(3).ListIndex = -1 Then
         MsgBox "無資料可供選取!"
         cmdItem.SetFocus
         Exit Sub
      Else
         For intX = 0 To ListBox1(3).ListCount - 1
            If ListBox1(3).Selected(intX) = True Then
              strTempA = strTempA & ";" & Trim(ListBox1(3).List(intX))
            End If
         Next intX
         If strTempA <> "" Then strTempA = Mid(strTempA, 2)
         bFound = False
      End If
   End If
'------------------------------------------
   If strTempA <> "" And bFound = False Then
      tmpArr1 = Empty
      tmpArr1 = Split(strTempA, ";")
      For intX = 0 To UBound(tmpArr1)
         If Trim(tmpArr1(intX)) <> "" Then
            strNewItem = IIf(p_idx = 2, Trim(Left(tmpArr1(intX), 4)), Trim(Left(tmpArr1(intX), 10)))
            If InStr(txtDB(intL(p_idx)), strNewItem) > 0 Then
               MsgBox Trim(tmpArr1(intX)) & "已存在於清單中！"
               cmdAdd(p_idx).SetFocus
               bFound = True
            End If
         End If
      Next intX
      If bFound = False Then
         For intX = 0 To UBound(tmpArr1)
            If Trim(tmpArr1(intX)) <> "" Then
               ii = ListBox1(p_idx).ListCount
               ListBox1(p_idx).AddItem Trim(tmpArr1(intX)), ii
               strNewItem = IIf(p_idx = 2, Trim(Left(tmpArr1(intX), 4)), Trim(Left(tmpArr1(intX), 10)))
               txtDB(intL(p_idx)) = txtDB(intL(p_idx)) & ";" & strNewItem
            End If
         Next intX
         If Mid(txtDB(intL(p_idx)), 1, 1) = ";" Then txtDB(intL(p_idx)) = Mid(txtDB(intL(p_idx)), 2)
         
         '清除來源
         If p_idx = 2 Then
            For intX = 0 To UBound(tmpArr1)
               If Trim(tmpArr1(intX)) <> "" Then
                  strNewItem = Trim(Left(tmpArr1(intX), 4))
                  ii = 0
                  For intI = 0 To ListBox1(3).ListCount - 1
                     If Trim(Left(ListBox1(3).List(ii), 4)) = strNewItem Then
                        ListBox1(3).RemoveItem ii
                        ii = ii - 1
                     End If
                     ii = ii + 1
                  Next intI
               End If
            Next intX
         Else
            Text1(p_idx) = ""
            Call Text1_Validate(p_idx, bFound)
         End If
      End If
   End If
End Sub

Private Sub RemovelstNo(p_idx As Integer)
Dim strTempA As String, ii As Integer

   If ListBox1(p_idx).ListCount > 0 Then  'ListBox1=>Form 2.0物件
      ii = 0
      Do While ii < ListBox1(p_idx).ListCount
         If ListBox1(p_idx).Selected(ii) = True Then
            strTempA = ListBox1(p_idx).List(ii)
            strExc(1) = IIf(p_idx = 2, Trim(Left(strTempA, 4)), Trim(Left(strTempA, 10)))
            txtDB(intL(p_idx)) = Replace(txtDB(intL(p_idx)), ";" & strExc(1), "")
            txtDB(intL(p_idx)) = Replace(txtDB(intL(p_idx)), strExc(1), "")
            If Left(txtDB(intL(p_idx)), 1) = ";" Then
               txtDB(intL(p_idx)) = Mid(txtDB(intL(p_idx)), 2)
            End If
            If Right(txtDB(intL(p_idx)), 1) = ";" Then
               txtDB(intL(p_idx)) = Mid(txtDB(intL(p_idx)), 1, Len(txtDB(intL(p_idx))) - 1)
            End If
            ListBox1(p_idx).RemoveItem ii
            If p_idx = 2 Then '新增到來源List的最下面
               ListBox1(3).AddItem strTempA, ListBox1(3).ListCount
            End If
            '因若屬性為單選時會自動選取上一個項目會導致全部被刪除，故移除後需將索引設為-1(無勾選)
            If ListBox1(p_idx).MultiSelect = 0 Then
               ListBox1(p_idx).ListIndex = -1
               Exit Do
            End If
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
      If txtDB(intL(p_idx)) = ";" Then txtDB(intL(p_idx)) = ""
   End If
End Sub

Private Function TxtValidate() As Boolean
Dim bolTmp As Boolean, intErr As Integer
   TxtValidate = False
   
   intErr = 5
   If Trim(txtDB(5)) = "" Or lblFM2(0) = "" Then
      MsgBox "需求人員不可空白!", vbExclamation, "檢核資料"
      GoTo EXITSUB
   Else
      Call txtDB_Validate(intErr, bolTmp)
      If bolTmp = True Then
         Exit Function
      End If
   End If
   If Trim(txtDB(6)) = "" And Trim(txtDB(7)) = "" And Trim(txtDB(10)) = "" Then
      MsgBox "代理人、申請人和固定請款對象編號不可皆為空白!", vbExclamation, "檢核資料"
      Text1(0).SetFocus
      Call Text1_GotFocus(0)
   End If
   If GetTextLength(txtDB(6)) > 50 Then
      MsgBox "代理人編號數量超過最大值！", vbExclamation, "檢核資料"
      Exit Function
   End If
   If GetTextLength(txtDB(6)) > 250 Then
      MsgBox "客戶編號數量超過最大值！", vbExclamation, "檢核資料"
      Exit Function
   End If
   'Added by Lydia 2025/06/24
   If Trim(txtDB(40)) <> "" And Trim(txtDB(10)) = "" Then
      MsgBox "只有個案設定請輸入固定請款對象！", vbExclamation, "檢核資料"
      Exit Function
   End If
   'end 2025/06/24
   intErr = 15
   If Trim(txtDB(15)) <> "" And Trim(txtDB(16)) <> "" And Trim(txtDB(15)) >= Trim(txtDB(16)) Then
      MsgBox "發文起始日期不可大於終止日期!", vbExclamation, "檢核資料"
      GoTo EXITSUB
   End If
   If (txtDB(15) <> "" Or txtDB(16) <> "") And Check1(2).Value = 0 Then
      MsgBox "未請款發文日期已輸入，但沒有勾選【未請款】的清單!", vbExclamation, "檢核資料"
      If txtDB(16) <> "" Then
         intErr = 15
      Else
         intErr = 16
      End If
      GoTo EXITSUB
   End If
   intErr = 31
   If Trim(txtDB(31)) <> "" And Check1(2).Value = 0 Then
      MsgBox "預估請款金額已輸入，但沒有勾選【未請款】的清單!", vbExclamation, "檢核資料"
      GoTo EXITSUB
   End If
   
   intErr = 18
   If Trim(txtDB(18)) <> "" And Trim(txtDB(19)) <> "" And Trim(txtDB(18)) >= Trim(txtDB(19)) Then
      MsgBox "法限起始日期不可大於終止日期!", vbExclamation, "檢核資料"
      GoTo EXITSUB
   End If
   If (txtDB(18) <> "" Or txtDB(19) <> "") And Check1(3).Value = 0 Then
      MsgBox "未收文法限日期已輸入，但沒有勾選【未收文】的清單!", vbExclamation, "檢核資料"
      If txtDB(18) <> "" Then
         intErr = 18
      Else
         intErr = 19
      End If
      GoTo EXITSUB
   End If
   
   
   intErr = 36
   If Trim(txtDB(36)) <> "" And Trim(txtDB(37)) <> "" And Trim(txtDB(36)) >= Trim(txtDB(37)) Then
      MsgBox "請款起始日期不可大於終止日期!", vbExclamation, "檢核資料"
      GoTo EXITSUB
   End If
   If (txtDB(36) <> "" Or txtDB(37) <> "") And Check1(4).Value = 0 Then
      MsgBox "未付款之請款日期已輸入，但沒有勾選【未付款】的清單!", vbExclamation, "檢核資料"
      If txtDB(36) <> "" Then
         intErr = 36
      Else
         intErr = 37
      End If
      GoTo EXITSUB
   End If
   intErr = 38
   If txtDB(38) <> "" And Check1(4).Value = 0 Then
      MsgBox "沒有勾選【未付款】的清單!", vbExclamation, "檢核資料"
      GoTo EXITSUB
   End If
   
   intErr = 23
   If Trim(txtDB(23)) = "" Then
      MsgBox "欄位抬頭不可空白!", vbExclamation, "檢核資料"
      GoTo EXITSUB
   Else
      Call txtDB_Validate(intErr, bolTmp)
      If bolTmp = True Then
         Exit Function
      End If
   End If
      
   If Trim(Combo2.Text) = "" Then
      MsgBox "顯示日期格式不可空白!", vbExclamation, "檢核資料"
      GoTo EXITSUB
   Else
      Call Combo2_Validate(bolTmp)
      If bolTmp = True Then
         Exit Function
      End If
   End If
   
   If Trim(txtDB(33)) <> "" Or Trim(txtDB(34)) <> "" Then
      If Left(Combo1.Text, 2) <> "1." And Left(Combo1.Text, 2) <> "2." Then
         MsgBox "日期範圍請選擇！", vbExclamation, "檢核資料"
         Combo1.SetFocus
         Exit Function
      Else
         intErr = 33
         If Trim(txtDB(33)) <> "" And Trim(txtDB(34)) <> "" And Trim(txtDB(33)) >= Trim(txtDB(34)) Then
            MsgBox "起始日期不可大於終止日期!", vbExclamation, "檢核資料"
            GoTo EXITSUB
         End If
      End If
   ElseIf Trim(Combo1) <> "" And Trim(txtDB(33)) = "" And Trim(txtDB(34)) = "" Then
         intErr = 33
         MsgBox "請輸入起始日期和終止日期!", vbExclamation, "檢核資料"
         GoTo EXITSUB
   End If
   If Trim(Combo1) <> "" And Check2(0).Value = 0 Then
     If MsgBox("指定" & Trim(Mid(Combo1.Text, 3)) & "：" & txtDB(33) & "~" & txtDB(34) & vbCrLf & "是否要含閉卷／銷卷？", vbInformation + vbYesNo + vbDefaultButton1, "檢核資料") = vbYes Then
        Check2(0).Value = 1
     End If
   End If
   
   intR = 0
   For Each oObj In Check1
      If oObj.Value = 1 Then
         intR = intR + 1
      End If
   Next
   If intR = 0 Then
      MsgBox "清單選項不可空白!", vbExclamation, "檢核資料"
      Exit Function
   End If
   
   If Trim(txtDB(26)) = "" And (Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1 Or Check1(3).Value = 1) Then
      MsgBox "案件清單顯示ITEM不可空白!", vbExclamation, "檢核資料"
      Exit Function
   End If
   
   intErr = 11
   If Trim(txtDB(11)) <> "" And txtDB(11) <> contSpec Then
      If GetTextLength(txtDB(11)) > 200 Then
         MsgBox "指定聯絡人字串超過最大值！", vbExclamation, "檢核資料"
         Exit Function
      End If
      If MsgBox("指定聯絡人：" & Replace(txtDB(11), ";", vbCrLf) & vbCrLf & vbCrLf & "是否加入條件？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
         GoTo EXITSUB
      End If
   End If
   
   '檢查週期性需求的日期範圍是否有修改
   If strPreNo <> "" And Trim(txtDB(5).Text & txtDB(6).Text & txtDB(7).Text & txtDB(10).Text) = Trim(strFER05 & strFER06 & strFER07 & strFER10) Then
      If Trim(txtDB(15).Tag & txtDB(16).Tag & txtDB(18).Tag & txtDB(19).Tag & txtDB(33).Tag & txtDB(34).Tag & txtDB(36).Tag & txtDB(37).Tag) <> "" _
          And Trim(txtDB(15).Text & txtDB(16).Text & txtDB(18).Text & txtDB(19).Text & txtDB(33).Text & txtDB(34).Text & txtDB(36).Text & txtDB(37).Text) = Trim(txtDB(15).Tag & txtDB(16).Tag & txtDB(18).Tag & txtDB(19).Tag & txtDB(33).Tag & txtDB(34).Tag & txtDB(36).Tag & txtDB(37).Tag) Then
         strExc(1) = ""
         If Trim(txtDB(15).Text & txtDB(16).Text) <> "" Then
            strExc(1) = strExc(1) & vbCrLf & "未請款之發文日期：" & IIf(Trim(txtDB(15)) = "", "(空白)", txtDB(15)) & " - " & IIf(Trim(txtDB(16)) = "", "(空白)", txtDB(16))
         End If
         If Trim(txtDB(18).Text & txtDB(19).Text) <> "" Then
            strExc(1) = strExc(1) & vbCrLf & "未收文之法定期限：" & IIf(Trim(txtDB(18)) = "", "(空白)", txtDB(18)) & " - " & IIf(Trim(txtDB(19)) = "", "(空白)", txtDB(19))
         End If
         If Trim(txtDB(36).Text & txtDB(37).Text) <> "" Then
            strExc(1) = strExc(1) & vbCrLf & "未付款之請款日期：" & IIf(Trim(txtDB(36)) = "", "(空白)", txtDB(36)) & " - " & IIf(Trim(txtDB(37)) = "", "(空白)", txtDB(37))
         End If
         If Trim(txtDB(33).Text & txtDB(34).Text) <> "" Then
            strExc(1) = strExc(1) & vbCrLf & Combo1.Text & "之日期範圍：" & IIf(Trim(txtDB(33)) = "", "(空白)", txtDB(33)) & " - " & IIf(Trim(txtDB(34)) = "", "(空白)", txtDB(34))
         End If
         If strExc(1) <> "" Then
            If MsgBox("匯入記錄：" & strPreNo & "，以下日期條件未變更：" & vbCrLf & strExc(1) & vbCrLf & vbCrLf & "請確認是否繼續作業？", vbExclamation + vbYesNo + vbDefaultButton2, "檢核資料") = vbNo Then
               Exit Function
            End If
         End If
      End If
   End If
   
   
   TxtValidate = True
   
   Exit Function
   
EXITSUB:
   If intErr > 0 Then
      txtDB(intErr).SetFocus
      txtDB_GotFocus intErr
   End If
End Function

'抓符合主要條件的案號(BP01~BP04)
Private Function ProcCaseNo(ByVal pFaList As String, ByVal pCuList As String, ByVal pCuType As String, ByVal pBCList As String, ByVal pContList As String, _
        Optional ByVal bolPA57 As Boolean = False, Optional ByVal bolNot605 As Boolean, Optional isPA16 As String = "") As String
Dim strConPA As String, strQ1 As String
Dim tmpArrA As Variant, intJ As Integer
Dim tmpArrB As Variant, intK As Integer
   
   strConPA = ""
   
   '代理人編號pFaList：多筆編號，用;做區隔
   If pFaList <> "" Then
      tmpArrA = Empty
      tmpArrA = Split(pFaList, ";")
      strQ1 = ""
      For intJ = 0 To UBound(tmpArrA)
         If Trim(tmpArrA(intJ)) <> "" Then
            If intJ = 0 Then
               If intJ = UBound(tmpArrA) Then
                  strConPA = strConPA & " AND PA75='" & Trim(tmpArrA(intJ)) & "'"
                  strQ1 = "Y"
               Else
                  strConPA = strConPA & " AND (PA75='" & Trim(tmpArrA(intJ)) & "' OR "
               End If
            Else
               strConPA = strConPA & "PA75='" & Trim(tmpArrA(intJ)) & "' OR "
            End If
         End If
      Next intJ
      If strQ1 <> "Y" Then
        strConPA = Mid(strConPA, 1, Len(strConPA) - 3) & ")"
      End If
   End If
   
   '申請人編號pCuList：多筆編號，用;做區隔 ; pCuType=1 限第一申請人
   If pCuList <> "" Then
      tmpArrA = Empty
      tmpArrA = Split(pCuList, ";")
      strQ1 = ""
      For intJ = 0 To UBound(tmpArrA)
         If Trim(tmpArrA(intJ)) <> "" Then
            If intJ = 0 Then
               If intJ = UBound(tmpArrA) Then
                  If pCuType = "1" Then '限第一申請人
                     strConPA = strConPA & " AND PA26='" & Trim(tmpArrA(intJ)) & "'"
                  Else
                     strConPA = strConPA & " AND INSTR(PA26||PA27||PA28||PA29||PA30,'" & Trim(tmpArrA(intJ)) & "')>0"
                  End If
                  strQ1 = "Y"
               Else
                  If pCuType = "1" Then '限第一申請人
                     strConPA = strConPA & " AND (PA26='" & Trim(tmpArrA(intJ)) & "' OR "
                  Else
                     strConPA = strConPA & " AND (INSTR(PA26||PA27||PA28||PA29||PA30,'" & Trim(tmpArrA(intJ)) & "')>0 OR "
                  End If
               End If
            Else
               If pCuType = "1" Then '限第一申請人
                  strConPA = strConPA & "PA26='" & Trim(tmpArrA(intJ)) & "' OR "
               Else
                  strConPA = strConPA & "INSTR(PA26||PA27||PA28||PA29||PA30,'" & Trim(tmpArrA(intJ)) & "')>0 OR "
               End If
            End If
         End If
      Next intJ
      If strQ1 <> "Y" Then
        strConPA = Mid(strConPA, 1, Len(strConPA) - 3) & ")"
      End If
   End If
   
   '固定請款對象編號pBCList：目前只有一筆編號,先用多筆編號的方式
   If pBCList <> "" Then
      tmpArrA = Empty
      tmpArrA = Split(pBCList, ";")
      strQ1 = ""
      For intJ = 0 To UBound(tmpArrA)
         If Trim(tmpArrA(intJ)) <> "" Then
            If intJ = 0 Then
               If intJ = UBound(tmpArrA) Then
                  'Modified by Lydia 2025/06/24 +只有個案設定=Y
                  'strConPA = strConPA & " AND INSTR(NVL(PA88, NVL(FA30, NVL(CU57, NVL(PA75, PA26)))),'" & Trim(tmpArrA(intJ)) & "')> 0"
                  strConPA = strConPA & " AND INSTR(" & IIf(txtDB(40) = "Y", "PA88||','", "NVL(PA88, NVL(FA30, NVL(CU57, NVL(PA75, PA26))))") & ",'" & Trim(tmpArrA(intJ)) & "')> 0"
                  strQ1 = "Y"
               Else
                  'Modified by Lydia 2025/06/24 +只有個案設定=Y
                  'strConPA = strConPA & " AND (INSTR(NVL(PA88, NVL(FA30, NVL(CU57, NVL(PA75, PA26)))),'" & Trim(tmpArrA(intJ)) & "')> 0 OR "
                  strConPA = strConPA & " AND (INSTR(" & IIf(txtDB(40) = "Y", "PA88||','", "NVL(PA88, NVL(FA30, NVL(CU57, NVL(PA75, PA26))))") & ",'" & Trim(tmpArrA(intJ)) & "')> 0 OR "
               End If
            Else
               'Modified by Lydia 2025/06/24 +只有個案設定=Y
               'strConPA = strConPA & "INSTR(NVL(PA88, NVL(FA30, NVL(CU57, NVL(PA75, PA26)))),'" & Trim(tmpArrA(intJ)) & "')> 0 OR "
               strConPA = strConPA & "INSTR(" & IIf(txtDB(40) = "Y", "PA88||','", "NVL(PA88, NVL(FA30, NVL(CU57, NVL(PA75, PA26))))") & ",'" & Trim(tmpArrA(intJ)) & "')> 0 OR "
            End If
         End If
      Next intJ
      If strQ1 <> "Y" Then
        strConPA = Mid(strConPA, 1, Len(strConPA) - 3) & ")"
      End If
   End If
   
   '指定聯絡人pContList：輸入姓名不分大小寫也不用加上稱謂，用空白或;區隔KeyWord
   If pContList <> "" Then
      pContList = Trim(Replace(pContList, "  ", " "))
      Do While Mid(pContList, 1, 1) = ";"
         pContList = Mid(pContList, 2)
      Loop
      tmpArrA = Empty
      tmpArrA = Split(pContList, ";")
      strQ1 = ""
      For intJ = 0 To UBound(tmpArrA)
         If Trim(tmpArrA(intJ)) <> "" Then
            tmpArrB = Empty
            tmpArrB = Split(Trim(tmpArrA(intJ)), " ")
            For intK = 0 To UBound(tmpArrB)
               If intJ = 0 And intK = 0 Then
                  If intK = UBound(tmpArrB) Then
                     strConPA = strConPA & " AND INSTR(LOWER(PA52||PA55||PA51||PA54||PA53||PA55||' '),LOWER('" & Trim(tmpArrB(intK)) & "')) >0"
                     strQ1 = "Y"
                  Else
                     strConPA = strConPA & " AND (INSTR(LOWER(PA52||PA55||PA51||PA54||PA53||PA55||' '),LOWER('" & Trim(tmpArrB(intK)) & "')) >0 OR "
                  End If
               Else
                  strConPA = strConPA & "INSTR(LOWER(PA52||PA55||PA51||PA54||PA53||PA55||' '),LOWER('" & Trim(tmpArrB(intK)) & "')) >0 OR "
               End If
            Next intK
         End If
      Next intJ
      If strQ1 <> "Y" Then
        strConPA = Mid(strConPA, 1, Len(strConPA) - 3) & ")"
      End If
   End If
   '是否含閉卷／銷卷bolPA57
   If bolPA57 = False Then
      strConPA = strConPA & " AND PA57||PA108 IS NULL "
   End If
   '排除年費不請款bolNot605
   If bolNot605 = True Then
      strConPA = strConPA & " AND NOT EXISTS(SELECT N1.* FROM NEXTPROGRESS N1 WHERE N1.NP02=PA01 AND N1.NP03=PA02 AND N1.NP04=PA03 AND N1.NP05=PA04 AND N1.NP06='N' " & _
                         "AND N1.NP09||N1.NP22 IN (SELECT MAX(N2.NP09||N2.NP22) FROM NEXTPROGRESS N2 WHERE N2.NP02=PA01 AND N2.NP03=PA02 AND N2.NP04=PA03 AND N2.NP05=PA04 AND N2.NP07='605' )) "
   End If
   'isPA16: 1-只印已核准, 2-只印未核准
   Select Case isPA16
      Case "1"
         strConPA = strConPA & " AND PA16='1' "
      Case "2"
         strConPA = strConPA & " AND PA16 IS NULL "
   End Select
   
   'Added by Lydia 2024/06/25 Y45622區分有無專利號之案號，但是FCP-063122在113/5/27,113/6/25處於已准但是未有專利號數
   If InStr(txtDB(35), "沒有專利號") > 0 Then
      strConPA = strConPA & " AND PA22 IS NULL "
   End If
   If InStr(txtDB(35), "有專利號") > 0 And InStr(txtDB(35), "沒有專利號") = 0 Then
      strConPA = strConPA & " AND PA22 IS NOT NULL "
   End If
   'end 2024/06/25
   
   '日期範圍種類：針對專利基本檔：1.提申日期、2.公告日期
   If Trim(Combo1) <> "" And Trim(txtDB(33) & txtDB(34)) <> "" Then
      '因為PCT案提申日也記錄在PA10，改用新案之代理人提申日來判斷
      If Left(Combo1, 2) = "1." Then
         'EX.P-131748 >>  AND ((PA01='P' AND CP47>=20230901 AND CP47<=20230930) OR (PA01<>'P' AND PA10>=20230901 AND PA10<=20230930))
          strQ1 = ""
         If txtDB(33) <> "" Then
            strQ1 = strQ1 & " AND CP47>=" & DBDATE(txtDB(33))
         End If
         If txtDB(34) <> "" Then
            strQ1 = strQ1 & " AND CP47<=" & DBDATE(txtDB(34))
         End If
         If strQ1 <> "" Then
           strConPA = strConPA & " AND ((PA01='P'" & strQ1 & ") OR (PA01<>'P'" & Replace(UCase(strQ1), "CP47", "PA10") & "))"
         End If
      ElseIf Left(Combo1, 2) = "2." Then
         If txtDB(33) <> "" Then
            strConPA = strConPA & " AND PA14>=" & DBDATE(txtDB(33))
         End If
         If txtDB(34) <> "" Then
            strConPA = strConPA & " AND PA14<=" & DBDATE(txtDB(34))
         End If
      End If
   End If
    
   strQ1 = "SELECT PA01 AS BP01, PA02 AS BP02, PA03 AS BP03, PA04 AS BP04,PA75 AS BP05, PA26 AS BP06,CP47 AS BP07 FROM PATENT, FAGENT, CUSTOMER, " & _
           "(SELECT CP01,CP02,CP03,CP04,CP47 FROM CASEPROGRESS WHERE CP31='Y' AND CP159=0) VTB1 " & _
           "WHERE SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
           "AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) " & strConPA

   ProcCaseNo = strQ1
   
   Exit Function
   
End Function

'******產生EXCEL SQL******
Private Function ProcExcel() As Boolean
Dim strTmp As String
Dim intA As Integer, TmpField As Variant
Dim strCon(0 To 5) As String
Dim strConEx1 As String, strConEx2 As String, strConEx3 As String
Dim strConGrp(0 To 5) As String
Dim strConITS As String 'Added by Lydia 2024/12/26 個案各項指示
Dim strConV(1 To 3) As String  'Added by Lydia 2024/12/26 案件清單之未付款、未發文、未請款(串聯為一個欄位)

   strTmp = ProcCaseNo(txtDB(6), txtDB(7), IIf(Check3(0).Value = 1, "1", ""), txtDB(10), Replace(txtDB(11), contSpec, ""), IIf(Check2(0).Value = 1, True, False), IIf(Check2(1).Value = 1, True, False), IIf(Check2(2).Value = 1, "1", IIf(Check2(3).Value = 1, "2", "")))
   intR = 1
   Set rsRead = ClsLawReadRstMsg(intR, strTmp)
   If intR = 1 Then
      Set rsQD = PUB_CreateRecordset(rsRead, , , , Me.Name, maxSeq)
      strConEx1 = ""
      Erase strCon
      TmpField = Empty
      TmpField = Split(txtDB(26), ";")
      For intA = 0 To UBound(TmpField)
         If Trim(TmpField(intA)) <> "" Then
            strTmp = "select * from FCPEListITEM where FEI01||FEI02=" & CNULL(Trim(TmpField(intA)))
            intR = 1
            Set rsRead = ClsLawReadRstMsg(intR, strTmp)
            If intR = 1 Then
               For intQ = 0 To 4 '0=案件清單,1=未發文,2=未請款,3=未收文,4=未付款
                  If Check1(intQ).Visible = True And Check1(intQ).Value = 1 Then
                     '取得欄位內容
                     strCon(intQ) = strCon(intQ) & ProcFieldValue(CStr(intQ), "" & rsRead.Fields("FEI07"), "" & rsRead.Fields("FEI08")) & " AS " & rsRead.Fields("FEI07")
                     strConGrp(intQ) = strConGrp(intQ) & ", " & rsRead.Fields("FEI07")
                     If InStr(strConEx1 & ";", Trim(TmpField(intA))) = 0 Then
                        '欄位抬頭用|串聯
                        strTitle = strTitle & "|" & IIf(txtDB(23) = "2", "" & rsRead.Fields("FEI05"), "" & rsRead.Fields("FEI06"))
                        strConEx1 = strConEx1 & Trim(TmpField(intA)) & ";"
                        'Added by Lydia 2024/12/26
                        If InStr("" & rsRead.Fields("FEI03"), "各項指示") > 0 Then
                           '去掉日期條件
                           strConITS = strConITS & IIf(strConITS <> "", ",", "") & Replace(rsRead.Fields("FEI07"), "W", "")
                        End If
                        If InStr("" & rsRead.Fields("FEI03"), "案件性質") > 0 And InStr("" & rsRead.Fields("FEI07"), "V") > 0 Then
                           '未付款之案件性質、未收文之案件性質、未請款之案件性質
                           strConV(Mid("" & rsRead.Fields("FEI07"), 2, 1)) = "Y"
                        End If
                        'end 2024/12/26
                     End If
                  End If
               Next intQ
            End If
         End If
      Next intA
      
      strTitle = Mid(strTitle, 2)
   Else
      MsgBox "查無案件資料！", vbInformation + vbOKOnly, "查詢結果"
      strFileName = ""
      ProcExcel = True
      Exit Function
   End If
   
   strConEx1 = ""
   '判斷有無實審
   strConEx1 = "select cp01,cp02,cp03,cp04,nvl(count(*),0) cnt1 from caseprogress where (cp01,cp02,cp03,cp04) in (select r001,r002,r003,r004 " & _
           " from rdatafactory where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "') and cp10 in ('1204','1215') and cp159=0 group by cp01,cp02,cp03,cp04"
   '抓目前下一程序(一旦銷閉卷就不顯示性質and pa57||pa108 is null)
   strConEx2 = "select np02,np03,np04,np05,np07,np08,np09,np23,decode(np07,'605',lastyear(pa72)+1) nyr,cpm03 as vcpm03,cpm04 as vcpm04,cpm10 as vcpm10" & _
            " from rdatafactory, patent, nextprogress, casepropertymap  where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
            " and pa57||pa108 is null and r001=np02(+) and r002=np03(+) and r003=np04(+) and r004=np05(+) and np02=cpm01(+) and np07=cpm02(+)" & _
            " and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+)" & _
            " and np06 is null and np02 in ('P','FCP') and not (np02 in ('P','PS','CFP','CPS','FCP','FG') and np07 in ('997','998','995','996','999','411','1204','1503')) "
   '抓最後一道收文
   strConEx3 = "select cp01 as lcp01,cp02 as lcp02,cp03 as lcp03,cp04 as lcp04,cp10 as lcp10,cp14 as lcp14 from caseprogress where cp09 in (" & _
            " select substr(maxno,9,9) mno1 from (select cp01,cp02,cp03,cp04,max(cp05||cp09) maxno from caseprogress where (cp01,cp02,cp03,cp04) in (" & _
            " select r001,r002,r003,r004  from rdatafactory where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "') and cp159=0 and cp09 <'D' group by cp01,cp02,cp03,cp04))"
   'Added by Lydia 2024/12/26
   If strConITS <> "" Then  '個案各項指示
      strConITS = " SELECT SUBSTR(ITS02,1,LENGTH(ITS02)-9) AS IPS01, SUBSTR(ITS02,LENGTH(ITS02)-8,6) AS IPS02, SUBSTR(ITS02,LENGTH(ITS02)-2,1) AS IPS03,SUBSTR(ITS02,LENGTH(ITS02)-1,2) AS IPS04,ITS04,ITS06" & _
               " From RDATAFACTORY, INSTRUCTIONS WHERE id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "' AND R001||R002||R003||R004=ITS02(+) AND ITS05 IS NULL AND ITS03 IN (" & GetAddStr(strConITS) & ")"
   End If
   If strConV(1) <> "" Then  '未付款之案件性質
      strConV(1) = " SELECT R001 AS V101 ,R002 AS V102,R003 AS V103 ,R004 AS V104," & _
            " LISTAGG(A1K01||'、'||SQLDATEW(A1K02+19110000)||'、'||A1K08,'；')  WITHIN GROUP (ORDER BY A1K01) AS V105" & _
            " FROM RDATAFACTORY,PATENT,ACC1K0 K WHERE ID='" & strUserNum & "' AND FORMNAME ='" & Me.Name & "' AND SEQNO='" & maxSeq & "'" & _
            " AND PA01(+)=R001 AND PA02(+)=R002 AND PA03(+)=R003 AND PA04(+)=R004 AND PA57||PA108 IS NULL AND A1K13(+)=PA01 AND A1K14(+)=PA02 AND A1K15(+)=PA03 AND A1K16(+)=PA04 AND NVL(A1K12,0)=0 AND A1K25 IS NULL AND A1K17 IS NULL" & _
            " AND A1K29 IS NULL AND A1K01 IS NOT NULL GROUP BY R001,R002,R003,R004"
   End If
   If strConV(2) <> "" Then  '未發文之案件性質
      strConV(2) = " SELECT PA01 AS V201 ,PA02 AS V202,PA03 AS V203 ,PA04 AS V204," & _
            " LISTAGG(DECODE(PA09,'000',CPM03,CPM04),'、')  WITHIN GROUP (ORDER BY CP10) AS V205" & _
            " FROM RDATAFACTORY,PATENT,CASEPROGRESS,CASEPROPERTYMAP M1 WHERE ID='" & strUserNum & "' AND FORMNAME ='" & Me.Name & "' AND SEQNO='" & maxSeq & "'" & _
            " AND PA01(+)=R001 AND PA02(+)=R002 AND PA03(+)=R003 AND PA04(+)=R004 AND PA57||PA108 IS NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP158=0 AND CP159=0 AND CP09 <'D' " & _
            " AND CPM01(+)=CP01 AND CPM02(+)=CP10 GROUP BY PA01,PA02,PA03,PA04"
   End If
   If strConV(3) <> "" Then  '未請款之案件性質:含C類來函預估請款金額
      strConV(3) = " SELECT PA01 AS V301 ,PA02 AS V302,PA03 AS V303 ,PA04 AS V304," & _
            " LISTAGG(DECODE(PA09,'000',CPM03,CPM04),'、')  WITHIN GROUP (ORDER BY CP10) AS V305" & _
            " FROM RDATAFACTORY,PATENT,CASEPROGRESS,CASEPROPERTYMAP M1 WHERE ID='" & strUserNum & "' AND FORMNAME ='" & Me.Name & "' AND SEQNO='" & maxSeq & "'" & _
            " AND PA01(+)=R001 AND PA02(+)=R002 AND PA03(+)=R003 AND PA04(+)=R004 AND PA57||PA108 IS NULL AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
            " AND (CP27>20030000 AND CP20||CP57||CP60 IS NULL AND CP10<>'901' AND CP10<>'902' AND CP10<>'935' AND CP09<>'A99017314' AND ( CP01<>'P' OR CP61 IS NOT NULL OR CP09<'B'" & _
            " OR (CP01='P' AND EXISTS(SELECT * FROM CASEPROPERTYMAP M2 WHERE CPM01='FCP' AND M2.CPM02=CP10 AND M2.CPM18>0))) ) GROUP BY PA01,PA02,PA03,PA04"
   End If
   'end 2024/12/26
   intA = 0 '記錄WorkSheet量
   '0=案件清單
   strConSql(0) = ""
   If strCon(0) <> "" Then
      intA = intA + 1
      'vtb1>>判斷有無實審, vtb2>>抓目前下一程序(一旦銷閉卷就不顯示性質and pa57||pa108 is null), vtb3>>抓最後一道收文
      strConSql(0) = " select " & Mid(strCon(0), 2) & IIf(Check3(1).Value = 1, ", PA26 AS GRPNO", "") & _
               " from rdatafactory,patent,fagent f1 ,customer c1,nation pn1,nation fn1, nation cn1, caseprogress vc1, caseprogress vc2, PatentYearFee," & _
               " (select USXR02 from USXRATE,(SELECT max(USXR01) LstDate FROM USXRATE WHERE USXR01<=to_char(sysdate,'yyyymmdd')-19110000) where USXR01=LstDate)" & _
               " ,(" & strConEx1 & ") vtb1 ,(" & strConEx2 & ") vtb2 ,(" & strConEx3 & ") vtb3"
      'Added by Lydia 2024/12/26
      '個案各項指示
      If strConITS <> "" Then strConSql(0) = strConSql(0) & ",(" & strConITS & ") VITS01"
      '未付款之案件性質
      If strConV(1) <> "" Then strConSql(0) = strConSql(0) & ",(" & strConV(1) & ") V1"
      '未收文之案件性質
      If strConV(2) <> "" Then strConSql(0) = strConSql(0) & ",(" & strConV(2) & ") V2"
      '未請款之案件性質
      If strConV(3) <> "" Then strConSql(0) = strConSql(0) & ",(" & strConV(3) & ") V3"
      'end 2024/12/26
      strConSql(0) = strConSql(0) & " where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
               " and r001=pa01(+) and r002=pa02(+) and r003=pa03(+) and r004=pa04(+) and pa01 is not null" & _
               " and vtb1.cp01(+)=pa01 and vtb1.cp02(+)=pa02 and vtb1.cp03(+)=pa03 and vtb1.cp04(+)=pa04 and vtb2.np02(+)=pa01 and vtb2.np03(+)=pa02 and vtb2.np04(+)=pa03 and vtb2.np05(+)=pa04" & _
               " and vtb3.lcp01(+)=pa01 and vtb3.lcp02(+)=pa02 and vtb3.lcp03(+)=pa03 and vtb3.lcp04(+)=pa04" & _
               " and vc1.cp01(+)=pa01 and vc1.cp02(+)=pa02 and vc1.cp03(+)=pa03 and vc1.cp04(+)=pa04 and vc1.cp10(+)='416'" & _
               " and vc2.cp01(+)=pa01 and vc2.cp02(+)=pa02 and vc2.cp03(+)=pa03 and vc2.cp04(+)=pa04 and vc2.cp10(+)='435'" & _
               " and f1.fa01(+)=substr(pa75,1,8) and f1.fa02(+)=substr(pa75,9,1)" & _
               " and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9,1)" & _
               " and pa09=pn1.na01(+) and fn1.na01(+) = f1.fa10 and cn1.na01(+)=c1.cu10" & _
               " and yf01(+)=pa09 and yf02(+)=pa08 and yf03(+)='Y00000000' and yf04(+)=np07 and yf05(+)=nYr"
      'Added by Lydia 2024/12/26
      '個案各項指示
      If strConITS <> "" Then strConSql(0) = strConSql(0) & " AND R001=IPS01(+) AND R002=IPS02(+) AND R003=IPS03(+) AND R004=IPS04(+) "
      '未付款之案件性質
      If strConV(1) <> "" Then strConSql(0) = strConSql(0) & " AND R001=V101(+) AND R002=V102(+) AND R003=V103(+) AND R004=V104(+) "
      '未收文之案件性質
      If strConV(2) <> "" Then strConSql(0) = strConSql(0) & " AND R001=V201(+) AND R002=V202(+) AND R003=V203(+) AND R004=V204(+) "
      '未請款之案件性質
      If strConV(3) <> "" Then strConSql(0) = strConSql(0) & " AND R001=V301(+) AND R002=V302(+) AND R003=V303(+) AND R004=V304(+) "
      'end 2024/12/26
      '因為下一程序有多筆，但又不顯示下一程序資料
      strConSql(0) = "SELECT " & Mid(strConGrp(0), 2) & IIf(Check3(1).Value = 1, ", GRPNO", "") & " FROM (" & strConSql(0) & ") GROUP BY " & Mid(strConGrp(0), 2) & IIf(Check3(1).Value = 1, ", GRPNO", "")
      strConSql(0) = strConSql(0) & " order by " & IIf(Check3(1).Value = 1, "GRPNO,", "") & IIf(txtDB(25) <> "", txtDB(25), IIf(Val(LblCnt(2)) < 2, "1", "1, 2 "))
   End If
   '4=未付款>>比照案件清單+acc1k0
   strConSql(4) = ""
   If strCon(4) <> "" Then
      intA = intA + 1
      'vtb1>>判斷有無實審, vtb2>>抓目前下一程序(一旦銷閉卷就不顯示性質and pa57||pa108 is null), vtb3>>抓最後一道收文
      '改用Y54686的未付款
      'strConSql(4) = " select " & Mid(strCon(4), 2) & IIf(Check3(1).Value = 1, ", PA26 AS GRPNO", "") & _
               " from rdatafactory,patent,fagent f1 ,customer c1 ,nation pn1,nation fn1, nation cn1, caseprogress vc1, caseprogress vc2, PatentYearFee," & _
               " (select k.* from rdatafactory,acc1k0 k where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
               " and a1k13(+)=r001 and a1k14(+)=r002 and a1k15(+)=r003 and a1k16(+)=r004" & _
               " and a1k29 is null and nvl(a1k12,0)=0 and a1k25 is null and a1k17 is null) vtby, (select USXR02 from USXRATE,(SELECT max(USXR01) LstDate FROM USXRATE WHERE USXR01<=to_char(sysdate,'yyyymmdd')-19110000) where USXR01=LstDate)" & _
               " ,(" & strConEx1 & ") vtb1 ,(" & strConEx2 & ") vtb2 ,(" & strConEx3 & ") vtb3"
      strConSql(4) = " select " & Mid(strCon(4), 2) & IIf(Check3(1).Value = 1, ", PA26 AS GRPNO", "") & _
               " from rdatafactory,patent,fagent f1 ,customer c1 ,nation pn1,nation fn1, nation cn1, caseprogress vc1, caseprogress vc2, PatentYearFee," & _
               " (SELECT A1K01,A1K02,A1K13,A1K14,A1K15,A1K16,A1K08,A1K18,A1K11" & _
               " ,TRUNC(SUM(DECODE(SUBSTR(A1L04,-2),'99',0,A1L05-A1L07) / A1K10)) A1K_FEE" & _
               " ,TRUNC(SUM(DECODE(SUBSTR(A1L04,-2),'99',A1L05-A1L07,0) / A1K10)) A1K_EXP" & _
               " ,SUBSTR(MIN(A1L02||NVL(A1J16,A1J03)),4) A1K_ITEM from rdatafactory,acc1k0,acc1l0,acc1j0 where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
               " and a1k13(+)=r001 and a1k14(+)=r002 and a1k15(+)=r003 and a1k16(+)=r004" & _
               " and a1k29 is null and nvl(a1k12,0)=0 and a1k25 is null and a1k17 is null and a1l01(+)=a1k01 and a1j01(+)=a1l03 and a1j02(+)=a1l04" & IIf(Trim(txtDB(36)) <> "", " AND A1K02>=" & TransDate(txtDB(36), 1), "") & IIf(Trim(txtDB(37)) <> "", " AND A1K02<=" & TransDate(txtDB(37), 1), "") & _
               " group by a1k01,a1k02,a1k13,a1k14,a1k15,a1k16,a1k08,a1k18,a1k11) vtby, (select USXR02 from USXRATE,(SELECT max(USXR01) LstDate FROM USXRATE WHERE USXR01<=to_char(sysdate,'yyyymmdd')-19110000) where USXR01=LstDate)" & _
               " ,(" & strConEx1 & ") vtb1 ,(" & strConEx2 & ") vtb2 ,(" & strConEx3 & ") vtb3"
      strConSql(4) = strConSql(4) & " where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
               " and r001=pa01(+) and r002=pa02(+) and r003=pa03(+) and r004=pa04(+) and pa01 is not null" & _
               " and vtb1.cp01(+)=pa01 and vtb1.cp02(+)=pa02 and vtb1.cp03(+)=pa03 and vtb1.cp04(+)=pa04 and vtb2.np02(+)=pa01 and vtb2.np03(+)=pa02 and vtb2.np04(+)=pa03 and vtb2.np05(+)=pa04" & _
               " and vtb3.lcp01(+)=pa01 and vtb3.lcp02(+)=pa02 and vtb3.lcp03(+)=pa03 and vtb3.lcp04(+)=pa04" & _
               " and vc1.cp01(+)=pa01 and vc1.cp02(+)=pa02 and vc1.cp03(+)=pa03 and vc1.cp04(+)=pa04 and vc1.cp10(+)='416'" & _
               " and vc2.cp01(+)=pa01 and vc2.cp02(+)=pa02 and vc2.cp03(+)=pa03 and vc2.cp04(+)=pa04 and vc2.cp10(+)='435'" & _
               " and f1.fa01(+)=substr(pa75,1,8) and f1.fa02(+)=substr(pa75,9,1)" & _
               " and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9,1)" & _
               " and pa09=pn1.na01(+) and fn1.na01(+) = f1.fa10 and cn1.na01(+)=c1.cu10" & _
               " and yf01(+)=pa09 and yf02(+)=pa08 and yf03(+)='Y00000000' and yf04(+)=np07 and yf05(+)=nYr" & _
               " and a1k13(+)=pa01 and a1k14(+)=pa02 and a1k15(+)=pa03 and a1k16(+)=pa04 " & IIf(txtDB(38) = "Y", " and a1k01 is not null", "")
      '因為下一程序有多筆，但又不顯示下一程序資料
      strConSql(4) = "SELECT " & Mid(strConGrp(4), 2) & IIf(Check3(1).Value = 1, ", GRPNO", "") & " FROM (" & strConSql(4) & ") GROUP BY " & Mid(strConGrp(4), 2) & IIf(Check3(1).Value = 1, ", GRPNO", "")
      strConSql(4) = strConSql(4) & " order by " & IIf(Check3(1).Value = 1, "GRPNO,", "") & IIf(txtDB(25) <> "", txtDB(25), IIf(Val(LblCnt(2)) < 2, "1", "1, 2 "))
   End If
   '1-未發文
   strConSql(1) = ""
   If strCon(1) <> "" Then
      If txtDB(20) = "" Then
         intA = intA + 1
      End If
      'vtb1>>判斷有無實審, vtb2>>抓目前下一程序(一旦銷閉卷就不顯示性質and pa57||pa108 is null)
      strConSql(1) = " select " & Mid(strCon(1), 2) & ",'未發文' as 狀態, a.cp27 as ord1 " & IIf(Check3(1).Value = 1, ", PA26 AS GRPNO", "") & _
               " from rdatafactory,patent,fagent f1 ,customer c1 ,nation pn1,nation fn1, nation cn1, caseprogress a, casepropertymap m1" & _
               ",(" & strConEx1 & ") vtb1 ,(" & strConEx2 & ") vtb2"
      strConGrp(1) = Mid(strConGrp(1), 2) & ",狀態,ord1 " & IIf(Check3(1).Value = 1, ",GRPNO", "")
      'Modified by Lydia 2025/03/17 and a.cp27 is null>> and a.cp158=0 and a.cp159=0
      strConSql(1) = strConSql(1) & " where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
               " and r001=pa01(+) and r002=pa02(+) and r003=pa03(+) and r004=pa04(+) and pa01 is not null" & _
               " and vtb1.cp01(+)=pa01 and vtb1.cp02(+)=pa02 and vtb1.cp03(+)=pa03 and vtb1.cp04(+)=pa04 and vtb2.np02(+)=pa01 and vtb2.np03(+)=pa02 and vtb2.np04(+)=pa03 and vtb2.np05(+)=pa04" & _
               " and f1.fa01(+)=substr(pa75,1,8) and f1.fa02(+)=substr(pa75,9,1)" & _
               " and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9,1)" & _
               " and pa09=pn1.na01(+) and fn1.na01(+) = f1.fa10 and cn1.na01(+)=c1.cu10" & _
               " and a.cp01(+)=pa01 and a.cp02(+)=pa02 and a.cp03(+)=pa03 and a.cp04(+)=pa04 and a.cp158=0 and a.cp159=0 and a.cp09 is not null and a.cp09 < 'D'" & _
               " and m1.cpm01(+)=a.cp01 and m1.cpm02(+)=a.cp10"
      If txtDB(20) <> "Y" Then
         '因為下一程序有多筆，但又不顯示下一程序資料
         strConSql(1) = "SELECT " & strConGrp(1) & IIf(Check3(1).Value = 1, ", GRPNO", "") & " FROM (" & strConSql(1) & ") GROUP BY " & strConGrp(1) & IIf(Check3(1).Value = 1, ", GRPNO", "")
         strConSql(1) = strConSql(1) & " order by " & IIf(Check3(1).Value = 1, "GRPNO,", "") & IIf(txtDB(25) <> "", txtDB(25), "1, ord1")
      End If
   End If
   '2-未請款
   strConSql(2) = ""
   If strCon(2) <> "" Then
      If txtDB(20) = "" Then
         intA = intA + 1
      End If
      'vtb1>>判斷有無實審, vtb2>>抓目前下一程序(一旦銷閉卷就不顯示性質and pa57||pa108 is null)
      strConSql(2) = " select " & Mid(strCon(2), 2) & ",'未請款' as 狀態, a.cp27 as ord1 " & IIf(Check3(1).Value = 1, ", PA26 AS GRPNO", "") & _
               " from rdatafactory,patent,fagent f1 ,customer c1 ,nation pn1, nation fn1, nation cn1, caseprogress a, casepropertymap m1" & _
               ",(" & strConEx1 & ") vtb1,(" & strConEx2 & ") vtb2"
      strConGrp(2) = Mid(strConGrp(2), 2) & ",狀態,ord1 " & IIf(Check3(1).Value = 1, ",GRPNO", "")
      'X48279案件&請款：排除P-95132的其他>>and a.cp09<>'A99017314'
      'Modified by Lydia 2024/11/26 +年費,規費
      'strConSql(2) = strConSql(2) & " where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
               " and r001=pa01(+) and r002=pa02(+) and r003=pa03(+) and r004=pa04(+) and pa01 is not null" & _
               " and vtb1.cp01(+)=pa01 and vtb1.cp02(+)=pa02 and vtb1.cp03(+)=pa03 and vtb1.cp04(+)=pa04 and vtb2.np02(+)=pa01 and vtb2.np03(+)=pa02 and vtb2.np04(+)=pa03 and vtb2.np05(+)=pa04" & _
               " and f1.fa01(+)=substr(pa75,1,8) and f1.fa02(+)=substr(pa75,9,1)" & _
               " and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9,1)" & _
               " and pa09=pn1.na01(+) and fn1.na01(+) = f1.fa10 and cn1.na01(+)=c1.cu10" & _
               " and a.cp01(+)=pa01 and a.cp02(+)=pa02 and a.cp03(+)=pa03 and a.cp04(+)=pa04 and a.cp20||a.cp57||a.cp60 is null and a.cp10<>'901' and a.cp10<>'902'" & _
               " and a.cp09<>'A99017314'"
      'Modified by Lydia 2025/03/31 +排除935案件轉至本所
      strConSql(2) = strConSql(2) & " ,PatentYearFee, (select USXR02 from USXRATE,(SELECT max(USXR01) LstDate FROM USXRATE WHERE USXR01<=to_char(sysdate,'yyyymmdd')-19110000) where USXR01=LstDate) " & _
               " where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
               " and r001=pa01(+) and r002=pa02(+) and r003=pa03(+) and r004=pa04(+) and pa01 is not null" & _
               " and vtb1.cp01(+)=pa01 and vtb1.cp02(+)=pa02 and vtb1.cp03(+)=pa03 and vtb1.cp04(+)=pa04 and vtb2.np02(+)=pa01 and vtb2.np03(+)=pa02 and vtb2.np04(+)=pa03 and vtb2.np05(+)=pa04" & _
               " and f1.fa01(+)=substr(pa75,1,8) and f1.fa02(+)=substr(pa75,9,1)" & _
               " and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9,1)" & _
               " and pa09=pn1.na01(+) and fn1.na01(+) = f1.fa10 and cn1.na01(+)=c1.cu10" & _
               " and a.cp01(+)=pa01 and a.cp02(+)=pa02 and a.cp03(+)=pa03 and a.cp04(+)=pa04 and a.cp20||a.cp57||a.cp60 is null and a.cp10<>'901' and a.cp10<>'902' and a.cp10<>'935'" & _
               " and a.cp09<>'A99017314' and yf01(+)=pa09 and yf02(+)=pa08 and yf03(+)='Y00000000' and yf04(+)=np07 and yf05(+)=nYr"
               
      '區分是否預估請款金額(Y:是，含C類來函)
      strConSql(2) = strConSql(2) & " and a.cp27>20030000 " & IIf(txtDB(31) = "Y", " ", "and (a.cp09<'B' or ( a.cp09>'B' and a.cp10='204' and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10='1225')))") & _
                " and ( a.cp01<>'P' or a.cp61 is not null or a.cp09<'B'" & _
                " or (a.cp01='P' and exists(select * from casepropertymap m2 where cpm01='FCP' and m2.cpm02=cp10 and m2.cpm18>0)))" & _
                IIf(Trim(txtDB(15)) = "", "", " and a.cp27>=" & DBDATE(txtDB(15))) & IIf(Trim(txtDB(16)) = "", "", " and a.cp27<=" & DBDATE(txtDB(16))) & _
                " and m1.cpm01(+)=a.cp01 and m1.cpm02(+)=a.cp10"
      If txtDB(20) <> "Y" Then
         '因為下一程序有多筆，但又不顯示下一程序資料
         strConSql(2) = "SELECT " & strConGrp(2) & IIf(Check3(1).Value = 1, ", GRPNO", "") & " FROM (" & strConSql(2) & ") GROUP BY " & strConGrp(2) & IIf(Check3(1).Value = 1, ", GRPNO", "")
         strConSql(2) = strConSql(2) & " order by " & IIf(Check3(1).Value = 1, "GRPNO,", "") & IIf(txtDB(25) <> "", txtDB(25), "1, ord1")
      End If
   End If
   '3-未收文
   strConSql(3) = ""
   If strCon(3) <> "" Then
      If txtDB(20) = "" Then
         intA = intA + 1
      End If
      'vtb1>>判斷有無實審, 目前下一程序(一旦銷閉卷就不顯示性質and pa57||pa108 is null)
      strConSql(3) = " select " & Mid(strCon(3), 2) & ",'未收文' as 狀態, np09 as ord1 " & IIf(Check3(1).Value = 1, ", PA26 AS GRPNO", "") & _
                  " from (select np02,np03,np04,np05,np07,np08,np09,np23,np10,decode(np07,'605',lastyear(pa72)+1) nYr" & _
                  " from rdatafactory,nextprogress,patent where id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
                  " and pa01(+)=r001 and pa02(+)=r002 and pa03(+)=r003 and pa04(+)=r004 and pa01 is not null and np02(+)=pa01 and np03(+)=pa02 and np04(+)=pa03 and np05(+)=pa04" & _
                  " and pa57||pa108 is null and np06 is null and np02 in ('P','FCP')" & _
                  " and not (np02 in ('P','PS','CFP','CPS','FCP','FG')" & _
                  " and np07 IN ('997','998','995','996','999','411','1204','1503'))" & _
                  IIf(Trim(txtDB(18)) = "", "", " and np09>=" & DBDATE(txtDB(18))) & IIf(Trim(txtDB(19)) = "", "", " and np09<=" & DBDATE(txtDB(19))) & _
                  " ) N,patent,fagent f1 ,customer c1,nation pn1,nation fn1, nation cn1, casepropertymap m1" & _
                  " ,(" & strConEx1 & ") vtb1"
      strConGrp(3) = Mid(strConGrp(3), 2) & ",狀態,ord1 " & IIf(Check3(1).Value = 1, ",GRPNO", "")
      'Modified by Lydia 2024/11/26 +年費,規費
      'strConSql(3) = strConSql(3) & " where np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and pa01 is not null" & _
                  " and vtb1.cp01(+)=pa01 and vtb1.cp02(+)=pa02 and vtb1.cp03(+)=pa03 and vtb1.cp04(+)=pa04" & _
                  " and f1.fa01(+)=substr(pa75,1,8) and f1.fa02(+)=substr(pa75,9,1)" & _
                  " and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9,1)" & _
                  " and pa09=pn1.na01(+) and fn1.na01(+) = f1.fa10 and cn1.na01(+)=c1.cu10" & _
                  " and m1.cpm01(+)=np02 and m1.cpm02(+)=np07"
      strConSql(3) = strConSql(3) & ",PatentYearFee, (select USXR02 from USXRATE,(SELECT max(USXR01) LstDate FROM USXRATE WHERE USXR01<=to_char(sysdate,'yyyymmdd')-19110000) where USXR01=LstDate) " & _
                  " where np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and pa01 is not null" & _
                  " and vtb1.cp01(+)=pa01 and vtb1.cp02(+)=pa02 and vtb1.cp03(+)=pa03 and vtb1.cp04(+)=pa04" & _
                  " and f1.fa01(+)=substr(pa75,1,8) and f1.fa02(+)=substr(pa75,9,1)" & _
                  " and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9,1)" & _
                  " and pa09=pn1.na01(+) and fn1.na01(+) = f1.fa10 and cn1.na01(+)=c1.cu10" & _
                  " and m1.cpm01(+)=np02 and m1.cpm02(+)=np07 " & _
                  " and yf01(+)=pa09 and yf02(+)=pa08 and yf03(+)='Y00000000' and yf04(+)=np07 and yf05(+)=nYr"
      If txtDB(20) <> "Y" Then
         '因為下一程序有多筆，但又不顯示下一程序資料
         strConSql(3) = "SELECT " & strConGrp(3) & IIf(Check3(1).Value = 1, ", GRPNO", "") & " FROM (" & strConSql(3) & ") GROUP BY " & strConGrp(3) & IIf(Check3(1).Value = 1, ", GRPNO", "")
         strConSql(3) = strConSql(3) & " order by " & IIf(Check3(1).Value = 1, "GRPNO,", "") & IIf(txtDB(25) <> "", txtDB(25), "1, ord1")
      End If
   End If
   '未發文/未請款/未收文是否匯出同一工作表
   strTmp = ""
   If txtDB(20) = "Y" And strConSql(1) & strConSql(2) & strConSql(3) <> "" Then
      strTitle = strTitle & "|狀態"
      intA = intA + 1
      strSql = IIf(strConSql(1) <> "", " Union " & strConSql(1), "") & IIf(strConSql(2) <> "", " Union " & strConSql(2), "") & IIf(strConSql(3) <> "", " Union " & strConSql(3), "")
      strSql = Mid(strSql, 7)
      '因為下一程序有多筆，但又不顯示下一程序資料
      strExc(1) = ""
      If strConSql(1) <> "" Then
         strExc(1) = "1"
      ElseIf strConSql(2) <> "" Then
         strExc(1) = "2"
      ElseIf strConSql(3) <> "" Then
         strExc(1) = "3"
      End If
      strSql = "SELECT " & strConGrp(Val(strExc(1))) & IIf(Check3(1).Value = 1, ", GRPNO", "") & " FROM (" & strSql & ") GROUP BY " & strConGrp(Val(strExc(1))) & IIf(Check3(1).Value = 1, ", GRPNO", "")
      strSql = strSql & " order by " & IIf(Check3(1).Value = 1, "GRPNO,", "") & IIf(txtDB(25) <> "", txtDB(25), "1, ord1")
      strTmp = strSql '傳入合併語法
   End If

   '行事曆：固定欄位
   strConSql(5) = ""
   If Check1(5).Value = 1 Then
       intA = intA + 1
       strConSql(5) = " SELECT PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) AS 本所案號, SC04 AS 事由,SQLDATET(SC01) AS 管制日期" & _
                   " ,PA75 AS 代理人編號, PA26 AS 申請人編號 From RDATAFACTORY, PATENT, STAFF_CALENDAR" & _
                   " WHERE id='" & strUserNum & "' and formname ='" & Me.Name & "' and seqno='" & maxSeq & "'" & _
                   " AND PA01(+)=R001 AND PA02(+)=R002 AND PA03(+)=R003 AND PA04(+)=R004 AND PA01 IS NOT NULL" & _
                   " AND SC05(+)=PA01 AND SC06(+)=PA02 AND SC07(+)=PA03 AND SC08(+)=PA04 AND SC01>0 AND SC18 IS NULL" & _
                   " ORDER BY 3,1"
   End If
   
   If ProcExcelSave(intA, strTmp) = True Then
      ProcExcel = True
   Else
      Exit Function
   End If
   
   
End Function

Private Function ProcFieldValue(ByVal pType As String, ByVal PField As String, ByVal pFieldType As String) As String
Dim strMid As String
Dim strDType As String

   ProcFieldValue = ""
   strDType = Left(Combo2, 1)
   
   '是否可以直接帶入基本檔欄位
   If pFieldType = "Y" Then
      If Right(UCase(PField), 1) = "W" Then
         Select Case strDType
            Case "2" '民國年
               ProcFieldValue = ", SQLDATET(" & Mid(UCase(PField), 1, Len(PField) - 1) & ")"
            Case "3" '西元年(DD.MM.YYYY)
               ProcFieldValue = ", TO_CHAR(TO_DATE(" & Mid(UCase(PField), 1, Len(PField) - 1) & ",'YYYYMMDD'),'DD.MM.YYYY')"
            Case Else  '西元年(YYYY/MM/DD)
               ProcFieldValue = ", SQLDATEW(" & Mid(UCase(PField), 1, Len(PField) - 1) & ")"
         End Select
      Else
         ProcFieldValue = ", " & PField
      End If
   Else
      '0=案件清單, 1-未發文, 2=未請款, 3=未收文, 4=未付款
      strMid = ""
      Select Case PField
         Case "PA02"  '本所案號O/R:
            strMid = ", PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04)"
         Case "PA09N"  '申請國家-英文
            strMid = ", PN1.NA04"
         Case "PA09C"  '申請國家-中文
            strMid = ", PN1.NA03"
         Case "PA75N"  '代理人編號+名稱
            strMid = ", RTRIM(PA75||' '||DECODE(F1.FA05,NULL,NVL(F1.FA04,F1.FA06),RTRIM(F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)))"
         Case "PA75E"  '代理人名稱(英文)
            strMid = ", RTRIM(DECODE(F1.FA05,NULL,NVL(F1.FA04,F1.FA06),RTRIM(F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)))"
         Case "PA26N"  '申請人編號+名稱
            strMid = ", RTRIM(PA26||' '||DECODE(C1.CU05,NULL,NVL(C1.CU04,C1.CU06),RTRIM(C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)))"
         Case "CP10N"  '收文性質
            If pType = "1" Or pType = "2" Or pType = "3" Then    '未發文/未請款/未收文
               If txtDB(23) = "1" Then  '英文
                  strMid = ", DECODE(CPM02,'605',DECODE(nYr,1,'1st',2,'2nd',3,'3rd',nYr||'th')||' annuity fee',NVL(NVL(CPM10,DECODE(PA09,'000',CPM03,CPM04)),' '))"
               Else
                  strMid = ", DECODE(CPM02,'605','第'||lastyear(pa72)||'年')||DECODE(PA09,'000',CPM03,CPM04)"
               End If
            Else
               strMid = ", ''" '空白
            End If
         Case "NP07N"  '下一道程序
            If txtDB(23) = "1" Then  '英文
               strMid = ", DECODE(NP07,'605',DECODE(nYr,1,'1st',2,'2nd',3,'3rd',nYr||'th')||' annuity fee',NVL(NVL(CPM10,DECODE(PA09,'000',CPM03,CPM04)),' '))"
            Else
               strMid = ", DECODE(NP07,'605','第'||nYr||'年')||DECODE(PA09,'000',CPM03,CPM04)"
            End If
            If pType <> "3" Then
               strMid = Replace(strMid, "CPM", "VCPM")
            End If
         Case "NP07NP09"  '下一道程序之法限
            Select Case strDType
               Case "2" '民國年
                  strMid = ", SQLDATET(NP09)"
               Case "3" '西元年(DD.MM.YYYY)
                  strMid = ", TO_CHAR(TO_DATE(NP09,'YYYYMMDD'),'DD.MM.YYYY')"
               Case Else  '西元年(YYYY/MM/DD)
                  strMid = ", SQLDATEW(NP09)"
            End Select
         Case "CP13N" '智權人員
            If pType = "1" Or pType = "2" Then   '未發文/未請款
               strMid = ", getstaffnamelist(a.CP13)"
            ElseIf pType = "3" Then  '未收文
               strMid = ", getstaffnamelist(NP10)"
            Else '案件清單,未付款
               strMid = ", ''" '空白
            End If
         Case "CP14N" '承辦人員
            If pType = "1" Or pType = "2" Then  '未發文/未請款
               strMid = ", getstaffnamelist(a.CP14)"
            ElseIf pType = "3" Then  '未收文
               strMid = ", getstaffnamelist(NP10)"
            Else '案件清單,未付款
               strMid = ", ''" '空白
            End If
         Case "NA51N" '該區承辦人員
            strMid = ", getstaffnamelist(fn1.NA51)"
         Case "NA16N" '該區FCP程序人員
            strMid = ", getstaffnamelist(DECODE(PA01,'FCP', fn1.NA16, fn1.NA79))"
         Case "NP09CP07"   '法定期限
            If pType = "1" Or pType = "2" Then  '未發文/未請款
               Select Case strDType
                  Case "2" '民國年
                     strMid = ", DECODE(a.CP27,NULL,SQLDATET(a.CP07))"
                  Case "3" '西元年(DD.MM.YYYY)
                     strMid = ",  DECODE(a.CP27,NULL,TO_CHAR(TO_DATE(a.CP07,'YYYYMMDD'),'DD.MM.YYYY'))"
                  Case Else  '西元年(YYYY/MM/DD)
                     strMid = ", DECODE(a.CP27,NULL,SQLDATEW(a.CP07))"
               End Select
            Else  '未收文,案件清單,未付款(預設都有抓案件之下一程序)
               Select Case strDType
                  Case "2" '民國年
                     strMid = ", SQLDATET(NP09)"
                  Case "3" '西元年(DD.MM.YYYY)
                     strMid = ", TO_CHAR(TO_DATE(NP09,'YYYYMMDD'),'DD.MM.YYYY')"
                  Case Else  '西元年(YYYY/MM/DD)
                     strMid = ", SQLDATEW(NP09)"
               End Select
            End If
         Case "NP08CP06"   '本所期限
            If pType = "1" Or pType = "2" Then  '未發文/未請款
               Select Case strDType
                  Case "2" '民國年
                     strMid = ", SQLDATET(a.CP06)"
                  Case "3" '西元年(DD.MM.YYYY)
                     strMid = ", TO_CHAR(TO_DATE(a.CP06,'YYYYMMDD'),'DD.MM.YYYY')"
                  Case Else  '西元年(YYYY/MM/DD)
                     strMid = ", SQLDATEW(a.CP06)"
               End Select
            Else '未收文,案件清單,未付款(預設都有抓案件之下一程序)
               Select Case strDType
                  Case "2" '民國年
                     strMid = ", SQLDATET(NP08)"
                  Case "3" '西元年(DD.MM.YYYY)
                     strMid = ", TO_CHAR(TO_DATE(NP08,'YYYYMMDD'),'DD.MM.YYYY')"
                  Case Else  '西元年(YYYY/MM/DD)
                     strMid = ", SQLDATEW(NP08)"
               End Select
            End If
         Case "SA01E"  '案件狀態(Status)-英文
            strMid = ", decode(pa57||pa108,null,decode(pa16,'1','Granted',decode(nvl(cnt1,0),0,'Pending','Under Examination')),'Abandoned')"
         Case "SA01C"  '案件狀態(Status)-中文
            strMid = ", decode(pa57||pa108,null,decode(pa16,'1','已獲准（Granted）',decode(nvl(cnt1,0),0,'未提實審 (Pending)','審查中（Under Examination）')),'已閉卷（Abandoned）')"
         Case "SA07E"  '案件狀態(Status)已准+發證-英文
            'Modified by Lydia 2024/12/16 從ProcCaseNo取得P案代理人提申日R007
            'strMid = ", decode(pa57||pa108,null,decode(pa16,'1',decode(pa22,null,'Allowed','Granted'),decode(nvl(cnt1,0),0,'Pending','Under Examination')),'Abandoned')"
            strMid = ", decode(pa57||pa108,null,decode(pa16,'1',decode(pa22,null,'Allowed','Granted'),decode(nvl(cnt1,0),0,decode(pa01,'P',decode(r007,null,'To be filed','Pending'),decode(pa10,null,'To be filed','Pending')) ,'Under Examination')),'Abandoned')"
         Case "SA07C"  '案件狀態(Status)已准+發證-中文
            'Modified by Lydia 2024/12/16 從ProcCaseNo取得P案代理人提申日R007
            'strMid = ", decode(pa57||pa108,null,decode(pa16,'1',decode(pa22,null,'已獲准（Allowed）','已發證（Granted）'),decode(nvl(cnt1,0),0,'未提實審 (Pending)','審查中（Under Examination）')),'已閉卷（Abandoned）')"
            strMid = ", decode(pa57||pa108,null,decode(pa16,'1',decode(pa22,null,'已獲准（Allowed）','已發證（Granted）'),decode(nvl(cnt1,0),0,decode(pa01,'P',decode(r007,null,'未提申 (To be filed)','未提實審 (Pending)'),decode(pa10,null,'未提申 (To be filed)','未提實審 (Pending)')) ,'Under Examination')),'已閉卷（Abandoned）')"
         Case "PA16"   '案件狀態(准/駁)
            If txtDB(23) = "1" Then  '英文
               strMid = ", decode(pa16,'1','Granted','2','Abandoned')"
            Else
               strMid = ", decode(pa16,'1','准','2','駁')"
            End If
         Case "PA52PA53PA51"  '第1聯絡人(PP)
            If txtDB(23) = "1" Then
               strMid = ", NVL(PA52,NVL(PA53,PA51))"
            Else
               strMid = ", NVL(PA51,NVL(PA52,PA53))"
            End If
         Case "PA55PA56PA54"  '第2聯絡人(PA)
            If txtDB(23) = "1" Then
               strMid = ", NVL(PA55,NVL(PA56,PA54))"
            Else
               strMid = ", NVL(PA54,NVL(PA55,PA56))"
            End If
         Case "PA27CNT"   '多人申請案
            strMid = ", decode(pa27,'','','多人申請')"
         Case "PA26E"  '第一申請人(英文)名稱
            strMid = ", DECODE(CU05,NULL,NVL(CU04,CU06),RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))"
         Case "SA02" '提實審日期 ;參考X71102(Amkor+JDevice
            If pType = "0" Then '案件清單
               Select Case strDType
                  Case "2" '民國年
                     strMid = ", DECODE(PA08,'2','NA',SQLDATET(DECODE(PA09,'000',NVL(VC1.CP27,VC2.CP27),VC1.CP47)))"
                  Case "3" '西元年(DD.MM.YYYY)
                     strMid = ", DECODE(PA08,'2','NA', TO_CHAR(TO_DATE(DECODE(PA09,'000',NVL(VC1.CP27,VC2.CP27),VC1.CP47))),'DD.MM.YYYY')"
                     '", TO_CHAR(TO_DATE(NP09,'YYYYMMDD'),'DD.MM.YYYY')"
                  Case Else  '西元年(YYYY/MM/DD)
                     strMid = ", DECODE(PA08,'2','NA',SQLDATEW(DECODE(PA09,'000',NVL(VC1.CP27,VC2.CP27),VC1.CP47)))"
               End Select
            Else
               strMid = ", ''" '空白
            End If
         Case "CP27W" '發文日期
            If pType = "2" Then '未發文
               Select Case strDType
                  Case "2" '民國年
                     strMid = ", SQLDATET(a.CP27)"
                  Case "3" '西元年(DD.MM.YYYY)
                     strMid = ", TO_CHAR(TO_DATE(a.CP27,'YYYYMMDD'),'DD.MM.YYYY')"
                  Case Else  '西元年(YYYY/MM/DD)
                     strMid = ", SQLDATEW(a.CP27)"
               End Select
            Else
               strMid = ", ''" '空白
            End If
         Case "CP05W" '收文日期
            If pType = "1" Or pType = "2" Then '未發文/未請款
               Select Case strDType
                  Case "2" '民國年
                     strMid = ", SQLDATET(a.CP05)"
                  Case "3" '西元年(DD.MM.YYYY)
                     strMid = ", TO_CHAR(TO_DATE(a.CP05,'YYYYMMDD'),'DD.MM.YYYY')"
                  Case Else  '西元年(YYYY/MM/DD)
                     strMid = ", SQLDATEW(a.CP05)"
               End Select
            Else
               strMid = ", ''" '空白
            End If
         Case "LCP10M" '最新一道收文性質
            If pType = "0" Or pType = "4" Then  '案件清單/未付款
               strMid = ", DECODE(vtb3.LCP10, NULL, '', GetCP10Desc(PA01, vtb3.LCP10, PA09))"
            Else
               strMid = ", ''" '空白
            End If
         Case "LCP10A" '最新一道收文之承辦工程師
            If pType = "0" Or pType = "4" Then  '案件清單/未付款
               strMid = ", getstaffnamelist(vtb3.LCP14)"
            Else
               strMid = ", ''" '空白
            End If
         Case "PA08E"  '專利種類-英文
            strMid = ", DECODE(PA08,'1','Invention','2','Utility Model','3','Design')"
         Case "PA08C"  '專利種類-中文
            strMid = ", DECODE(PA08,'1','發明','2','新型','3','設計')"
         Case "PA26PA30N"  '多申請人1~5+英文名稱 (Applicant(s))
            strMid = ", getcustomernamelist(PA26||';'||PA27||';'||PA28||';'||PA29||';'||PA30)"
         Case "PA26PA30E"  '多申請人英文名稱不含編號
            strMid = ", getcustomernamelist(PA26||';'||PA27||';'||PA28||';'||PA29||';'||PA30,'2')"
         Case "PA27N"   '申請人2編號+英文名稱 (Applicant(2))
            strMid = ", getcustomernamelist(PA27)"
         Case "PA28N"   '申請人3編號+英文名稱 (Applicant(3))
            strMid = ", getcustomernamelist(PA28)"
         Case "PA29N"   '申請人4編號+英文名稱 (Applicant(4))
            strMid = ", getcustomernamelist(PA29)"
         Case "PA30N"   '申請人5編號+英文名稱 (Applicant(5))
            strMid = ", getcustomernamelist(PA30)"
         Case "PA150N"   '工程師組別
             strMid = ", CST16(PA150)"
         Case "PD06"   '優先權號
             strMid = ", GETPRIORITY2(PA01,PA02,PA03,PA04)"
         Case "PD05W"   '優先權日=>西元年
             strMid = ", GETPRIORITYPD05(PA01,PA02,PA03,PA04)"
         Case "PD06PD05"   '優先權號+優先權日
             strMid = ", GETPRIORITYPD06PD05(PA01,PA02,PA03,PA04)"
         Case "PA88"   '個案固定請款對象
             strMid = ", decode(PA88,NULL,NULL,decode(substr(PA88,1,1),'X',getcustomernamelist(PA88),getfagentnamelist(PA88)))"
         Case "PA133"   '個案D/N固定列印對象
             strMid = ", decode(PA133,NULL,NULL,decode(substr(PA133,1,1),'X',getcustomernamelist(PA133),getfagentnamelist(PA133)))"
         Case "CU57"   'X編號固定請款對象
             strMid = ", decode(C1.CU57,NULL,NULL,decode(substr(C1.CU57,1,1),'X',getcustomernamelist(C1.CU57),getfagentnamelist(C1.CU57)))"
         Case "CU105"   'X編號D/N固定列印對象
             strMid = ", decode(C1.CU105,NULL,NULL,decode(substr(C1.CU105,1,1),'X',getcustomernamelist(C1.CU105),getfagentnamelist(C1.CU105)))"
         Case "FA30"   'X編號固定請款對象
             strMid = ", decode(FA30,NULL,NULL,decode(substr(FA30,1,1),'X',getcustomernamelist(FA30),getfagentnamelist(FA30)))"
         Case "FA71"   'X編號D/N固定列印對象
             strMid = ", decode(FA71,NULL,NULL,decode(substr(FA71,1,1),'X',getcustomernamelist(FA71),getfagentnamelist(FA71)))"
         Case "PA26CU10"   '申請人1國籍-中文
             strMid = ", cn1.NA03"
         Case "PA75FA10"   '申請人1國籍-中文
             strMid = ", fn1.NA03"
         Case "SA03"  '下一年度年費之年度
            strMid = ", nYr"
         Case "SA04T"  '下一年年費規費台幣金額
            'Modified by Lydia 2024/11/26 Y23008抓未收文的規費和服務費+ p_Type= 2,3
            If pType = "0" Or pType = "4" Or pType = "3" Or pType = "2" Then
               strMid = ", DECODE(NP07,NULL,NULL,YF07)"
            Else
               strMid = ", ''" '空白
            End If
         Case "SA05T"  '下一年年費服務費台幣金額
            'Modified by Lydia 2024/11/26 Y23008抓未收文的規費和服務費+ p_Type= 2,3
            If pType = "0" Or pType = "4" Or pType = "3" Or pType = "2" Then
               strMid = ", DECODE(NP07,NULL,NULL,YF06)"
            Else
               strMid = ", ''" '空白
            End If
         Case "SA04U"  '下一年年費規費美金金額
            'Modified by Lydia 2024/11/26 Y23008抓未收文的規費和服務費+ p_Type= 2,3
            If pType = "0" Or pType = "4" Or pType = "3" Or pType = "2" Then
               strMid = ", DECODE(NP07,NULL,NULL,trunc(nvl(yf07,0)/USXR02))"
            Else
               strMid = ", ''" '空白
            End If
         Case "SA05U"  '下一年年費服務費美金金額
            'Modified by Lydia 2024/11/26 Y23008抓未收文的規費和服務費+ p_Type= 2,3
            If pType = "0" Or pType = "4" Or pType = "3" Or pType = "2" Then
               strMid = ", DECODE(NP07,NULL,NULL,trunc(nvl(yf06,0)/USXR02))"
            Else
               strMid = ", ''" '空白
            End If
         Case "SA06T"  '下一年年費台幣金額
            'Modified by Lydia 2024/11/26 Y23008抓未收文的規費和服務費+ p_Type= 2,3
            If pType = "0" Or pType = "4" Or pType = "3" Or pType = "2" Then
               strMid = ", DECODE(NP07,NULL,NULL,nvl(YF07,0)+nvl(YF06,0)) "
            Else
               strMid = ", ''" '空白
            End If
         Case "SA06U"  '下一年年費美金金額
            'Modified by Lydia 2024/11/26 Y23008抓未收文的規費和服務費+ p_Type= 2,3
            If pType = "0" Or pType = "4" Or pType = "3" Or pType = "2" Then
               strMid = ", DECODE(NP07,NULL,NULL,TRUNC((NVL(YF07,0)+NVL(YF06,0))/USXR02))"
            Else
               strMid = ", ''" '空白
            End If
         Case "USRATE"  '美金匯率
            'Modified by Lydia 2024/11/26 + p_Type= 2,3
            If pType = "0" Or pType = "4" Or pType = "3" Or pType = "2" Then
               strMid = ", USXR02"
            Else
               strMid = ", ''" '空白
            End If
         Case "A1K02" '未付款帳單日期
            If pType = "4" Then
               Select Case strDType
                  Case "2" '民國年
                     strMid = ", SQLDATET(" & PField & ")"
                  Case "3" '西元年(DD.MM.YYYY)
                     strMid = ", TO_CHAR(TO_DATE(" & PField & "+19110000,'YYYYMMDD'),'DD.MM.YYYY')"
                  Case Else  '西元年(YYYY/MM/DD)
                     strMid = ", SQLDATEW(" & PField & "+19110000)"
               End Select
            Else
               strMid = ", ''" '空白
            End If
         Case "A1K_FEE" '未付款帳單服務費
            If pType = "4" Then
               strMid = ", A1K_FEE"
            Else
               strMid = ", ''" '空白
            End If
         Case "A1K_EXP" '未付款帳單規費
            If pType = "4" Then
               strMid = ", A1K_EXP"
            Else
               strMid = ", ''" '空白
            End If
         Case "A1K_ITEM" '未付款帳單第1個請款項目
            If pType = "4" Then
               strMid = ", A1K_ITEM"
            Else
               strMid = ", ''" '空白
            End If
         Case "A1K_DG" '請款金額與未付款之差額
            If pType = "4" Then
               strMid = ", A1K08-(A1K_FEE+A1K_EXP)"
            Else
               strMid = ", ''" '空白
            End If
         Case "PA0506" '專利名稱(英文+中文)
            strMid = ", LTRIM(RTRIM(PA06||' '||PA05))"
         'Added by Lydia 2024/12/26
         'Modified by Lydia 2025/05/20 +A00狀態
         Case "C00", "A00" '各項指示:以類別為Key
            If pType = "0" Then '目前只在案件清單
               strMid = ", VITS01.ITS06"
            Else
               strMid = ", '未設定'" '空白
            End If
            
         Case Else
            strMid = ", '未設定'" '空白
      End Select
      ProcFieldValue = strMid
   End If

End Function

'******產生EXCEL檔******
Private Function ProcExcelSave(ByVal pNum As Integer, ByVal pSQL As String) As Boolean
Dim xlsReport
Dim wksReport
Dim iRound As Integer, nRows As Integer, nPages As Integer, strNCols As String
Dim bolMerge As Boolean, strMsg As String
Dim strWksName As String, strWksNameD As String, strGrpOld As String
Dim tmpArr1 As Variant
Dim tmpArray As Variant

On Error GoTo ErrHnd
   
   ProcExcelSave = False
   
   For iRound = LBound(strConSql) To UBound(strConSql)
      strQuery = ""
      If strConSql(iRound) <> "" Then
         strQuery = strConSql(iRound)
         '工作表名稱
         Select Case iRound
            Case 0: strWksName = "案件清單"
            Case 1: strWksName = "未發文"
            Case 2: strWksName = "未請款"
            Case 3: strWksName = "未收文"
            Case 4: strWksName = "未付款"
            Case 5: strWksName = "行事曆"
         End Select
         If iRound >= 1 And iRound <= 3 Then
            If txtDB(20) = "Y" Then  '未發文/未請款/未收文是否匯出同一工作表(Sheet)
              If bolMerge = False Then
                 bolMerge = True
                 strQuery = pSQL
                 strWksName = Trim(Mid(IIf(Check1(1).Value = 1, " or 未發文", "") & IIf(Check1(2).Value = 1, " or 未請款", "") & IIf(Check1(3).Value = 1, " or 未收文", ""), 4))
              Else
                 strQuery = ""
              End If
            End If
         End If
         If strQuery <> "" Then
            intQ = 1
            Set rsQD = ClsLawReadRstMsg(intQ, strQuery)
            If intQ = 0 Then
               strMsg = strMsg & "、" & strWksName
            Else
               strGrpOld = ""
               '整批輸入資料
               rsQD.MoveFirst
               Do While Not rsQD.EOF
                  If Check3(1).Value = 1 Then '不同申請人分成不同Sheet
                     strWksNameD = strWksName & "_" & rsQD.Fields("GRPNO")
                  Else
                     strWksNameD = strWksName
                  End If
                  If strGrpOld <> strWksNameD Then '切換工作表
                     nPages = nPages + 1
                     '增加工作表
                     If pNum < nPages Then
                        xlsReport.Worksheets.add After:=wksReport '插入sheet
                        pNum = pNum + 1
                     End If
                     If nPages = 1 Then
                        Set xlsReport = CreateObject("Excel.Application")
                        xlsReport.SheetsInNewWorkbook = pNum
                        xlsReport.Workbooks.add
                        Set wksReport = xlsReport.Worksheets(nPages)
                        wksReport.Activate
                        xlsReport.Visible = True
                        
                        tmpArr1 = Empty
                        '行事曆固定欄位
                        If iRound = 5 Then
                           tmpArr1 = Split("本所案號|事由|管制日期|代理人編號|申請人編號", "|")
                        Else
                           tmpArr1 = Split(strTitle, "|")
                        End If
                        ReDim tmpArray(1 To UBound(tmpArr1) + 1)
                     Else
                        tmpArr1 = Empty
                        '行事曆固定欄位
                        If iRound = 5 Then
                           tmpArr1 = Split("本所案號|事由|管制日期|代理人編號|申請人編號", "|")
                        Else
                           tmpArr1 = Split(strTitle, "|")
                        End If
                        ReDim tmpArray(1 To UBound(tmpArr1) + 1)
                        
                        For intQ = 1 To UBound(tmpArray) '調整為能使文字全部顯示之欄寬(前一工作表)
                           wksReport.Columns(Pub_NumberToSystem26(intQ) & ":" & Pub_NumberToSystem26(intQ)).EntireColumn.AutoFit
                        Next
                        Set wksReport = xlsReport.Worksheets(nPages)
                        wksReport.Activate
                     End If
                     xlsReport.Worksheets(nPages).Name = strWksNameD  '工作表名稱
                      '設定欄位名稱及欄寬
                     nRows = 1
                     For intQ = 1 To UBound(tmpArr1) + 1
                         strNCols = Pub_NumberToSystem26(intQ)
                         wksReport.Range(strNCols & nRows).Value = tmpArr1(intQ - 1)
                         wksReport.Range(strNCols & ":" & strNCols).ColumnWidth = 20
                         wksReport.Range(strNCols & ":" & strNCols).Font.Name = "Arial"
                         wksReport.Range(strNCols & ":" & strNCols).Font.Size = 10
                         'wksReport.Range(strNCols & nRows).HorizontalAlignment = xlCenter  '不用置中
                     Next intQ
                     wksReport.Range("A" & nRows & ":" & strNCols & nRows).Font.Bold = True
                     
                     nRows = nRows + 1
                     wksReport.Range("A" & nRows).Select
                     xlsReport.ActiveWindow.FreezePanes = True '凍結窗格
                  End If
                  strGrpOld = strWksNameD
                  
                  For intQ = 1 To UBound(tmpArray)
                     tmpArray(intQ) = "" & rsQD.Fields(intQ - 1)
                  Next intQ
                  wksReport.Range("A" & nRows & ":" & Pub_NumberToSystem26(UBound(tmpArray)) & nRows).NumberFormatLocal = "@"
                  wksReport.Range("A" & nRows & ":" & Pub_NumberToSystem26(UBound(tmpArray)) & nRows).Value = tmpArray
                  '當儲存格格式為通用格式，選取儲存格的值再代入到選取的儲存格=>自動化為數值欄位
                  'wksReport.Range("A" & nRows & ":" & Pub_NumberToSystem26(UBound(tmpArray)) & nRows).Value = wksReport.Range("A" & nRows & ":" & Pub_NumberToSystem26(UBound(tmpArray)) & nRows).Value
                  nRows = nRows + 1
                  
                  rsQD.MoveNext
               Loop
            End If 'rsQD
         End If
      End If  'If strQuery <> "" Then
   Next iRound
   
   If nPages > 0 Then
      For intQ = 1 To UBound(tmpArray) '調整為能使文字全部顯示之欄寬(目前工作表)
         wksReport.Columns(Pub_NumberToSystem26(intQ) & ":" & Pub_NumberToSystem26(intQ)).EntireColumn.AutoFit
      Next
      xlsReport.Sheets(1).Select '選擇工作表
      '判斷版本
      If Val(xlsReport.Version) < 12 Then
         xlsReport.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
      Else
         xlsReport.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
      End If
      xlsReport.Workbooks.Close
      xlsReport.Quit
      Set wksReport = Nothing
      Set xlsReport = Nothing
   Else
      strFileName = "" '無資料可產生清單
   End If
   If strMsg <> "" Then
      MsgBox "以下清單查無資料：" & Mid(strMsg, 2), vbInformation + vbOKOnly, "查詢結果"
   End If
   
   ProcExcelSave = True
   Exit Function

ErrHnd:

   MsgBox Err.Description, , "Excel檔案產生失敗"
   strFileName = ""
End Function

Private Sub Command1_Click()

On Error GoTo ErrHandle
  If MsgBox("確定將清單ITEM的No.14改成No.13？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
     strQuery = "select * from fcpelistitem where fei01='A' and fei02>12 order by 1,2"
     intQ = 1
     Set rsQD = ClsLawReadRstMsg(intQ, strQuery)
     If intQ = 1 Then
        rsQD.MoveFirst
        cnnConnection.BeginTrans
        Do While Not rsQD.EOF
           strSql = "Update FcpElistItem set fei02='" & Val("" & rsQD.Fields("fei02")) - 1 & "' where fei01='" & rsQD.Fields("fei01") & "' and fei02='" & rsQD.Fields("fei02") & "' "
           cnnConnection.Execute strSql, intI
           strSql = "Update FcpElistRec set fer26=replace(fer26,'" & rsQD.Fields("fei01") & rsQD.Fields("fei02") & "','" & rsQD.Fields("fei01") & Val("" & rsQD.Fields("fei02")) - 1 & "') "
           cnnConnection.Execute strSql, intI
           rsQD.MoveNext
        Loop
        cnnConnection.CommitTrans
     End If
     MsgBox "OK!"
  End If
  Exit Sub
  
ErrHandle:
  If Err.Number <> 0 Then
     cnnConnection.RollbackTrans
     MsgBox "修正失敗:" & vbCrLf & Err.Description
  End If
End Sub


'Memo by Lydia 2024/02/21 先保留(尚未加入不同申請人不同工作表)
'******產生EXCEL檔******
Private Function ProcExcelSave_old(ByVal pNum As Integer, ByVal pSQL As String) As Boolean
Dim xlsReport
Dim wksReport
Dim iRound As Integer, nRows As Integer, nPages As Integer
Dim bolMerge As Boolean, strMsg As String
Dim tmpArr1 As Variant, strNCols As String
Dim tmpArray As Variant
Dim strGrp As String

On Error GoTo ErrHnd
   
   ProcExcelSave_old = False
   
   For iRound = LBound(strConSql) To UBound(strConSql)
      strQuery = ""
      If strConSql(iRound) <> "" Then
         strQuery = strConSql(iRound)
         '工作表名稱
         Select Case iRound
            Case 0: strNCols = "案件清單"
            Case 1: strNCols = "未發文"
            Case 2: strNCols = "未請款"
            Case 3: strNCols = "未收文"
            Case 4: strNCols = "未付款"
            Case 5: strNCols = "行事曆"
         End Select
         If iRound >= 1 And iRound <= 3 Then
            If txtDB(20) = "Y" Then  '未發文/未請款/未收文是否匯出同一工作表(Sheet)
              If bolMerge = False Then
                 bolMerge = True
                 strQuery = pSQL
                 strNCols = Trim(Mid(IIf(Check1(1).Value = 1, " or 未發文", "") & IIf(Check1(2).Value = 1, " or 未請款", "") & IIf(Check1(3).Value = 1, " or 未收文", ""), 4))
              Else
                 strQuery = ""
              End If
            End If
         End If
         If strQuery <> "" Then
            intQ = 1
            Set rsQD = ClsLawReadRstMsg(intQ, strQuery)
            If intQ = 0 Then
               strMsg = strMsg & "、" & strNCols
            Else
               nPages = nPages + 1
               '增加工作表
               If pNum < nPages Then
                  xlsReport.Worksheets.add After:=wksReport '插入sheet
                  pNum = pNum + 1
               End If
               If nPages = 1 Then
                  Set xlsReport = CreateObject("Excel.Application")
                  xlsReport.SheetsInNewWorkbook = pNum
                  xlsReport.Workbooks.add
                  Set wksReport = xlsReport.Worksheets(nPages)
                  wksReport.Activate
                  xlsReport.Visible = True
                  
                  tmpArr1 = Empty
                  tmpArr1 = Split(strTitle, "|")
                  ReDim tmpArray(1 To UBound(tmpArr1) + 1)
               Else
                  For intQ = 1 To UBound(tmpArray) '調整為能使文字全部顯示之欄寬(前一工作表)
                     wksReport.Columns(Pub_NumberToSystem26(intQ) & ":" & Pub_NumberToSystem26(intQ)).EntireColumn.AutoFit
                  Next
                  Set wksReport = xlsReport.Worksheets(nPages)
                  wksReport.Activate
               End If
               If Check3(1).Value = 1 Then '不同申請人分成不同Sheet
                  strNCols = strNCols & "_" & rsQD.Fields("GRPNO")
               End If
               xlsReport.Worksheets(nPages).Name = strNCols  '工作表名稱
                '設定欄位名稱及欄寬
               nRows = 1
               For intQ = 1 To UBound(tmpArr1) + 1
                   strNCols = Pub_NumberToSystem26(intQ)
                   wksReport.Range(strNCols & nRows).Value = tmpArr1(intQ - 1)
                   wksReport.Range(strNCols & ":" & strNCols).ColumnWidth = 20
                   wksReport.Range(strNCols & ":" & strNCols).Font.Name = "Arial"
                   wksReport.Range(strNCols & ":" & strNCols).Font.Size = 10
                   'wksReport.Range(strNCols & nRows).HorizontalAlignment = xlCenter  '不用置中
               Next intQ
               wksReport.Range("A" & nRows & ":" & strNCols & nRows).Font.Bold = True
               
               nRows = nRows + 1
               wksReport.Range("A" & nRows).Select
               xlsReport.ActiveWindow.FreezePanes = True '凍結窗格

               '整批輸入資料
               rsQD.MoveFirst
               Do While Not rsQD.EOF
                  For intQ = 1 To UBound(tmpArray)
                     tmpArray(intQ) = "" & rsQD.Fields(intQ - 1)
                  Next intQ
                  wksReport.Range("A" & nRows & ":" & Pub_NumberToSystem26(UBound(tmpArray)) & nRows).NumberFormatLocal = "@"
                  wksReport.Range("A" & nRows & ":" & Pub_NumberToSystem26(UBound(tmpArray)) & nRows).Value = tmpArray
                  '當儲存格格式為通用格式，選取儲存格的值再代入到選取的儲存格=>自動化為數值欄位
                  'wksReport.Range("A" & nRows & ":" & Pub_NumberToSystem26(UBound(tmpArray)) & nRows).Value = wksReport.Range("A" & nRows & ":" & Pub_NumberToSystem26(UBound(tmpArray)) & nRows).Value
                  nRows = nRows + 1
                  rsQD.MoveNext
               Loop
            End If 'rsQD
         End If
      End If  'If strQuery <> "" Then
   Next iRound
   
   If nPages > 0 Then
      For intQ = 1 To UBound(tmpArray) '調整為能使文字全部顯示之欄寬(目前工作表)
         wksReport.Columns(Pub_NumberToSystem26(intQ) & ":" & Pub_NumberToSystem26(intQ)).EntireColumn.AutoFit
      Next
      xlsReport.Sheets(1).Select '選擇工作表
      '判斷版本
      If Val(xlsReport.Version) < 12 Then
         xlsReport.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
      Else
         xlsReport.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
      End If
      xlsReport.Workbooks.Close
      xlsReport.Quit
      Set wksReport = Nothing
      Set xlsReport = Nothing
   Else
      strFileName = "" '無資料可產生清單
   End If
   If strMsg <> "" Then
      MsgBox "以下清單查無資料：" & Mid(strMsg, 2), vbInformation + vbOKOnly, "查詢結果"
   End If
   
   ProcExcelSave_old = True
   Exit Function

ErrHnd:

   MsgBox Err.Description, , "Excel檔案產生失敗"
   strFileName = ""
End Function

'Added by Lydia 2025/01/17
Public Function GetNowFER01() As String
Dim strA1 As String, intA As Integer
Dim rsAD As New ADODB.Recordset

   GetNowFER01 = ""
   strA1 = "select decode(mno,null,to_char(to_number(substr(to_char(sysdate,'yyyymmdd'),1,4))-1911)||'0001', to_number(mno)+1) mnowno from (" & _
            " select max(fer01) mno from fcpelistrec where fer03>=substr(to_char(sysdate,'yyyymmdd'),1,4)||'0100') "
   intA = 1
   'PKEY：民國年+流水號4碼
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
      GetNowFER01 = rsAD.Fields("mnowno")
   End If
   Set rsAD = Nothing
   
End Function
