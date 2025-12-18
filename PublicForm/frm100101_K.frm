VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_K 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料查詢"
   ClientHeight    =   5730
   ClientLeft      =   110
   ClientTop       =   750
   ClientWidth     =   9570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9570
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   13
      Left            =   8520
      MaxLength       =   7
      TabIndex        =   97
      Top             =   4110
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   14
      Left            =   8520
      MaxLength       =   7
      TabIndex        =   96
      Top             =   4410
      Width           =   930
   End
   Begin VB.Frame Frame2 
      Height          =   370
      Left            =   4350
      TabIndex        =   90
      Top             =   2940
      Visible         =   0   'False
      Width           =   5170
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   10
         Left            =   1120
         MaxLength       =   6
         TabIndex        =   92
         Top             =   0
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   19
         Left            =   4180
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   91
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "外文核稿人："
         Height          =   180
         Index           =   25
         Left            =   -810
         TabIndex        =   95
         Top             =   60
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "外文核完日："
         Height          =   180
         Index           =   41
         Left            =   3090
         TabIndex        =   94
         Top             =   60
         Width           =   1080
      End
      Begin MSForms.ComboBox Combo4 
         Height          =   320
         Left            =   1140
         TabIndex        =   93
         Top             =   0
         Width           =   1890
         VariousPropertyBits=   679495711
         DisplayStyle    =   3
         Size            =   "3334;564"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.TextBox textCP143 
      Height          =   270
      Left            =   8520
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   780
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   77
      Top             =   2610
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   74
      Top             =   1080
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   17
      Left            =   7350
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   71
      Top             =   2610
      Width           =   915
   End
   Begin VB.CommandButton cmd 
      Caption         =   "承辦歷程"
      Height          =   375
      Index           =   2
      Left            =   7050
      TabIndex        =   70
      Top             =   60
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   67
      Top             =   1410
      Width           =   930
   End
   Begin VB.CommandButton cmd 
      Caption         =   "智權人員補充資料記錄(&S)"
      Height          =   375
      Index           =   0
      Left            =   6930
      TabIndex        =   65
      Top             =   1890
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.TextBox txtEP36 
      Height          =   264
      Left            =   8520
      MaxLength       =   7
      TabIndex        =   61
      Top             =   1350
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txtEP37 
      Height          =   264
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   60
      Top             =   2280
      Width           =   930
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "存檔"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6180
      Style           =   1  '圖片外觀
      TabIndex        =   59
      Top             =   60
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   11
      Top             =   4100
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   10
      Top             =   2310
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1710
      Width           =   480
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4380
      Width           =   360
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   7
      Top             =   480
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   6
      Top             =   780
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   18
      Left            =   5460
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2000
      Width           =   930
   End
   Begin VB.CommandButton cmd 
      Caption         =   "回前畫面"
      Height          =   375
      Index           =   1
      Left            =   8340
      TabIndex        =   0
      Top             =   60
      Width           =   1125
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   1570
      Left            =   30
      TabIndex        =   86
      Top             =   4110
      Width           =   4090
      _ExtentX        =   7214
      _ExtentY        =   2769
      _Version        =   393216
      TabHeight       =   360
      TabCaption(0)   =   "條款"
      TabPicture(0)   =   "frm100101_K.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txt1(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "商標描述中文"
      TabPicture(1)   =   "frm100101_K.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "商標描述英文"
      TabPicture(2)   =   "frm100101_K.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt1(6)"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txt1 
         Height          =   1290
         Index           =   11
         Left            =   50
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   89
         Top             =   240
         Width           =   3980
      End
      Begin VB.TextBox txt1 
         Height          =   1290
         Index           =   0
         Left            =   -74940
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   88
         Top             =   240
         Width           =   3980
      End
      Begin VB.TextBox txt1 
         Height          =   1290
         Index           =   6
         Left            =   -74940
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   87
         Top             =   240
         Width           =   3980
      End
   End
   Begin VB.Label LblEP39 
      AutoSize        =   -1  'True
      Caption         =   "核稿完成日："
      Height          =   180
      Left            =   7440
      TabIndex        =   99
      Top             =   4140
      Width           =   1080
   End
   Begin VB.Label LblEP42 
      AutoSize        =   -1  'True
      Caption         =   "判發完成日："
      Height          =   180
      Left            =   7410
      TabIndex        =   98
      Top             =   4460
      Width           =   1080
   End
   Begin VB.Label lblFee 
      AutoSize        =   -1  'True
      Caption         =   "lblFee"
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
      Height          =   180
      Left            =   3030
      TabIndex        =   85
      Top             =   3320
      Width           =   510
   End
   Begin VB.Label lblCertType 
      AutoSize        =   -1  'True
      Caption         =   "lblCertType"
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
      Height          =   180
      Left            =   3030
      TabIndex        =   84
      Top             =   2880
      Width           =   1010
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   720
      Left            =   5460
      TabIndex        =   83
      Top             =   4950
      Width           =   3770
      VariousPropertyBits=   -1466941409
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6641;1270"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtEP12 
      Height          =   720
      Left            =   5460
      TabIndex        =   12
      Top             =   3330
      Width           =   3770
      VariousPropertyBits=   -1466941409
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6641;1270"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "查名齊備日："
      Height          =   180
      Index           =   2
      Left            =   7380
      TabIndex        =   82
      Top             =   830
      Width           =   1100
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   18
      Left            =   5550
      TabIndex        =   80
      Top             =   2610
      Visible         =   0   'False
      Width           =   840
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1482;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblEApp 
      Caption         =   "電子送件"
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
      Height          =   180
      Left            =   3030
      TabIndex        =   79
      Top             =   1140
      Visible         =   0   'False
      Width           =   890
   End
   Begin VB.Label Label1 
      Caption         =   "會稿完成日："
      Height          =   180
      Index           =   5
      Left            =   4350
      TabIndex        =   78
      Top             =   2640
      Width           =   1080
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   14
      Left            =   6420
      TabIndex        =   76
      Top             =   1110
      Width           =   1190
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2090;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "核稿人："
      Height          =   180
      Index           =   6
      Left            =   4700
      TabIndex        =   75
      Top             =   1110
      Width           =   740
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   35
      Left            =   8310
      TabIndex        =   73
      Top             =   2670
      Width           =   860
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1508;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "判發人："
      Height          =   180
      Index           =   52
      Left            =   6570
      TabIndex        =   72
      Top             =   2670
      Width           =   740
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   10
      Left            =   5490
      TabIndex        =   69
      Top             =   1440
      Visible         =   0   'False
      Width           =   950
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1667;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "完稿日："
      Height          =   180
      Index           =   26
      Left            =   4700
      TabIndex        =   68
      Top             =   1440
      Width           =   740
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Index           =   0
      Left            =   4530
      TabIndex        =   66
      Top             =   4980
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   20
      Left            =   5490
      TabIndex        =   64
      Top             =   3750
      Visible         =   0   'False
      Width           =   780
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2408;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "智權人員齊備日："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   44
      Left            =   7050
      TabIndex        =   63
      Top             =   1380
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "客戶會稿日："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   45
      Left            =   6570
      TabIndex        =   62
      Top             =   2340
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "(N:不通知,不發文)"
      Height          =   180
      Index           =   22
      Left            =   5880
      TabIndex        =   18
      ToolTipText     =   "(N:  不通知, 自動內部收文)"
      Top             =   4440
      Width           =   2960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(Y/N)"
      Height          =   180
      Index           =   30
      Left            =   6000
      TabIndex        =   36
      Top             =   1740
      Width           =   410
   End
   Begin VB.Label Label1 
      Caption         =   "承辦期限："
      Height          =   180
      Index           =   46
      Left            =   4470
      TabIndex        =   58
      Top             =   570
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "承辦備註："
      Height          =   180
      Index           =   31
      Left            =   4530
      TabIndex        =   57
      Top             =   3320
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "取消收文日："
      Height          =   180
      Index           =   29
      Left            =   4350
      TabIndex        =   56
      Top             =   4700
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "發文日："
      Height          =   180
      Index           =   7
      Left            =   4700
      TabIndex        =   55
      Top             =   4140
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "是否會稿："
      Height          =   180
      Index           =   27
      Left            =   4470
      TabIndex        =   54
      Top             =   1740
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "會稿日："
      Height          =   180
      Index           =   24
      Left            =   4700
      TabIndex        =   53
      Top             =   2340
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "齊備日："
      Height          =   180
      Index           =   23
      Left            =   4700
      TabIndex        =   52
      Top             =   810
      Width           =   740
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   23
      Left            =   690
      TabIndex        =   51
      Top             =   3810
      Width           =   920
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   19
      Left            =   1050
      TabIndex        =   50
      Top             =   3230
      Width           =   1170
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2064;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   17
      Left            =   1050
      TabIndex        =   49
      Top             =   2940
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   15
      Left            =   1050
      TabIndex        =   48
      Top             =   2640
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   13
      Left            =   1530
      TabIndex        =   47
      Top             =   2360
      Width           =   1170
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2064;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   11
      Left            =   1530
      TabIndex        =   46
      Top             =   2070
      Width           =   600
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1058;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   9
      Left            =   1110
      TabIndex        =   45
      Top             =   1770
      Width           =   3050
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "5371;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   7
      Left            =   1110
      TabIndex        =   44
      Top             =   1490
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3228;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   5
      Left            =   1140
      TabIndex        =   43
      Top             =   1190
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   3
      Left            =   1140
      TabIndex        =   42
      Top             =   920
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "(N：不算)"
      Height          =   260
      Index           =   32
      Left            =   2270
      TabIndex        =   41
      Top             =   2070
      Width           =   1070
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   28
      Left            =   5460
      TabIndex        =   40
      Top             =   4700
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   12
      Left            =   5490
      TabIndex        =   39
      Top             =   2340
      Visible         =   0   'False
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "123"
      Size            =   "2408;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   8
      Left            =   5490
      TabIndex        =   38
      Top             =   840
      Visible         =   0   'False
      Width           =   920
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2408;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   6
      Left            =   5490
      TabIndex        =   37
      Top             =   1710
      Visible         =   0   'False
      Width           =   560
      VariousPropertyBits=   27
      Caption         =   "lblF"
      Size            =   "979;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "點數："
      Height          =   260
      Index           =   11
      Left            =   150
      TabIndex        =   35
      Top             =   3810
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   260
      Index           =   13
      Left            =   150
      TabIndex        =   34
      Top             =   3230
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   260
      Index           =   14
      Left            =   150
      TabIndex        =   33
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   260
      Index           =   15
      Left            =   150
      TabIndex        =   32
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "專利/商標種類："
      Height          =   260
      Index           =   16
      Left            =   150
      TabIndex        =   31
      Top             =   2360
      Width           =   1370
   End
   Begin VB.Label Label1 
      Caption         =   "是否算案件數："
      Height          =   260
      Index           =   17
      Left            =   150
      TabIndex        =   30
      Top             =   2070
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   260
      Index           =   18
      Left            =   150
      TabIndex        =   29
      Top             =   1770
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   260
      Index           =   19
      Left            =   150
      TabIndex        =   28
      Top             =   1490
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "收文日："
      Height          =   260
      Index           =   20
      Left            =   150
      TabIndex        =   27
      Top             =   1190
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   260
      Index           =   21
      Left            =   150
      TabIndex        =   26
      Top             =   900
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   260
      Index           =   8
      Left            =   1410
      TabIndex        =   25
      Top             =   620
      Width           =   740
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   0
      Left            =   710
      TabIndex        =   24
      Top             =   620
      Width           =   630
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1111;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   1
      Left            =   2190
      TabIndex        =   23
      Top             =   620
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "目次："
      Height          =   260
      Index           =   1
      Left            =   150
      TabIndex        =   22
      Top             =   620
      Width           =   540
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   22
      Left            =   5520
      TabIndex        =   21
      Top             =   4470
      Visible         =   0   'False
      Width           =   270
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "476;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   21
      Left            =   1080
      TabIndex        =   20
      Top             =   3540
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2487;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否通知客戶："
      Height          =   180
      Index           =   3
      Left            =   4170
      TabIndex        =   19
      Top             =   4440
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   260
      Index           =   12
      Left            =   150
      TabIndex        =   17
      Top             =   3530
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   2
      Left            =   5490
      TabIndex        =   16
      Top             =   570
      Width           =   900
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2408;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
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
      Height          =   180
      Left            =   3030
      TabIndex        =   15
      Top             =   1490
      Width           =   950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "指定會稿日："
      Height          =   180
      Index           =   39
      Left            =   4350
      TabIndex        =   14
      Top             =   2030
      Width           =   1080
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   33
      Left            =   5490
      TabIndex        =   13
      Top             =   2010
      Width           =   840
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2408;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國外案承辦人："
      Height          =   180
      Index           =   9
      Left            =   2940
      TabIndex        =   4
      Top             =   6525
      Width           =   1260
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   27
      Left            =   4260
      TabIndex        =   3
      Top             =   6525
      Width           =   270
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2408;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   25
      Left            =   4470
      TabIndex        =   2
      Top             =   6225
      Width           =   270
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2408;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國外案本所案號："
      Height          =   180
      Index           =   10
      Left            =   2955
      TabIndex        =   1
      Top             =   6225
      Width           =   1440
   End
End
Attribute VB_Name = "frm100101_K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ; lbl1(index)、txt1(10)改為txtEP12、txtCP64
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create By Sindy 2012/5/21 從frm100101_F拆出來
Option Explicit

Dim i As Integer, j As Integer, s As Integer, strSql As String
'紀錄作用按鍵
Public cmdState As Integer
Public CmdFormName As String
Dim m_SqlGrpStr1 As String, m_SqlGrpStr2 As String, m_SqlGrpStr3 As String, m_SqlGrpStr4 As String, m_SqlGrpStr5 As String
'Add by Sindy 2023/12/22
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
'2023/12/22 END


Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
'         Me.Hide
         If frm090201_b_1.Process(lbl1(3)) Then
            frm090201_b_1.cmdOK(1).Visible = False
            frm090201_b_1.Frame1.Visible = False
            frm090201_b_1.Frame2.Enabled = False
            frm090201_b_1.Show vbModal
         End If
         Unload frm090201_b_1
         Set frm090201_b_1 = Nothing
'         Me.Show
      Case 1
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         CmdFormName = ""
      'Add By Sindy 2018/4/19
      Case 2 '承辦歷程
         frm100101_F_2.Hide
         frm100101_F_2.m_EEP01 = lbl1(3) '總收文號
         frm100101_F_2.SetParent Me
         If frm100101_F_2.QueryData = True Then
            frm100101_F_2.Show
            Me.Hide
         End If
         Exit Sub
   End Select
End Sub

Private Sub cmd_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Private Sub CmdSave_Click()
Dim Cancel As Boolean
   
On Error GoTo ErrHnd
   
   'Modify By Sindy 2018/8/28 智權人員會稿日 ==> 改成客戶會稿日, 不開放維護
   'If txtEP36 = "" And txtEP37 = "" Then
   If txtEP36 = "" Then
      MsgBox "請輸入資料！", vbExclamation
      txtEP36.SetFocus
      Exit Sub
   Else
      Cancel = False
      txtEP36_Validate Cancel
      If Cancel = True Then
         txtEP36.SetFocus
         Exit Sub
      End If
'      txtEP37_Validate Cancel
'      If Cancel = True Then
'         txtEP37.SetFocus
'         Exit Sub
'      End If
      
      cnnConnection.BeginTrans
      cnnConnection.Execute "begin user_data.user_formname:='" & Me.Name & "';end;"
'      strSql = "update engineerprogress" & _
'                        " set ep36=" & IIf(Trim(txtEP36) = "", "null", DBDATE(txtEP36)) & _
'                            " ,ep37=" & IIf(Trim(txtEP37) = "", "null", DBDATE(txtEP37)) & _
'                   " where ep02='" & Trim(Lbl1(3)) & "'"
      strSql = "update engineerprogress" & _
                        " set ep36=" & IIf(Trim(txtEP36) = "", "null", DBDATE(txtEP36)) & _
                   " where ep02='" & Trim(lbl1(3)) & "'"
      cnnConnection.Execute strSql
      cnnConnection.Execute "begin user_data.user_formname:=Null;end;"
      cnnConnection.CommitTrans
      'MsgBox "存檔完成！", vbExclamation
      CmdFormName = ""
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
   End If
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.Execute "begin user_data.user_formname:=Null;end;"
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
   
   Frame2.BorderStyle = 0 'Add By Sindy 2024/12/6
End Sub

'Add By Sindy 2021/9/3 抓法務顧問資料
Sub Process_LH(strText As String)
Dim strCP10 As String, strCP14 As String, strCP13 As String, strSales As String

   '專利/商標種類
   Label1(16).Visible = False
   lbl1(13).Visible = False
   '是否通知客戶
   Label1(3).Visible = False
   txt1(9).Visible = False
   Label1(22).Visible = False
                       'Modified by Morgan 2022/12/6 修正LA案件永遠顯示閉卷問題
                       strSql = "SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(LC05,NVL(LC06,LC07)),EP09,CP26,EP07,'' as ptm,EP04,decode(LC15,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,LC08,cp49,EP27,EP31,cp13,ep05,LC15 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,LC11 as cuno,cp111,cp112,ep28,ep32,ep33,na03,ep36,ep37,ep38,cp143,cp64,cp144,cp14,s5.ST04,cp01,CP118,ep40" & _
                                " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,LawCase,nation WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr3 & ") and LC15=na01(+) "
   strSql = strSql + " UNION all SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,HC06,EP09,CP26,EP07,'',EP04,decode('000','000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,hc09,'*',EP27,EP31,cp13,ep05,'000' as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,HC05 as cuno,cp111,cp112,ep28,ep32,ep33,na03,ep36,ep37,ep38,cp143,cp64,cp144,cp14,s5.ST04,cp01,CP118,ep40 " & _
                                " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,HireCase,nation WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr4 & ") and '000'=na01(+) "
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         .MoveFirst
         For i = 0 To 29
            '2011/6/2 刪除lbll(24)
            If i <> 24 And i <> 4 And i <> 16 And i <> 26 And i <> 29 Then
               lbl1(i) = CheckStr(.Fields(i))
            End If
         Next i
         
         strCP10 = "" & .Fields("cp10")
         strCP14 = "" & .Fields("cp14")
         strCP13 = "" & .Fields("cp13")
         If .Fields("st04") = "1" Then
            strSales = strCP13
         Else
            strSales = ChkMailId(strCP13)
         End If
         
         'Add By Sindy 2018/4/20 電子送件
         If Not IsNull(.Fields("CP118")) Then
            lblEApp.Visible = True
         Else
            lblEApp.Visible = False
         End If
         '增加顯示判發人
         If "" & adoRecordset.Fields("EP40") <> "" Then
            txt1(17).Text = adoRecordset.Fields("EP40")
            lbl1(35).Caption = GetPrjSalesNM(txt1(17).Text)
         Else
            txt1(17).Text = ""
            lbl1(35).Caption = ""
         End If
         '2018/4/20 END
         
         lbl1(33) = ""
         lbl1(33) = ChangeWStringToTString(CheckStr(.Fields("ep28")))
            
         'If txtEP36.Visible = True Then
            txtEP36 = ChangeWStringToTString(CheckStr(.Fields("EP36")))
            txtEP37 = ChangeWStringToTString(CheckStr(.Fields("EP37")))
         'End If
         
         If IsNull(.Fields(31).Value) <> 0 Then
            Me.lblClose.Caption = ""
         Else
            Me.lblClose.Caption = "已閉卷"
         End If

         txtEP12.Text = "" & .Fields("ep12")
         txtCP64.Text = "" & .Fields("cp64")
      End If
       
   End With
   CheckOC
   
   txt1(1).Text = lbl1(6).Caption
   txt1(2).Text = ChangeWStringToTString(lbl1(8).Caption)
   txt1(3).Text = ChangeWStringToTString(lbl1(10).Caption)
   txt1(4).Text = ChangeWStringToTString(lbl1(12).Caption)
   txt1(7).Text = ChangeWStringToTString(lbl1(18).Caption) '會稿完成日
   txt1(8).Text = ChangeWStringToTString(lbl1(20).Caption)
   txt1(9).Text = lbl1(22).Caption
   Me.txt1(12).Text = ChangeTDateStringToTString(Me.lbl1(2).Caption)
   Me.txt1(12).Locked = False
   txt1(18) = lbl1(33)
   Me.txt1(12).Locked = True
   '核稿人
   txt1(5).Text = lbl1(14).Caption
   lbl1(14).Caption = GetPrjSalesNM(txt1(5).Text)
   
   For i = 0 To 18
      'If i <> 0 And i <> 5 And i <> 6 And i <> 7 And i <> 13 And i <> 14 And i <> 15 And i <> 16 And i <> 17 Then
      If i <> 0 And i <> 10 And i <> 6 And i <> 15 And i <> 16 And i <> 17 Then
'         If i <> 10 Then
'            txt1(i).Enabled = False
'         Else
            txt1(i).Locked = True
'         End If
      End If
   Next i
End Sub

Sub Process(strText As String)
Dim strCP10 As String, strCP14 As String, strCP13 As String, strSales As String
Dim oLbl As Object
Dim oTxt1 As Object
Dim strCP01 As String, strSys As String
   
   pub_QL05 = ";總收文號：" & strText & "(承辦進度)" 'Add By Sindy 2025/8/7
   
   m_SqlGrpStr1 = SQLGrpStr("", 1)
   m_SqlGrpStr2 = SQLGrpStr("", 2)
   m_SqlGrpStr3 = SQLGrpStr("", 3)
   m_SqlGrpStr4 = SQLGrpStr("", 4)
   m_SqlGrpStr5 = SQLGrpStr("", 5)
   
   'Modify By Sindy 2013/9/27
   cmdSave.Visible = False
   Label1(44).Visible = False
   txtEP36.Visible = False
   txtEP36.Enabled = False
'   Label1(45).Visible = False
'   txtEP37.Visible = False
'   txtEP37.Enabled = False
   '2013/9/27 END
   
   '***** 清除欄位值 *****
   For Each oLbl In lbl1
      oLbl.Caption = ""
   Next
   Me.lblClose.Caption = ""
   For Each oTxt1 In txt1
      oTxt1.Text = ""
   Next
   txtEP36 = ""
   txtEP37 = ""
   Label1(2).Visible = False: textCP143.Visible = False
   '***** END *****
   
   'Add By Sindy 2021/9/3
   strSql = "SELECT cp01,cp09 FROM CASEPROGRESS WHERE CP09='" & strText & "'"
   CheckOC2
   adoRecordset1.CursorLocation = adUseClient
   adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
      strCP01 = adoRecordset1.Fields("cp01")
   End If
   CheckOC2
   strSys = CheckSys(strCP01)
   If InStr("3,4,7,8", strSys) > 0 Then '法務,顧服
      Call Process_LH(strText)
      Exit Sub
   End If
   '2021/9/3 END
   'Add By Sindy 2024/6/13 + ,tm72,tm137,tm138
   'Modify By Sindy 2024/12/6 + ,s4.st02 ep03nm
   'Modify by Sindy 2025/7/28 +,EP39,EP42
                       strSql = "SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(TM05,NVL(TM06,TM07)),EP09,CP26,EP07,decode(tm10,'000',ptm03,ptm04),EP04,decode(tm10,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,TM29,cp49,EP27,EP31,cp13,ep05,tm10 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,tm23 as cuno,cp111,cp112,ep28,ep32,ep33,na03,ep36,ep37,ep38,cp143,cp64,cp144,cp14,s5.ST04,cp01,CP118,ep40,TM136,CP141,CP142,CP164,tm72,tm137,tm138,s4.st02 ep03nm,EP39,EP42" & _
                                " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,TRADEMARK,nation WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr2 & ") and tm10=na01(+) "
   strSql = strSql + " UNION all SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(SP05,NVL(SP06,SP07)),EP09,CP26,EP07,'',EP04,decode(sp09,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,SP15,'*',EP27,EP31,cp13,ep05,sp09 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,sp08 as cuno,cp111,cp112,ep28,ep32,ep33,na03,ep36,ep37,ep38,cp143,cp64,cp144,cp14,s5.ST04,cp01,CP118,ep40,'' TM136,CP141,CP142,CP164,'' tm72,'' tm137,'' tm138,s4.st02 ep03nm,EP39,EP42" & _
                                " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,SERVICEPRACTICE,nation WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr5 & ") and sp09=na01(+) "
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         pub_QL05 = ";本所案號：" & .Fields(7) & pub_QL05 'Add By Sindy 2025/8/7
         If pub_QL04 <> "" Then InsertQueryLog (.RecordCount) 'Add By Sindy 2025/8/7
         .MoveFirst
         For i = 0 To 29
            '2011/6/2 刪除lbll(24)
            If i <> 24 And i <> 4 And i <> 16 And i <> 26 And i <> 29 Then
               lbl1(i) = CheckStr(.Fields(i))
            End If
         Next i
         
         strCP10 = "" & .Fields("cp10")
         strCP14 = "" & .Fields("cp14")
         
         'Add By Sindy 2024/12/6
         If Left(PUB_GetST03(strCP14), 2) <> "P2" Then '外商
            Me.Frame2.Visible = True '外文核稿人
            If "" & IsNull(.Fields("EP03")) Then
               Combo4.Text = ""
            Else
               Combo4.Text = .Fields("EP03") & " ==> " & .Fields("EP03nm")
            End If
            If "" & IsNull(.Fields("EP33")) Then
               txt1(19) = ""
            Else
               txt1(19) = ChangeWStringToTString(.Fields("EP33"))
            End If
         End If
         '2024/12/6 END
         
         strCP13 = "" & .Fields("cp13")
         If .Fields("st04") = "1" Then
            strSales = strCP13
         Else
            strSales = ChkMailId(strCP13)
         End If
         
         'Add By Sindy 2018/4/20 電子送件
         If Not IsNull(.Fields("CP118")) Then
            lblEApp.Visible = True
         Else
            lblEApp.Visible = False
         End If
         '增加顯示判發人
         If "" & adoRecordset.Fields("EP40") <> "" Then
            txt1(17).Text = adoRecordset.Fields("EP40")
            lbl1(35).Caption = GetPrjSalesNM(txt1(17).Text)
         Else
            txt1(17).Text = ""
            lbl1(35).Caption = ""
         End If
         '2018/4/20 END
         
         lbl1(33) = ""
         lbl1(33) = ChangeWStringToTString(CheckStr(.Fields("ep28")))
            
         'If txtEP36.Visible = True Then
            txtEP36 = ChangeWStringToTString(CheckStr(.Fields("EP36")))
            txtEP37 = ChangeWStringToTString(CheckStr(.Fields("EP37")))
         'End If
         
         'Added by Lydia 2019/01/30 查名齊備日
         'Label1(2).Visible = False: textCP143.Visible = False
         textCP143.Tag = ChangeWStringToTString(CheckStr("" & .Fields("CP143")))
         textCP143.Text = textCP143.Tag
         If "" & .Fields("CP01") = "T" And "" & .Fields("m_country") = "000" And "" & .Fields("CP10") = 申請 Then
             Label1(2).Visible = True: textCP143.Visible = True
         End If
         'end 2019/01/30
                  
         'Add By Sindy 2024/8/7
         If Left(PUB_GetST03(strUserNum), 1) = "F" Then
            Label1(44).Visible = False
            txtEP36.Visible = False
         Else
         '2024/8/7 END
            'Modify By Sindy 2013/9/27
            If CmdFormName = UCase("frm210134") Then
               cmdSave.Visible = True
               Label1(44).Visible = True
               txtEP36.Visible = True
               txtEP36.Enabled = True
   '            Label1(45).Visible = True
   '            txtEP37.Visible = True
   '            txtEP37.Enabled = True
            Else
               'Modify By Sindy 2020/7/29
               'If PUB_ChkSalesCaDateQLim(strUserNum, strCP13, SystemNumber(Me.lbl1(7).Caption, 1)) = True Then
               'Modify By Sindy 2023/12/22 +, , bolSpecMan, strSpecCode
               Call PUB_SetFormSaleDept(strUserNum, , , , , bolSpecMan, strSpecCode, , , , , , True)
               If PUB_ChkSalePerLimit(strCP13, strUserNum, False, bolSpecMan, strSpecCode) = True Then
               '2020/7/29 END
                  Label1(44).Visible = True
                  txtEP36.Visible = True
                  txtEP36.Enabled = False
   '               Label1(45).Visible = True
   '               txtEP37.Visible = True
   '               txtEP37.Enabled = False
               End If
            End If
            '2013/9/27 END
         End If
         
         If IsNull(.Fields(31).Value) <> 0 Then
            Me.lblClose.Caption = ""
         Else
            Me.lblClose.Caption = "已閉卷"
         End If
         
         'Add By Sindy 2024/1/15
         Me.lblCertType.Caption = ""
         If .Fields("CP01") = "T" And .Fields("m_country") = "000" And strCP10 = "717" Then
            If "" & .Fields("TM136").Value = "1" Then
               lblCertType = "電子註冊證"
            End If
         End If
         Me.LblFee.Caption = ""
         If "" & .Fields("CP141").Value = "3" Then '有指定日期送件
            Me.LblFee.Caption = "指定" & ChangeWStringToTDateString("" & .Fields("CP142").Value) & _
                                IIf("" & .Fields("CP164").Value = "1", "當天", IIf("" & .Fields("CP164").Value = "2", "之前", IIf("" & .Fields("CP164").Value = "3", "之後", ""))) & "送件"
         End If
         '2024/1/15 END
         
         'Add By Sindy 2025/7/28
         If IsNull(.Fields("EP39")) Then
            txt1(13) = ""
         Else
            txt1(13) = ChangeWStringToTString(.Fields("EP39"))
         End If
         If IsNull(.Fields("EP42")) Then
            txt1(14) = ""
         Else
            txt1(14) = ChangeWStringToTString(.Fields("EP42"))
         End If
         '2025/7/28 END
         
         '91.08.14 增加若是商標案則加秀條款欄位* 就是商標   nick  start
         If CheckStr(.Fields(32).Value) <> "*" Then
            txt1(11).Text = CheckStr(.Fields(32))
            txt1(11).Visible = True 'Add By Sindy 2024/7/31
            SSTab2.TabVisible(0) = True
         Else
            txt1(11).Text = ""
            txt1(11).Visible = False 'Add By Sindy 2024/7/31
            SSTab2.TabVisible(0) = False
         End If
         
         'Add By Sindy 2024/6/13
         If "" & .Fields("tm72") <> "" Then
            txt1(0).Text = "" & .Fields("tm137")
            txt1(0).Visible = True 'Add By Sindy 2024/7/31
            SSTab2.TabVisible(1) = True
            txt1(6).Text = "" & .Fields("tm138")
            txt1(6).Visible = True 'Add By Sindy 2024/7/31
            SSTab2.TabVisible(2) = True
            If txt1(0).Text <> "" Then
               SSTab2.Tab = 1
            ElseIf txt1(6).Text <> "" Then
               SSTab2.Tab = 2
            Else
               SSTab2.Tab = 1
            End If
         Else
            txt1(0).Text = ""
            txt1(0).Visible = False 'Add By Sindy 2024/7/31
            SSTab2.TabVisible(1) = False
            txt1(6).Text = ""
            txt1(6).Visible = False 'Add By Sindy 2024/7/31
            SSTab2.TabVisible(2) = False
         End If
         '2024/6/13 END
         
         '智權人員補充資料記錄
         'Modified by Lydia 2018/12/10 開放T台灣案管控文件齊備
         'If (.Fields("CP01") = "T" Or .Fields("CP01") = "FCT") And _
            .Fields("m_country") = "000" And _
            InStr(TMdebate, strCP10) > 0 And _
            Val(DBDATE(lbl1(5))) >= Val(TMdebateStarDT) Then
         'Modified by Lydia 2022/07/15 + T大陸案之齊備日管控;  TC案之文件齊備日管控
         'If ((.Fields("CP01") = "T" Or .Fields("CP01") = "FCT") And .Fields("m_country") = "000" And InStr(TMdebate, strCP10) > 0 And Val(DBDATE(lbl1(5))) >= Val(TMdebateStarDT)) Or _
              (.Fields("CP01") = "T" And .Fields("m_country") = "000" And Val(DBDATE(lbl1(5))) >= Val(T案收文齊備啟用日)) Then
         'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
         If ((.Fields("CP01") = "T" Or .Fields("CP01") = "FCT") And InStr("000,020", .Fields("m_country")) > 0 And InStr(TMdebate, strCP10) > 0 And Not (.Fields("CP01") = "FCT" And InStr(FCT_NotTMdebate, strCP10) > 0) And Val(DBDATE(lbl1(5))) >= Val(TMdebateStarDT)) Or _
              (.Fields("CP01") = "T" And InStr("000,020", .Fields("m_country")) > 0 And Val(DBDATE(lbl1(5))) >= Val(T案收文齊備啟用日)) Or _
              (.Fields("CP01") = "TC" And InStr("000,020", .Fields("m_country")) > 0) Then
            cmd(0).Visible = True
         Else
            cmd(0).Visible = False
         End If

         txtEP12.Text = "" & .Fields("ep12")
         txtCP64.Text = "" & .Fields("cp64")
         strSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01(+) AND CP02=CM02(+) AND CP03=CM03(+) AND CP04=CM04(+) AND CP14=ST01(+) AND CP31='Y' and cp09='" & strText & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
         CheckOC2
         adoRecordset1.CursorLocation = adUseClient
         adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            lbl1(25) = CheckStr(adoRecordset1.Fields(0))
            lbl1(27) = CheckStr(adoRecordset1.Fields(1))
         Else
            lbl1(25) = ""
            lbl1(27) = ""
         End If
         CheckOC2
      Else
         If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
      End If
       
   End With
   CheckOC
   
   txt1(1).Text = lbl1(6).Caption
   txt1(2).Text = ChangeWStringToTString(lbl1(8).Caption)
   txt1(3).Text = ChangeWStringToTString(lbl1(10).Caption)
   txt1(4).Text = ChangeWStringToTString(lbl1(12).Caption)
   txt1(7).Text = ChangeWStringToTString(lbl1(18).Caption) '會稿完成日
   txt1(8).Text = ChangeWStringToTString(lbl1(20).Caption)
   txt1(9).Text = lbl1(22).Caption
   Me.txt1(12).Text = ChangeTDateStringToTString(Me.lbl1(2).Caption)
   Me.txt1(12).Locked = False
   txt1(18) = lbl1(33)
   Me.txt1(12).Locked = True
   '核稿人
   txt1(5).Text = lbl1(14).Caption
   lbl1(14).Caption = GetPrjSalesNM(txt1(5).Text)
   
   For i = 0 To 18
      If i <> 10 And i <> 15 And i <> 16 Then
'         If i <> 10 Then
            txt1(i).Locked = True
'         Else
'            'Modified by Lydia 2021/12/23
'            'txt1(i).Locked = True
'            txtEP12.Locked = True
'         End If
      End If
   Next i
   'Modified by Lydia 2021/12/23
   txtEP12.Locked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_K = Nothing
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtEP36_GotFocus()
   InverseTextBox txtEP36
End Sub

Private Sub txtEP36_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtEP36_Validate(Cancel As Boolean)
   If txtEP36 = "" Then Exit Sub
   If ChkDate(txtEP36) = False Then
      Call txtEP36_GotFocus
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub txtEP37_GotFocus()
   InverseTextBox txtEP37
End Sub

Private Sub txtEP37_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtEP37_Validate(Cancel As Boolean)
   If txtEP37 = "" Then Exit Sub
   If ChkDate(txtEP37) = False Then
      Call txtEP37_GotFocus
      Cancel = True
      Exit Sub
   End If
End Sub
