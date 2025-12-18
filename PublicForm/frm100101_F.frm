VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_F 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料查詢"
   ClientHeight    =   5820
   ClientLeft      =   110
   ClientTop       =   760
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9360
   Begin VB.TextBox txtPA162 
      Height          =   270
      Left            =   5010
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "存檔"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7320
      Style           =   1  '圖片外觀
      TabIndex        =   129
      Top             =   70
      Width           =   780
   End
   Begin VB.CommandButton cmd 
      Caption         =   "承辦歷程"
      Height          =   330
      Index           =   2
      Left            =   8160
      TabIndex        =   128
      Top             =   540
      Width           =   1125
   End
   Begin VB.CommandButton cmd 
      Caption         =   "回前畫面"
      Height          =   330
      Index           =   1
      Left            =   8160
      TabIndex        =   127
      Top             =   70
      Width           =   1125
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   11
      Left            =   8310
      MaxLength       =   7
      TabIndex        =   125
      Top             =   4200
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   10
      Left            =   8310
      MaxLength       =   7
      TabIndex        =   123
      Top             =   3000
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   24
      Left            =   8250
      MaxLength       =   1
      TabIndex        =   112
      Top             =   2670
      Width           =   270
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   23
      Left            =   7320
      MaxLength       =   6
      TabIndex        =   108
      Top             =   3930
      Width           =   920
   End
   Begin VB.CommandButton cmd 
      Caption         =   "開庭/面詢紀錄上傳"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   6240
      TabIndex        =   104
      Top             =   540
      Width           =   1860
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   21
      Left            =   8310
      MaxLength       =   7
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   1755
      Width           =   930
   End
   Begin VB.TextBox txtEP38 
      Height          =   270
      Left            =   8310
      MaxLength       =   7
      TabIndex        =   24
      Top             =   3630
      Width           =   930
   End
   Begin VB.TextBox txtEP37 
      Enabled         =   0   'False
      Height          =   270
      Left            =   8310
      MaxLength       =   7
      TabIndex        =   20
      Top             =   2070
      Width           =   930
   End
   Begin VB.TextBox txtEP36 
      Height          =   270
      Left            =   8310
      MaxLength       =   7
      TabIndex        =   12
      Top             =   1470
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   5010
      MaxLength       =   7
      TabIndex        =   25
      Top             =   3915
      Width           =   870
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   5190
      MaxLength       =   7
      TabIndex        =   23
      Top             =   3630
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   5010
      MaxLength       =   6
      TabIndex        =   21
      Top             =   2970
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   5010
      MaxLength       =   7
      TabIndex        =   19
      Top             =   2670
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   5010
      MaxLength       =   7
      TabIndex        =   18
      Top             =   1755
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   5010
      MaxLength       =   1
      TabIndex        =   17
      Top             =   2070
      Width           =   480
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   5010
      MaxLength       =   6
      TabIndex        =   16
      Top             =   1160
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   5570
      MaxLength       =   1
      TabIndex        =   15
      Top             =   4440
      Width           =   360
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   12
      Left            =   5010
      MaxLength       =   7
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   14
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   0
      Top             =   675
      Width           =   480
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   13
      Left            =   8250
      MaxLength       =   7
      TabIndex        =   13
      Top             =   890
      Width           =   855
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   5010
      MaxLength       =   7
      TabIndex        =   11
      Top             =   1500
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   15
      Left            =   3200
      MaxLength       =   3
      TabIndex        =   10
      Top             =   3230
      Width           =   525
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   17
      Left            =   8250
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1190
      Width           =   360
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   18
      Left            =   5010
      MaxLength       =   7
      TabIndex        =   8
      Top             =   2355
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   19
      Left            =   8310
      MaxLength       =   7
      TabIndex        =   7
      Top             =   3315
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   20
      Left            =   8250
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2400
      Width           =   270
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "專利相關案件"
      Height          =   330
      Left            =   90
      TabIndex        =   5
      Top             =   4470
      Width           =   1680
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   5190
      MaxLength       =   6
      TabIndex        =   22
      Top             =   3300
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "工作時數："
      Height          =   180
      Index           =   37
      Left            =   7410
      TabIndex        =   135
      Top             =   4790
      Width           =   960
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   30
      Left            =   8400
      TabIndex        =   134
      Top             =   4790
      Width           =   590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1041;459"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "核稿時數："
      Height          =   180
      Index           =   2
      Left            =   7410
      TabIndex        =   133
      Top             =   4520
      Width           =   960
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   29
      Left            =   8400
      TabIndex        =   132
      Top             =   4520
      Width           =   590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1041;459"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否提供核准分割建議:        ( Y: 是 N:否 )"
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
      Index           =   0
      Left            =   2940
      TabIndex        =   131
      Top             =   610
      Visible         =   0   'False
      Width           =   3580
   End
   Begin VB.Label LblEP42 
      AutoSize        =   -1  'True
      Caption         =   "判發完成日："
      Height          =   180
      Left            =   7230
      TabIndex        =   126
      Top             =   4250
      Width           =   1080
   End
   Begin VB.Label LblEP39 
      AutoSize        =   -1  'True
      Caption         =   "核稿完成日："
      Height          =   180
      Left            =   7230
      TabIndex        =   124
      Top             =   3030
      Width           =   1080
   End
   Begin MSForms.TextBox txtCP144 
      Height          =   560
      Left            =   1020
      TabIndex        =   121
      Top             =   30
      Width           =   6170
      VariousPropertyBits=   -1467989985
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "10883;988"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCP99 
      Height          =   840
      Left            =   2160
      TabIndex        =   120
      Top             =   3756
      Width           =   1908
      VariousPropertyBits=   -1467989985
      ScrollBars      =   2
      Size            =   "3360;1482"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Left            =   5010
      TabIndex        =   118
      Top             =   1140
      Width           =   1890
      VariousPropertyBits=   679495705
      DisplayStyle    =   3
      Size            =   "3334;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo4 
      Height          =   300
      Left            =   5190
      TabIndex        =   119
      Top             =   3300
      Width           =   2000
      VariousPropertyBits=   679495705
      DisplayStyle    =   3
      Size            =   "3528;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtEP12 
      Height          =   960
      Left            =   840
      TabIndex        =   117
      Top             =   4830
      Width           =   3270
      VariousPropertyBits=   -1466941409
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "5768;1693"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   510
      Left            =   4890
      TabIndex        =   116
      Top             =   5250
      Width           =   4340
      VariousPropertyBits=   -1466941409
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "7646;900"
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
      TabIndex        =   115
      Top             =   1350
      Visible         =   0   'False
      Width           =   890
   End
   Begin VB.Label lblCM10 
      Caption         =   "一案兩請"
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
      TabIndex        =   114
      Top             =   1640
      Visible         =   0   'False
      Width           =   830
   End
   Begin VB.Label LblEP37 
      Alignment       =   1  '靠右對齊
      Caption         =   "客戶會稿日："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   7050
      TabIndex        =   100
      Top             =   2100
      Width           =   1260
   End
   Begin VB.Label LblEP41 
      AutoSize        =   -1  'True
      Caption         =   "核稿語文：      (1.英2.日)"
      Height          =   180
      Left            =   7350
      TabIndex        =   113
      Top             =   2700
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "判發人："
      Height          =   180
      Index           =   52
      Left            =   6570
      TabIndex        =   111
      Top             =   3960
      Width           =   740
   End
   Begin VB.Label LblEP38 
      Alignment       =   1  '靠右對齊
      Caption         =   "智權人員會稿完成日："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   6510
      TabIndex        =   110
      Top             =   3690
      Width           =   1800
   End
   Begin MSForms.Label lbl1 
      Height          =   230
      Index           =   35
      Left            =   8280
      TabIndex        =   109
      Top             =   3960
      Width           =   860
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1517;406"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否為複雜或特殊案件： "
      Height          =   180
      Index           =   51
      Left            =   4140
      TabIndex        =   107
      Top             =   4230
      Width           =   2030
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(Y:是)"
      Height          =   180
      Index           =   50
      Left            =   6650
      TabIndex        =   106
      Top             =   4230
      Width           =   470
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   24
      Left            =   6200
      TabIndex        =   105
      Top             =   4230
      Width           =   420
      VariousPropertyBits=   27
      Caption         =   "中-lblFM2"
      Size            =   "741;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "報價備註："
      Height          =   180
      Index           =   49
      Left            =   90
      TabIndex        =   103
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label LblCP143 
      AutoSize        =   -1  'True
      Caption         =   "申請文件齊備日："
      Height          =   180
      Left            =   6870
      TabIndex        =   102
      Top             =   1790
      Width           =   1440
   End
   Begin VB.Label LblEP36 
      Alignment       =   1  '靠右對齊
      Caption         =   "智權人員齊備日："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   6840
      TabIndex        =   99
      Top             =   1500
      Width           =   1470
   End
   Begin VB.Label Label1 
      Caption         =   "(N:  不通知)"
      Height          =   180
      Index           =   22
      Left            =   5960
      TabIndex        =   45
      ToolTipText     =   "(N:  不通知, 自動內部收文)"
      Top             =   4490
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(Y/N)"
      Height          =   180
      Index           =   30
      Left            =   5640
      TabIndex        =   63
      Top             =   2120
      Width           =   410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文核完日："
      Height          =   180
      Index           =   41
      Left            =   7230
      TabIndex        =   27
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Label LblEP32 
      AutoSize        =   -1  'True
      Caption         =   "是否暫停核稿：      (Y：暫停)"
      Height          =   180
      Left            =   6990
      TabIndex        =   26
      Top             =   2430
      Width           =   2340
   End
   Begin VB.Label LblCP106 
      AutoSize        =   -1  'True
      Caption         =   "是否提供圖檔：       (Y：是)"
      Height          =   180
      Left            =   6990
      TabIndex        =   32
      Top             =   1230
      Width           =   2210
   End
   Begin VB.Label Label1 
      Caption         =   "承辦期限："
      Height          =   180
      Index           =   46
      Left            =   4140
      TabIndex        =   98
      Top             =   900
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "承辦天數："
      Height          =   180
      Index           =   33
      Left            =   4110
      TabIndex        =   97
      Top             =   4790
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "進度備註："
      Height          =   240
      Index           =   31
      Left            =   4140
      TabIndex        =   96
      Top             =   5280
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "取消收文日："
      Height          =   180
      Index           =   29
      Left            =   4140
      TabIndex        =   95
      Top             =   5000
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文核稿人："
      Height          =   180
      Index           =   25
      Left            =   4140
      TabIndex        =   94
      Top             =   3350
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "發文日："
      Height          =   180
      Index           =   7
      Left            =   4140
      TabIndex        =   93
      Top             =   3960
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "    核稿人："
      Height          =   180
      Index           =   6
      Left            =   3930
      TabIndex        =   92
      Top             =   3000
      Width           =   1100
   End
   Begin VB.Label Label1 
      Caption         =   "會稿完成日："
      Height          =   180
      Index           =   5
      Left            =   4140
      TabIndex        =   91
      Top             =   3660
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "是否會稿："
      Height          =   180
      Index           =   27
      Left            =   4140
      TabIndex        =   90
      Top             =   2100
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "完稿日："
      Height          =   180
      Index           =   26
      Left            =   4140
      TabIndex        =   89
      Top             =   1790
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "會稿日："
      Height          =   180
      Index           =   24
      Left            =   4140
      TabIndex        =   88
      Top             =   2700
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "齊備日："
      Height          =   180
      Index           =   23
      Left            =   4140
      TabIndex        =   87
      Top             =   1500
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員："
      Height          =   180
      Index           =   4
      Left            =   4140
      TabIndex        =   86
      Top             =   1200
      Width           =   960
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   23
      Left            =   636
      TabIndex        =   85
      Top             =   4212
      Width           =   924
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   19
      Left            =   996
      TabIndex        =   84
      Top             =   3624
      Width           =   1104
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1947;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   17
      Left            =   996
      TabIndex        =   83
      Top             =   3336
      Width           =   1104
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1947;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   15
      Left            =   996
      TabIndex        =   82
      Top             =   3036
      Width           =   1416
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2498;466"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   13
      Left            =   1476
      TabIndex        =   81
      Top             =   2760
      Width           =   1176
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2064;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   11
      Left            =   1476
      TabIndex        =   80
      Top             =   2472
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
      Height          =   264
      Index           =   9
      Left            =   1056
      TabIndex        =   79
      Top             =   2172
      Width           =   2904
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "5115;459"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   7
      Left            =   1056
      TabIndex        =   78
      Top             =   1884
      Width           =   1836
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3228;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   5
      Left            =   1080
      TabIndex        =   77
      Top             =   1584
      Width           =   1596
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   3
      Left            =   1080
      TabIndex        =   76
      Top             =   1320
      Width           =   1596
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
      Height          =   180
      Index           =   32
      Left            =   2210
      TabIndex        =   75
      Top             =   2490
      Width           =   1070
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   28
      Left            =   5280
      TabIndex        =   74
      Top             =   5000
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2646;459"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   26
      Left            =   5100
      TabIndex        =   73
      Top             =   4790
      Width           =   590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1032;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   20
      Left            =   80
      TabIndex        =   72
      Top             =   5520
      Visible         =   0   'False
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
      Index           =   18
      Left            =   80
      TabIndex        =   71
      Top             =   5310
      Visible         =   0   'False
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
      Index           =   16
      Left            =   6150
      TabIndex        =   70
      Top             =   3350
      Width           =   990
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1746;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   230
      Index           =   14
      Left            =   5940
      TabIndex        =   69
      Top             =   3000
      Width           =   860
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1517;406"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   12
      Left            =   5040
      TabIndex        =   68
      Top             =   2700
      Visible         =   0   'False
      Width           =   1050
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1852;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   10
      Left            =   5010
      TabIndex        =   67
      Top             =   1790
      Visible         =   0   'False
      Width           =   980
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1720;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   8
      Left            =   5010
      TabIndex        =   66
      Top             =   1500
      Visible         =   0   'False
      Width           =   980
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1720;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   6
      Left            =   5150
      TabIndex        =   65
      Top             =   2070
      Visible         =   0   'False
      Width           =   560
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "979;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   4
      Left            =   5550
      TabIndex        =   64
      Top             =   1190
      Visible         =   0   'False
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
      Caption         =   "點數："
      Height          =   264
      Index           =   11
      Left            =   96
      TabIndex        =   62
      Top             =   4212
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   264
      Index           =   13
      Left            =   96
      TabIndex        =   61
      Top             =   3624
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   264
      Index           =   14
      Left            =   96
      TabIndex        =   60
      Top             =   3336
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   264
      Index           =   15
      Left            =   96
      TabIndex        =   59
      Top             =   3036
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "專利/商標種類："
      Height          =   264
      Index           =   16
      Left            =   96
      TabIndex        =   58
      Top             =   2760
      Width           =   1368
   End
   Begin VB.Label Label1 
      Caption         =   "是否算案件數："
      Height          =   264
      Index           =   17
      Left            =   96
      TabIndex        =   57
      Top             =   2472
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   264
      Index           =   18
      Left            =   96
      TabIndex        =   56
      Top             =   2172
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   264
      Index           =   19
      Left            =   96
      TabIndex        =   55
      Top             =   1884
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "收文日："
      Height          =   264
      Index           =   20
      Left            =   96
      TabIndex        =   54
      Top             =   1584
      Width           =   744
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   264
      Index           =   21
      Left            =   96
      TabIndex        =   53
      Top             =   1296
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   8
      Left            =   1356
      TabIndex        =   52
      Top             =   1020
      Width           =   744
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   0
      Left            =   648
      TabIndex        =   51
      Top             =   1020
      Width           =   636
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1111;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   1
      Left            =   2136
      TabIndex        =   50
      Top             =   996
      Width           =   1596
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
      Height          =   264
      Index           =   1
      Left            =   96
      TabIndex        =   49
      Top             =   1020
      Width           =   540
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   22
      Left            =   5570
      TabIndex        =   48
      Top             =   4460
      Visible         =   0   'False
      Width           =   450
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "794;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   264
      Index           =   21
      Left            =   996
      TabIndex        =   47
      Top             =   3924
      Width           =   1104
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1947;466"
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
      Left            =   4140
      TabIndex        =   46
      Top             =   4470
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   264
      Index           =   12
      Left            =   96
      TabIndex        =   44
      Top             =   3924
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   31
      Left            =   6780
      TabIndex        =   43
      Top             =   4790
      Width           =   560
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "988;459"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "支援時數："
      Height          =   180
      Index           =   28
      Left            =   5780
      TabIndex        =   42
      Top             =   4790
      Width           =   960
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   2
      Left            =   5090
      TabIndex        =   41
      Top             =   900
      Width           =   960
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1693;459"
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
      TabIndex        =   40
      Top             =   1910
      Width           =   950
   End
   Begin VB.Label Label1 
      Caption         =   "承辦備註："
      Height          =   180
      Index           =   34
      Left            =   90
      TabIndex        =   39
      Top             =   4860
      Width           =   740
   End
   Begin VB.Label Label1 
      Caption         =   "Claims完稿日："
      Height          =   180
      Index           =   35
      Left            =   7010
      TabIndex        =   38
      Top             =   960
      Width           =   1220
   End
   Begin VB.Label Label1 
      Caption         =   "收卷註記：              (Y：收到卷宗)"
      Height          =   180
      Index           =   36
      Left            =   90
      TabIndex        =   37
      Top             =   770
      Width           =   2760
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "基數："
      Height          =   180
      Left            =   2850
      TabIndex        =   36
      Top             =   2780
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "加乘註記："
      Height          =   180
      Left            =   2196
      TabIndex        =   35
      Top             =   3336
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "加乘註記/基數修改理由："
      Height          =   180
      Left            =   2100
      TabIndex        =   34
      Top             =   3564
      Width           =   2028
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   32
      Left            =   3470
      TabIndex        =   33
      Top             =   2780
      Width           =   410
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "714;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "預定會稿日："
      Height          =   180
      Index           =   39
      Left            =   3950
      TabIndex        =   31
      Top             =   2390
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "會稿加乘："
      Height          =   180
      Index           =   40
      Left            =   2490
      TabIndex        =   30
      Top             =   3030
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   33
      Left            =   5040
      TabIndex        =   29
      Top             =   2400
      Width           =   1080
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1905;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   260
      Index           =   34
      Left            =   3470
      TabIndex        =   28
      Top             =   3030
      Width           =   410
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "714;450"
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
   Begin MSForms.Label Label2 
      Height          =   860
      Left            =   2160
      TabIndex        =   122
      Top             =   3750
      Width           =   1910
      Caption         =   "Label2"
      Size            =   "3360;1508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm100101_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ; lbl1(index)、Combo2、Combo4、txt1(22)改為txtCP144、txt1(16)改為txtCP99、Label2
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer, s As Integer, strSql As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Public CmdFormName As String
'Add by Sindy 2023/12/22
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
'2023/12/22 END


'92.04.16 nick
Public Sub PubShowNextData()
    'add by nickc 2006/04/27
    If cmdState = 100 Then
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
        CmdFormName = "" 'Add By Sindy 2010/10/29
    'Added by Sindy 2013/5/16
    ElseIf cmdState = 200 Then
         frm100101_F_2.Hide
         frm100101_F_2.m_EEP01 = lbl1(3) '總收文號
         frm100101_F_2.SetParent Me
         If frm100101_F_2.QueryData = True Then
            frm100101_F_2.Show
            Me.Hide
         End If
         Exit Sub
    'Added by Morgan 2012/4/16
    ElseIf cmdState = 300 Then
         With frm090201_2_4
            .m_Key = lbl1(3)
            .cmdOK(0).Visible = False
            .cmdAddAtt.Visible = False
            .cmdRemAtt.Visible = False
            .Show vbModal
         End With
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
   '92.04.16 nick 紀錄作用按鍵
   If Index = 1 Then
      cmdState = 100
   'Added by Sindy 2013/5/16
   ElseIf Index = 2 Then
      cmdState = 200
   'end 2013/5/16
   'Added by Morgan 2012/4/16
   ElseIf Index = 0 Then
      cmdState = 300
   'end 2012/4/16
   End If
   PubShowNextData
   Exit Sub
   '92.04.16 nick 以下無效
   Me.Hide
End Sub

'2006/04/27  nick  加畫面顯示其他國外案
Private Sub Cmd1_Click()
   'CmdState = 200
   cmdState = -1
   If fnSaveParentForm(Me) = False Then
       Me.Enabled = True
       Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_F_1.Show
   frm100101_F_1.StrMenu (lbl1(7).Caption)
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2010/10/29
Private Sub CmdSave_Click()
Dim Cancel As Boolean
   
On Error GoTo ErrHnd
   
   'Modify By Sindy 2018/8/28 智權人員會稿日 ==> 改成客戶會稿日, 不開放維護
   'If txtEP36 = "" And txtEP37 = "" And txtEP38 = "" Then
   If txtEP36 = "" And txtEP38 = "" Then
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
      txtEP38_Validate Cancel
      If Cancel = True Then
         txtEP38.SetFocus
         Exit Sub
      End If
      
      cnnConnection.BeginTrans
      cnnConnection.Execute "begin user_data.user_formname:='" & Me.Name & "';end;"
'      strSql = "update engineerprogress" & _
'                        " set ep36=" & IIf(Trim(txtEP36) = "", "null", DBDATE(txtEP36)) & _
'                            " ,ep37=" & IIf(Trim(txtEP37) = "", "null", DBDATE(txtEP37)) & _
'                            " ,ep38=" & IIf(Trim(txtEP38) = "", "null", DBDATE(txtEP38)) & _
'                   " where ep02='" & Trim(lbl1(3)) & "'"
      strSql = "update engineerprogress" & _
                        " set ep36=" & IIf(Trim(txtEP36) = "", "null", DBDATE(txtEP36)) & _
                            " ,ep38=" & IIf(Trim(txtEP38) = "", "null", DBDATE(txtEP38)) & _
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
Dim ADORECORDSET66 As New ADODB.Recordset

   bolToEndByNick = False
   MoveFormToCenter Me
   '92.04.16 nick
   cmdState = -1
   
   'Add By Sindy 2013/5/16
'   If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      Me.cmd(2).Visible = True  '承辦歷程
'   Else
'      Me.cmd(2).Visible = False
'   End If
   '2013/5/16 End
   'Modify by Amy 2014/09/22 取消工程師輸入本所期限
   txt1(13).Visible = False
   Label1(35).Visible = False '承辦人輸入本所期限：
End Sub

Sub Process(strText As String)
Dim m_SqlGrpStr1 As String, m_SqlGrpStr2 As String, m_SqlGrpStr3 As String, m_SqlGrpStr4 As String, m_SqlGrpStr5 As String
Dim strCP10 As String, strCP14 As String, strCP13 As String, strSales As String
   
   pub_QL05 = ";總收文號：" & strText & "(承辦進度)" 'Add By Sindy 2025/8/7
   m_SqlGrpStr1 = SQLGrpStr("", 1)
   m_SqlGrpStr2 = SQLGrpStr("", 2)
   m_SqlGrpStr3 = SQLGrpStr("", 3)
   m_SqlGrpStr4 = SQLGrpStr("", 4)
   m_SqlGrpStr5 = SQLGrpStr("", 5)
   
   'Modify By Sindy 2013/9/27
   cmdSave.Visible = False
   LblEP36.Visible = False
   txtEP36.Visible = False
   txtEP36.Enabled = False
'   Label1(45).Visible = False
'   txtEP37.Visible = False
'   txtEP37.Enabled = False
   LblEP38.Visible = False
   txtEP38.Visible = False
   txtEP38.Enabled = False
   '2013/9/27 END
   
   For i = 0 To 28 '29
       If i <> 24 Then lbl1(i) = ""   '2011/6/2 刪除lbll(24)
   Next i
   Me.lbl1(31).Caption = ""
   Me.lblClose.Caption = ""
   
   'Add By Sindy 2010/10/29
   'If txtEP36.Visible = True Then
       txtEP36 = ""
       txtEP37 = ""
       txtEP38 = ""
   'End If
   '2010/10/29 End
   
   For i = 0 To 14
       txt1(i) = ""
   Next i
   
   'Add By Sindy 2014/9/2
   txt1(23).Text = ""
   lbl1(35).Caption = ""
   txt1(5).Text = ""
   '2014/9/2 End
   
'   '2011/5/26 ADD BY SONIA cp144報價備註只有專利處及電腦中心可以看
'   If Pub_StrUserSt03 = "M51" Or Mid(Pub_StrUserSt03, 1, 2) = "P1" Then
'      '報價備註
'      Label1(49).Visible = True
'      txtCP144.Visible = True
'      'Add By Sindy 2016/2/22
'      '承辦備註
'      Label1(34).Visible = True
'      txtEP12.Visible = True
'      '2016/2/22 END
'   Else
'      '報價備註
'      Label1(49).Visible = False
'      txtCP144.Visible = False
'      'Add By Sindy 2016/2/22
'      '承辦備註
'      Label1(34).Visible = False
'      txtEP12.Visible = False
'      '2016/2/22 END
'   End If
'   '2011/5/26 END
   
Dim arrCaseNo '本所案號
   '預設為商標案件
   'm_blnTMCase = False
   'Modify By Sindy 2010/10/29 增加ep36,ep37,ep38
   '2011/5/24 modify by sonia 增加cp143申請文件齊備日
   '2011/5/26 modify by sonia 增加cp64,cp144報價備註
   'Modified by Morgan 2012/4/16 +CP14
   'Modified by Morgan 2012/7/23 +CP147
   'Modify by Sindy 2013/9/5 +EP40
   'Modify by Sindy 2015/3/16 +EP41
   'Modify by Sindy 2023/10/30 +,EP39,EP42,CP113,CP114
                        strSql = "SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(PA05,NVL(PA06,PA07)),EP09,CP26,EP07,DECODE(PA09,'000',PTM03,PTM04),EP04,decode(pa09,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,PA57,'*',EP27,EP31,cp13,ep05,pa09 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,pa26 as cuno,cp111,cp112,ep28,ep32,ep33,na03,ep36,ep37,ep38,cp143,cp64,cp144,cp14,s5.ST04,CP147,EP40,EP41,CP118,EP39,EP42,CP113,CP114 " & _
                                       " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT,nation WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr1 & ") and pa09=na01(+) "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(TM05,NVL(TM06,TM07)),EP09,CP26,EP07,decode(tm10,'000',ptm03,ptm04),EP04,decode(tm10,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,TM29,cp49,EP27,EP31,cp13,ep05,tm10 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,tm23 as cuno,cp111,cp112,ep28,ep32,ep33,na03,ep36,ep37,ep38,cp143,cp64,cp144,cp14,s5.ST04,CP147,EP40,EP41,CP118,EP39,EP42,CP113,CP114 " & _
                                       " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,TRADEMARK,nation WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr2 & ") and tm10=na01(+) "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(LC05,NVL(LC06,LC07)),EP09,CP26,EP07,'',EP04,decode(lc15,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,LC08,'*',EP27,EP31,cp13,ep05,lc15 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,lc11 as cuno,cp111,cp112,ep28,ep32,ep33,na03,ep36,ep37,ep38,cp143,cp64,cp144,cp14,s5.ST04,CP147,EP40,EP41,CP118,EP39,EP42,CP113,CP114 " & _
                                       " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,LAWCASE,nation WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr3 & ") and lc15=na01(+) "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,HC06,EP09,CP26,EP07,'',EP04,CPM03,EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,HC09,'*',EP27,EP31,cp13,ep05,'000' as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,hc05 as cuno,cp111,cp112,ep28,ep32,ep33,na03,ep36,ep37,ep38,cp143,cp64,cp144,cp14,s5.ST04,CP147,EP40,EP41,CP118,EP39,EP42,CP113,CP114 " & _
                                       " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,HIRECASE,nation WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr4 & ") and '000'=na01(+) "
   strSql = strSql + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(SP05,NVL(SP06,SP07)),EP09,CP26,EP07,'',EP04,decode(sp09,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,SP15,'*',EP27,EP31,cp13,ep05,sp09 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,sp08 as cuno,cp111,cp112,ep28,ep32,ep33,na03,ep36,ep37,ep38,cp143,cp64,cp144,cp14,s5.ST04,CP147,EP40,EP41,CP118,EP39,EP42,CP113,CP114 " & _
                                       " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,SERVICEPRACTICE,nation WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & strText & "' and cp01 in (" & m_SqlGrpStr5 & ") and sp09=na01(+) "
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         pub_QL05 = ";本所案號：" & .Fields(7) & pub_QL05 'Add By Sindy 2025/8/7
         If pub_QL04 <> "" Then InsertQueryLog (.RecordCount) 'Add By Sindy 2025/8/7
         .MoveFirst
         lbl1(24) = "" & .Fields("CP147") 'Added by Morgan 2012/7/23
         For i = 0 To 30
            'Add By Sindy 2023/11/30
            If i = 29 Then
               lbl1(i) = "" & .Fields("cp114")
            ElseIf i = 30 Then
               lbl1(i) = "" & .Fields("cp113")
            '2023/11/30 END
            ElseIf i <> 24 Then
               lbl1(i) = CheckStr(.Fields(i)) '2011/6/2 刪除lbll(24)
            End If
         Next i
         
         strCP10 = "" & .Fields("cp10")
         'Added by Morgan 2012/4/16
         strCP14 = "" & .Fields("cp14")
         strCP13 = "" & .Fields("cp13")
         If .Fields("st04") = "1" Then
            strSales = strCP13
         Else
            strSales = ChkMailId(strCP13)
         End If
         'end 2012/4/16
         
         'Added by Morgan 2013/5/7 電子送件
         If Not IsNull(.Fields("CP118")) Then
            lblEApp.Visible = True
         Else
            lblEApp.Visible = False
         End If
         'end 2013/5/7
         'Add By Sindy 2015/3/13 一案兩請
         strExc(1) = SystemNumber(lbl1(7).Caption, 1)
         strExc(2) = SystemNumber(lbl1(7).Caption, 2)
         strExc(3) = SystemNumber(lbl1(7).Caption, 3)
         strExc(4) = SystemNumber(lbl1(7).Caption, 4)
         strSql = "select * from casemap where cm01='" & strExc(1) & "' and cm02='" & strExc(2) & "' and cm03='" & strExc(3) & "' and cm04='" & strExc(4) & "' and cm10='3'" & _
                  " Union select * from casemap where cm05='" & strExc(1) & "' and cm06='" & strExc(2) & "' and cm07='" & strExc(3) & "' and cm08='" & strExc(4) & "' and cm10='3'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            lblCM10.Visible = True
         Else
            lblCM10.Visible = False
         End If
         '2015/3/13 END
         
         lbl1(33) = ""
         lbl1(34) = ""
   
         lbl1(34) = CheckStr(.Fields("cp111"))
         lbl1(33) = ChangeWStringToTString(CheckStr(.Fields("ep28")))
            
         'Add By Sindy 2010/10/29
         'If txtEP36.Visible = True Then
            txtEP36 = ChangeWStringToTString(CheckStr(.Fields("EP36")))
            txtEP37 = ChangeWStringToTString(CheckStr(.Fields("EP37")))
            txtEP38 = ChangeWStringToTString(CheckStr(.Fields("EP38")))
         'End If
         '2010/10/29 End
         
         'Modify By Sindy 2013/9/27
         If CmdFormName = UCase("frm210134") Then
            cmdSave.Visible = True
            LblEP36.Visible = True
            txtEP36.Visible = True
            txtEP36.Enabled = True
'            Label1(45).Visible = True
'            txtEP37.Visible = True
'            txtEP37.Enabled = True
            LblEP38.Visible = True
            txtEP38.Visible = True
            txtEP38.Enabled = True
         Else
            'Modify By Sindy 2020/7/29
            'If PUB_ChkSalesCaDateQLim(strUserNum, strCP13, SystemNumber(Me.lbl1(7).Caption, 1)) = True Then
            'Modify By Sindy 2023/12/22 +, , bolSpecMan, strSpecCode
            Call PUB_SetFormSaleDept(strUserNum, , , , , bolSpecMan, strSpecCode, , , , , , True)
            If PUB_ChkSalePerLimit(strCP13, strUserNum, False, bolSpecMan, strSpecCode) = True Then
            '2020/7/29 END
               LblEP36.Visible = True
               txtEP36.Visible = True
               txtEP36.Enabled = False
'               Label1(45).Visible = True
'               txtEP37.Visible = True
'               txtEP37.Enabled = False
               LblEP38.Visible = True
               txtEP38.Visible = True
               txtEP38.Enabled = False
            End If
         End If
         '2013/9/27 End
         
         lbl1(32).Caption = "" & .Fields("CP97").Value
         txt1(15).Text = "" & .Fields("cp98").Value
         txtCP99.Text = "" & .Fields("cp99").Value
         'add by nickc 2007/10/11
         Label2.Caption = txtCP99.Text
         
         txt1(17).Text = "" & .Fields("cp106").Value
         
         If IsNull(.Fields(31).Value) <> 0 Then
            Me.lblClose.Caption = ""
         Else
            Me.lblClose.Caption = "已閉卷"
         End If
         
         'Add By Sindy 2016/2/22 承辦備註
         txtEP12.Text = "" & .Fields("ep12")
         '2016/2/22 END
         
         If IsNull(.Fields("EP27")) Then
            txt1(14) = ""
         Else
            txt1(14) = "Y"
         End If
         If IsNull(.Fields("EP31")) Then
            txt1(13) = ""
         Else
            txt1(13) = ChangeWStringToTString(.Fields("EP31"))
         End If
         If IsNull(.Fields("EP33")) Then
            txt1(19) = ""
         Else
            txt1(19) = ChangeWStringToTString(.Fields("EP33"))
         End If
         'Add By Sindy 2023/10/30
         If IsNull(.Fields("EP39")) Then
            txt1(10) = ""
         Else
            txt1(10) = ChangeWStringToTString(.Fields("EP39"))
         End If
         If IsNull(.Fields("EP42")) Then
            txt1(11) = ""
         Else
            txt1(11) = ChangeWStringToTString(.Fields("EP42"))
         End If
         '2023/10/30 END
         If IsNull(.Fields("EP32")) Then
            txt1(20) = ""
         Else
            txt1(20) = .Fields("EP32")
         End If
         txt1(20).Tag = txt1(20).Text
         '2011/5/24 add by sonia
         txtCP64.Text = "" & .Fields("cp64")
         If IsNull(.Fields("cp143")) Then
            txt1(21) = ""
         Else
            txt1(21) = ChangeWStringToTString(.Fields("cp143"))
         End If
         '2011/5/24 end
         txtCP144.Text = "" & .Fields("cp144") '2011/5/26 add by sonia 報價備註
         
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
      '        strSQL = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(lbl1(7), 1) & "' and ibf02='" & SystemNumber(lbl1(7), 2) & "' and ibf03='" & SystemNumber(lbl1(7), 3) & "' and ibf04='" & SystemNumber(lbl1(7), 4) & "' and ibf05='1' "
      '        CheckOC2
      '        adoRecordset1.CursorLocation = adUseClient
      '        adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
      '        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
      '            cmdPic.Caption = "已設定代表圖(&I)"
      '            cmdPic.BackColor = &HC0FFC0
      '        Else
      '            cmdPic.Caption = "未設定代表圖(&I)"
      '            cmdPic.BackColor = &HC0C0FF
      '        End If
      '        CheckOC2
         'Modified by Morgan 2024/6/5
         'Me.lbl1(31).Caption = IIf(IsNull(.Fields("CP15")), "0", "" & .Fields("CP15"))
         PUB_SettHour lbl1(3).Caption, strExc(1)
         lbl1(31) = strExc(1)
         'end 2024/6/5
         'Add By Sindy 2013/9/5 增加顯示判發人
         If "" & adoRecordset.Fields("EP40") <> "" Then
            txt1(23).Text = adoRecordset.Fields("EP40")
            lbl1(35).Caption = GetPrjSalesNM(txt1(23).Text)
         Else
            txt1(23).Text = ""
            lbl1(35).Caption = ""
         End If
         '2013/9/5 END
         'Add By Sindy 2015/3/16
         txt1(24).Text = "" & adoRecordset.Fields("EP41")
         If txt1(24).Text = "2" Then
            Label1(25).Caption = "日文核稿人："
            Label1(41).Caption = "日文核完日："
         End If
         '2015/3/16 END
         'Add By Sindy 2014/3/21 核稿人應顯示EP04欄位值
         If "" & adoRecordset.Fields("EP04") <> "" Then
            txt1(5).Text = adoRecordset.Fields("EP04")
            lbl1(14).Caption = GetPrjSalesNM(txt1(5).Text) 'Add By Sindy 2023/10/19
         Else
            txt1(5).Text = ""
            lbl1(14).Caption = "" 'Add By Sindy 2023/10/19
         End If
         '2014/3/21 END
      Else
         If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
      End If
   End With
   CheckOC
   
   txt1(0).Text = lbl1(4).Caption
   lbl1(4).Caption = GetPrjSalesNM(txt1(0).Text)
   If Trim(txt1(0).Text) = "" Then
      Combo2.Text = ""
   Else
      Combo2.Text = Trim(txt1(0).Text) & " ==> " & GetPrjSalesNM(Trim(txt1(0).Text))
   End If
   txt1(6).Text = lbl1(16).Caption
   If Trim(txt1(6).Text) = "" Then
      Combo4.Text = ""
   Else
      Combo4.Text = Trim(txt1(6).Text) & " ==> " & GetPrjSalesNM(Trim(txt1(6).Text))
   End If
   
   txt1(1).Text = lbl1(6).Caption
   txt1(2).Text = ChangeWStringToTString(lbl1(8).Caption)
   txt1(3).Text = ChangeWStringToTString(lbl1(10).Caption)
   txt1(4).Text = ChangeWStringToTString(lbl1(12).Caption)
   'Modify By Sindy 2014/3/21 Mark起來因為不應該從核判表讀取,應顯示EP04欄位值
'   If Len(lbl1(14).Caption) = 0 Then
'      strSql = "select pp04 from caseprogress,promoterproofreader where cp09='" & lbl1(3) & "' and cp01=pp01(+) and cp14=pp02(+) and cp10=pp03(+) "
'      CheckOC2
'      adoRecordset1.CursorLocation = adUseClient
'      adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If adoRecordset1.RecordCount <> 0 Then
'         txt1(5) = CheckStr(adoRecordset1.Fields(0))
'      End If
'   Else
'      txt1(5).Text = lbl1(14).Caption
'   End If
   '2014/3/21 END
   
   '會稿完成日
   txt1(7).Text = ChangeWStringToTString(lbl1(18).Caption)
   txt1(8).Text = ChangeWStringToTString(lbl1(20).Caption)
   txt1(9).Text = lbl1(22).Caption
   Me.txt1(12).Text = ChangeTDateStringToTString(Me.lbl1(2).Caption)
   Me.txt1(12).Locked = False
   txt1(18) = lbl1(33) '預定會稿日
   
   '若為個人工作管理
   Label1(36).Visible = True
   txt1(14).Visible = True
   txt1(14).Enabled = True
    If Me.lbl1(7).Caption <> "" Then
        arrCaseNo = Split(Me.lbl1(7).Caption, "-")
        If arrCaseNo(0) = "P" Or arrCaseNo(0) = "CFP" Then
            '承辦期限
            If Me.txt1(12).Text <> "" Then Me.txt1(12).Locked = True
        End If
    End If
   'C類中P,CFP不可輸發文日
   If Left(UCase(strText), 1) = "C" Then
       If SystemNumber(Me.lbl1(7).Caption, 1) = "P" Or SystemNumber(Me.lbl1(7).Caption, 1) = "CFP" Then
           Label1(3).Visible = False
           txt1(9).Visible = False
           Label1(22).Visible = False
           txt1(8).Enabled = False
           txt1(8).TabStop = False
       Else
          Label1(3).Visible = True
          txt1(9).Visible = True
          Label1(22).Visible = True
          txt1(8).Enabled = True
          txt1(8).TabStop = True
       End If
   Else
      Label1(3).Visible = False
      txt1(9).Visible = False
      Label1(22).Visible = False
      txt1(8).Enabled = False
      txt1(8).TabStop = False
   End If
   
   For i = 0 To 22
       '2011/6/2 modify by sonia
       'txt1(i).Enabled = False
       'Modified by Morgan 2016/5/23 +16
       If i <> 16 And i <> 22 Then
         txt1(i).Enabled = False
       Else
         'Mark by Lydia 2021/12/23 已預設備註Locked = True
         'txt1(i).Locked = True
       End If
       '2011/6/2 end
   Next i
   txt1(23).Enabled = False 'Add By Sindy 2013/9/5
   txt1(24).Enabled = False 'Add By Sindy 2015/3/16
   
   'Add by Morgan 2008/8/15
   'Modify By Sindy 2023/9/15 +FMP案
   If ((SystemNumber(lbl1(7).Caption, 1) = "FCP" Or SystemNumber(lbl1(7).Caption, 1) = "FG" _
        Or Left(PUB_GetST03(strCP14), 1) = "F")) Then
      Label1(34).Caption = "作業備註" 'Add By Sindy 2024/1/25 承辦備註
      cmd(0).Visible = False
      Label1(25).Caption = "外文核稿人："
      Label1(41).Caption = "外文完成日："
      LblEP41.Visible = False: txt1(24).Visible = False '核稿語文
      LblEP32.Visible = False: txt1(20).Visible = False '是否暫停核稿
      LblEP36.Visible = False: txtEP36.Visible = False '智權人員齊備日
      LblEP37.Visible = False: txtEP37.Visible = False '客戶會稿日
      LblEP38.Visible = False: txtEP38.Visible = False '智權人員會稿完成日
      LblCP106.Visible = False: txt1(17).Visible = False '是否提供圖檔
      LblCP143.Visible = False: txt1(21).Visible = False '申請文件齊備日
      'Modify By Sindy 2025/3/18 +外專繪圖同新案翻譯歷程操作方式
      If (strCP10 = "201" Or strCP10 = "931") Then
         Label1(5).Caption = "核稿期限：" 'EP08
         If strSrvDate(1) < FCP核完日改用EP39 Then 'Add By Sindy 2023/10/27
            LblEP39.Visible = False: txt1(10).Visible = False
            Label1(41).Caption = "核稿完成日：" 'EP33 'Add by Morgan 2008/12/1
         End If
         'Modify By Sindy 2023/9/15
         Label1(35).Visible = True
         Label1(35).Caption = "Claims完稿日：" 'EP31
         txt1(13).Visible = True
         Label1(6).Caption = "核稿工程師："
         '2023/9/15 END
      End If
      'Add By Sindy 2023/11/23
      '分割建議
      strExc(0) = "select pa162,DST05 from caseprogress a,patent,divsugtext" & _
                  " where cp09='" & lbl1(3) & "' and dst09='" & lbl1(3) & "'" & _
                  " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
                  " and pa08 in ('1','2')" & _
                  " and dst01(+)=pa01 and dst02(+)=pa02 and dst03(+)=pa03 and dst04(+)=pa04 and pa162='Y'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         '中說請款修正定稿文字
         strExc(0) = "select AMD05,nvl(CP27,0) CP27 from caseprogress a,patent,Amendedtext" & _
                     " where cp09='" & lbl1(3) & "' and amd09='" & lbl1(3) & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
                     " and amd01(+)=pa01 and amd02(+)=pa02 and amd03(+)=pa03 and amd04(+)=pa04"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Label1(49).Caption = "請款修正定稿文字:"
            txtCP144 = "" & RsTemp(0)
         Else
            Label1(49).Visible = False
            txtCP144.Visible = False
         End If
      Else
         Label1(0).Visible = True
         txtPA162.Visible = True
         txtPA162 = "" & RsTemp(0)
         Label1(49).Caption = "建議定稿文字:"
         txtCP144 = "" & RsTemp(1)
      End If
      '2023/11/23 END
   Else
      'Modify by Amy 2014/09/22 取消工程師輸入本所期限
      'Label1(35).Visible = True
      'Txt1(13).Visible = True
      'Txt1(13).Enabled = True
   End If
   
   'Added by Morgan 2012/4/16
   'Mofified by Morgan 2012/6/4 +213現場勘察,408面詢
   If SystemNumber(lbl1(7).Caption, 1) = "P" And (strCP10 = "211" Or strCP10 = "212" Or strCP10 = "226" Or strCP10 = "213" Or strCP10 = "408") Then
      cmd(0).Visible = True
      '查詢權限僅電腦中心、智權人員(若智權人員離職才開放其區主管)、承辦人、游經理、王副總
      'Modified by Lydia 2017/03/28 chkmailid取得可能不只一人
      'If Pub_StrUserSt03 = "M51" Or strUserNum = strCP13 Or strUserNum = strSales Or strUserNum = strCP14 Or PUB_GetST05(strUserNum) = "72" Or PUB_GetST05(strUserNum) = "71" Then
      'Modified by Morgan 2025/2/4 +P10部門
      'Modified by Morgan 2025/6/25 +79075
      If strUserNum = "79075" Or Pub_StrUserSt03 = "P10" Or Pub_StrUserSt03 = "M51" Or strUserNum = strCP13 Or InStr(strSales, strUserNum) > 0 Or strUserNum = strCP14 Or PUB_GetST05(strUserNum) = "72" Or PUB_GetST05(strUserNum) = "71" Then
         cmd(0).Enabled = True
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_F = Nothing
End Sub

'Add By Sindy 2010/11/25
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2010/10/29
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

Private Sub txtEP38_GotFocus()
   InverseTextBox txtEP38
End Sub

Private Sub txtEP38_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtEP38_Validate(Cancel As Boolean)
   If txtEP38 = "" Then Exit Sub
   If ChkDate(txtEP38) = False Then
      Call txtEP38_GotFocus
      Cancel = True
      Exit Sub
   End If
End Sub

'Added by Lydia 2016/10/27 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   OpenIme
End Sub
