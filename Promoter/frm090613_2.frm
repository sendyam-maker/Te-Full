VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090613_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件處理時間統計查詢"
   ClientHeight    =   6180
   ClientLeft      =   420
   ClientTop       =   2145
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9315
   Begin VB.CommandButton cmd1 
      Caption         =   "相關國內外案件"
      Height          =   330
      Left            =   180
      TabIndex        =   36
      Top             =   5010
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   285
      Left            =   4245
      TabIndex        =   28
      Top             =   2370
      Width           =   5100
      Begin VB.OptionButton opCP112 
         Caption         =   "不適用   會稿加乘註記 及 預定會稿日 "
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   30
         Top             =   60
         Width           =   3300
      End
      Begin VB.OptionButton opCP112 
         Caption         =   "適用"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   45
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   ")"
         Height          =   180
         Index           =   45
         Left            =   4320
         TabIndex        =   32
         Top             =   15
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "("
         Height          =   180
         Index           =   44
         Left            =   0
         TabIndex        =   31
         Top             =   15
         Width           =   60
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1170
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1395
      Width           =   1605
   End
   Begin VB.CommandButton cmd 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7920
      TabIndex        =   2
      Top             =   50
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "文件檔案(&F)"
      Height          =   400
      Index           =   0
      Left            =   5430
      TabIndex        =   1
      Top             =   50
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "申請書(&L)"
      Height          =   400
      Index           =   1
      Left            =   6750
      TabIndex        =   0
      Top             =   50
      Visible         =   0   'False
      Width           =   1080
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   315
      Left            =   5070
      TabIndex        =   9
      Top             =   1080
      Width           =   1890
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "3334;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   5280
      TabIndex        =   37
      Top             =   3585
      Visible         =   0   'False
      Width           =   900
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1587;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   20
      Left            =   8250
      TabIndex        =   35
      Top             =   3300
      Width           =   270
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "476;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   19
      Left            =   8340
      TabIndex        =   34
      Top             =   3630
      Width           =   900
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1587;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   18
      Left            =   5070
      TabIndex        =   33
      Top             =   2655
      Width           =   870
      VariousPropertyBits=   671105049
      MaxLength       =   7
      Size            =   "1535;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   17
      Left            =   8280
      TabIndex        =   27
      Top             =   1080
      Width           =   360
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "635;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   795
      Index           =   16
      Left            =   2160
      TabIndex        =   26
      Top             =   4650
      Width           =   1980
      VariousPropertyBits=   671105051
      Size            =   "3492;1402"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   15
      Left            =   3255
      TabIndex        =   25
      Top             =   3630
      Width           =   525
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "926;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   5070
      TabIndex        =   24
      Top             =   1395
      Width           =   930
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1640;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   13
      Left            =   5070
      TabIndex        =   23
      Top             =   510
      Width           =   915
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1614;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   14
      Left            =   1140
      TabIndex        =   22
      Top             =   802
      Width           =   480
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "847;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   12
      Left            =   5070
      TabIndex        =   21
      Top             =   802
      Width           =   930
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1640;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   660
      Index           =   11
      Left            =   690
      TabIndex        =   20
      Top             =   5490
      Width           =   3450
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6085;1164"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   5655
      TabIndex        =   19
      Top             =   4515
      Width           =   360
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "635;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   5100
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   1275
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "2249;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   5070
      TabIndex        =   17
      Top             =   2025
      Width           =   480
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "847;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   5070
      TabIndex        =   16
      Top             =   1710
      Width           =   915
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1614;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   5070
      TabIndex        =   15
      Top             =   2955
      Width           =   915
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1614;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   5070
      TabIndex        =   14
      Top             =   3300
      Width           =   915
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1614;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   5340
      TabIndex        =   13
      Top             =   3930
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   5070
      TabIndex        =   12
      Top             =   4230
      Width           =   870
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1535;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   600
      Index           =   10
      Left            =   5070
      TabIndex        =   11
      Top             =   4860
      Width           =   3345
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "5900;1058"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo4 
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      Top             =   3600
      Width           =   1995
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "3528;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否暫停核稿："
      Height          =   180
      Index           =   42
      Left            =   6990
      TabIndex        =   41
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(Y：暫停)"
      Height          =   180
      Index           =   43
      Left            =   8520
      TabIndex        =   117
      Top             =   3360
      Width           =   780
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   35
      Left            =   3540
      TabIndex        =   116
      Top             =   3330
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
      Height          =   255
      Index           =   34
      Left            =   5310
      TabIndex        =   115
      Top             =   2685
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
      AutoSize        =   -1  'True
      Caption         =   "會稿加乘註記："
      Height          =   180
      Index           =   40
      Left            =   2310
      TabIndex        =   114
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "預定會稿日："
      Height          =   180
      Index           =   39
      Left            =   4035
      TabIndex        =   113
      Top             =   2715
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(Y：是)"
      Height          =   180
      Index           =   38
      Left            =   8685
      TabIndex        =   112
      Top             =   1140
      Width           =   600
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   33
      Left            =   3540
      TabIndex        =   111
      Top             =   2985
      Width           =   645
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1138;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "加乘註記修改理由："
      Height          =   180
      Left            =   2250
      TabIndex        =   110
      Top             =   4470
      Width           =   1620
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "加乘註記："
      Height          =   180
      Left            =   2250
      TabIndex        =   109
      Top             =   3690
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "計件值："
      Height          =   180
      Left            =   2835
      TabIndex        =   108
      Top             =   3015
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "收卷註記：              (Y：收到卷宗)"
      Height          =   180
      Index           =   36
      Left            =   120
      TabIndex        =   107
      Top             =   862
      Width           =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人輸入本所期限："
      Height          =   180
      Index           =   35
      Left            =   3195
      TabIndex        =   106
      Top             =   570
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "條款："
      Height          =   180
      Index           =   34
      Left            =   120
      TabIndex        =   105
      Top             =   5505
      Width           =   735
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
      Left            =   3120
      TabIndex        =   104
      Top             =   2085
      Width           =   945
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   32
      Left            =   5145
      TabIndex        =   103
      Top             =   825
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
      Caption         =   "支援時數："
      Height          =   180
      Index           =   28
      Left            =   6105
      TabIndex        =   102
      Top             =   5580
      Width           =   960
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   31
      Left            =   7125
      TabIndex        =   101
      Top             =   5580
      Width           =   585
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1032;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   12
      Left            =   120
      TabIndex        =   100
      Top             =   4260
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "是否通知客戶："
      Height          =   180
      Index           =   3
      Left            =   4230
      TabIndex        =   99
      Top             =   4545
      Width           =   1350
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   21
      Left            =   1050
      TabIndex        =   98
      Top             =   4260
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2487;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   22
      Left            =   5640
      TabIndex        =   97
      Top             =   4530
      Visible         =   0   'False
      Width           =   390
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "688;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "目次："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   96
      Top             =   1140
      Width           =   540
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   2190
      TabIndex        =   95
      Top             =   1110
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
      Height          =   255
      Index           =   1
      Left            =   705
      TabIndex        =   94
      Top             =   1110
      Width           =   630
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1111;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   8
      Left            =   1410
      TabIndex        =   93
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "國外案承辦人："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   92
      Top             =   5850
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   180
      Index           =   21
      Left            =   120
      TabIndex        =   91
      Top             =   1455
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "收文日："
      Height          =   180
      Index           =   20
      Left            =   120
      TabIndex        =   90
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   19
      Left            =   120
      TabIndex        =   89
      Top             =   2085
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   18
      Left            =   120
      TabIndex        =   88
      Top             =   2415
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "是否算案件數："
      Height          =   180
      Index           =   17
      Left            =   120
      TabIndex        =   87
      Top             =   2715
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "專利/商標種類："
      Height          =   180
      Index           =   16
      Left            =   120
      TabIndex        =   86
      Top             =   3015
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   15
      Left            =   120
      TabIndex        =   85
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   14
      Left            =   120
      TabIndex        =   84
      Top             =   3690
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   83
      Top             =   3990
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "點數："
      Height          =   180
      Index           =   11
      Left            =   120
      TabIndex        =   82
      Top             =   4665
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "國外案本所案號："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   81
      Top             =   5850
      Visible         =   0   'False
      Width           =   1515
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   80
      Top             =   1140
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
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   6
      Left            =   5100
      TabIndex        =   79
      Top             =   2040
      Visible         =   0   'False
      Width           =   435
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "767;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   8
      Left            =   5100
      TabIndex        =   78
      Top             =   1425
      Visible         =   0   'False
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
      Height          =   255
      Index           =   10
      Left            =   5100
      TabIndex        =   77
      Top             =   1710
      Visible         =   0   'False
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
      Height          =   255
      Index           =   12
      Left            =   5100
      TabIndex        =   76
      Top             =   2985
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
      Height          =   255
      Index           =   14
      Left            =   6060
      TabIndex        =   75
      Top             =   3330
      Width           =   885
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1561;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   16
      Left            =   5520
      TabIndex        =   74
      Top             =   3630
      Width           =   1260
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2222;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   18
      Left            =   5700
      TabIndex        =   73
      Top             =   3960
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
      Height          =   255
      Index           =   20
      Left            =   5070
      TabIndex        =   72
      Top             =   4260
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
      Height          =   300
      Index           =   24
      Left            =   5100
      TabIndex        =   71
      Top             =   4860
      Visible         =   0   'False
      Width           =   3360
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "5927;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   26
      Left            =   5220
      TabIndex        =   70
      Top             =   5580
      Width           =   585
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1032;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   28
      Left            =   5445
      TabIndex        =   69
      Top             =   5820
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
      Height          =   180
      Index           =   32
      Left            =   2295
      TabIndex        =   68
      Top             =   2715
      Width           =   1065
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1170
      TabIndex        =   67
      Top             =   1410
      Width           =   1800
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3175;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   5
      Left            =   1170
      TabIndex        =   66
      Top             =   1710
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
      Height          =   255
      Index           =   7
      Left            =   1170
      TabIndex        =   65
      Top             =   2055
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
      Height          =   255
      Index           =   9
      Left            =   1170
      TabIndex        =   64
      Top             =   2385
      Width           =   3045
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "5371;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   11
      Left            =   1560
      TabIndex        =   63
      Top             =   2685
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
      Height          =   255
      Index           =   13
      Left            =   1560
      TabIndex        =   62
      Top             =   2985
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
      Height          =   255
      Index           =   15
      Left            =   1050
      TabIndex        =   61
      Top             =   3330
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
      Height          =   255
      Index           =   17
      Left            =   1050
      TabIndex        =   60
      Top             =   3660
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
      Height          =   255
      Index           =   19
      Left            =   1050
      TabIndex        =   59
      Top             =   3990
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
      Height          =   255
      Index           =   23
      Left            =   720
      TabIndex        =   58
      Top             =   4665
      Width           =   915
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1614;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1620
      TabIndex        =   57
      Top             =   5820
      Visible         =   0   'False
      Width           =   2040
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3598;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   30
      Left            =   1650
      TabIndex        =   56
      Top             =   5820
      Visible         =   0   'False
      Width           =   2490
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4392;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員："
      Height          =   180
      Index           =   4
      Left            =   4200
      TabIndex        =   55
      Top             =   1140
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "齊備日："
      Height          =   180
      Index           =   23
      Left            =   4230
      TabIndex        =   54
      Top             =   1455
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "會稿日："
      Height          =   180
      Index           =   24
      Left            =   4230
      TabIndex        =   53
      Top             =   3015
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "完稿日："
      Height          =   180
      Index           =   26
      Left            =   4230
      TabIndex        =   52
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "是否會稿："
      Height          =   180
      Index           =   27
      Left            =   4230
      TabIndex        =   51
      Top             =   2085
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "會稿完成日："
      Height          =   180
      Index           =   5
      Left            =   4230
      TabIndex        =   50
      Top             =   3975
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "核稿人："
      Height          =   180
      Index           =   6
      Left            =   4230
      TabIndex        =   49
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "發文日："
      Height          =   180
      Index           =   7
      Left            =   4230
      TabIndex        =   48
      Top             =   4275
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文核稿人："
      Height          =   180
      Index           =   25
      Left            =   4230
      TabIndex        =   47
      Top             =   3690
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "取消收文日："
      Height          =   180
      Index           =   29
      Left            =   4230
      TabIndex        =   46
      Top             =   5820
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "備註："
      Height          =   180
      Index           =   31
      Left            =   4230
      TabIndex        =   45
      Top             =   4920
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "承辦天數："
      Height          =   180
      Index           =   33
      Left            =   4230
      TabIndex        =   44
      Top             =   5580
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "承辦期限："
      Height          =   180
      Index           =   46
      Left            =   4200
      TabIndex        =   43
      Top             =   862
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否提供圖檔："
      Height          =   180
      Index           =   37
      Left            =   7020
      TabIndex        =   42
      Top             =   1140
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文核完日："
      Height          =   180
      Index           =   41
      Left            =   7290
      TabIndex        =   40
      Top             =   3690
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(Y/N)"
      Height          =   180
      Index           =   30
      Left            =   5730
      TabIndex        =   39
      Top             =   2085
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "(N:  不通知, 自動內部收文)"
      Height          =   180
      Index           =   22
      Left            =   6045
      TabIndex        =   38
      ToolTipText     =   "(N:  不通知, 自動內部收文)"
      Top             =   4560
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "國外案承辦人："
      Height          =   180
      Index           =   9
      Left            =   540
      TabIndex        =   7
      Top             =   6990
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "國外案本所案號："
      Height          =   180
      Index           =   10
      Left            =   540
      TabIndex        =   6
      Top             =   6615
      Width           =   1470
   End
   Begin MSForms.Label lbl1 
      Height          =   180
      Index           =   25
      Left            =   2055
      TabIndex        =   5
      Top             =   6630
      Width           =   3225
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
      Index           =   27
      Left            =   1875
      TabIndex        =   4
      Top             =   6990
      Width           =   3390
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2408;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   29
      Left            =   4260
      TabIndex        =   3
      Top             =   6240
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
End
Attribute VB_Name = "frm090613_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/08 改成Form2.0 ; lbl1(index)、txt1(index)、Combo2、Combo4
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
'Modified by Morgan 2013/4/19 原105條件要再加125
Option Explicit

Dim s As Integer, i As Integer
Dim Fobj As FileSystemObject
Dim StrCP10nick As String
Dim AdoRs As New ADODB.Recordset
Dim StrNewSQL As String
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


Private Sub cmd_Click()
Me.Hide
frm090613_1.Show
Unload Me
End Sub

'2006/04/27  nick  加畫面顯示其他國外案
Private Sub Cmd1_Click()
Me.Hide
Screen.MousePointer = vbHourglass
frm090613_2_1.Show
frm090613_2_1.StrMenu (lbl1(7).Caption)
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim Wxdc As New Word.Application
StrCP10nick = Trim(lbl1(15).Caption)
If StrCP10nick = "修正" Then
   Set AdoRs = New ADODB.Recordset
   StrNewSQL = "SELECT DECODE(PA09,'000',CPM03,CPM04) FROM CASEPROGRESS,CASEPROPERTYMAP,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105','125') "
   StrNewSQL = StrNewSQL & " UNION all  SELECT DECODE(TM10,'000',CPM03,CPM04) FROM CASEPROGRESS,CASEPROPERTYMAP,TRADEMARK WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105','125')  "
   StrNewSQL = StrNewSQL & " UNION all  SELECT DECODE(LC15,'000',CPM03,CPM04) FROM CASEPROGRESS,CASEPROPERTYMAP,LAWCASE WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105','125') "
   StrNewSQL = StrNewSQL & " UNION all  SELECT CPM03                          FROM CASEPROGRESS,CASEPROPERTYMAP,HIRECASE WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105','125') "
   StrNewSQL = StrNewSQL & " UNION all  SELECT DECODE(SP09,'000',CPM03,CPM04) FROM CASEPROGRESS,CASEPROPERTYMAP,SERVICEPRACTICE WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP01='" & SystemNumber(Trim(lbl1(7).Caption), 1) & "' AND CP02='" & SystemNumber(Trim(lbl1(7).Caption), 2) & "' AND CP03='" & SystemNumber(Trim(lbl1(7).Caption), 3) & "' AND CP04='" & SystemNumber(Trim(lbl1(7).Caption), 4) & "' AND CP10 IN ('101','102','103','104','105','125') "
   AdoRs.CursorLocation = adUseClient
   AdoRs.Open StrNewSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRs.RecordCount <> 0 Then
      StrCP10nick = CheckStr(AdoRs.Fields(0))
   Else
      s = MsgBox("此案號無申請案之文件檔！！", , "沒有檔案！！")
      AdoRs.Close
      Set AdoRs = Nothing
      Exit Sub
   End If
   AdoRs.Close
   Set AdoRs = Nothing
End If
Dim DFileName As String    '應該存放檔名
Dim DSFileName As String   '範本檔名
Dim DFilePath As String    '應該存放路徑
DFileName = ChangeFileName(Trim(lbl1(7).Caption), Trim(lbl1(15).Caption), Trim(lbl1(3).Caption))
DFilePath = GetDocFilePath(DFileName) & "\"
DFileName = DFileName & ".doc"
DSFileName = SystemNumber(Trim(lbl1(7).Caption), 1) & Trim(lbl1(15).Caption) & ".doc"
Select Case Index
Case 0
     If Len(Dir(DFilePath & DFileName)) <> 0 Then
        Set Wxdc = CreateObject("word.application")
        Wxdc.Visible = True
        Wxdc.Documents.Open DFilePath & DFileName
     Else
        If Len(Dir(SMPPath & "\" & DSFileName)) <> 0 Then
            Wxdc.Visible = True
            Wxdc.Documents.Open SMPPath & "\" & DSFileName
            Wxdc.Documents(1).SaveAs DFilePath & DFileName
        Else
            s = MsgBox("檔名：" & SMPPath & DSFileName & "不存在", , "錯誤發生")
            Exit Sub
        End If
     End If
Case 1
      If Len(Dir(SMPPath & "\apply" & DSFileName)) <> 0 Then
          Wxdc.Visible = True
          Wxdc.Documents.Open SMPPath & "\apply" & DSFileName
          Wxdc.Documents(1).SaveAs DocTempPath & "\apply" & DSFileName
      Else
          s = MsgBox("檔名：" & SMPPath & "\apply" & DSFileName & "不存在", , "錯誤發生")
          Exit Sub
      End If
Case Else
End Select


End Sub

Private Sub Combo1_Click()
Screen.MousePointer = vbHourglass
Me.Enabled = False
StrMenu
Me.Enabled = True
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
MoveFormToCenter Me
'StrMenu
'Modify by Amy 2014/09/22 取消工程師輸入本所期限
Label1(35).Visible = False
txt1(13).Visible = False
End Sub

Sub StrMenu()
'''''''''''edit by nickc 2007/08/22
'''''''''''strSQL = "SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,S2.ST02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",NVL(PA05,NVL(PA06,PA07))," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",DECODE(PA09,'000',PTM03,PTM04),S3.ST02,decode(pa09,'000',cpm03,cpm04),S4.ST02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',0,''," & SQLDate("CP57") & ",CP10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE CP09=EP02(+) AND  CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & "' AND CP01 IN (" & SQLGrpStr("", 1) & ") "
'''''''''''strSQL = strSQL + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,S2.ST02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",NVL(TM05,NVL(TM06,TM07))," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",decode(tm10,'000',ptm03,ptm04),S3.ST02,decode(tm10,'000',cpm03,cpm04),S4.ST02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',0,''," & SQLDate("CP57") & ",CP10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,TRADEMARK WHERE CP09=EP02(+) AND  CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND tm08=PTM02(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & "' AND CP01 IN (" & SQLGrpStr("", 2) & ") "
'''''''''''strSQL = strSQL + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,S2.ST02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",NVL(LC05,NVL(LC06,LC07))," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",'',S3.ST02,decode(lc15,'000',cpm03,cpm04),S4.ST02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',0,''," & SQLDate("CP57") & ",CP10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,LAWCASE WHERE CP09=EP02(+) AND  CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & "' AND CP01 IN (" & SQLGrpStr("", 3) & ") "
'''''''''''strSQL = strSQL + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,S2.ST02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",HC06," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",'',S3.ST02,CPM03,S4.ST02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',0,''," & SQLDate("CP57") & ",CP10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,HIRECASE WHERE CP09=EP02(+) AND  CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & "' AND CP01 IN (" & SQLGrpStr("", 4) & ") "
'''''''''''strSQL = strSQL + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,S2.ST02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",NVL(SP05,NVL(SP06,SP07))," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",'',S3.ST02,decode(sp09,'000',cpm03,cpm04),S4.ST02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',0,''," & SQLDate("CP57") & ",CP10 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,SERVICEPRACTICE WHERE CP09=EP02(+) AND  CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & "' AND CP01 IN (" & SQLGrpStr("", 5) & ") "
''''''''''strSQL = "SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,s2.st02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",NVL(PA05,NVL(PA06,PA07))," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",decode(pa01,'CFP',ptm03,DECODE(PA09,'000',PTM03,PTM04)),s3.st02,decode(pa09,'000',cpm03,cpm04),s4.st02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10, '' As CP49 ,'' As TMXX,cp97,cp98,cp99,ep32,SQLDatet(ep33) as ep33 " & _
''''''''''               " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE CP09=EP02(+) AND cP01=Pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & "' and cp01 in (" & SQLGrpStr("", 1) & ") "
''''''''''strSQL = strSQL + " UNION all select EP01,S1.ST02," & SQLDate("CP48") & ",CP09,s2.st02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",NVL(TM05,NVL(TM06,TM07))," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",decode(tm10,'000',ptm03,ptm04),s3.st02,decode(tm10,'000',cpm03,cpm04),s4.st02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10, CP49 ,'' As TMXX,cp97,cp98,cp99,ep32,SQLDatet(ep33) as ep33 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,TRADEMARK WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & _
''''''''''               "' and cp01 in (" & SQLGrpStr("", 2) & ") "
''''''''''strSQL = strSQL + " UNION all select EP01,S1.ST02," & SQLDate("CP48") & ",CP09,s2.st02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",NVL(LC05,NVL(LC06,LC07))," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",'',s3.st02,decode(lc15,'000',cpm03,cpm04),s4.st02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10, '' As CP49 ,'' As TMXX,cp97,cp98,cp99,ep32,SQLDatet(ep33) as ep33 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,LAWCASE WHERE CP09=EP02(+) AND cp01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & "' and cp01 in (" & SQLGrpStr("", 3) & ") "
''''''''''strSQL = strSQL + " UNION all select EP01,S1.ST02," & SQLDate("CP48") & ",CP09,s2.st02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",HC06," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",'',s3.st02,CPM03,s4.st02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10, '' As CP49 ,'' As TMXX,cp97,cp98,cp99,ep32,SQLDatet(ep33) as ep33 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,HIRECASE WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & "' and cp01 in (" & SQLGrpStr("", 4) & ") "
''''''''''strSQL = strSQL + " UNION all select EP01,S1.ST02," & SQLDate("CP48") & ",CP09,s2.st02," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04," & SQLDate("EP06") & ",NVL(SP05,NVL(SP06,SP07))," & SQLDate("EP09") & ",CP26," & SQLDate("EP07") & ",'',s3.st02,decode(sp09,'000',cpm03,cpm04),s4.st02," & SQLDate("CP06") & "," & SQLDate("EP08") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10, '' As CP49 ,'' As TMXX,cp97,cp98,cp99,ep32,SQLDatet(ep33) as ep33 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,SERVICEPRACTICE WHERE CP09=EP02(+) AND cP01=sP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND CP09='" & Combo1.Text & "' and cp01 in (" & SQLGrpStr("", 5) & ") "
''''''''''
''''''''''CheckOC
''''''''''With adoRecordset
''''''''''    .CursorLocation = adUseClient
''''''''''    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
''''''''''    If .RecordCount <> 0 And .RecordCount > 0 Then
''''''''''        .MoveFirst
''''''''''        For i = 0 To 36 ' edit by nickc 2007/08/22   29
''''''''''            lbl1(i) = CheckStr(.Fields(i))
''''''''''        Next i
''''''''''        'strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01 AND CP02=CM02 AND CP03=CM03 AND CP04=CM04 AND CP14=ST01(+) AND CP31='Y' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
''''''''''        '92.04.03 nick add left join
''''''''''        'strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01 AND CP02=CM02 AND CP03=CM03 AND CP04=CM04 AND CP14=ST01(+) AND CP31='Y' and cp09='" & Combo1.Text & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
''''''''''        'edit by nickc 2006/04/26
''''''''''        'strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01(+) AND CP02=CM02(+) AND CP03=CM03(+) AND CP04=CM04(+) AND CP14=ST01(+) AND CP31='Y' and cp09='" & Combo1.Text & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
'''''''''''edit by nickc 2006/04/27 改按鈕
'''''''''''        strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS C1,caseprogress C2,STAFF WHERE cm01=c2.cp01(+) and cm02=c2.cp02(+) and cm03=c2.cp03(+) and cm04=c2.cp04(+) and C1.CP01=CM05(+) AND C1.CP02=CM06(+) AND C1.CP03=CM07(+) AND C1.CP04=CM08(+) AND C2.CP14=ST01(+) and cm10='0' AND C1.CP31='Y' and C1.cp09='" & Combo1.Text & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
'''''''''''        CheckOC2
'''''''''''        adoRecordset1.CursorLocation = adUseClient
'''''''''''        adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'''''''''''        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'''''''''''            lbl1(25) = CheckStr(adoRecordset1.Fields(0))
'''''''''''            lbl1(27) = CheckStr(adoRecordset1.Fields(1))
'''''''''''        Else
'''''''''''            lbl1(25) = ""
'''''''''''            lbl1(27) = ""
'''''''''''        End If
''''''''''        CheckOC2
''''''''''        '計算承辦天數
''''''''''        If Len(lbl1(12)) <> 0 And Len(lbl1(8)) <> 0 And Val(lbl1(12)) <> 0 And Val(lbl1(8)) <> 0 Then
''''''''''            lbl1(26) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(lbl1(12))), ChangeTStringToWString(ChangeTDateStringToTString(lbl1(8))))
''''''''''        Else
''''''''''            If Len(lbl1(10)) <> 0 And Len(lbl1(8)) <> 0 And Val(lbl1(10)) <> 0 And Val(lbl1(8)) <> 0 Then
''''''''''                lbl1(26) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(lbl1(10))), ChangeTStringToWString(ChangeTDateStringToTString(lbl1(8))))
''''''''''            Else
''''''''''                lbl1(26) = "0"
''''''''''            End If
''''''''''        End If
''''''''''    End If
''''''''''End With
''''''''''CheckOC
Dim m_SqlGrpStr1 As String, m_SqlGrpStr2 As String, m_SqlGrpStr3 As String, m_SqlGrpStr4 As String, m_SqlGrpStr5 As String
Dim m_CP10 As String
m_SqlGrpStr1 = SQLGrpStr("", 1)
m_SqlGrpStr2 = SQLGrpStr("", 2)
m_SqlGrpStr3 = SQLGrpStr("", 3)
m_SqlGrpStr4 = SQLGrpStr("", 4)
m_SqlGrpStr5 = SQLGrpStr("", 5)

Dim arrCaseNo '本所案號
'預設為商標案件
'm_blnTMCase = False
                     strSql = "SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(PA05,NVL(PA06,PA07)),EP09,CP26,EP07,DECODE(PA09,'000',PTM03,PTM04),EP04,decode(pa09,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,PA57,'*',EP27,EP31,cp13,ep05,pa09 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,pa26 as cuno,cp111,cp112,ep28,ep32,ep33,na03 " & _
                                    " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT,nation WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & Combo1.Text & "' and cp01 in (" & m_SqlGrpStr1 & ") and pa09=na01(+) "
strSql = strSql + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(TM05,NVL(TM06,TM07)),EP09,CP26,EP07,decode(tm10,'000',ptm03,ptm04),EP04,decode(tm10,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,TM29,cp49,EP27,EP31,cp13,ep05,tm10 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,tm23 as cuno,cp111,cp112,ep28,ep32,ep33,na03 " & _
                                   " FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,PATENTTRADEMARKMAP,TRADEMARK,nation WHERE EP02=CP09(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & Combo1.Text & "' and cp01 in (" & m_SqlGrpStr2 & ") and tm10=na01(+) "
strSql = strSql + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(LC05,NVL(LC06,LC07)),EP09,CP26,EP07,'',EP04,decode(lc15,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,LC08,'*',EP27,EP31,cp13,ep05,lc15 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,lc11 as cuno,cp111,cp112,ep28,ep32,ep33,na03                 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,LAWCASE,nation WHERE EP02=CP09(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and cp01=cpm01(+) and cp10=cpm02(+)  AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & Combo1.Text & "' and cp01 in (" & m_SqlGrpStr3 & ") and lc15=na01(+) "
strSql = strSql + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,HC06,EP09,CP26,EP07,'',EP04,CPM03,EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,HC09,'*',EP27,EP31,cp13,ep05,'000' as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,hc05 as cuno,cp111,cp112,ep28,ep32,ep33,na03                      FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,HIRECASE,nation WHERE EP02=CP09(+) AND CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & Combo1.Text & "' and cp01 in (" & m_SqlGrpStr4 & ") and '000'=na01(+) "
strSql = strSql + " UNION all  SELECT EP01,S1.ST02," & SQLDate("CP48") & ",CP09,EP13," & SQLDate("CP05") & ",EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04,EP06,NVL(SP05,NVL(SP06,SP07)),EP09,CP26,EP07,'',EP04,decode(sp09,'000',cpm03,cpm04),EP03," & SQLDate("CP06") & ",EP08," & SQLDate("CP07") & ",CP27,S5.ST02,EP11,CP18,EP12,'',Nvl(EP35,0),''," & SQLDate("CP57") & ",CP10,CP15,SP15,'*',EP27,EP31,cp13,ep05,sp09 as m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99,cp106,sp08 as cuno,cp111,cp112,ep28,ep32,ep33,na03                FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,CASEPROPERTYMAP,SERVICEPRACTICE,nation WHERE EP02=CP09(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND EP05=S1.ST01(+) AND EP13=S2.ST01(+) AND EP04=S3.ST01(+) AND EP03=S4.ST01(+) AND CP13=S5.ST01(+) AND EP02='" & Combo1.Text & "' and cp01 in (" & m_SqlGrpStr5 & ") and sp09=na01(+) "

CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        For i = 0 To 29
            lbl1(i) = CheckStr(.Fields(i))
        Next i

        m_CP10 = "" & .Fields("cp10").Value
        lbl1(33) = ""
        lbl1(34) = ""
        
        'Modify by Morgan 2008/12/3 本功能只有查詢原本帶出資料內容就好
        'If (m_CP10 = "101" Or m_CP10 = "102") And (SystemNumber(lbl1(7).Caption, 1) = "P" Or SystemNumber(lbl1(7).Caption, 1) = "CFP") Then
        '    If UCase(CheckStr(.Fields("cp112"))) = "Y" Then
        '        lbl1(33) = ChangeWStringToTString(CheckStr(.Fields("ep28")))
        '        lbl1(34) = CheckStr(.Fields("cp111"))
        '    End If
        '    Select Case UCase(CheckStr(.Fields("cp112")))
        '    Case "Y"
        '        opCP112(0).Value = True
        '    Case "N"
        '        opCP112(1).Value = True
        '    Case Else
        '        opCP112(0).Value = False
        '        opCP112(1).Value = False
        '    End Select
        '    If ProState = "2" Then
        '        Frame1.Enabled = True
        '    End If
        'Else
        '    Frame1.Enabled = False
        '    If (SystemNumber(lbl1(7).Caption, 1) = "P" Or SystemNumber(lbl1(7).Caption, 1) = "CFP") Then
        '        opCP112(1).Value = True
        '    End If
        'End If
        lbl1(33) = ChangeWStringToTString(CheckStr(.Fields("ep28")))
        Select Case UCase(CheckStr(.Fields("cp112")))
            Case "Y"
               opCP112(0).Value = True
               lbl1(34) = CheckStr(.Fields("cp111"))
            Case "N"
               opCP112(1).Value = True
            Case Else
               opCP112(0).Value = False
               opCP112(1).Value = False
        End Select
        'end 2008/12/3
        
        lbl1(32).Caption = "" & .Fields("CP97").Value
        txt1(15).Text = "" & .Fields("cp98").Value
        txt1(16).Text = "" & .Fields("cp99").Value
        txt1(17).Text = "" & .Fields("cp106").Value

        If IsNull(.Fields(31).Value) <> 0 Then
            Me.lblClose.Caption = ""
        Else
            Me.lblClose.Caption = "已閉卷"
        End If
        '91.08.14 增加若是商標案則加秀條款欄位* 就是商標   nick  start
        If CheckStr(.Fields(32).Value) <> "*" Then
            txt1(11).Text = CheckStr(.Fields(32))
            txt1(11).Visible = True
            Label1(34).Visible = True
        Else
            txt1(11).Text = ""
            txt1(11).Visible = False
            Label1(34).Visible = False
        End If
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
         If IsNull(.Fields("EP32")) Then
            txt1(20) = ""
         Else
            txt1(20) = .Fields("EP32")
         End If
         txt1(20).Tag = txt1(20).Text
        strSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01(+) AND CP02=CM02(+) AND CP03=CM03(+) AND CP04=CM04(+) AND CP14=ST01(+) AND CP31='Y' and cp09='" & Combo1.Text & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
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
        Me.lbl1(31).Caption = IIf(IsNull(.Fields("CP15")), "0", "" & .Fields("CP15"))
    Else
        For i = 0 To 29
            lbl1(i) = ""
        Next i
        Me.lbl1(31).Caption = ""
        Me.lblClose.Caption = ""
        
        For i = 0 To 14
            txt1(i) = ""
        Next i
    End If
    
End With
CheckOC
txt1(0).Text = lbl1(4).Caption
txt1(6).Text = lbl1(16).Caption

Dim tmpInti As Integer


Combo2.Enabled = False
'Modified by Lydia 2022/02/08 Form 2.0 不支援.Text寫法,並且不知用處;先隱藏
'Combo2.Text = ""
'Combo4.Text = ""
'For tmpInti = 0 To Combo2.ListCount - 1
'    If Trim(txt1(0).Text) = Trim(Mid(Combo2.List(tmpInti), 1, InStr(1, Combo2.List(tmpInti), "=") - IIf(InStr(1, Combo2.List(tmpInti), "=") = 0, 0, 1))) Then
'        Combo2.Text = Combo2.List(tmpInti)
'    End If
'Next tmpInti
'
'For tmpInti = 0 To Combo4.ListCount - 1
'    If Trim(txt1(6).Text) = Trim(Mid(Combo4.List(tmpInti), 1, InStr(1, Combo4.List(tmpInti), "=") - IIf(InStr(1, Combo4.List(tmpInti), "=") = 0, 0, 1))) Then
'        Combo4.Text = Combo4.List(tmpInti)
'    End If
'Next tmpInti
Combo2.Clear
Combo4.Clear
'end 2022/02/08

    txt1(19).Enabled = False



lbl1(4).Caption = GetPrjSalesNM(txt1(0).Text)
txt1(1).Text = lbl1(6).Caption
txt1(2).Text = ChangeWStringToTString(lbl1(8).Caption)
txt1(3).Text = ChangeWStringToTString(lbl1(10).Caption)
txt1(4).Text = ChangeWStringToTString(lbl1(12).Caption)
If Len(lbl1(14).Caption) = 0 Then
   strSql = "select pp04 from caseprogress,promoterproofreader where cp09='" & Combo1.Text & "' and cp01=pp01(+) and cp14=pp02(+) and cp10=pp03(+) "
   CheckOC2
   adoRecordset1.CursorLocation = adUseClient
   adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset1.RecordCount <> 0 Then
      txt1(5) = CheckStr(adoRecordset1.Fields(0))
   End If
Else
   txt1(5).Text = lbl1(14).Caption
End If
lbl1(14).Caption = GetPrjSalesNM(txt1(5).Text)

lbl1(16).Caption = GetPrjSalesNM(txt1(6).Text)
'會稿完成日
txt1(7).Text = ChangeWStringToTString(lbl1(18).Caption)
txt1(8).Text = ChangeWStringToTString(lbl1(20).Caption)
txt1(9).Text = lbl1(22).Caption
txt1(10).Text = lbl1(24).Caption
Me.txt1(12).Text = ChangeTDateStringToTString(Me.lbl1(2).Caption)
Me.txt1(12).Locked = False
txt1(18) = lbl1(33)

'若為個人工作管理
   Label1(36).Visible = True
   'Modify by Amy 2014/09/22 取消工程師輸入本所期限
   'Label1(35).Visible = True
   'txt1(13).Visible = True
   'txt1(13).Enabled = True
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
If Left(UCase(Combo1.Text), 1) = "C" Then
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

For i = 0 To 20
      txt1(i).Enabled = False
Next i
Combo4.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set Fobj = New FileSystemObject
Fobj.DeleteFile DocTempPath & "\*.doc", False
Set Fobj = Nothing
Set frm090613_2 = Nothing
End Sub

