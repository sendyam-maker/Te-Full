VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11k0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯費資料輸入"
   ClientHeight    =   5508
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   8364
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5508
   ScaleWidth      =   8364
   Begin VB.TextBox txtEP09 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7185
      TabIndex        =   29
      Top             =   1770
      Width           =   945
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3285
      Left            =   90
      TabIndex        =   28
      Top             =   2190
      Width           =   8205
      _ExtentX        =   14478
      _ExtentY        =   5800
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   16777088
      TabCaption(0)   =   "108.8.15前完稿"
      TabPicture(0)   =   "Frmacc11k0.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "108.8.15(含)後完稿"
      TabPicture(1)   =   "Frmacc11k0.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CFP"
      TabPicture(2)   =   "Frmacc11k0.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFF80&
         Height          =   2925
         Left            =   30
         TabIndex        =   68
         Top             =   330
         Width           =   8115
         Begin VB.TextBox txtCFP_TF 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   23
            Left            =   1545
            MaxLength       =   6
            TabIndex        =   71
            Top             =   210
            Width           =   855
         End
         Begin VB.TextBox txtCFP_TF 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   21
            Left            =   5472
            MaxLength       =   6
            TabIndex        =   70
            Top             =   210
            Width           =   855
         End
         Begin VB.TextBox txtCFP_TF 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   22
            Left            =   1545
            TabIndex        =   69
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblCFP_TransType 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "(語種)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2550
            TabIndex        =   77
            Top             =   270
            Width           =   570
         End
         Begin VB.Label lblCFP_WordCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "原文字數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   96
            TabIndex        =   76
            Top             =   276
            Width           =   816
         End
         Begin VB.Label lblCFP_WordCountAcc 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "原文字數-財務"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3960
            TabIndex        =   75
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "計算公式：( 費率單位：NT$/千字 )"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   74
            Top             =   1020
            Width           =   3225
         End
         Begin VB.Label lblCFP_TransFeeRule 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "應付翻譯費 = 原文字數*翻譯費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   240
            TabIndex        =   73
            Top             =   1308
            Width           =   2988
         End
         Begin VB.Label lblCFP_TransRate 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "翻譯費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   96
            TabIndex        =   72
            Top             =   660
            Width           =   816
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF80&
         Height          =   2925
         Left            =   -74970
         TabIndex        =   54
         Top             =   330
         Width           =   8115
         Begin VB.TextBox txtTFNew 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   1545
            TabIndex        =   58
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtTFNew 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   3960
            TabIndex        =   57
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtTF21 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5370
            MaxLength       =   6
            TabIndex        =   56
            Top             =   210
            Width           =   855
         End
         Begin VB.TextBox txtTF 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   23
            Left            =   1545
            MaxLength       =   6
            TabIndex        =   55
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "英文翻譯費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   67
            Top             =   660
            Width           =   1260
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "日文翻譯費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2580
            TabIndex        =   66
            Top             =   660
            Width           =   1260
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "英文翻譯費 = 英文原文字數*英文翻譯費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   65
            Top             =   1320
            Width           =   3915
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "應付翻譯費 = 英(日)文翻譯費*相似折扣%*瑕疵折扣%*加成比率%"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   64
            Top             =   1860
            Width           =   6030
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "日文翻譯費 = 日文原文字數*日文翻譯費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   63
            Top             =   1590
            Width           =   3915
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "計算公式：( 費率單位：NT$/千字 )"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   62
            Top             =   1020
            Width           =   3225
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "原文字數-財務"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3960
            TabIndex        =   61
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "原文字數-承辦"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   60
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "(語種)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2550
            TabIndex        =   59
            Top             =   270
            Width           =   570
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFF80&
         Height          =   2925
         Left            =   -74970
         TabIndex        =   31
         Top             =   330
         Width           =   8145
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFC0FF&
            Caption         =   "以英文字數計費(依當月即期賣出匯率,重新檢視英文翻譯費率 )"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   150
            TabIndex        =   53
            Top             =   150
            Width           =   5715
         End
         Begin VB.TextBox txtTF 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   22
            Left            =   1665
            TabIndex        =   39
            Top             =   1140
            Width           =   855
         End
         Begin VB.TextBox txtTF 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   1665
            TabIndex        =   38
            Top             =   780
            Width           =   855
         End
         Begin VB.TextBox txtTF 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   4140
            TabIndex        =   37
            Top             =   780
            Width           =   855
         End
         Begin VB.TextBox txtTF 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   17
            Left            =   6750
            TabIndex        =   36
            Top             =   780
            Width           =   855
         End
         Begin VB.TextBox txtTF 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   21
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   35
            Top             =   435
            Width           =   855
         End
         Begin VB.TextBox txtTF 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   7065
            TabIndex        =   34
            Top             =   435
            Width           =   855
         End
         Begin VB.TextBox txtTF 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4950
            TabIndex        =   33
            Top             =   435
            Width           =   855
         End
         Begin VB.TextBox txtTF 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3015
            MaxLength       =   6
            TabIndex        =   32
            Top             =   435
            Width           =   855
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "英文翻譯費率2"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   52
            Top             =   1200
            Width           =   1365
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "= (中文字數+數學式數)*英文翻譯費率+(中文字數+數學式數)*中文打字費率"
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
            Left            =   810
            TabIndex        =   51
            Top             =   2070
            Width           =   6450
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "英文翻譯費 = 英文字數*英文翻譯費率2*80% 或"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   50
            Top             =   1800
            Width           =   4350
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "應付翻譯費 = 英(日)文翻譯費*相似折扣%*瑕疵折扣%*加成比率%"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   49
            Top             =   2640
            Width           =   6030
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "日文翻譯費 = 日文字數*日文翻譯費率+(中文字數+數學式數)*中文打字費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   48
            Top             =   2340
            Width           =   6930
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "計算公式：( 費率單位：NT$/千字 )"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   1500
            Width           =   3225
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "英文翻譯費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   46
            Top             =   840
            Width           =   1260
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "日文翻譯費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2700
            TabIndex        =   45
            Top             =   840
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "中文打字費率"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5310
            TabIndex        =   44
            Top             =   840
            Width           =   1260
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "英文字數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   43
            Top             =   495
            Width           =   840
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "中文字數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   42
            Top             =   495
            Width           =   840
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "數學式字數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5985
            TabIndex        =   41
            Top             =   495
            Width           =   1050
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "日文字數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4095
            TabIndex        =   40
            Top             =   495
            Width           =   840
         End
      End
   End
   Begin VB.TextBox txtTF 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4590
      TabIndex        =   26
      Top             =   1770
      Width           =   1485
   End
   Begin VB.TextBox txtTF 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   1485
      TabIndex        =   20
      Top             =   1770
      Width           =   1530
   End
   Begin VB.TextBox txtTF 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   18
      Left            =   6588
      TabIndex        =   9
      Top             =   1320
      Width           =   945
   End
   Begin VB.TextBox txtTF 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4410
      TabIndex        =   6
      Top             =   75
      Width           =   1125
   End
   Begin VB.TextBox txtTF 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   4050
      TabIndex        =   8
      Top             =   1320
      Width           =   945
   End
   Begin VB.TextBox txtTF 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   1485
      TabIndex        =   7
      Top             =   1320
      Width           =   945
   End
   Begin VB.TextBox txtCaseNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1260
      MaxLength       =   3
      TabIndex        =   0
      Top             =   75
      Width           =   510
   End
   Begin VB.TextBox txtCaseNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   2
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtCaseNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2745
      MaxLength       =   1
      TabIndex        =   3
      Top             =   75
      Width           =   240
   End
   Begin VB.TextBox txtCaseNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   2970
      MaxLength       =   2
      TabIndex        =   4
      Top             =   75
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   3285
      Picture         =   "Frmacc11k0.frx":0054
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   75
      Width           =   350
   End
   Begin VB.TextBox txtCaseNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1755
      MaxLength       =   6
      TabIndex        =   1
      Top             =   75
      Width           =   1005
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "案件性質："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   5496
      TabIndex        =   79
      Top             =   1008
      Width           =   1020
   End
   Begin MSForms.Label lblProperty 
      Height          =   192
      Left            =   6564
      TabIndex        =   78
      Top             =   1008
      Width           =   1668
      VariousPropertyBits=   268435475
      Caption         =   "lblProperty"
      Size            =   "2942;339"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label28 
      BackStyle       =   0  '透明
      Caption         =   "完稿日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   30
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "應付翻譯費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3330
      TabIndex        =   27
      Top             =   1830
      Width           =   1050
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   25
      Top             =   105
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "案件名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   24
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "收文號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3735
      TabIndex        =   23
      Top             =   105
      Width           =   645
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "相似折扣                 %"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   270
      TabIndex        =   22
      Top             =   1380
      Width           =   2025
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "單據編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   21
      Top             =   1800
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   8000
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "加成比率                 %"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   5508
      TabIndex        =   19
      Top             =   1380
      Width           =   1044
   End
   Begin MSForms.Label lblEng 
      Height          =   192
      Left            =   6564
      TabIndex        =   18
      Top             =   120
      Width           =   1500
      VariousPropertyBits=   268435475
      Caption         =   "lblEng"
      Size            =   "2646;339"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "承辦人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   5676
      TabIndex        =   17
      Top             =   120
      Width           =   816
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "瑕疵折扣                %"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2925
      TabIndex        =   16
      Top             =   1380
      Width           =   1965
   End
   Begin MSForms.Label lblCName3 
      Height          =   192
      Left            =   1692
      TabIndex        =   15
      Top             =   1020
      Width           =   3744
      VariousPropertyBits=   268435475
      Caption         =   "lblCName3"
      Size            =   "6604;339"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label lblCName2 
      Height          =   195
      Left            =   1710
      TabIndex        =   14
      Top             =   720
      Width           =   6500
      VariousPropertyBits=   268435475
      Caption         =   "lblCName2"
      Size            =   "11465;344"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label lblCName1 
      Height          =   195
      Left            =   1710
      TabIndex        =   13
      Top             =   420
      Width           =   6500
      VariousPropertyBits=   268435475
      Caption         =   "lblCName1"
      Size            =   "11465;344"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   12
      Top             =   1020
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "英："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   11
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "中："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   10
      Top             =   420
      Width           =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   180
      X2              =   8000
      Y1              =   1710
      Y2              =   1710
   End
End
Attribute VB_Name = "Frmacc11k0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改 lblEng/lblCName1/lblCName2/lblCName3
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Add by Morgan 2007/5/21
Option Explicit

Dim m_EP09 As String, m_TF27 As String 'Added by Morgan 2019/8/15

Public Function FormSave() As Boolean
   If txtTF(1) = "" Then
      MsgBox "案號輸入錯誤！", vbExclamation
      txtCaseNo(1).SetFocus
      Exit Function
   
   'Added by Morgan 2017/2/10
   End If
   
   'Added by Morgan 2025/11/10
   If txtCaseNo(0) = "CFP" Then
      If Val("" & txtCFP_TF(22)) = 0 Then
         MsgBox "尚未設定" & lblCFP_TransRate & "，不可存檔！", vbExclamation
         Exit Function
      
      ElseIf Val(txtCFP_TF(21)) = 0 Then
         MsgBox "尚未輸入" & lblCFP_WordCountAcc & "，不可存檔！", vbExclamation
         txtCFP_TF(21).SetFocus
         Exit Function
      
      End If
   Else
   'end 2025/11/10
   
      'Added by Morgan 2019/8/14
      If Val(m_EP09) >= 20190815 Then
         If Val(txtTF21) = 0 Then
            MsgBox "尚未輸入原文字數，不可存檔！", vbExclamation
            Exit Function
         'Added by Morgan 2023/5/2
         ElseIf m_TF27 = "" Then
            MsgBox "尚未設定原文語種，不可存檔！" & vbCrLf & vbCrLf & "外專->資料處理->新案建檔->翻譯", vbExclamation
            Exit Function
         'end 2023/5/2
         ElseIf Val(txtTFNew(15)) = 0 And m_TF27 = "1" Then
            MsgBox "尚未設定英文翻譯費率，不可存檔！", vbExclamation
            Exit Function
         
         ElseIf Val(txtTFNew(16)) = 0 And m_TF27 = "2" Then
            MsgBox "尚未設定日文翻譯費率，不可存檔！", vbExclamation
            Exit Function
            
         End If
         
         
         If Val(txtTF21) <> Val(txtTF(23)) Then
            If MsgBox("輸入的原文字數【" & txtTF21 & "】與承辦的原文字數【" & txtTF(23) & "】不同，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
               Exit Function
            End If
         End If
         
      Else
      'end 2019/8/14
      
         If Check1.Value = vbChecked Then
            If Val(txtTF(21)) = 0 Then
               MsgBox "請輸入英文字數！", vbExclamation
               txtTF(21).SetFocus
               Exit Function
            End If
         'end 2017/2/10
         
         ElseIf Val(txtTF(2)) = 0 Then
            MsgBox "請輸入中文字數！", vbExclamation
            txtTF(2).SetFocus
            Exit Function
         End If
         
      End If
   End If
   
   SetFee 'Added by Morgan 2017/2/10
   
   'Added by Morgan 2015/9/23
   '檢查翻譯費用是否超過比例
   If PUB_ChkTranslationFee(txtTF(1), Format(txtTF(0))) = False Then Exit Function
   'end 2015/9/23
   
On Error GoTo ErrHnd
   
   'Added by Morgan 2025/11/10
   If txtCaseNo(0) = "CFP" Then
      strSql = "Update transfee set tf21=" & CNULL(txtCFP_TF(21), True) & " where tf01='" & txtTF(1) & "'"
      
   'end 2025/11/10
   'Added by Morgan 2019/8/15
   ElseIf Val(m_EP09) >= 20190815 Then
      strSql = "Update transfee set tf21=" & CNULL(txtTF21, True) & " where tf01='" & txtTF(1) & "'"
   Else
   'end 2019/8/15
   
      strSql = "Update transfee set tf02=" & CNULL(txtTF(2), True) & ",tf21=" & CNULL(txtTF(21), True) & " where tf01='" & txtTF(1) & "'"
   End If 'Added by Morgan 2019/8/15
      

   
   adoTaie.Execute strSql, intI
   If intI = 1 Then
      FormSave = True
      txtCaseNo(0).Tag = txtCaseNo(0).Text
      txtCaseNo(1).Tag = txtCaseNo(1).Text
      txtCaseNo(2).Tag = txtCaseNo(2).Text
      txtCaseNo(3).Tag = txtCaseNo(3).Text
      txtCaseNo(4).Tag = txtCaseNo(4).Text
   End If
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      If Val(txtTF(22)) = 0 Then
         MsgBox "尚未設定英文翻譯費率2，不可以英文字數計費！", vbExclamation
         Check1.Value = vbUnchecked
         Exit Sub
      End If
   End If
   
   SetText True
End Sub

Private Sub SetText(Optional pSetFocus As Boolean = False)
   
   If Check1.Value = vbChecked Then
      txtTF(21).Enabled = True
      txtTF(21).BackColor = &HFFFFFF
      
      txtTF(2).Text = ""
      txtTF(2).Enabled = False
      txtTF(2).BackColor = &HE0E0E0
   Else
      txtTF(2).Enabled = True
      txtTF(2).BackColor = &HFFFFFF
      
      txtTF(21).Text = ""
      txtTF(21).Enabled = False
      txtTF(21).BackColor = &HE0E0E0
   End If
   
   If pSetFocus Then
      If Check1.Value = vbChecked Then
         txtTF(21).SetFocus
      Else
         txtTF(2).SetFocus
      End If
   End If
End Sub

Public Sub Command1_Click()
   If SearchCheck = True Then
      If ReadData = False Then
         txtCaseNo(1).SetFocus
         txtCaseNo_GotFocus 2
      'Else
      '   txtTF(2).SetFocus
      End If
   End If
End Sub

Private Function SearchCheck(Optional p_bMsg As Boolean = True) As Boolean
   If txtCaseNo(3) = "" Then txtCaseNo(3) = "0"
   If txtCaseNo(4) = "" Then txtCaseNo(4) = "00"
   If Len(txtCaseNo(0)) = 0 Then
      If p_bMsg = True Then
         MsgBox "本所案號不可空白", , "USER 輸入錯誤"
         txtCaseNo(0).SetFocus
      End If
      Exit Function
   ElseIf Len(txtCaseNo(1)) = 0 Then
      If p_bMsg = True Then
         MsgBox "本所案號不可空白", , "USER 輸入錯誤"
         txtCaseNo(1).SetFocus
         txtCaseNo_GotFocus 1
      End If
      Exit Function
   ElseIf txtCaseNo(0) = "TF" And Len(txtCaseNo(2)) = 0 Then
      If p_bMsg = True Then
         MsgBox "本所案號不可空白", , "USER 輸入錯誤"
         txtCaseNo(2).SetFocus
         txtCaseNo_GotFocus 2
      End If
      Exit Function
   End If
   SearchCheck = True
End Function

Public Function FormCheck() As Boolean
   
   If txtTF(7).Text <> "" Then
      MsgBox "已轉應付，不可異動！"
      Exit Function
   ElseIf txtTF(1).Text = "" Then
      MsgBox "案號資料錯誤！"
      Exit Function
   ElseIf txtTF(1).Tag <> txtTF(1).Text Then
      MsgBox "案號有改，請重新查詢！"
      Exit Function
   End If
   FormCheck = True
End Function

Private Function ReadData(Optional p_bNoMsg As Boolean) As Boolean
   Dim strMsg As String
   
   'Modified by Morgan 2018/9/13 +cp14 not like 'F%
   'Modified by Morgan 2025/11/7 +CFP案
   strExc(0) = "select '' v,sqldatet(ep09) ddate,cp09,decode(pa09,'020',cpm04,cpm03) cp10N,pa05,pa06,pa07,cp14||' '||st02 st02,ep09,TF.*,SP.*" & _
      " from caseprogress,patent,staff,transfee TF,Staff_PayRate SP,engineerprogress,casepropertymap" & _
      " where cp01='" & txtCaseNo(0) & "' and cp02='" & txtCaseNo(1) & txtCaseNo(2) & "' and cp03='" & txtCaseNo(3) & "' and cp04='" & txtCaseNo(4) & "'" & _
      " and cp14 like 'F%' and (cp10 in ('201','927') or cp01='CFP')" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and st01(+)=cp14 and tf01(+)=cp09 and (tf03 is not null or (cp01='CFP' and tf23>0)) and spr01(+)=cp14 and ep02(+)=cp09 and ep09>0" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 order by ep09 desc,cp09 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      strMsg = "資料不存在！"
   Else
      With RsTemp
      
      'Added by Morgan 2025/11/7
      'CFP案可能會有多筆(來函也會要翻譯)
      If .RecordCount > 1 Then
         frm880023.p_iChoice = 2
         frm880023.SetGridHead RsTemp, "v,完稿日,收文號,案件性質", "200,1200,1200,1600", flexAlignCenterCenter & "," & flexAlignCenterCenter & "," & flexAlignCenterCenter
         frm880023.Show vbModal
         strExc(1) = frm880023.p_sReturn
         Set frm880023 = Nothing
         If strExc(1) = "" Then
            Exit Function
         Else
            .Find "cp09='" & strExc(1) & "'"
         End If
      End If
      'end 2025/11/7
      
      'Modified by Morgan 2017/2/10 +英文字數檢查
      If strSaveConfirm = MsgText(3) And (Val("" & .Fields("tf02")) > 0 Or Val("" & .Fields("tf21")) > 0) Then
         strMsg = "已輸中文/原文字數，請以修改模式進行！"
      ElseIf strSaveConfirm <> MsgText(3) And (IsNull(.Fields("tf02")) And IsNull(.Fields("tf21"))) Then
         strMsg = "尚未輸中文/原文字數，請以新增模式進行！"
      Else
         lblCName1 = "" & .Fields("pa05")
         lblCName2 = "" & .Fields("pa06")
         lblCName3 = "" & .Fields("pa07")
         lblEng = "" & .Fields("st02")
         lblProperty = "" & .Fields("cp10N") 'Added by Morgan 2025/11/10
         txtTF(1) = "" & .Fields("tf01")
         txtTF(2) = "" & .Fields("tf02")
         txtTF(3) = "" & .Fields("tf03")
         txtTF(4) = "" & .Fields("tf04")
         txtTF(5) = "" & .Fields("tf05")
         txtTF(6) = "" & .Fields("tf06")
         txtTF(7) = "" & .Fields("tf07")
         
         
         'Modified by Morgan 2019/8/14
         'txtTF(21) = "" & .Fields("tf21") 'Added by Morgan 2017/2/10
         m_EP09 = "" & .Fields("ep09")
         txtEP09 = ChangeWStringToTDateString(m_EP09)
         
         'Added by Morgan 2025/11/10
         If txtCaseNo(0) = "CFP" Then
            SSTab1.TabVisible(2) = True
            SSTab1.TabVisible(0) = False
            SSTab1.TabVisible(1) = False
            
            txtCFP_TF(23) = "" & .Fields("TF23") '中/原文字數
            txtCFP_TF(21) = "" & .Fields("TF21") '中/原文字數-財務
            txtCFP_TF(22) = "" & .Fields("TF22") 'CFP翻譯費率
            
            If "" & .Fields("TF27") = "5" Then
               lblCFP_WordCount = "中文字數"
               strExc(1) = "中翻" & Left(Pub_GetTransFeeL("2", "" & .Fields("TF28")), 1)
               lblCFP_TransType = "(" & strExc(1) & ")"
               lblCFP_TransRate = strExc(1) & "費率"
               If txtCFP_TF(22) = "" Then
                  If .Fields("TF28") = "3" Then '中翻日
                     txtCFP_TF(22) = "" & .Fields("SPR15")
                  ElseIf .Fields("TF28") = "4" Then '中翻德
                     txtCFP_TF(22) = "" & .Fields("SPR16")
                  End If
               End If
            Else
               lblCFP_WordCount = "原文字數"
               strExc(1) = Pub_GetTransFeeL("1", "" & .Fields("TF27"))
               lblCFP_TransType = "(" & Left(strExc(1), 1) & "翻中)"
               lblCFP_TransRate = strExc(1) & "翻譯費率"
               If txtCFP_TF(22) = "" Then
                  If .Fields("TF27") = "2" Then '日翻中
                     txtCFP_TF(22) = "" & .Fields("SPR13")
                  ElseIf .Fields("TF27") = "3" Then '德翻中
                     txtCFP_TF(22) = "" & .Fields("SPR14")
                  End If
               End If
            End If
            lblCFP_WordCountAcc = lblCFP_WordCount & "-財務"
            lblCFP_TransFeeRule = "應付翻譯費 = " & lblCFP_WordCount & "*" & lblCFP_TransRate
         'end 2025/11/10
         
         ElseIf Val(m_EP09) >= 20190815 Then
            SSTab1.TabVisible(1) = True
            SSTab1.TabVisible(0) = False
            SSTab1.TabVisible(2) = False 'Added by Morgan 2025/11/10
            txtTF21 = "" & .Fields("tf21")
            '原文字數-承辦
            txtTF(23) = "" & .Fields("tf23")
            '原文語種
            m_TF27 = "" & .Fields("tf27")
            If .Fields("tf27") = "1" Then
               Label35 = "(英文)"
            ElseIf .Fields("tf27") = "2" Then
               Label35 = "(日文)"
            ElseIf .Fields("tf27") = "3" Then
               Label35 = "(德文)"
            Else
               Label35 = "(?文)"
            End If
            
         Else
            SSTab1.TabVisible(0) = True
            SSTab1.TabVisible(1) = False
            SSTab1.TabVisible(2) = False 'Added by Morgan 2025/11/10
            txtTF(21) = "" & .Fields("tf21")
         End If
         'end 2019/8/14
         
         If Not IsNull(.Fields("tf07")) Then
            txtTF(15) = "" & .Fields("tf15")
            txtTF(16) = "" & .Fields("tf16")
            txtTF(17) = "" & .Fields("tf17")
            txtTF(22) = "" & .Fields("tf22") 'Added by Morgan 2017/2/10
            
            'Added by Morgan 2019/8/15
            txtTF(23) = "" & .Fields("tf23")
            txtTFNew(15) = "" & .Fields("tf15")
            txtTFNew(16) = "" & .Fields("tf16")
            'end 2019/8/15
         Else
            txtTF(15) = "" & .Fields("spr02")
            txtTF(16) = "" & .Fields("spr03")
            txtTF(17) = "" & .Fields("spr04")
            txtTF(22) = "" & .Fields("spr11") 'Added by Morgan 2017/2/10
            
            'Added by Morgan 2019/8/15
            txtTFNew(15) = "" & .Fields("spr12")
            txtTFNew(16) = "" & .Fields("spr13")
            'end 2019/8/15
         End If
         txtTF(18) = "" & .Fields("tf18")
         If strSaveConfirm <> MsgText(3) Then
            txtTF(1).Tag = txtTF(1).Text
            txtCaseNo(0).Tag = txtCaseNo(0).Text
            txtCaseNo(1).Tag = txtCaseNo(1).Text
            txtCaseNo(2).Tag = txtCaseNo(2).Text
            txtCaseNo(3).Tag = txtCaseNo(3).Text
            txtCaseNo(4).Tag = txtCaseNo(4).Text
         End If
         
         'Modified by Morgan 2017/2/10
         If Val(txtTF(21)) > 0 Then
            Check1.Value = vbChecked
         Else
            Check1.Value = vbUnchecked
         End If
         'end 2017/2/10
                           
         SetFee
         ReadData = True
      End If
      End With
   End If
   If ReadData = False Then
      If p_bNoMsg = False Then
         MsgBox strMsg
      End If
   End If
End Function

Private Sub SetFee()
   Dim dblFee As Double
   
'Modified by Morgan 2017/2/10 改用函數
'   '日文翻譯費=日文字數*日文翻譯費率+(中文字數+數學式數)*中文打字費率
'   If Val(txtTF(3)) > 0 Then
'      '翻譯費,打字費要個別四捨五入
'      dblFee = Round(Val(txtTF(3)) * Val(txtTF(16)) / 1000) + Round((Val(txtTF(2)) + Val(txtTF(4))) * (Val(txtTF(17)) / 1000))
'
'   'Added by Morgan 2017/2/10
'   '英文翻譯費2=英文字數*英文翻譯費率2*80%
'   ElseIf Val(txtTF(21)) > 0 Then
'      dblFee = Round(Val(txtTF(21)) * (Val(txtTF(22)) / 1000) * 0.8)
'   'end 2017/2/10
'
'   '英文翻譯費=(中文字數+數學式數)*英文翻譯費率+(中文字數+數學式數)*中文打字費率
'   Else
'      '翻譯費,打字費要個別四捨五入
'      dblFee = Round((Val(txtTF(2)) + Val(txtTF(4))) * Val(txtTF(15)) / 1000) + Round((Val(txtTF(2)) + Val(txtTF(4))) * Val(txtTF(17)) / 1000)
'   End If
'   '翻譯費=原翻譯費*相似折扣%*瑕疵折扣%*加成比率%
'   dblFee = Round(dblFee * Val(IIf(txtTF(5) = "", 100, txtTF(5))) / 100 * Val(IIf(txtTF(6) = "", 100, txtTF(6))) / 100 * Val(IIf(txtTF(18) = "", 100, txtTF(18))) / 100)
   
   'Added by Morgan 2025/11/10
   If txtCaseNo(0) = "CFP" Then
      dblFee = PUB_GetTransFeeNew(Val(txtCFP_TF(21)), Val(txtCFP_TF(22)), 0, 0, 0)
   Else
   'end 2025/11/10
   
      'Added by Morgan 2019/8/14
      '108.8.15 以後完稿案件改以原文字數計算翻譯費並取消中文打字費
      If Val(m_EP09) >= 20190815 Then
         dblFee = PUB_GetTransFeeNew(Val(txtTF21), IIf(m_TF27 = "1", Val(txtTFNew(15)), Val(txtTFNew(16))), Val(txtTF(5)), Val(txtTF(6)), Val(txtTF(18)))
      Else
      'end 2019/8/14
         dblFee = PUB_GetTransFee(Val(txtTF(2)), Val(txtTF(3)), Val(txtTF(4)), Val(txtTF(15)), Val(txtTF(16)), Val(txtTF(17)), Val(txtTF(5)), Val(txtTF(6)), Val(txtTF(18)), Val(txtTF(21)), Val(txtTF(22)))
         
      End If 'Added by Morgan 2019/8/14
      
   End If
'end 2017/2/10

   txtTF(0) = Format(dblFee, "#,##0")
End Sub

Public Sub FormClear()
   Dim oText As TextBox
   For Each oText In txtTF
      oText = ""
   Next
   lblCName1 = ""
   lblCName2 = ""
   lblCName3 = ""
   lblEng = ""
End Sub

Private Sub Form_Activate()
   txtCaseNo(1).SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, Me.Width, Me.Height
   'tool3_enabled
   FormClear
   FormEnable
   txtCaseNo(0) = "FCP"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   PUB_SendMailCache 'Added by Lydia 2019/07/03
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11k0 = Nothing
End Sub

Private Sub SSTab1_GotFocus()
   'Added by Morgan 2019/8/15
   If SSTab1.TabVisible(1) = True Then
      If txtTF21.Enabled Then txtTF21.SetFocus
   Else
      If txtTF(2).Enabled Then txtTF(2).SetFocus
   End If
   'end 2019/8/15
End Sub

Private Sub txtCaseNo_Change(Index As Integer)
   If Index = 0 Then
      If txtCaseNo(0).Text = "TF" Then
         txtCaseNo(1).Text = ""
         txtCaseNo(1).MaxLength = 5
         txtCaseNo(2).Text = ""
         txtCaseNo(2).Visible = True
      Else
         txtCaseNo(1).Text = ""
         txtCaseNo(1).MaxLength = 6
         txtCaseNo(2).Text = ""
         txtCaseNo(2).Visible = False
      End If
   End If
   If txtTF(1) <> "" Then
      FormClear
   End If
End Sub

Private Sub txtCaseNo_GotFocus(Index As Integer)
   TextInverse txtCaseNo(Index)
End Sub

Private Sub txtCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Command1_Click
   Else
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtCaseNo_Validate(Index As Integer, Cancel As Boolean)
   If Index = 4 Then
      If SearchCheck(False) = True Then
         ReadData
      End If
   End If
End Sub

Private Sub txtTF_GotFocus(Index As Integer)
   TextInverse txtTF(Index)
End Sub

Public Function FormDelete() As Boolean
   If FormCheck = False Then
      Exit Function
   End If
On Error GoTo ErrHnd
   'Modified by Morgan 2020/9/1
   'strSql = "Update transfee set tf02=null,tf11=null where tf01='" & txtTF(1) & "'"
   strSql = "Update transfee set tf02=null,tf21=null where tf01='" & txtTF(1) & "'"
   'end 2020/9/1
   adoTaie.Execute strSql, intI
   If intI = 1 Then
      FormDelete = True
      txtCaseNo(0).Tag = ""
      txtTF(1).Tag = ""
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function


Public Sub FormEnable()
   Dim bolLock As Boolean
   '新增
   If strSaveConfirm = MsgText(3) Then
      txtCaseNo(0).Locked = False
      txtCaseNo(1).Locked = False
      txtCaseNo(2).Locked = False
      txtCaseNo(3).Locked = False
      Command1.Enabled = False
      bolLock = False
      SetText 'Added by Morgan 2017/2/10
   '修改
   ElseIf strSaveConfirm = MsgText(4) Then
      txtCaseNo(0).Locked = True
      txtCaseNo(1).Locked = True
      txtCaseNo(2).Locked = True
      txtCaseNo(3).Locked = True
      Command1.Enabled = False
      bolLock = False
   '瀏覽
   Else
      txtCaseNo(0).Locked = False
      txtCaseNo(1).Locked = False
      txtCaseNo(2).Locked = False
      txtCaseNo(3).Locked = False
      Command1.Enabled = True
      bolLock = True
   End If
   
   txtTF(2).Locked = bolLock
   txtTF21.Locked = bolLock 'Added by Morgan 2019/8/15
   Check1.Enabled = Not bolLock 'Added by Morgan 2017/2/10
End Sub

Private Sub txtTF_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtTF_Validate(Index As Integer, Cancel As Boolean)
   'Modified by Morgan 2017/2/10
   'If Index = 2 And Not txtTF(2).Locked Then
   If (Index = 2 Or Index = 21) And txtTF(Index).Enabled Then
   'end 2017/2/10
      SetFee
   End If
End Sub

Private Sub txtTF21_Validate(Cancel As Boolean)
   If txtTF21.Enabled And txtTF21.Locked = False Then SetFee 'Added by Morgan 2019/8/15
End Sub
