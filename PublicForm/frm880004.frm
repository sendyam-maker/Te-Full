VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880004 
   BorderStyle     =   1  '單線固定
   Caption         =   "再確認"
   ClientHeight    =   6984
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   8784
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6984
   ScaleWidth      =   8784
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame Frame8 
      BorderStyle     =   0  '沒有框線
      Height          =   4665
      Left            =   870
      TabIndex        =   92
      Top             =   1890
      Width           =   6975
      Begin VB.CommandButton cmdRemove8 
         Caption         =   "移除 ->"
         Height          =   285
         Left            =   5790
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   2610
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd8 
         Caption         =   "<- 新增"
         Height          =   285
         Left            =   5790
         TabIndex        =   107
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdExit8 
         Caption         =   "結束"
         Height          =   315
         Left            =   5880
         TabIndex        =   103
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdQuery8 
         Caption         =   "查詢"
         Default         =   -1  'True
         Height          =   315
         Left            =   4200
         TabIndex        =   98
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   100
         Top             =   540
         Width           =   2115
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   3390
         MaxLength       =   2
         TabIndex        =   97
         Top             =   127
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   96
         Top             =   127
         Width           =   345
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   95
         Top             =   127
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   94
         Top             =   127
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "審定號／"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   99
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號："
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   93
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "本所案號　　　　　審定/申請號　　案件名稱"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   5
         Left            =   300
         TabIndex        =   115
         Top             =   2010
         Width           =   4995
      End
      Begin MSForms.ListBox lstData8 
         Height          =   2250
         Left            =   150
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   2280
         Width           =   5580
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "9842;3969"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cmbTM05 
         Height          =   285
         Left            =   1380
         TabIndex        =   113
         Top             =   990
         Width           =   5385
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "9499;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblfm2 
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   112
         Top             =   1710
         Width           =   1545
         Size            =   "2725;503"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblfm2 
         Height          =   285
         Index           =   2
         Left            =   1350
         TabIndex        =   111
         Top             =   1710
         Width           =   1545
         Size            =   "2725;503"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblfm2 
         Height          =   285
         Index           =   1
         Left            =   2370
         TabIndex        =   110
         Top             =   1380
         Width           =   4305
         Size            =   "7594;503"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblfm2 
         Height          =   285
         Index           =   0
         Left            =   1350
         TabIndex        =   109
         Top             =   1380
         Width           =   945
         Size            =   "1667;503"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         Caption         =   "承辦人員："
         Height          =   285
         Index           =   4
         Left            =   3570
         TabIndex        =   106
         Top             =   1710
         Width           =   945
      End
      Begin VB.Label Label11 
         Caption         =   "智權人員："
         Height          =   285
         Index           =   3
         Left            =   300
         TabIndex        =   105
         Top             =   1710
         Width           =   945
      End
      Begin VB.Label Label11 
         Caption         =   "申請人1："
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   104
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label Label11 
         Caption         =   "案件名稱："
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   102
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label Label11 
         Caption         =   "申請案號："
         Height          =   285
         Index           =   0
         Left            =   390
         TabIndex        =   101
         Top             =   780
         Width           =   1005
      End
   End
   Begin VB.Frame Frame7 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame6"
      Height          =   3195
      Left            =   4680
      TabIndex        =   64
      Top             =   -30
      Width           =   7455
      Begin VB.CommandButton cmdExit7 
         Caption         =   "結束(&X)"
         Height          =   400
         Left            =   6600
         TabIndex        =   65
         Top             =   0
         Width           =   800
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "客戶C類來函和B類收文的相關設定：依照上述的規則讀取符合的第一次設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   88
         Top             =   2850
         Width           =   6615
      End
      Begin VB.Label Label9 
         Caption         =   "112/1/9 取消：(Y/X編號為8碼的優先於編號6碼，若已有8碼的設定則不再讀取6碼的設定) "
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   13
         Left            =   150
         TabIndex        =   87
         Top             =   480
         Visible         =   0   'False
         Width           =   7275
      End
      Begin VB.Label Label9 
         Caption         =   "只抓Y+X1~X5設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   38
         Left            =   1050
         TabIndex        =   86
         Top             =   1740
         Width           =   5655
      End
      Begin VB.Label Label9 
         Caption         =   "３ Y or X："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   72
         Top             =   2250
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "１ 有個案：符合1.1個案+ (Y+X)就不用再抓1.2個案+Y or X"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   71
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "複數備註讀取的規則分成三階段，依序符合階段就不再抓資料："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   5460
      End
      Begin VB.Label Label9 
         Caption         =   "２ Y+X："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   69
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "先抓Y設定再抓X1~X5設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   20
         Left            =   1140
         TabIndex        =   68
         Top             =   2250
         Width           =   3735
      End
      Begin VB.Label Label9 
         Caption         =   "1.1個案+ (Y+X)：(Y+X)要抓Y+X1~X5設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   1080
         TabIndex        =   67
         Top             =   1020
         Width           =   5415
      End
      Begin VB.Label Label9 
         Caption         =   "1.2個案+ Y or X：先抓Y設定再抓X1~X5設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   1080
         TabIndex        =   66
         Top             =   1290
         Width           =   5655
      End
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame6"
      Height          =   6615
      Left            =   1248
      TabIndex        =   49
      Top             =   72
      Width           =   7455
      Begin VB.CommandButton cmdExit6 
         Caption         =   "結束(&X)"
         Height          =   400
         Left            =   6600
         TabIndex        =   61
         Top             =   0
         Width           =   800
      End
      Begin VB.Label Label9 
         Caption         =   "案號-送件日.FIX.、案號-送件日.FIX.SEQ. 或 .FIG.PDF)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   225
         Index           =   37
         Left            =   2610
         TabIndex        =   85
         Top             =   2700
         Width           =   4815
      End
      Begin VB.Label Label9 
         Caption         =   "(同時存放在English_Vers)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   36
         Left            =   4140
         TabIndex        =   84
         Top             =   840
         Width           =   2085
      End
      Begin VB.Label Label9 
         Caption         =   "同時存放外文提申本之修正(*.FIX.ORI.PDF)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   35
         Left            =   1770
         TabIndex        =   83
         Top             =   3630
         Width           =   4695
      End
      Begin VB.Label Label9 
         Caption         =   "例如：FCP058901.RES.doc、FCP058901.RES.PDF"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   34
         Left            =   1800
         TabIndex        =   82
         Top             =   5280
         Width           =   5655
      End
      Begin VB.Label Label9 
         Caption         =   "存放相似比對結果檔案(*.RES.DOC/DOCX或PDF檔)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   33
         Left            =   1800
         TabIndex        =   81
         Top             =   5040
         Width           =   5415
      End
      Begin VB.Label Label9 
         Caption         =   "例如：FCP053555-1080107.fix_u.doc、FCP053555-1080107.Fig.pdf"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   32
         Left            =   1800
         TabIndex        =   80
         Top             =   2970
         Width           =   5295
      End
      Begin VB.Label Label9 
         Caption         =   "存放最終版中說或圖式PDF (案號-送件日.FIX_U、"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   225
         Index           =   31
         Left            =   1800
         TabIndex        =   79
         Top             =   2505
         Width           =   5535
      End
      Begin VB.Label Label9 
         Caption         =   "存放設計說明書(*.DES.DOC/DOCX檔、或TXT檔)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   30
         Left            =   1800
         TabIndex        =   78
         Top             =   1950
         Width           =   5535
      End
      Begin VB.Label Label9 
         Caption         =   "例如：FCP058901.SEP.doc、FCP058901.SEP.xls"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   29
         Left            =   1800
         TabIndex        =   77
         Top             =   4800
         Width           =   5655
      End
      Begin VB.Label Label9 
         Caption         =   "路徑:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   28
         Left            =   1800
         TabIndex        =   76
         Top             =   4320
         Width           =   5175
      End
      Begin VB.Label Label9 
         Caption         =   "中說發文後刪除檔案。"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   27
         Left            =   1800
         TabIndex        =   75
         Top             =   5535
         Width           =   5655
      End
      Begin VB.Label Label9 
         Caption         =   "存放提供翻譯參考用之說明書(*.SEP.)，不限制檔案類型；"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   26
         Left            =   1800
         TabIndex        =   74
         Top             =   4575
         Width           =   5415
      End
      Begin VB.Label Label9 
         Caption         =   "SIMILAR_RESULT："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   73
         Top             =   4335
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   $"frm880004.frx":0000
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   12
         Left            =   240
         TabIndex        =   63
         Top             =   6030
         Width           =   7215
      End
      Begin VB.Label Label9 
         Caption         =   "、電子送件專用檔(*.ZIP)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   11
         Left            =   1800
         TabIndex        =   62
         Top             =   570
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "例如：FCP058901.fix.doc、FCP058901.fix_u.doc"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   1800
         TabIndex        =   60
         Top             =   2250
         Width           =   3975
      End
      Begin VB.Label Label9 
         Caption         =   "、中說修正本(*.FIX_U.DOC、*.COR_U.DOC或DOCX檔、或TXT檔)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   1800
         TabIndex        =   59
         Top             =   1500
         Width           =   5625
      End
      Begin VB.Label Label9 
         Caption         =   "、中說圖檔(*.FIG.PDF)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   58
         Top             =   1755
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "存放中說替換本(*.FIX.DOC、*.COR.DOC或DOCX檔、或TXT檔)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   57
         Top             =   1245
         Width           =   5535
      End
      Begin VB.Label Label9 
         Caption         =   "例如：FCP058901.fix.ori.doc、FCP058901.fix.ori.pdf"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   1770
         TabIndex        =   56
         Top             =   3870
         Width           =   4485
      End
      Begin VB.Label Label9 
         Caption         =   "存放外文本Word檔(*.ORI.DOC或DOCX檔、或TXT檔)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   1770
         TabIndex        =   55
         Top             =   3405
         Width           =   4695
      End
      Begin VB.Label Label9 
         Caption         =   "例如：FCP058901.fix.ori.pdf"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   54
         Top             =   840
         Width           =   2325
      End
      Begin VB.Label Label9 
         Caption         =   "存放外文提申本(*.ORI.PDF、*.FIX_*.PDF)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   53
         Top             =   300
         Width           =   4695
      End
      Begin VB.Label Label9 
         Caption         =   "English_Vers："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   450
         TabIndex        =   52
         Top             =   3405
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "電子送件暫存區："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "專利案件："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   50
         Top             =   1245
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame5"
      Height          =   1995
      Left            =   5430
      TabIndex        =   40
      Top             =   1770
      Width           =   7095
      Begin VB.CommandButton Cmd5 
         Caption         =   "確定(&O)"
         Height          =   400
         Left            =   5760
         TabIndex        =   41
         Top             =   120
         Width           =   912
      End
      Begin MSForms.ListBox ListBox5 
         Height          =   1140
         Left            =   120
         TabIndex        =   45
         Top             =   690
         Width           =   6855
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "12091;2011"
         MatchEntry      =   0
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         Caption         =   "分類＆說明　記錄日期　 內容"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   44
         Top             =   420
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "選取項目將設為失效指示"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3600
         TabIndex        =   42
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame4"
      Height          =   2775
      Left            =   0
      TabIndex        =   22
      Top             =   3480
      Width           =   5175
      Begin VB.CommandButton cmdIn 
         Caption         =   "確定(&O)"
         Height          =   400
         Left            =   3840
         TabIndex        =   21
         Top             =   2040
         Width           =   912
      End
      Begin VB.TextBox txtIn 
         Height          =   300
         Index           =   1
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   20
         Top             =   2400
         Width           =   465
      End
      Begin VB.TextBox txtIn 
         Height          =   300
         Index           =   0
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2040
         Width           =   700
      End
      Begin MSForms.Label Label6 
         Height          =   300
         Left            =   2550
         TabIndex        =   91
         Top             =   2040
         Width           =   1470
         VariousPropertyBits=   27
         Caption         =   "Label6"
         Size            =   "2593;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   39
         Top             =   1680
         Width           =   765
         VariousPropertyBits=   27
         Caption         =   "lblData(8)"
         Size            =   "1349;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         Caption         =   "案件性質："
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   38
         Top             =   1120
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   37
         Top             =   1395
         Width           =   2325
         VariousPropertyBits=   27
         Caption         =   "lblData(7)"
         Size            =   "4101;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         Caption         =   "智權人員："
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   1400
         Width           =   975
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   35
         Top             =   1125
         Width           =   2805
         VariousPropertyBits=   27
         Caption         =   "lblData(6)"
         Size            =   "4948;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   34
         Top             =   840
         Width           =   4005
         VariousPropertyBits=   27
         Caption         =   "lblData(5)"
         Size            =   "7064;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   33
         Top             =   555
         Width           =   1005
         VariousPropertyBits=   27
         Caption         =   "lblData(4)"
         Size            =   "1773;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   32
         Top             =   285
         Width           =   765
         VariousPropertyBits=   27
         Caption         =   "lblData(2)"
         Size            =   "1349;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   31
         Top             =   0
         Width           =   3045
         VariousPropertyBits=   27
         Caption         =   "lblData(1)"
         Size            =   "5371;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblData 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   30
         Top             =   0
         Width           =   1005
         VariousPropertyBits=   27
         Caption         =   "lblData(0)"
         Size            =   "1773;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         Caption         =   "智權報價點數："
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "案件名稱："
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   560
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "申請國家："
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   280
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "申請人："
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "簽核點數："
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   24
         Top             =   2423
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "提高點數簽核主管："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   2063
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   2505
      Left            =   5280
      TabIndex        =   7
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdPath 
         Height          =   330
         Index           =   2
         Left            =   4800
         Picture         =   "frm880004.frx":009E
         Style           =   1  '圖片外觀
         TabIndex        =   48
         Top             =   2032
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Text            =   "\\Pat1\OA_SCAN"
         Top             =   2040
         Visible         =   0   'False
         Width           =   4665
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Text            =   "\\Pat1\OA_SCAN"
         Top             =   1413
         Visible         =   0   'False
         Width           =   4665
      End
      Begin VB.CommandButton cmdPath 
         Height          =   330
         Index           =   1
         Left            =   4815
         Picture         =   "frm880004.frx":01A0
         Style           =   1  '圖片外觀
         TabIndex        =   17
         Top             =   1405
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Text            =   "\\Pat1\OA_SCAN"
         Top             =   786
         Visible         =   0   'False
         Width           =   4665
      End
      Begin VB.CommandButton cmdPath 
         Height          =   330
         Index           =   0
         Left            =   4815
         Picture         =   "frm880004.frx":02A2
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   778
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txtPcnt 
         Height          =   300
         Index           =   1
         Left            =   3120
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtPcnt 
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "確定(&O)"
         Height          =   400
         Index           =   2
         Left            =   4200
         TabIndex        =   11
         Top             =   70
         Width           =   912
      End
      Begin VB.Label lbl5 
         AutoSize        =   -1  'True
         Caption         =   "迅達附件存放路徑:"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   1794
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lbl5 
         AutoSize        =   -1  'True
         Caption         =   "捷恩凱附件存放路徑:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1167
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lbl5 
         AutoSize        =   -1  'True
         Caption         =   "舜禹附件存放路徑:"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "圖示頁數："
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   143
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "說明書頁數："
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   143
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   5175
      Begin VB.CommandButton cmdOut 
         Caption         =   "確定(&O)"
         Height          =   400
         Left            =   4200
         TabIndex        =   4
         Top             =   480
         Width           =   912
      End
      Begin VB.Label Label10 
         Caption         =   "文字貼上：Ctrl+V 文字複製：Ctrl+C"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   210
         TabIndex        =   90
         Top             =   360
         Width           =   3285
      End
      Begin MSForms.TextBox txtMemo 
         Height          =   990
         Left            =   120
         TabIndex        =   89
         Top             =   570
         Width           =   3975
         VariousPropertyBits=   -1466941413
         MaxLength       =   500
         Size            =   "7011;1746"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "備註內容會以中文25字折行也可以按Enter換行,最多6行"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "回前畫面(&U)"
         Height          =   400
         Index           =   1
         Left            =   3900
         TabIndex        =   3
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "確定(&O)"
         Height          =   400
         Index           =   0
         Left            =   2940
         TabIndex        =   2
         Top             =   120
         Width           =   912
      End
      Begin VB.TextBox txtReKey 
         Height          =   264
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   4872
      End
      Begin VB.Label Label1 
         Caption         =   "由於改變之欄位為重要欄位，請再一次確認"
         Height          =   480
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   4875
      End
   End
End
Attribute VB_Name = "frm880004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ;Frame2(txtMeo)、Frame4(Label6、lblData(index))、Frame5(ListBox5)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit
'Modified by Lydia 2017/10/17 改到basUpdate
'Public txtTemp As TextBox, bolTheSame As Boolean
'Modified by Morgan 2021/12/9
'Public txtTemp As TextBox
Public txtTemp As Object

'Added by Lydia 2015/06/17
'Memo by Lydia 2017/07/26 iStiu 空白或1:再確認畫面;
'                               2:客戶請款明細表(frm210146)的列印備註;
'                               3:撰寫信函(frm090401)的FC翻譯郵件,執行時要輸入的資料;
'                               4:轉定稿(basQuery.PUB_Cache2Letter):期限通知報價點數-提高點數;
'                               5:各項指示(frm12040159)在加入有效分類時,確認同類指示是否有效;
'                               6.工程師上傳作業(frm090905): 歸檔說明(Added by Lydia 2018/05/18)
'                               7.核駁及審查意見通知函備註(frm060507): 顯示備註的優先順序,參考basUpdate.PUB_GetIncomMemoNew (Added by Lydia 2018/08/29)
'                               8.多案案號輸入: 內商主管機關來電處理記錄(frm020108)和FCT聯絡單列印(frm1106)共用 'Added by Lydia 2022/07/28
'end 2017/07/26
Public iStiu As Integer
Public mPreForm As Form
'Added by Lydia 2015/11/11 抓本機電腦符合外翻名稱的資料夾路徑
Dim oFileSys As New FileSystemObject
Dim FT14 As String '指定外翻的員工編號
'選擇資料夾
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
(lpBI As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
(ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
(ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
   hwndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type
'end 2015/11/11 -------------
'Added by Lydia 2016/05/23 定稿報價輸入簽核主管和點數
Public m_LCV01 As String '定稿變數暫存檔PK
Public m_LCV02 As String '定稿變數暫存檔PK
Public m_LCV03 As String '定稿變數暫存檔PK
Public m_LCV06 As String '智權人員確認
Public m_LC07 As String '定稿報價確認人員(智權人員)
'Added by Lydia 2017/07/26 各項指示
Public m_TempList As String '傳入資料
Dim varTempArr As Variant '暫存變數陣列
'Added by Lydia 2022/07/28
Dim intQ As Integer, intK As Integer
Dim tmpObj As Object
Dim strPCase(1 To 4) As String

Private Sub cmdok_Click(Index As Integer)
    'Modified by Lydia 2015/11/02 改成select
    'If Index = 0 Then
    Select Case Index
        Case 0
            If txtReKey.Text <> txtTemp.Text Then
               ShowMsg MsgText(9201)
            Else
               Unload Me
               bolTheSame = True
            End If
    'Else
        Case 1
            txtTemp.Text = txtTemp.Tag
            Unload Me
    'End If
        'Added by Lydia 2015/11/02
        Case 2
        'Mark by Lydia 2025/03/13 已不再使用
'            mPreForm.mPcnt1 = Val(txtPcnt(0))
'            mPreForm.mPcnt2 = Val(txtPcnt(1))
'            'Added by Lydia 2015/11/11 指定外翻路徑
''            If txtPath(0).Text = "" Then
''                MsgBox "請指定" & Replace(lbl5(0).Caption, ":", "") & " !", vbExclamation
''                Exit Sub
''            ElseIf txtPath(1).Text = "" Then
''                MsgBox "請指定" & Replace(lbl5(1).Caption, ":", "") & " !", vbExclamation
''                Exit Sub
''            End If
'            'Modified by Lydia 2017/09/28 +迅達F5698
'            'For intI = 0 To 1
'            'Remove by Lydia 2018/03/30 改從卷宗區抓資料
'            'Modified by Lydia 2018/04/30 P案請回復之前做法,由給舜禹翻譯資料夾帶檔案
'            If Left(m_TempList, 1) = "P" Then
'                For intI = 0 To 2
'                    If txtPath(intI) = "" Then
'                        MsgBox "請指定" & Replace(lbl5(intI).Caption, ":", "") & " !", vbExclamation
'                        txtPath(intI).SetFocus
'                        Exit Sub
'                    Else
'                       If oFileSys.FolderExists(txtPath(intI)) = False Then
'                           MsgBox lbl5(intI).Caption & " [ " & txtPath(intI) & " ] 不存在，請確認！", vbCritical
'                           txtPath(intI).SetFocus
'                           Exit Sub
'                       End If
'                    End If
'                Next intI
'                'Modified by Lydia 2018/01/04 F5588-> 外翻_舜禹
'                If FT14 = 外翻_舜禹 Then
'                   mPreForm.strLoadPath = txtPath(0).Text
'                 'Modified by Lydia 2018/01/04 F5653-> 外翻_捷恩凱
'                ElseIf FT14 = 外翻_捷恩凱 Then
'                   mPreForm.strLoadPath = txtPath(1).Text
'                'Added by Lydia 2017/09/28
'                'Modified by Lydia 2018/01/04 F5698-> 外翻_迅達
'                ElseIf FT14 = 外翻_迅達 Then
'                   mPreForm.strLoadPath = txtPath(2).Text
'                'end 2017/09/28
'                End If
'                'end 2015/11/11
'            End If 'end 2018/04/30
'
'            Unload Me
        'end ---- Mark by Lydia 2025/03/13 已不再使用
    End Select
End Sub

Private Sub Form_Load()
Me.Height = 2100: txtMemo.Text = "" 'Added by Lydia 2015/06/17 預設表單高度

Me.Width = 5310  'Memo by Lydia 2017/09/28 預設表單寬度

MoveFormToCenter Me
bolTheSame = False
'Modified by Lydia 2015/11/02 改成case
''Added by Lydia 2015/06/17 切換Frame
'If iStiu = 2 Then
'  Frame1.Visible = False
'  Frame2.Top = 0
'  If Len(mPreForm.rptMemo) > 0 Then txtMemo.Text = mPreForm.rptMemo
'  'txtMemo.SetFocus
'Else
'  Frame2.Visible = False
'End If
Frame1.Visible = False: Frame2.Visible = False: Frame3.Visible = False
Frame4.Visible = False 'Added by Lydia 2016/05/23
Frame5.Visible = False 'Added by Lydia 2017/07/26
Frame6.Visible = False 'Added by Lydia 2018/05/18
Frame7.Visible = False 'Added by Lydia 2018/08/29
Frame8.Visible = False 'Added by Lydia 2022/07/28

    Select Case iStiu
        Case 2
            Frame2.Visible = True: Frame2.Top = 0
            If Len(mPreForm.rptMemo) > 0 Then txtMemo.Text = mPreForm.rptMemo
        Case 3 'Added by Lydia 2015/11/02 +外翻_翻譯案件工作確認單
'        'Mark by Lydia 2025/03/13 已不再使用
'            Frame3.Visible = True: Frame3.Top = 0
'            'Added by Lydia 2017/09/28 預設高度和靠左對齊
'            'Modified by Lydia 2018/04/30
'            'Me.Height = Frame3.Height + 600
'            If Left(m_TempList, 1) = "P" Then
'                 Me.Height = Frame3.Height + 600
'                 For intI = 0 To 2
'                      lbl5(intI).Visible = True
'                      txtPath(intI).Visible = True
'                      cmdPath(intI).Visible = True
'                 Next
'            Else
'                 Me.Height = 1800
'                 For intI = 0 To 2
'                      lbl5(intI).Visible = False
'                      txtPath(intI).Visible = False
'                      cmdPath(intI).Visible = False
'                 Next
'            End If
'            'end 2018/04/30
'            Frame3.Left = 0
'            'end 2017/09/28
'            cmdOK(2).Default = True
'            If mPreForm.mPcnt1 > 0 Then txtPcnt(0) = mPreForm.mPcnt1
'            If mPreForm.mPcnt2 > 0 Then txtPcnt(1) = mPreForm.mPcnt2
'            'Added by Lydia 2015/11/11 抓本機電腦符合外翻名稱的資料夾路徑
'            '讀取前次設定路徑
'            'Modified by Lydia 2018/01/04 F5588-> 外翻_舜禹
'            'Remove by Lydia 2018/03/30 改從卷宗區抓資料
'            'Modified by Lydia 2018/04/30 P案請回復之前做法,由給舜禹翻譯資料夾帶檔案
'            If Left(m_TempList, 1) = "P" Then
'                    txtPath(0).Tag = 外翻_舜禹
'                    txtPath(0).Text = GetSetting("TAIE", txtPath(0).Tag, UCase(mPreForm.Name) & "Dir", "")  '江蘇舜禹翻譯
'                    'Modified by Lydia 2018/01/04 F5653-> 外翻_捷恩凱
'                    txtPath(1).Tag = 外翻_捷恩凱
'                    txtPath(1).Text = GetSetting("TAIE", txtPath(1).Tag, UCase(mPreForm.Name) & "Dir", "")  '南京捷恩凱
'                    'Added by Lydia 2017/09/28
'                    'Modified by Lydia 2018/01/04 F5698-> 外翻_迅達
'                    txtPath(2).Tag = 外翻_迅達
'                    txtPath(2).Text = GetSetting("TAIE", txtPath(2).Tag, UCase(mPreForm.Name) & "Dir", "")  '迅達翻譯社
'                    'end 2017/09/28
'            End If  'end 2018/04/30
'            FT14 = mPreForm.strLoadPath
'            'end 2015/11/11
        'end--- Mark by Lydia 2025/03/13 已不再使用
        'Added by Lydia 2016/05/23 智權人員調整報價，請加入若為年費(605)超過8點,維持費(606)及延展費(607)超過10點時，發E-mail給主管批示；
        Case 4             '在產生報價定稿時，由程序輸入提高點數簽核主管及簽核點數
            Frame4.Visible = True: Frame4.Top = 100
            Me.Height = 3600
            txtIn(0).Text = "": txtIn(1).Text = "": Label6.Caption = ""
            GetData4
        'end 2016/05/23
        'Added by Lydia 2017/07/26 各項指示有效清單
        Case 5
            Frame5.Visible = True: Frame5.Top = 0: Frame5.Left = 0
            Me.Height = Frame5.Height + 360
            Me.Width = Frame5.Width + 120
            Me.Caption = "各項指示有效清單"
            GetList5
        'end 2017/07/26
        'Added by Lydia 2018/05/18 工程師上傳作業
        Case 6
            Frame6.Visible = True: Frame6.Top = 0: Frame6.Left = 0
            Me.Height = Frame6.Height + 360
            Me.Width = Frame6.Width + 120
            Me.Caption = "工程師上傳作業-歸檔說明"
            Label9(28).Caption = "路徑:" & Pub_GetSpecMan("FCP相似比對結果暫存") 'Added by Lydia 2024/12/31 翻譯參考用之word版說明書(*.SEP)和相似比對結果(*.RES)放在一起
        'end 2018/05/18
        'Added by Lydia 2018/08/29
        Case 7
            Frame7.Visible = True: Frame7.Top = 0: Frame7.Left = 0
            Me.Height = Frame7.Height + 360
            Me.Width = Frame7.Width + 120
            'Modified by Lydia 2021/11/12 優先順序說明=>規則說明
            Me.Caption = "核駁及審查意見通知函備註-規則說明"
        'end 2018/08/29
        'Added by Lydia 2022/07/28
        Case 8
            Frame8.Visible = True: Frame8.Top = 0: Frame8.Left = 0
            Me.Height = Frame8.Height + 360
            Me.Width = Frame8.Width + 120
            Me.Caption = "多案案號輸入"
            Call ClearData8
            lstData8.Tag = ""
            Option1(0).Value = True
            If m_TempList <> "" Then  '載入上次輸入的案號
                varTempArr = Split(m_TempList, ",")
                '照原順序排
                For intK = UBound(varTempArr) To LBound(varTempArr) Step -1
                    If Trim(varTempArr(intK)) <> "" Then
                        Call ChgCaseNo(Trim(varTempArr(intK)), strPCase)
                        If strPCase(1) <> "" And strPCase(2) <> "" Then
                            Text1(0) = strPCase(1): Text1(1) = strPCase(2)
                            Text1(2) = strPCase(3): Text1(3) = strPCase(4)
                            Call DoQuery8
                            If lblfm2(0).Caption <> "" Then
                                Call cmdAdd8_Click
                            End If
                        End If
                    End If
                Next intK
            End If
        'end 2022/07/28
        Case Else
            Frame1.Visible = True
            cmdOK(0).Default = True 'Added by Lydia 2015/10/30
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm880004 = Nothing
End Sub

Private Sub txtReKey_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
'Added by Lydia 2015/06/17
Private Sub txtMemo_GotFocus()
   TextInverse txtMemo
End Sub
'Added by Lydia 2015/06/17
Private Sub cmdOut_Click()
Dim nRow As Integer
Dim strRes As String 'Added by Lydia 2016/11/16
'Memo by Lydia 2016/11/16 原本備註內容用Enter換行，現在加上超過中文25字會自動折行，行數最多6行。

    'Frame2 傳回值
    strExc(1) = Trim(txtMemo.Text)
     Do While strExc(1) <> ""
        intI = InStr(strExc(1), vbCrLf)
        'Modified by Lydia 2016/11/16 +自動折行(中文25字)
        'If intI > 0 Then
        '   strExc(0) = Left(strExc(1), intI - 1)
        '   strExc(1) = Mid(strExc(1), intI + 2)
        '   nRow = nRow + 1
        If intI > 0 Or GetTextLength(strExc(1)) > 50 Then
           'Modified by Lydia 2016/11/22 無Enter
           'strExc(0) = Left(strExc(1), intI - 1)
           'strExc(1) = Mid(strExc(1), intI + 2)
           strExc(0) = Left(strExc(1), IIf(intI > 0, intI - 1, Len(strExc(1))))
           strExc(1) = IIf(intI > 0, Mid(strExc(1), intI + 2), "")
           If GetTextLength(strExc(0)) > 50 Then
              strExc(3) = PUB_StrToStr(strExc(0), 50)
              strExc(0) = Replace(strExc(0), strExc(3), "")
              strExc(1) = strExc(0) & vbCrLf & strExc(1)
              strRes = strRes & strExc(3) & vbCrLf
              'nRow = nRow + 1 'Remove by Lydia 2016/11/22
           Else
              strRes = strRes & strExc(0) & vbCrLf
           End If
           nRow = nRow + 1
        'end 2016/11/16
        Else
           strRes = strRes & strExc(1) 'Added by Lydia 2016/11/16
           strExc(1) = ""
        End If
     Loop
     'Added by Lydia 2024/10/22 排除資策會Excel的格式
     If Val(m_TempList) > 0 Then
        mPreForm.rptMemo = Trim(strRes)
        Unload Me
     Else
     'ends 2024/10/22
        If nRow > 5 Then
           MsgBox "備註最多6行!"
        ElseIf CheckLengthIsOK(txtMemo, txtMemo.MaxLength) = True Then
               'Modified by Lydia 2016/11/16
               'mPreForm.rptMemo = Trim(txtMemo.Text)
               mPreForm.rptMemo = Trim(strRes)
               Unload Me
        End If
     End If
End Sub
'Added by Lydia 2015/11/02
Private Sub txtPCnt_GotFocus(Index As Integer)
   TextInverse txtPcnt(Index)
End Sub

Private Sub txtPCnt_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
'end 2015/11/02

'Added by Lydia 2015/11/11 設定本機電腦符合外翻名稱的資料夾路徑
Private Sub cmdPath_Click(Index As Integer)
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = "請選擇" & Replace(lbl5(Index).Caption, ":", "")
   
   With tBrowseInfo
       .hwndOwner = Me.hWnd
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS
   End With
   
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   
   If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       SaveSetting "TAIE", txtPath(Index).Tag, UCase(mPreForm.Name) & "Dir", sBuffer
       txtPath(Index).Text = sBuffer
   End If
End Sub
'Added by Lydia 2016/05/23
Private Sub txtIn_GotFocus(Index As Integer)
   TextInverse txtIn(Index)
End Sub
Private Sub txtIn_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtIn_Validate(Index As Integer, Cancel As Boolean)
Dim StrStr As String

Select Case Index
Case 0
   If txtIn(Index) = "" Then
      MsgBox "請輸入在職的員工編號!", vbCritical
      Cancel = True
      txtIn(Index).SetFocus
   Else
      StrStr = GetStaffName(txtIn(Index))
      If StrStr <> "" Then
         Label6.Caption = StrStr
      Else
         Label6.Caption = ""
         MsgBox "請輸入在職的員工編號!", vbCritical
         Cancel = True
         txtIn(Index).SetFocus
         txtIn_GotFocus Index
      End If
   End If
Case 1
   If Val(txtIn(Index).Text) < 0 Or Trim(txtIn(Index)) = "" Then
      txtIn(Index).Text = ""
      MsgBox "請輸入正數或零!", vbCritical
      Cancel = True
      txtIn(Index).SetFocus
      txtIn_GotFocus Index
   End If
End Select
End Sub
Private Function CheckInData() As Boolean
Dim StrStr As String
Dim intA As Integer
Dim Cancel As Boolean
CheckInData = False

   For intA = 0 To 1
      txtIn_Validate intA, Cancel
      If Cancel Then
         txtIn(intA).SetFocus
         txtIn_GotFocus intA
         Exit Function
      End If
   Next
   
   If m_LC07 <> "" Then
      StrStr = GetFLOW001Person(m_LC07, Flow_接洽單)
      If StrStr <> Trim(txtIn(0).Text) Then
         If MsgBox(GetStaffName(m_LC07) & "的預設簽核主管為" & StrStr & " " & GetStaffName(StrStr) & "，請確認輸入的主管是否正確?", vbCritical + vbYesNo) = vbNo Then
            txtIn(0).SetFocus
            txtIn_GotFocus 0
            Exit Function
         End If
      End If
   End If
   
   If m_LCV06 <> "" And txtIn(1) <> "" And Val(m_LCV06) <> Val(txtIn(1)) Then
      If MsgBox("簽核點數:" & txtIn(1).Text & " ，智權報價點數:" & m_LCV06 & " ，請確認簽核點數是否正確?", vbCritical + vbYesNo) = vbNo Then
         txtIn(1).SetFocus
         txtIn_GotFocus 1
         Exit Function
      End If
   End If

CheckInData = True
End Function

Private Sub cmdIn_Click()
Dim strUpd As String
Dim intS As Integer
Dim strA1 As String, strA2 As String
Dim rsAD As New ADODB.Recordset

   If CheckInData Then
      '抓費用金額
      strA1 = "select b.lcv03,b.lcv04,b.lcv06,a.lcv04 dot1,a.lcv06 dot2 from lettercachevar a,lettercachevar b where b.lcv01='" & m_LCV01 & "' and b.lcv02='" & m_LCV02 & "' and b.lcv03='" & Replace(m_LCV03, "點數", "") & "' " & _
             "and a.lcv01(+)=b.lcv01 and a.lcv02(+)=b.lcv02 and a.lcv03(+)=b.lcv03||'點數'"
      intS = 1
      Set rsAD = ClsLawReadRstMsg(intS, strA1)
      '依輸入點數，更新費用金額
      If intS = 1 Then
         strA1 = Val("" & rsAD.Fields("lcv04")) - Val("" & rsAD.Fields("dot1")) * 1000
         strA2 = Val(strA1) + Val(txtIn(1)) * 1000
         If Val(txtIn(1)) = 0 Then strA2 = "0"
         strUpd = "update lettercachevar set lcv09='" & Trim(txtIn(0).Text) & "', lcv10=" & strA2 & _
                 " where lcv01='" & m_LCV01 & "' and lcv02='" & m_LCV02 & "' and lcv03='" & rsAD.Fields("lcv03") & "' "
         cnnConnection.Execute strUpd, intS
      End If
      strUpd = "update lettercachevar set lcv09='" & Trim(txtIn(0).Text) & "', lcv10=" & IIf(Trim(txtIn(1)) = "", "0", Trim(txtIn(1))) & _
               " where lcv01='" & m_LCV01 & "' and lcv02='" & m_LCV02 & "' and lcv03='" & m_LCV03 & "' "
      cnnConnection.Execute strUpd, intS
      '先更新簽核主管,等到所有費用項目簽核後再更新金額
      strUpd = "update lettercachevar set lcv09='" & Trim(txtIn(0).Text) & "'" & _
               " where lcv01='" & m_LCV01 & "' and lcv02='" & m_LCV02 & "' and lcv03 in ('費用合計','點數合計') "
      cnnConnection.Execute strUpd, intS
      'email通知智權人員
      If Val(txtIn(1)) = 0 Then
         strUpd = "主管不同意調整點數，將維持以程序點數通知客戶！"
      'Added by Lydia 2016/08/31 主管調整點數
      ElseIf Val(txtIn(1)) <> Val(m_LCV06) Then
         strUpd = "您報價" & m_LCV06 & "點，但主管改為 " & txtIn(1) & " 點！"
      'end 2016/08/31
      Else
         strUpd = "主管同意調整為 " & txtIn(1) & " 點！"
      End If
      'Modified by Lydia 2016/10/11 因為報價項目可能不只一項,將信函合併
      'PUB_SendMail strUserNum, m_LC07, "", lblData(4).Caption & "催" & lblData(6).Caption & "期限通知報價點數，" & strUpd, vbCrLf & "同主旨"
      strKeyPoint = strKeyPoint & vbCrLf & "催" & lblData(6).Caption & "期限通知報價點數，" & strUpd & vbCrLf & String(100, "-") & vbCrLf
      
      Unload Me
   End If
End Sub

Private Sub GetData4()
Dim oLabel As Label
Dim ii As Integer
Dim rsA As New ADODB.Recordset
Dim strR As String

   For Each oLabel In lblData
      oLabel.Caption = ""
   Next
   If m_LCV01 <> "" Then
      strR = "SELECT PA26,NVL(NVL(CU04,NVL(CU06,RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90))),PA26) 申請人" & _
      ", NA03 申請國家, NP02||'-'||NP03||DECODE(NP04||NP05,'000','','-'||NP04||'-'||NP05) 本所案號" & _
      ", NVL(PA05,NVL(PA07,PA06)) 案件名稱, DECODE(NA01,'000',CPM03,CPM04) 案件性質" & _
      " ,NP10||' '||ST02 智權人員" & _
      " From LETTERCACHE, NEXTPROGRESS, CASEPROGRESS, PATENT, CUSTOMER, Nation, CASEPROPERTYMAP, staff" & _
      " WHERE LC01='" & m_LCV01 & "' AND LC02='" & m_LCV02 & "' AND NP01(+)=LC01 AND NP22(+)=LC02" & _
      " AND CP09(+)=NP01 AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA01 IS NOT NULL" & _
      " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
      " AND NA01(+)=PA09 AND CPM01(+)=NP02 AND CPM02=DECODE(LC11,CP66,CP10,NP07)" & _
      " AND ST01(+)=NP10  "
      ii = 1
      Set rsA = ClsLawReadRstMsg(ii, strR)
      If ii = 1 Then
         lblData(0) = "" & rsA.Fields("PA26")
         lblData(1) = "" & rsA.Fields("申請人")
         lblData(2) = "" & rsA.Fields("申請國家")
         lblData(4) = "" & rsA.Fields("本所案號")
         lblData(5) = "" & rsA.Fields("案件名稱")
         lblData(6) = IIf(m_LCV03 <> "費用點數", Replace(m_LCV03, "點數", ""), "" & rsA.Fields("案件性質"))
         lblData(7) = "" & rsA.Fields("智權人員")
      End If
         lblData(8) = m_LCV06
   End If
End Sub
'end 2016/05/23

'Added by Lydia 2017/07/26
Private Sub Cmd5_Click()
    GetList5_Ans
    Unload Me
End Sub

'Added by Lydia 2017/07/26 各項指示有效清單
Private Sub GetList5()
Dim intA As Integer
Dim strA1 As String
   ListBox5.Clear
   
   If m_TempList <> "" Then
      varTempArr = Empty
      varTempArr = Split(m_TempList, "||")
      
      For intA = 0 To UBound(varTempArr) - 1
         strA1 = Trim(varTempArr(intA))
         If strA1 <> "" Then
            ListBox5.AddItem Mid(strA1, InStr(strA1, "<@>") + 3) '區隔PKey
            'ListBox5.ItemData(intA) = intA 'Mark by Lydia 2022/01/07 Form 2.0沒有ItemData屬性
         End If
      Next intA
   End If
End Sub

'Added by Lydia 2017/07/26 取得勾選的各項指示有效清單
Private Sub GetList5_Ans()
Dim intA As Integer
Dim strA1 As String
   
   If ListBox5.ListCount > 0 Then
      For intA = 0 To ListBox5.ListCount - 1
         If ListBox5.Selected(intA) = True Then
            strA1 = strA1 & Mid(varTempArr(intA), 1, InStr(varTempArr(intA), "<@>") - 1) & ","
         End If
      Next intA
   End If
   
   mPreForm.Tag = strA1
End Sub

'Added by Lydia 2017/09/28
Private Sub txtPath_GotFocus(Index As Integer)
  TextInverse txtPath(Index)
End Sub

'Added by Lydia 2018/05/18
Private Sub cmdExit6_Click()
    Unload Me
End Sub

'Added by Lydia 2018/08/29
Private Sub cmdExit7_Click()
    Unload Me
End Sub

'Added by Lydia 2022/07/28
Private Sub cmdExit8_Click()
    mPreForm.cmdInput.Tag = lstData8.Tag
    Unload Me
End Sub

'Added by Lydia 2022/07/28
Private Sub Option1_Click(Index As Integer)
   If Index = 0 Then
      Text1(0).Enabled = True
      Text1(1).Enabled = True
      Text1(2).Enabled = True
      Text1(3).Enabled = True
      Text1(4).Enabled = False
   ElseIf Index = 1 Then
      Text1(0).Enabled = False
      Text1(1).Enabled = False
      Text1(2).Enabled = False
      Text1(3).Enabled = False
      Text1(4).Enabled = True
   End If
End Sub

'Added by Lydia 2022/07/28
Private Sub Text1_GotFocus(Index As Integer)
    TextInverse Text1(Index)
End Sub

'Added by Lydia 2022/07/28
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
     KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2022/07/28
Private Sub cmdQuery8_Click()
  
  If Option1(0).Value = True Then
     If Trim(Text1(0)) = "" Or Trim(Text1(1)) = "" Then
         If Trim(Text1(0)) = "" Then
             intQ = 0
         Else
             intQ = 1
         End If
         MsgBox "請輸入本所案號!!", vbCritical
         Text1(intQ).SetFocus
         Text1_GotFocus intQ
         Exit Sub
     Else
         Text1(0) = Trim(Text1(0))
         Text1(1) = Trim(Text1(1))
         Text1(2) = Left(Trim(Text1(2)) & "0", 1)
         Text1(3) = Left(Trim(Text1(3)) & "00", 2)
     End If
  Else
     If Trim(Text1(4)) = "" Then
         MsgBox "請輸入審定號／申請案號!!", vbCritical
         Text1(4).SetFocus
         Text1_GotFocus 4
         Exit Sub
     End If
     Text1(4) = Trim(Text1(4))
  End If
   
  Call DoQuery8
End Sub

'Added by Lydia 2022/07/28
Private Sub DoQuery8()
Dim rsA As New ADODB.Recordset
Dim strR As String
Dim strTmpA As String

   Call ClearData8(False)
   
   If Option1(0).Value = True Then
       strR = strR & " and tm01='" & Text1(0) & "' and tm02='" & Text1(1) & "' and tm03='" & Text1(2) & "' and tm04='" & Text1(3) & "' "
   ElseIf Option1(1).Value = True Then
       strR = strR & " and (tm12 ='" & Text1(4) & "' or tm15 ='" & Text1(4) & "') "
   End If
  
   strR = "select tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm23, nvl(cu05,nvl(cu04,cu06)) cname1,nvl(tm15,tm12) tm1512 " & _
             "From trademark, customer where substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & strR
   intQ = 0
   Set rsA = ClsLawReadRstMsg(intQ, strR)
   If intQ = 1 Then
       lblfm2(0) = "" & rsA.Fields("tm23")
       lblfm2(1) = "" & rsA.Fields("cname1")
       
       Text1(0) = rsA.Fields("tm01")
       Text1(1) = rsA.Fields("tm02")
       Text1(2) = rsA.Fields("tm03")
       Text1(3) = rsA.Fields("tm04")
       Text1(4) = "" & rsA.Fields("tm1512")
       For Each tmpObj In Text1
          tmpObj.Tag = tmpObj.Text
       Next

       If "" & rsA.Fields("tm07") <> "" Then
           cmbTM05.AddItem "外：" & rsA.Fields("tm07"), 0
       End If
       If "" & rsA.Fields("tm06") <> "" Then
           cmbTM05.AddItem "英：" & rsA.Fields("tm06"), 0
       End If
       cmbTM05.AddItem "中：" & rsA.Fields("tm05"), 0
       cmbTM05.ListIndex = 0
       
       Select Case Trim(Text1(0).Text)
           Case "T"
               'T案，則以PUB_GetAKindSalesNo抓智權人員
               strTmpA = PUB_GetAKindSalesNo(Text1(0), Text1(1), Text1(2), Text1(3))
           Case "FCT"
               'FCT案，則以PUB_GetFCTSalesNo抓智權人員
               strTmpA = PUB_GetFCTSalesNo(Text1(0), Text1(1), Text1(2), Text1(3))
       End Select
       lblfm2(2) = GetStaffName(strTmpA, True)
       '抓該案號之最後承辦人非程序(ST03<>'P22')者，若離職則抓部門主管
       'Added by Lydia 2023/12/25
       If strSrvDate(1) >= 新部門啟用日 Then
           strR = "select cp05,cp09,cp14,st04 as stype, st02 as s1name,nvl(a0924,a0909) as s2no, getstaffnamelist(nvl(a0924,a0909)) as s2name " & _
                  "from caseprogress,staff,acc090,acc090new where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "' " & _
                  "and cp14=st01(+) and st03<>'P22' and st03=a0901(+) and st93=a0921(+) "
       Else
       'end 2023/12/25
           strR = "select cp14,s1.st04 stype,s1.st02 s1name,a0909 s2no,s2.st02 s2name " & _
                     "from caseprogress,acc090,staff s1,staff s2 where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "' " & _
                     "and cp14=s1.st01(+) and s1.st03=a0901(+) and a0909=s2.st01 and cp57 is null and s1.st03 <>'P22' order by cp05 desc,cp09 desc"
       End If
       intQ = 1
       Set rsA = ClsLawReadRstMsg(intQ, strR)
       If intQ = 1 Then
          If rsA.Fields("stype") = "1" Then
             lblfm2(3).Caption = "" & rsA.Fields("s1name")
          Else
             lblfm2(3).Caption = "" & rsA.Fields("s2name")
          End If
        End If
   End If
End Sub

'Added by Lydia 2022/07/28
Private Sub cmdAdd8_Click()

    If lblfm2(0).Caption = "" Or Text1(0).Tag & Text1(1).Tag & Text1(2).Tag & Text1(3).Tag & Text1(4).Tag <> Text1(0) & Text1(1) & Text1(2) & Text1(3) & Text1(4) Then
        MsgBox "請先查詢本所案號／審定號／申請案號!!", vbCritical
        Exit Sub
    End If
    
    '檢查
    If InStr(m_LCV01 & ",", Text1(0) & Text1(1) & Text1(2) & Text1(3)) > 0 Or InStr(m_LCV01 & ",", Text1(4)) > 0 Then
        MsgBox Text1(0) & "-" & Text1(1) & "-" & Text1(2) & "-" & Text1(3) & "與前一畫面的案號重覆!!", vbCritical
        Exit Sub
    ElseIf InStr(lstData8.Tag & ",", Text1(0) & Text1(1) & Text1(2) & Text1(3)) > 0 Then
        If Option1(0).Value = True Then
            MsgBox Text1(0) & "-" & Text1(1) & "-" & Text1(2) & "-" & Text1(3) & "已存在清單!!", vbCritical
        ElseIf Option1(1).Value = True Then
            MsgBox Text1(4) & "已存在清單!!", vbCritical
        End If
        Exit Sub
    End If
    
    '加入: 本所案號 , 審定號／申請案號
    lstData8.Tag = Text1(0) & Text1(1) & Text1(2) & Text1(3) & "," & lstData8.Tag
    lstData8.AddItem PUB_StrToStr(Text1(0) & "-" & Text1(1) & "-" & Text1(2) & "-" & Text1(3), 20, True) & PUB_StrToStr(Text1(4), 20, True) & Mid(cmbTM05.List(0), 3), 0
    
    Call ClearData8
End Sub

'Added by Lydia 2022/07/28
Private Sub cmdRemove8_Click()
    lstData8.Tag = PUB_RemoveListBox2(lstData8, lstData8.Tag)
End Sub

'Added by Lydia 2022/07/28
Private Sub ClearData8(Optional ByVal bReset As Boolean = True)
    
    For Each tmpObj In Text1
        If bReset = True Then tmpObj.Text = ""
        tmpObj.Tag = ""
    Next
    For Each tmpObj In lblfm2
         tmpObj.Caption = ""
    Next
    cmbTM05.Clear
End Sub


