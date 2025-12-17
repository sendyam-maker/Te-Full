VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc44q0 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶扣繳明細核對表"
   ClientHeight    =   5890
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5890
   ScaleWidth      =   9180
   Begin VB.TextBox txtNote 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   885
      Left            =   450
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "產生資料中，暫時不要使用Excel..."
      Top             =   5700
      Visible         =   0   'False
      Width           =   8520
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "寄發，作業執行完畢通知信"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   345
      Left            =   4860
      TabIndex        =   47
      Top             =   4500
      Width           =   3885
   End
   Begin VB.CommandButton cmdMail 
      BackColor       =   &H00C0FFC0&
      Caption         =   "寄發催扣繳憑單MAIL"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   630
      Style           =   1  '圖片外觀
      TabIndex        =   46
      Top             =   4290
      Width           =   3810
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "作業說明"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7920
      Style           =   1  '圖片外觀
      TabIndex        =   39
      Top             =   1260
      Width           =   1005
   End
   Begin VB.CheckBox Check5 
      Caption         =   "寄測試信箱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   6810
      TabIndex        =   38
      Top             =   540
      Width           =   2235
   End
   Begin VB.ComboBox Combo2 
      Height          =   260
      Left            =   6135
      Style           =   2  '單純下拉式
      TabIndex        =   20
      Top             =   5250
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.ComboBox Combo1 
      Height          =   260
      Left            =   6135
      Style           =   2  '單純下拉式
      TabIndex        =   19
      Top             =   4920
      Width           =   2820
   End
   Begin VB.OptionButton Option1 
      Caption         =   "只列印有扣繳但不含已確認過的"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3810
      TabIndex        =   16
      Top             =   3480
      Width           =   3975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "只列印 當年度 有扣繳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3810
      TabIndex        =   15
      Top             =   3210
      Width           =   3975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "去年無扣單但今年有扣稅"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   14
      Top             =   3510
      Width           =   3135
   End
   Begin VB.CheckBox Check3 
      Caption         =   "目前無應收帳款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   13
      Top             =   3210
      Width           =   3135
   End
   Begin VB.CheckBox Check2 
      Caption         =   $"Frmacc44q0.frx":0000
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   12
      Top             =   2910
      Width           =   8715
   End
   Begin VB.CheckBox Check1 
      Caption         =   "稅額達 2,001 以上但未扣繳 (一收款單號合計達 2,001 含同收款單之所有收據)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   11
      Top             =   2640
      Width           =   8655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel(&E) (剔除代填繳款書客戶)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4860
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   3840
      Width           =   3810
   End
   Begin VB.TextBox txtType 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1830
      Width           =   612
   End
   Begin VB.TextBox txtComp 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2265
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "L"
      Top             =   1470
      Width           =   612
   End
   Begin VB.TextBox txtComp 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1305
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "1"
      Top             =   1470
      Width           =   612
   End
   Begin VB.TextBox txtCustNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1305
      MaxLength       =   9
      TabIndex        =   2
      Top             =   410
      Width           =   1572
   End
   Begin VB.TextBox txtCustNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3225
      MaxLength       =   9
      TabIndex        =   3
      Top             =   410
      Width           =   1572
   End
   Begin VB.CommandButton cmdLikeSearch 
      Caption         =   "搜尋"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8295
      TabIndex        =   1
      Top             =   90
      Width           =   675
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1110
      Width           =   612
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      MaxLength       =   5
      TabIndex        =   4
      Top             =   750
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "寄發扣繳核對函(一年一次)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   630
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   3840
      Width           =   3810
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   90
      TabIndex        =   31
      Top             =   2190
      Width           =   8985
      Begin VB.CheckBox Check6 
         Caption         =   "稅額達 2,001 以上但未扣繳 (單筆 2,001 含同收款單之所有收據)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   150
         Width           =   8235
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   320
      Left            =   4110
      TabIndex        =   6
      Top             =   1110
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   564
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   320
      Left            =   5690
      TabIndex        =   7
      Top             =   1110
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   564
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.ComboBox cboTitle 
      Height          =   320
      Left            =   1310
      TabIndex        =   0
      Top             =   80
      Width           =   6810
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12012;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "（年度扣繳核對資料請找電腦中心另外提供主機）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   3810
      TabIndex        =   45
      Top             =   1860
      Width           =   5040
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "四 檢查扣繳明細""複合""列印客戶住址/扣繳明細/抬頭等是否正確"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   90
      TabIndex        =   43
      Top             =   5610
      Width           =   6260
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "三 以mail通知的扣繳明細請先處理"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   90
      TabIndex        =   42
      Top             =   5400
      Width           =   3420
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "二 確認收據抬頭三個字以身份別是否正確"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   90
      TabIndex        =   41
      Top             =   5190
      Width           =   4130
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "一 催業務同仁已收款請速繳財務"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   90
      TabIndex        =   40
      Top             =   4980
      Width           =   3230
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "（每一萬號一個檔案）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   170
      Left            =   4830
      TabIndex        =   37
      Top             =   4230
      Width           =   4010
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "點陣印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   4890
      TabIndex        =   36
      Top             =   5280
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "複合\傳真列印A4,印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   3600
      TabIndex        =   35
      Top             =   4950
      Width           =   2490
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "（列印條件為整年度時請將收款期間條件刪除）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   290
      Left            =   3120
      TabIndex        =   34
      Top             =   1470
      Width           =   4430
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   5520
      TabIndex        =   33
      Top             =   1110
      Width           =   260
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "收款期間"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   3060
      TabIndex        =   32
      Top             =   1140
      Width           =   980
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印年度扣繳核對明細注意事項："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   90
      TabIndex        =   30
      Top             =   4770
      Width           =   3380
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印別                (1.單一 2.複合)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   29
      Top             =   1860
      Width           =   3360
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   2040
      TabIndex        =   28
      Top             =   1530
      Width           =   260
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   240
      TabIndex        =   27
      Top             =   1530
      Width           =   980
   End
   Begin MSForms.Label lblSales 
      Height          =   260
      Left            =   2430
      TabIndex        =   26
      Top             =   810
      Width           =   1190
      VariousPropertyBits=   19
      Caption         =   "SalesName"
      Size            =   "2090;450"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   2990
      TabIndex        =   25
      Top             =   440
      Width           =   260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   240
      TabIndex        =   24
      Top             =   1140
      Width           =   980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "客戶代號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   240
      TabIndex        =   23
      Top             =   440
      Width           =   980
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收據抬頭"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   240
      TabIndex        =   22
      Top             =   110
      Width           =   980
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   240
      TabIndex        =   21
      Top             =   780
      Width           =   980
   End
End
Attribute VB_Name = "Frmacc44q0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/4/15 Form2.0已修改 (Printer列印未改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
'Modify by Morgan 2006/11/17 格式大改
Option Explicit

Dim lngPageNo As Long '頁數
Dim lngXo As Long, lngYo As Long, lngX As Long, lngY As Long '列印位置
Dim bol1stPage As Boolean '是否首頁
Dim iChrPix As Integer '字寬
Dim iRowPix As Integer '列高
Dim PLeft() '欄位X位置
Dim stCustNo As String, stTitle As String, stUniNo As String, stCompName As String '表頭資料
Dim lngSum(1 To 5) As Long '小計
Dim adoquery As New ADODB.Recordset
'Dim adocheck As New ADODB.Recordset
Private Const DDollar As String = "###,###,###,##0" '金錢格式
Dim bolStarSign As Boolean
'Add By Sindy 2013/12/2
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer
'2013/12/2 END
Dim intTxtCounter As Integer 'Add By Sindy 2015/12/21
'Add By Sindy 2014/10/21
Dim stTel As String, stFax As String, stMail As String
Dim stCU15 As String, stAccNote As String, stSaleArea As String, stSales As String
Dim strPrinter As String, strEmp As String, strEMP_Tel As String
Dim longRow_Mail As Long, longRow_Fax As Long, longRow_Other As Long, bolPrintA4 As Boolean
Dim strTempFolder As String
Dim TempFileName As String, ff As Integer
'2014/10/21 END
Dim bolWriteData As Boolean 'Add By Sindy 2014/12/9 是否有寫入資料至Excel
Dim dblTotOther As Double, dblTotMail As Double, dblTotFax As Double 'Add By Sindy 2014/12/23
Dim bolMailSendErr As Boolean 'Add By Sindy 2016/12/12
Dim strEmailAttch As String 'Add By Sindy 2020/4/24
'Add By Sindy 2022/4/15
Dim xlsAnnuity_A4 As New Excel.Application
Dim wksAnnuity_A4 As New Worksheet
'2022/4/15 END
Dim stAccPerson As String, stTxtPerson As String  'Add by Amy 2024/05/17
Dim m_AccMail As String 'Add By Sindy 2024/10/8 扣繳信件信箱


'Add By Sindy 2016/2/24
Private Sub Check6_Click()
   Option1(0).Value = False
   Option1(1).Value = False
   '第一、二、三點不可重覆勾選
   If Check6.Value = 1 Then
      Check1.Value = 0
      Check2.Value = 0
      cmdMail.Enabled = False 'Add By Sindy 2019/3/20
   Else
      'Add By Sindy 2019/3/14
      cmdMail.Enabled = False
      If Check2.Value = 1 And Check6.Value = 0 And _
         Check1.Value = 0 And Check3.Value = 0 And _
         Check4.Value = 0 Then
         cmdMail.Enabled = True '只勾選(只印全年服務費超過 20,000 或 服務費未超過 20,000 但有扣繳的 或 當年有收款且有建信箱者)才可操作此功能
      End If
      '2019/3/14 END
   End If
End Sub

Private Sub Check1_Click()
   Option1(0).Value = False
   Option1(1).Value = False
   'Modify By Sindy 2014/12/23 第一點和第二點不可重覆勾選
   If Check1.Value = 1 Then
      Check2.Value = 0
      Check6.Value = 0 'Add By Sindy 2016/2/24
      cmdMail.Enabled = False 'Add By Sindy 2019/3/20
   Else
      'Add By Sindy 2019/3/14
      cmdMail.Enabled = False
      If Check2.Value = 1 And Check6.Value = 0 And _
         Check1.Value = 0 And Check3.Value = 0 And _
         Check4.Value = 0 Then
         cmdMail.Enabled = True '只勾選(只印全年服務費超過 20,000 或 服務費未超過 20,000 但有扣繳的 或 當年有收款且有建信箱者)才可操作此功能
      End If
      '2019/3/14 END
   End If
End Sub

Private Sub Check2_Click()
   Option1(0).Value = False
   Option1(1).Value = False
   'Modify By Sindy 2014/12/23 第一點和第二點不可重覆勾選
   If Check2.Value = 1 Then
      Check1.Value = 0
      Check6.Value = 0 'Add By Sindy 2016/2/24
   End If
   'Add By Sindy 2019/3/14
   cmdMail.Enabled = False
   If Check2.Value = 1 And Check6.Value = 0 And _
      Check1.Value = 0 And Check3.Value = 0 And _
      Check4.Value = 0 Then
      cmdMail.Enabled = True '只勾選(只印全年服務費超過 20,000 或 服務費未超過 20,000 但有扣繳的 或 當年有收款且有建信箱者)才可操作此功能
   End If
   '2019/3/14 END
End Sub

Private Sub Check3_Click()
   Option1(0).Value = False
   Option1(1).Value = False
   If Check3.Value = 1 Then
      cmdMail.Enabled = False 'Add By Sindy 2019/3/20
   End If
   'Add By Sindy 2019/3/14
   cmdMail.Enabled = False
   If Check2.Value = 1 And Check6.Value = 0 And _
      Check1.Value = 0 And Check3.Value = 0 And _
      Check4.Value = 0 Then
      cmdMail.Enabled = True '只勾選(只印全年服務費超過 20,000 或 服務費未超過 20,000 但有扣繳的 或 當年有收款且有建信箱者)才可操作此功能
   End If
   '2019/3/14 END
End Sub

Private Sub Check4_Click()
   Option1(0).Value = False
   Option1(1).Value = False
   If Check4.Value = 1 Then
      cmdMail.Enabled = False 'Add By Sindy 2019/3/20
   End If
   'Add By Sindy 2019/3/14
   cmdMail.Enabled = False
   If Check2.Value = 1 And Check6.Value = 0 And _
      Check1.Value = 0 And Check3.Value = 0 And _
      Check4.Value = 0 Then
      cmdMail.Enabled = True '只勾選(只印全年服務費超過 20,000 或 服務費未超過 20,000 但有扣繳的 或 當年有收款且有建信箱者)才可操作此功能
   End If
   '2019/3/14 END
End Sub

'Add By Sindy 2019/3/20
Private Sub cmdMail_Click()
   If FormCheck = True Then
      If txtType = "2" Then '2.複合
         MsgBox "寄發E-Mail，列印別只能是單一資料!!!", vbInformation
         txtType.SetFocus
         Exit Sub
      'Add By Sindy 2021/2/22
      ElseIf Check5.Value = 0 And Check7.Value = 0 Then
         If MsgBox("寄發完畢，要發通知信給您嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            Check7.Value = 1
         End If
      End If
      '2021/2/22 END
      
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Call Process(False, True)
      'FormClear
      Me.Enabled = True
      If MaskEdBox2 <> MsgText(29) Then
         PUB_SaveLastDate Me.Name, "MaskEdBox2", ChangeTDateStringToTString(MaskEdBox2)
      End If
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub Command1_Click()
   If FormCheck = True Then
      'Add By Sindy 2022/12/21
      If txtType <> "2" Then '2.複合 是列印紙本,無寄信的問題
      '2022/12/21 END
         'Add By Sindy 2016/12/12
         If MsgBox("傳送要數小時不能作業，且不能同時同一台電腦上操作「扣繳憑單查詢及列印」寄發Mail(因為會同樣產生pdf檔)，" & vbCrLf & _
                   "會造成程式互相衝突(任何會產生pdf檔的程式都請先不要操作)！是否繼續列印？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
            Exit Sub
         End If
         
         'Add By Sindy 2020/12/23
         If Check5.Value = 0 And Check7.Value = 0 Then
            If MsgBox("寄發完畢，要發通知信給您嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
               Check7.Value = 1
            End If
         End If
         '2020/12/23 END
      End If
      '2016/12/12 END
      
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Call Process(False)
      'FormClear
      Me.Enabled = True
      'Add By Sindy 2014/10/17
      If MaskEdBox2 <> MsgText(29) Then
         PUB_SaveLastDate Me.Name, "MaskEdBox2", ChangeTDateStringToTString(MaskEdBox2)
      End If
'      If PUB_GetLastDate(Me.Name, "MaskEdBox2") <> "" Then
'         MaskEdBox1.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox2"))
'      End If
      '2014/10/17 END
      Screen.MousePointer = vbDefault
   End If
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

'產生Excel檔
Private Sub Command2_Click()
   If FormCheck = True Then
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Call Process(True)
      'FormClear
      Me.Enabled = True
      'Add By Sindy 2014/11/12
      If MaskEdBox2 <> MsgText(29) Then
         PUB_SaveLastDate Me.Name, "MaskEdBox2", ChangeTDateStringToTString(MaskEdBox2)
      End If
      '2014/11/12 END
      Screen.MousePointer = vbDefault
   End If
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Command3_Click()
   Frmacc44q0_1.Show vbModal '強制回應表單
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
'   If KeyCode <> vbKeyEscape Then
'      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
'   End If
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 9300, 6310
   FormClear
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   
   'Add By Sindy 2014/10/21
   PUB_SetPrinter Me.Name, Combo1, strPrinter
'   PUB_SetPrinter Me.Name, Combo2
   Check5.Caption = "寄測試信箱（" & strUserNum & "）"
   '2014/10/21 END
   
   'Add By Sindy 2022/3/30
   If Dir(strExcelPath, vbDirectory) = "" Then
      MkDir strExcelPath
   End If
   'Add by Amy 2024/05/17 財務2個特殊設定拆成3個
   If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
       stAccPerson = Pub_GetSpecMan("財務處應收處理人員")
   Else
      stAccPerson = Pub_GetSpecMan("財務處總帳人員")
   End If
   stTxtPerson = stAccPerson '取第一個人
   If InStr(stTxtPerson, ";") > 0 Then stTxtPerson = Mid(stTxtPerson, 1, Val(InStr(stTxtPerson, ";")) - 1)
   'end 2024/05/17
   
   m_AccMail = "taieacc@taie.com.tw" 'Add By Sindy 2024/10/8 扣繳信件信箱
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2014/10/21
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
'   '若印表機變動, 則更新列印設定
'   If Me.Combo2.Text <> Me.Combo2.Tag Then
'      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
'   End If
   '2014/10/21 END
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set adoquery = Nothing
'   Set adocheck = Nothing
   Set Frmacc44q0 = Nothing
End Sub

Private Sub Option1_GotFocus(Index As Integer)
   Check1.Value = 0
   Check2.Value = 0
   Check3.Value = 0
   Check4.Value = 0
   Check6.Value = 0 'Add By Sindy 2016/2/24
   'Add By Sindy 2025/3/3
   If Index = 0 Then
      cmdMail.Enabled = True
   Else
   '2025/3/3 END
      cmdMail.Enabled = False 'Add By Sindy 2019/3/20
   End If
End Sub

Private Sub Text4_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text4.IMEMode = 2
   CloseIme
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Text4 = "" Then
      lblSales = ""
   Else
      lblSales = GetStaffName(Text4)
      If lblSales = "" Then
         MsgBox "智權人員不存在，請重新輸入！"
         Cancel = True
      End If
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      KeyAscii = 0
   End If
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   txtCustNo(0) = "X00001000"
   txtCustNo(1) = "X99999ZZZ"
   cboTitle.Clear
   Text4 = ""
   lblSales.Caption = ""
   
   'Modify By Sindy 2014/10/15 解開Mark
   'Remove by Morgan 2009/12/14 收款日期條件取消
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   'Modify By Sindy 2016/12/16 取消預設
'   'Add By Sindy 2014/10/17
'   If PUB_GetLastDate(Me.Name, "MaskEdBox2") <> "" Then
'      '讀取上次輸入的迄止日期+1日
'      MaskEdBox1.Text = ChangeTStringToTDateString(DBDATE(DateAdd("d", 1, Format(PUB_GetLastDate(Me.Name, "MaskEdBox2"), "####/##/##"))) - 19110000)
'   Else
'      MaskEdBox1.Text = CFDate(Left(Val(strSrvDate(1)) - 19110000, 3) & "0101")
'   End If
'   '2014/10/17 END
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   'Modify By Sindy 2016/12/16 取消預設
'   If Left(Val(strSrvDate(1)) - 19110000, 3) <> Left(MaskEdBox1.Text, 3) Then
'      MaskEdBox2.Text = CFDate(Left(MaskEdBox1.Text, 3) & "1231")
'   Else
'      MaskEdBox2.Text = CFDate(Left(Val(strSrvDate(1)) - 19110000, 3) & "1231")
'   End If
   MaskEdBox2.Mask = DFormat
   
   If Text5 = "" Then
      '預設年度斷4月
      If Val(Right(strSrvDate(2), 4)) >= 401 Then
         Text5 = strSrvDate(2) \ 10000
      Else
         Text5 = strSrvDate(2) \ 10000 - 1
      End If
   End If
   Check1.Value = 0
   'Modify By Sindy 2025/3/3 勾選欄位,要預設在第一個選項
   'Check2.Value = 1 'Modify By Sindy 2024/12/16 勾選欄位,要預設在第三個選項
   Check2.Value = 0
   'Check6.Value = 0 'Add By Sindy 2016/2/24
   Check6.Value = 1
   '2025/3/3 END
   Check3.Value = 0
   Check4.Value = 0
   Option1(0).Value = False
   Option1(1).Value = False
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   FormCheck = False
   
   If Text5 = "" Then
      MsgBox "扣繳年度不可空白！"
      Text5.SetFocus
      Exit Function
   End If
   If txtType = "" Then
      MsgBox "請輸入列印別！"
      txtType.SetFocus
      Exit Function
   End If
   'Add By Sindy 2013/12/30
   If txtComp(1) = "J" Then
      MsgBox "公司別不可為 J"
      txtComp(1).SetFocus
      Exit Function
   End If
   If txtComp(2) = "J" Then
      MsgBox "公司別不可為 J"
      txtComp(2).SetFocus
      Exit Function
   End If
   '2013/12/30 END
'   ElseIf cboTitle <> "" Then
'      FormCheck = True
'   ElseIf txtCustNo(0) <> "" And txtCustNo(1) <> "" Then
'      FormCheck = True
   'Add By Sindy 2014/12/9
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      If ChkDate(Replace(MaskEdBox1, "/", "")) = False Then
         MaskEdBox1.SetFocus
         Exit Function
      End If
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      If ChkDate(Replace(MaskEdBox2, "/", "")) = False Then
         MaskEdBox2.SetFocus
         Exit Function
      End If
   End If
   '2014/12/9 END
   If cboTitle = "" And (txtCustNo(0) = "" Or txtCustNo(1) = "") Then
      MsgBox "[收據抬頭]或[客戶代號]條件不可同時空白！"
      cboTitle.SetFocus
      Exit Function
   End If
   
   FormCheck = True
End Function

Private Sub cboTitle_Click()
   If cboTitle.ListIndex > 0 Then
      If txtCustNo(0).Text = "" Then
         txtCustNo(0).Text = Right(cboTitle.Text, 9)
      ElseIf txtCustNo(1).Text = "" Then
         txtCustNo(1).Text = Right(cboTitle.Text, 9)
      End If
      strExc(1) = cboTitle.List(cboTitle.ListIndex)
      cboTitle.List(0) = RTrim(Left(strExc(1), Len(strExc(1)) - 9))
   End If
   cboTitle.ListIndex = 0
End Sub

Private Sub cboTitle_GotFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'cboTitle.IMEMode = 1
   OpenIme
End Sub

Private Sub cboTitle_Validate(Cancel As Boolean)
   If CheckLen(Label1, cboTitle, 100) = MsgText(603) Then
      Cancel = True
      Exit Sub
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub cmdLikeSearch_Click()
   If cboTitle.Text = "" Then
      MsgBox "請輸入收據抬頭！", vbCritical
   Else
      'Modify By Sindy 2022/4/15 從 cboTitle_KeyPress 搬過來
      If txtCustNo(0) <> "" Or txtCustNo(1) <> "" Or cboTitle.ListCount > 0 Then
         txtCustNo(0) = "": txtCustNo(1) = ""
         Text4 = "": lblSales = ""
         cboTitle.Clear
      End If
      '2022/4/15 END
      'Modify by Morgan 2007/10/2 改呼叫共用函數
      'Modify by Sindy 2013/12/30
      PUB_AddItem2CboTitle cboTitle, txtCustNo(0), txtCustNo(1), Text5, True
      'end 2007/10/2
   End If
End Sub

Private Sub txtComp_GotFocus(Index As Integer)
    TextInverse txtComp(Index)
End Sub

Private Sub txtComp_KeyPress(Index As Integer, KeyAscii As Integer)
   'Modify By Sindy 2020/4/24 + And Chr(KeyAscii) <> "L"
   If KeyAscii <> 8 And (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And Chr(KeyAscii) <> "L" Then
       KeyAscii = 0
   End If
End Sub

Private Sub txtComp_Validate(Index As Integer, Cancel As Boolean)
    If txtComp(Index) <> "" Then
        If txtComp(1) > txtComp(2) Then
            MsgBox "公司別範圍錯誤！", vbCritical
            Cancel = True
            Call txtComp_GotFocus(Index)
        End If
    End If
End Sub

Private Sub txtCustNo_GotFocus(Index As Integer)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtCustNo(Index).IMEMode = 2
   CloseIme
   If Index = 1 Then
      If Len(txtCustNo(0)) = 6 Then txtCustNo(0) = txtCustNo(0) & "000" 'Add By Sindy 2013/12/4
      If txtCustNo(0) <> "" And txtCustNo(1) = "" Then
         'Modify By Sindy 2014/8/11 999=>ZZZ
         'txtCustNo(1) = Left(txtCustNo(0), 6) & "999" 'Add By Sindy 2013/12/4
         txtCustNo(1) = Left(txtCustNo(0), 6) & "ZZZ"
         txtCustNo(1).SelStart = 6
         txtCustNo(1).SelLength = 3
      Else
         TextInverse txtCustNo(Index)
      End If
   Else
      TextInverse txtCustNo(Index)
   End If
   
End Sub

Private Sub txtCustNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboTitle_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   If txtCustNo(0) <> "" Or txtCustNo(1) <> "" Or cboTitle.ListCount > 0 Then
'      txtCustNo(0) = "": txtCustNo(1) = ""
'      Text4 = "": lblSales = ""
'      cboTitle.Clear
'   End If
End Sub

'Add By Sindy 2013/12/24
Private Sub SetExcelWorksheets()
   'xlsAnnuity_A4.Visible = True
   xlsAnnuity_A4.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsAnnuity_A4.Workbooks.add
   Set wksAnnuity_A4 = xlsAnnuity_A4.Worksheets(1)
   xlsAnnuity_A4.ActiveWindow.Zoom = 75 'Add By Sindy 2016/4/25 瑞婷:畫面比例100%太大了,調整為75%
   '把Excel的警告訊息關掉
   xlsAnnuity_A4.DisplayAlerts = False
   'wksAnnuity_A4.PageSetup.Orientation = xlLandscape '橫印
   wksAnnuity_A4.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity_A4.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity_A4.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity_A4.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity_A4.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity_A4.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity_A4.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   '設定各欄位長度
   wksAnnuity_A4.Columns("A:A").ColumnWidth = 10
   wksAnnuity_A4.Columns("B:B").ColumnWidth = 10
   wksAnnuity_A4.Columns("C:C").ColumnWidth = 10
   wksAnnuity_A4.Columns("D:D").ColumnWidth = 11
   wksAnnuity_A4.Columns("E:E").ColumnWidth = 11
   wksAnnuity_A4.Columns("F:F").ColumnWidth = 14
   wksAnnuity_A4.Columns("G:G").ColumnWidth = 12
   wksAnnuity_A4.Columns("H:H").ColumnWidth = 12
   wksAnnuity_A4.Columns("I:I").ColumnWidth = 25
   'wksAnnuity_A4.Columns("J:J").ColumnWidth = 25 'Add By Sindy 2014/10/15
   intCounter = 1 'intCounter + 1
End Sub

'Add By Sindy 2015/12/21
Private Sub SetExcelWorksheetsTxt()
   xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   xlsAnnuity.ActiveWindow.Zoom = 75 'Add By Sindy 2016/4/25 瑞婷:畫面比例100%太大了,調整為75%
   '把Excel的警告訊息關掉
   xlsAnnuity.DisplayAlerts = False
   'wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 10
   wksAnnuity.Columns("B:B").ColumnWidth = 20
   wksAnnuity.Columns("C:C").ColumnWidth = 15
   intTxtCounter = 1
End Sub

'Add By Sindy 2015/12/21
'Txt表頭
Private Sub PrintHeadTxt_Excel(ByRef iRow As Integer)
Dim strTemp As String
   
   With wksAnnuity
      'iRow = iRow + 1
      .Range("A" & iRow).Value = "客戶編號"
      .Range("B" & iRow).Value = "客戶名稱"
      .Range("C" & iRow).Value = "寄送狀況"
      strTemp = "A" & iRow & ":C" & iRow
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
      End With
   End With
End Sub

'Add By Sindy 2015/12/21
Private Sub PrintDataTxt_Excel(ByRef iRow As Integer, ByVal stCustName As String, _
                               ByVal strData1 As String)
   iRow = iRow + 1
   '客戶編號
   If stCustNo <> "" Then
      wksAnnuity.Range("A" & iRow).Value = stCustNo
   End If
   '客戶名稱
   If stCustName <> "" Then
      wksAnnuity.Range("B" & iRow).Value = stCustName
   End If
   '寄送狀況
   If strData1 <> "" Then
      wksAnnuity.Range("C" & iRow).Value = strData1
   End If
   
   'Add By Sindy 2015/12/23 記錄清單
   'Modify By Sindy 2024/12/24 + ChgSQL
   strSql = "insert into ACCTMP44q0_1(FName,UserID,DT01,DT02,DT03,DT04,DT05)" & _
            " values('" & Me.Name & "','" & strUserNum & "'" & _
            ",'" & stCustNo & "','" & ChgSQL(Trim(stCustName)) & "','" & strData1 & "'" & _
            "," & strSrvDate(2) & "," & Right("000000" & ServerTime, 6) & ")"
   cnnConnection.Execute strSql
   '2015/12/23 END
End Sub

'Add By Sindy 2016/11/1 會計師資料
Private Function GetAcc49(ByVal strTitle As String, ByVal strCustID As String, ByRef strA4904 As String, _
                          ByRef strA4905 As String) As String
Dim rsR As New ADODB.Recordset
   
   GetAcc49 = ""
   strA4905 = "": strA4904 = ""
   'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
   strExc(0) = "SELECT * FROM acc490 WHERE a4901='" & ChgSQL(strTitle) & "'" & _
               " union " & _
               "SELECT * FROM acc490 WHERE a4901='" & Left(Trim(strCustID) & "000000", 9) & "'"
   intI = 1
   Set rsR = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strA4905 = "" & rsR.Fields("A4905") 'E-Mail
      strA4904 = "" & rsR.Fields("A4904") 'Fax
      GetAcc49 = "會計師資料:" & rsR.Fields("a4912").Value & rsR.Fields("a4902").Value & _
                  IIf(Trim("" & rsR.Fields("a4903").Value) <> "", vbCrLf & " 電話:" & rsR.Fields("a4903").Value, "") & _
                  IIf(Trim("" & rsR.Fields("a4904").Value) <> "", vbCrLf & " 傳真:" & rsR.Fields("a4904").Value, "") & _
                  IIf(Trim("" & rsR.Fields("a4905").Value) <> "", vbCrLf & " E-Mail:" & rsR.Fields("a4905").Value, "") & _
                  IIf(Trim("" & rsR.Fields("a4914").Value) <> "", vbCrLf & " 備註:" & rsR.Fields("a4914").Value, "")
   End If
End Function

'Add By Sindy 2025/9/5 信箱為NO的要另列清單
Private Sub PrintA4_Excel_MailNO()
Dim strFilePathName As String, stCustName As String, stCompNo As String
Dim stFaxAcc As String
Dim strGetCustNo As String
Dim strTemp As String
   
   '信箱為NO的要另列清單，帶出編號、收據抬頭、智權人員、電話、傳真、會計備註。
   strExc(0) = "select distinct T02,T15,T23,T17,T18,T21 from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
               " and Upper(T16)='NO'" & _
               " order by T23,T02"
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set xlsAnnuity = New Excel.Application
      TempFileName = PUB_Getdesktop & "\" & strSrvDate(2) & "財務信箱NO.xls"
      
      xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
      xlsAnnuity.Workbooks.add
      Set wksAnnuity = xlsAnnuity.Worksheets(1)
      xlsAnnuity.ActiveWindow.Zoom = 75 '瑞婷:畫面比例100%太大了,調整為75%
      '把Excel的警告訊息關掉
      xlsAnnuity.DisplayAlerts = False
      'wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
      wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
      wksAnnuity.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
      wksAnnuity.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
      wksAnnuity.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
      wksAnnuity.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
      wksAnnuity.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
      wksAnnuity.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
      '設定各欄位長度
      wksAnnuity.Columns("A:A").ColumnWidth = 10
      wksAnnuity.Columns("B:B").ColumnWidth = 25
      wksAnnuity.Columns("C:C").ColumnWidth = 10
      wksAnnuity.Columns("D:D").ColumnWidth = 10
      wksAnnuity.Columns("E:E").ColumnWidth = 10
      wksAnnuity.Columns("F:F").ColumnWidth = 30
      intTxtCounter = 1
      
      With wksAnnuity
         .Range("A" & intTxtCounter).Value = "編號"
         .Range("B" & intTxtCounter).Value = "收據抬頭"
         .Range("C" & intTxtCounter).Value = "智權人員"
         .Range("D" & intTxtCounter).Value = "電話"
         .Range("E" & intTxtCounter).Value = "傳真"
         .Range("F" & intTxtCounter).Value = "會計備註"
         strTemp = "A" & intTxtCounter & ":F" & intTxtCounter
         .Range(strTemp).Select
         With .Application.Selection.Borders(xlEdgeBottom)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
         End With
      End With
      
      With adoquery
         .MoveFirst
         Do While Not .EOF
            intTxtCounter = intTxtCounter + 1
            wksAnnuity.Range("A" & intTxtCounter).Value = "" & adoquery.Fields("T02")
            wksAnnuity.Range("B" & intTxtCounter).Value = "" & adoquery.Fields("T15")
            wksAnnuity.Range("C" & intTxtCounter).Value = "" & adoquery.Fields("T23")
            wksAnnuity.Range("D" & intTxtCounter).Value = "" & adoquery.Fields("T17")
            wksAnnuity.Range("E" & intTxtCounter).Value = "" & adoquery.Fields("T18")
            wksAnnuity.Range("F" & intTxtCounter).Value = "" & adoquery.Fields("T21")
            .MoveNext
         Loop
         strTemp = "A" & intTxtCounter & ":F" & intTxtCounter
         wksAnnuity.Range(strTemp).Select
         With wksAnnuity.Application.Selection.Borders(xlEdgeBottom)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
         End With
         xlsAnnuity.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
         '列印標題
         xlsAnnuity.ActiveSheet.PageSetup.PrintTitleRows = "$1:$6"
         With xlsAnnuity.ActiveSheet.PageSetup
            .Zoom = False
            '.FitToPagesTall = 1 '縮放成一頁高
            .FitToPagesWide = 1 '縮放成一頁寬
            .FitToPagesTall = 1000 'Added by Morgan 2022/4/8 預設為1,筆數多時會縮小
         End With
         'xlsAnnuity.Workbooks(1).PrintOut
         '判斷版本
         If Val(xlsAnnuity.Version) < 12 Then
             xlsAnnuity.Workbooks(1).SaveAs FileName:=TempFileName, FileFormat:=-4143
         Else
             xlsAnnuity.Workbooks(1).SaveAs FileName:=TempFileName, FileFormat:=56
         End If
         xlsAnnuity.Workbooks.Close
         xlsAnnuity.Quit
         Set xlsAnnuity = Nothing
      End With
   End If
   adoquery.Close
End Sub

'Modify By Sindy 2019/3/20 , Optional bolMail As Boolean = False
Private Sub Process(bolExcel As Boolean, Optional bolMail As Boolean = False)
   Dim stCon As String, stCon2 As String, stVTB As String, stVTB1K As String
   Dim st0k0Con As String, st0l0Con As String
   Dim stCustName As String, stCustAddr As String, stCompNo As String
   Dim strMainSql As String 'Add By Sindy 2014/10/21
   Dim PrinterIndex As Integer, i As Integer 'Add By Sindy 2014/10/21
   Dim stCustNo_OK As String 'Add By Sindy 2014/10/22
   Dim strExcelPath As String 'Add By Sindy 2014/12/9
   Dim m_A0K04 As String, dblTot As Double, bolPrintCheck As Boolean
   Dim st1k0Con As String, st0y0Con As String 'Add By Sindy 2015/11/20
   Dim intQ As Integer
   'Add By Sindy 2016/1/8
   Dim m_CU15 As String, m_CU115 As String, m_CU20 As String
   Dim m_CU16 As String, m_CU18 As String, m_CU158 As String
   Dim m_CU22 As String, m_CU80 As String, m_CU30 As String
   Dim m_CU31 As String, m_CU159 As String
   Dim m_CU12 As String '業務區
   Dim m_CU13 As String '智權人員
   '2016/1/8 END
   Dim m_CU168 As String 'Add By Sindy 2016/11/15
   Dim m_CU01 As String, m_CU02 As String
   Dim strA49 As String 'Add By Sindy 2016/11/1
   Dim m_A0m01 As String 'Add By Sindy 2016/11/3
   Dim m_A4905 As String 'Add By Sindy 2016/11/9
   Dim m_A4904 As String 'Add By Sindy 2016/11/17
   Dim Excel_kk As Integer 'Add By Sindy 2016/11/17
   Dim BolContinueReSend As Boolean 'Add By Sindy 2016/12/15
   Dim m_CU172 As String 'Add By Sindy 2017/3/16
   Dim hLocalFile As Long
   
On Error GoTo ErrHnd
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2020/12/30 清除查詢印表記錄檔欄位
   BolContinueReSend = False 'Add By Sindy 2016/12/15
   st1k0Con = "" 'Add By Sindy 2015/11/20
   dblTotOther = 0: dblTotMail = 0: dblTotFax = 0 'Add By Sindy 2014/12/23
   'Add By Sindy 2019/3/20
   cmdMail.Tag = ""
   If bolMail = True Then
      cmdMail.Tag = "EMail"
   End If
   '2019/3/20 END
   'Add By Sindy 2014/11/3
   If txtType = "2" Then '2.複合
      pub_QL05 = pub_QL05 & ";列印別:2.複合" 'Add By Sindy 2020/12/30
      strExcelPath = PUB_Getdesktop & "\複合\" 'Add By Sindy 2014/12/9
      MsgBox "複合資料，請確認傳真及地址是否正確。", vbInformation
   Else
   '2014/11/3 END
      pub_QL05 = pub_QL05 & ";列印別:1.單一" 'Add By Sindy 2020/12/30
      strExcelPath = PUB_Getdesktop & "\單一\" 'Add By Sindy 2014/12/9
      If bolExcel = False Then
         If MsgBox("此作業有發E-Mail功能，確定要執行嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
            Exit Sub
         End If
      End If
   End If
   'Add By Sindy 2014/12/9
   If Dir(strExcelPath, vbDirectory) = "" Then
      MkDir strExcelPath
   End If
   '2014/12/9 END
   
   TempFileName = "" 'Add By Sindy 2014/10/24
   pub_QL05 = pub_QL05 & ";扣繳年度:" & Text5 'Add By Sindy 2020/12/30
   stCon = " and a0k05='2' and nvl(a0k09,0)= 0 and a0k16=" & Val(Text5)
   st1k0Con = " and nvl(a1k12,0)= 0 and a1v09=" & Val(Text5) & " and A1K35 is not null" 'Add By Sindy 2015/11/20
   st0k0Con = stCon
   stCon2 = ""
   If cboTitle.Text <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";收據抬頭:" & cboTitle 'Add By Sindy 2020/12/30
      'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
      stCon = stCon & " and a0k04='" & ChgSQL(cboTitle.Text) & "'"
      st1k0Con = st1k0Con & " and a1k35='" & cboTitle.Text & "'" 'Add By Sindy 2015/11/20
   End If
   
   If txtCustNo(0) <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";客戶代號:" & txtCustNo(0) 'Add By Sindy 2020/12/30
      stCon2 = stCon2 & " and CuNo>='" & txtCustNo(0).Text & "'"
   End If
   If txtCustNo(1) <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";客戶代號:" & txtCustNo(1) 'Add By Sindy 2020/12/30
      stCon2 = stCon2 & " and CuNo<='" & txtCustNo(1).Text & "'"
   End If
   'Add By Sindy 2015/11/20
   If txtCustNo(0) <> MsgText(601) And txtCustNo(1) <> MsgText(601) Then
      st1k0Con = st1k0Con & " and ((a1k03>='" & txtCustNo(0).Text & "' and a1k03<='" & txtCustNo(1).Text & "') or (a1k27>='" & txtCustNo(0).Text & "' and a1k27<='" & txtCustNo(1).Text & "') or (a1k28>='" & txtCustNo(0).Text & "' and a1k28<='" & txtCustNo(1).Text & "'))"
   ElseIf txtCustNo(0) <> MsgText(601) Then
      st1k0Con = st1k0Con & " and (a1k03>='" & txtCustNo(0).Text & "' or a1k27>='" & txtCustNo(0).Text & "' or a1k28>='" & txtCustNo(0).Text & "')"
   ElseIf txtCustNo(1) <> MsgText(601) Then
      st1k0Con = st1k0Con & " and (a1k03>='" & txtCustNo(1).Text & "' or a1k27>='" & txtCustNo(1).Text & "' or a1k28>='" & txtCustNo(1).Text & "')"
   End If
   '2015/11/20 END
   
   If Text4 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";智權人員:" & Text4 'Add By Sindy 2020/12/30
      st0k0Con = st0k0Con & " and a0k20 = '" & Text4 & "'"
      st1k0Con = st1k0Con & " and cp13 = '" & Text4 & "'" 'Add By Sindy 2015/11/20
   End If
   
   st0k0Con = st0k0Con & " And a0k11<>'J'" 'Add By Sindy 2013/12/30
   If txtComp(1) <> "" Then
      pub_QL05 = pub_QL05 & ";公司別:" & txtComp(1) 'Add By Sindy 2020/12/30
      st0k0Con = st0k0Con & " And a0k11>='" & txtComp(1) & "'"
      st1k0Con = st1k0Con & " And a1v03>='" & txtComp(1) & "'" 'Add By Sindy 2015/11/20
   End If
   If txtComp(2) <> "" Then
      pub_QL05 = pub_QL05 & ";公司別:" & txtComp(2) 'Add By Sindy 2020/12/30
      st0k0Con = st0k0Con & " And a0k11<='" & txtComp(2) & "'"
      st1k0Con = st1k0Con & " And a1v03<='" & txtComp(2) & "'" 'Add By Sindy 2015/11/20
   End If
   
   st0l0Con = ""
   st0y0Con = "" 'Add By Sindy 2015/11/20
   'Modify By Sindy 2014/10/15 解開Mark
   'Remove by Morgan 2009/12/14 收款日期條件取消(另加條件:全年服務費超過 20,000 或 服務費 未超過 20,000 但有扣繳)
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      pub_QL05 = pub_QL05 & ";收款期間:" & MaskEdBox1 'Add By Sindy 2020/12/30
      st0l0Con = st0l0Con & " and a0l02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      st0y0Con = st0y0Con & " and a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & "" 'Add By Sindy 2015/11/20
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      pub_QL05 = pub_QL05 & ";收款期間:" & MaskEdBox2 'Add By Sindy 2020/12/30
      st0l0Con = st0l0Con & " and a0l02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      st0y0Con = st0y0Con & " and a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & "" 'Add By Sindy 2015/11/20
   End If
   
   'Add By Sindy 2020/12/30
   If Me.Check1.Value = 1 Then pub_QL05 = pub_QL05 & ";" & Check1.Caption
   If Me.Check2.Value = 1 Then pub_QL05 = pub_QL05 & ";" & Check2.Caption
   If Me.Check3.Value = 1 Then pub_QL05 = pub_QL05 & ";" & Check3.Caption
   If Me.Check4.Value = 1 Then pub_QL05 = pub_QL05 & ";" & Check4.Caption
   If Me.Check5.Value = 1 Then pub_QL05 = pub_QL05 & ";" & Check5.Caption
   If Me.Check6.Value = 1 Then pub_QL05 = pub_QL05 & ";" & Check6.Caption
   If Me.Check7.Value = 1 Then pub_QL05 = pub_QL05 & ";" & Check7.Caption
   '2020/12/30 END
   
   'Modify by Morgan 2007/9/28 改分列印別
   ''抓合併客戶編號(該年度收據日期最大的客戶編號),收據編號
   'stVTB = "select CuNo,a0k01" & _
      " from (select a0k16 Yr,a0k04 CuName,substrb(max((a0k02+1000000)||a0k03),8,9) CuNo" & _
      " From acc0k0 where 1=1" & stCon & _
      " group by a0k16,a0k04) V0, acc0k0 where a0k16(+)=Yr and a0k04(+)=CuName" & stCon2 & st0k0Con
   '單一
   If txtType = "1" Then
      'Modify By Sindy 2013/12/2 +a0k04
      stVTB = "SELECT a0k01,CuNo,a0k04 FROM (SELECT a0k16 Yr,a0k04 x1" & _
         " ,SUBSTR(MAX((a0k02+1000000)||a0k03),8,9) CuNo" & _
         " FROM acc0k0 WHERE 1=1" & stCon & " GROUP BY a0k16,a0k04" & _
         " HAVING COUNT(DISTINCT SUBSTR(a0k03,1,6))=1" & _
         ") X,(SELECT DISTINCT A0K04 z1 FROM (SELECT a0k03 y1" & _
         " FROM acc0k0 WHERE 1=1" & stCon & " GROUP BY a0k03" & _
         " HAVING COUNT(DISTINCT a0k04)>1" & _
         " ) Y, ACC0K0 WHERE A0K03(+)=y1" & stCon & ") Z,acc0k0" & _
         " WHERE z1(+)=x1 AND z1 IS NULL AND a0k16(+)=Yr AND a0k04(+)=x1" & stCon2 & st0k0Con
      'Add By Sindy 2015/11/20 執行單一時, 才需要加入Acc1k0國外請款資料
      stVTB1K = "SELECT a1k01,a1k28 CuNo,a1k35" & _
                " From acc1k0,acc1v0,caseprogress" & _
                " where 1=1" & st1k0Con & _
                " and a1k01=a1v02(+) and a1v01=cp09(+)"
      '2015/11/20 END
   '複合
   Else
      'Modify By Sindy 2013/12/2 +a0k04
      stVTB = "SELECT a0k01,CuNo,a0k04 FROM (" & _
         " SELECT Yr,X1,SUBSTR(MAX((a0k02+1000000)||a0k03),8,9) CuNo" & _
         " FROM (SELECT a0k16 Yr,a0k04 X1 FROM acc0k0 WHERE 1=1" & stCon & _
         " GROUP BY a0k16,a0k04 HAVING COUNT(DISTINCT SUBSTR(a0k03,1,6))>1" & _
         " UNION SELECT a0k16 Yr, A0K04 X1 FROM (SELECT a0k03 Y1" & _
         " FROM acc0k0 WHERE 1=1" & stCon & " GROUP BY a0k03 HAVING COUNT(DISTINCT a0k04)>1" & _
         " ) Y, ACC0K0 WHERE A0K03(+)=Y1" & stCon & ") W,acc0k0" & _
         " WHERE a0k16(+)=Yr AND a0k04(+)=X1" & stCon & " GROUP BY Yr,X1,SUBSTR(a0k03,1,6)) Z,acc0k0" & _
         " WHERE a0k16(+)=Yr AND a0k04(+)=X1 and substr(a0k03,1,6)=substr(CuNo,1,6)" & stCon2 & st0k0Con
   End If
   'end 2007/9/28
   
   'Add By Sindy 2022/12/21
   If txtType <> "2" Then '2.複合 是列印紙本,無接續的問題
   '2022/12/21 END
      'Add By Sindy 2015/12/24 檢查是否有資料做一半的,若有,詢問要繼續還是清空重來
      strExc(0) = "select T06,Count(*) from ACCTMP44q0" & _
                  " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                  " and T06='X'" & _
                  " group by T06"
      intI = 1
      Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(0) = "select T06,Count(*) from ACCTMP44q0" & _
                     " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                     " and T06<>'X'" & _
                     " group by T06"
         intI = 1
         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            intQ = MsgBox("此作業尚有未處理的資料，確定要繼續執行嗎？" & vbCrLf & vbCrLf & _
                          "（【注意】：若要繼續執行，畫面上的輸入條件必須和前次相同，以免抓取資料不一致！）", vbExclamation + vbYesNoCancel, "重要訊息！")
            If intQ = vbYes Then '繼續
               BolContinueReSend = True 'Add By Sindy 2016/12/15
               GoTo GoToRun
            ElseIf intQ = vbCancel Then '取消
               adoquery.Close
               Set adoquery = Nothing
               Exit Sub
            End If
         End If
      End If
   End If
   '2015/12/24 END
   
   'Added by Morgan 2011/11/15 寫暫存
   cnnConnection.Execute "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'", intI
   cnnConnection.Execute "delete from ACCTMP44q0_1 where FName='" & Me.Name & "' and UserID='" & strUserNum & "'", intI 'Add By Sindy 2015/12/24
   'Modify By Sindy 2013/12/2 +a0k04
   'T01 : 收據號碼
   'T15 : 收據抬頭
   'Modify By Sindy 2014/10/21 +T16.E-Mail T17.FAX1 T18.FAX2
   'Modify By Sindy 2015/1/15 要以收據抬頭為基準,抓傳真等資料,改在下面程式再處理
   'Modify By Sindy 2015/11/20 + 執行單一時, 才需要加入Acc1k0國外請款資料
   'cnnConnection.Execute "insert into ACCTMP44q0(T01,T02,T05,T14,T15) SELECT distinct A0k01 T01,CuNo T02,'" & Me.Name & "' T05,'" & strUserNum & "' T14,a0k04 T15 from (" & stVTB & ")", intI
   cnnConnection.Execute "insert into ACCTMP44q0(T01,T02,T05,T14,T15) SELECT distinct A0k01 T01,CuNo T02,'" & Me.Name & "' T05,'" & strUserNum & "' T14,a0k04 T15 from (" & stVTB & IIf(txtType = "1", " union " & stVTB1K, "") & ")", intI
   '2015/11/20 END
   'decode(cu16,null,cu17,cu16||';'||cu17) T17,decode(cu18,null,cu19,cu18||';'||cu19) T18
'   cnnConnection.Execute "insert into ACCTMP44q0(T01,T02,T05,T14,T15,T16,T17,T18) " & _
'                         "SELECT T01,T02,T05,T14,T15,decode(cu115,null,cu20,cu115) T16,cu16 T17,cu18 T18 FROM " & _
'                         " (SELECT distinct A0k01 T01,CuNo T02,'" & Me.Name & "' T05,'" & strUserNum & "' T14,a0k04 T15 from (" & stVTB & ")) X" & _
'                         ",customer" & _
'                         " WHERE substr(X.T02,1,8)=cu01(+) AND substr(X.T02,9,1)=cu02(+)" _
'                         , intI
   'Add By Sindy 2015/11/24
   strExc(0) = "select T01 from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "' and rownum=1"
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2020/12/30
      MsgBox "無資料可列印！"
      adoquery.Close
      Set adoquery = Nothing
      Exit Sub
   End If
   '2015/11/24 END
   
   'Add By Sindy 2019/3/22 排除沒有扣繳明細檔的資料
   strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
            " and not exists (select * from acc1v0 where a1v02=T01)"
   cnnConnection.Execute strSql, intI
   
   'Add By Sindy 2013/12/2
   If Check3.Value = 1 Or Check4.Value = 1 Or Option1(0).Value = True Or Option1(1).Value = True Then
      strExc(0) = "select t15 from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "' group by t15"
      intI = 1
      Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With adoquery
            .MoveFirst
            Do While Not .EOF
               If Check3.Value = 1 Then '目前無應收帳款
                  'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
                  'Modified by Lydia 2025/06/10 (a0k32 is null) 改用函數判斷：geta0k32type(a0k01)='1'
                  strExc(0) = "SELECT A0K01,cp09,cp79,a0k06,a0k07,a0k17,a0k18,a0s05" & _
                              " from ACCTMP44q0,acc0k0,acc0j0," & _
                              "(select a0s02, sum(nvl(a0s05,0)+nvl(a0s06, 0)+nvl(a0s07, 0)) as a0s05" & _
                                " From acc0s0" & _
                              " group by a0s02) acc0s0,caseprogress" & _
                              " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                              " and T15='" & ChgSQL(.Fields("T15")) & "'" & _
                              " and T01=a0k01(+)" & _
                              " and a0j13(+)=a0k01" & _
                              " and cp09(+)=a0j01" & _
                              " and a0k01 = a0s02(+)" & _
                              " and (a0k09 is null or a0k09 = 0)" & _
                              " and (a0k06+a0k07) > decode(a0s05, null, 0, a0s05)" & _
                              " and (a0k06+a0k07-decode(a0s05, null, 0, a0s05)) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & _
                              " and nvl(cp79, 0) > 0" & _
                              " and geta0k32type(a0k01)='1'"
                  'Add By Sindy 2015/11/20
                  'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
                  strExc(0) = strExc(0) & " union " & _
                              "SELECT A1K01,cp09,cp79,0 a0k06,0 a0k07,0 a0k17,0 a0k18,0 a0s05" & _
                              " From ACCTMP44q0, acc1k0, caseprogress" & _
                              " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                              " and T15='" & ChgSQL(.Fields("T15")) & "'" & _
                              " and T01=a1k01(+)" & _
                              " and cp60(+)=a1k01" & _
                              " and (a1k12 is null or a1k12 = 0)" & _
                              " and a1k29 is null" & _
                              " and nvl(cp79, 0) > 0"
                  '2015/11/20 END
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then '有應收帳
                     'Modify by Amy 2023/06/02 bug 收據抬頭 有單引號會錯
                     strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                              " and T15='" & ChgSQL(.Fields("T15")) & "'"
                     cnnConnection.Execute strSql, intI
                  End If
               ElseIf Check4.Value = 1 Then '去年無扣單但今年有扣稅
                  '去年有扣單
                  'Modify by Amy 2023/06/02 bug 收據抬頭 有單引號會錯
                  strExc(0) = "select a1v01,a1v02,a1v15" & _
                              " From acc0k0, acc0j0, acc1v0" & _
                              " where a0k04='" & ChgSQL(.Fields("T15")) & "' and a0k16=" & Val(Text5) - 1 & _
                              " and a0j13(+)=a0k01" & _
                              " and a1v01(+)=a0j01 and a1v02(+)=a0j13" & _
                              " and a1v15 is not null"
                  'Add By Sindy 2015/11/20
                  'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
                  strExc(0) = strExc(0) & " union " & _
                              "select a1v01,a1v02,a1v15" & _
                              " From acc1k0, caseprogress, acc1v0" & _
                              " where a1k35='" & ChgSQL(.Fields("T15")) & "'" & _
                              " and cp60(+)=a1k01" & _
                              " and a1v01(+)=cp09 and a1v02(+)=cp60" & _
                              " and a1v15 is not null" & _
                              " and a1v09=" & Val(Text5) - 1
                  '2015/11/20 END
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then '有扣單
                     'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
                     strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                              " and T15='" & ChgSQL(.Fields("T15")) & "'"
                     cnnConnection.Execute strSql, intI
                  Else '無扣單
                     '今年有扣稅
                     'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
                     strExc(0) = "select t01 from ACCTMP44q0,ACC1V0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                                 " and T15='" & ChgSQL(.Fields("T15")) & "'" & _
                                 " and T01=a1v02(+) and nvl(a1v06,0)>0"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 0 Then '不符合
                        'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
                        strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                                 " and T15='" & ChgSQL(.Fields("T15")) & "'"
                        cnnConnection.Execute strSql, intI
                     End If
                  End If
'               ElseIf Check5.Value = 1 Then '稅額達 2,000 但沒扣繳
'                  strExc(0) = "select t01 from ACCTMP44q0,ACC1V0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
'                              " and T15='" & .Fields("T15") & "'" & _
'                              " and T01=a1v02(+) and nvl(a1v04,0)>2000 and nvl(a1v06,0)=0"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 0 Then '不符合
'                     strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
'                              " and T15='" & .Fields("T15") & "'"
'                     cnnConnection.Execute strSql, intI
'                  End If
               'Add By Sindy 2013/12/20
               ElseIf Option1(0).Value = True Then '只列印今年有扣繳
                  'Modify by Amy 2023/06/02 bug 收據抬頭 有單引號會錯
                  '今年有扣稅
                  strExc(0) = "select t01 from ACCTMP44q0,ACC1V0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                              " and T15='" & ChgSQL(.Fields("T15")) & "'" & _
                              " and T01=a1v02(+) and nvl(a1v06,0)>0"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then '不符合
                     strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                              " and T15='" & ChgSQL(.Fields("T15")) & "'"
                     cnnConnection.Execute strSql, intI
                  End If
                  'end 2023/06/02
               ElseIf Option1(1).Value = True Then '只列印有扣繳但不含已確認過的
                  'Modify by Amy 2023/06/02 bug 收據抬頭 有單引號會錯
                  '今年有扣稅
                  strExc(0) = "select t01 from ACCTMP44q0,ACC1V0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                              " and T15='" & ChgSQL(.Fields("T15")) & "'" & _
                              " and T01=a1v02(+) and nvl(a1v06,0)>0"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then '不符合
                     strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                              " and T15='" & ChgSQL(.Fields("T15")) & "'"
                     cnnConnection.Execute strSql, intI
                  Else
                     '有確認過的
                     strExc(0) = "select A2801 from ACC280 where A2801='" & ChgSQL(.Fields("T15")) & "' and A2802=" & Text5
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then '不含
                        strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                                 " and T15='" & ChgSQL(.Fields("T15")) & "'"
                        cnnConnection.Execute strSql, intI
                     End If
                  End If
                  'end 2023/06/02
               '2013/12/20 END
               End If
               .MoveNext
            Loop
         End With
      End If
      adoquery.Close
   End If
'   strExc(0) = "select count(*) from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If RsTemp.Fields(0) = 0 Then
'         MsgBox "無資料可列印！"
'         GoTo ReadExit
'      End If
''   Else
''      MsgBox "無資料可列印！"
''      GoTo ReadExit
'   End If
   '2013/12/2 END
   
   
   '************************************************************************************
   'Add By Sindy 2015/1/15 以收據抬頭抓e-mail及傳真等資料,以鼎元光電和銀泰測試
   'T06.strType（E:Email F:Fax P:紙本 A:尚無處理到的資料或Run Excel結束後註記為A）
   'T07.客戶編號
   'T08.執行日期
   'T09.執行時間
   'T16.E-Mail
   'T17.Tel
   'T18.Fax
   'T19.貴公司
   'T20.客戶地址若客戶狀態有資料時優先抓
   'T21.會計備註
   'T22.業務區
   'T23.智權人員
   'T24.是否已讀取到客戶資料 Y.已讀取
   'T25.是否為境外公司
   'T26.原A0K03客戶編號 (105.6.4)
   'T27.收款單號 (105.11.3)
   'T28.會計傳真 (105.11.18)
   'T29.收據抬頭讀到的客戶編號 (108.3.25)
   '************************************************************************************
   'Add By Sindy 2016/1/7
   '先以收據抬頭抓資料
   strExc(0) = "select T15 from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T24 is null group by T15 order by T15 asc"
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      adoquery.MoveFirst
      Do While Not adoquery.EOF
         'Modify By Sindy 2016/6/4 + m_CU01, m_CU02
         'Modify By Sindy 2016/11/17 + m_CU168
         'Modify By Sindy 2017/3/16 + m_CU172
         If GetTitleCustData(adoquery.Fields("T15"), "", "", m_CU01, m_CU02, _
                            m_CU15, m_CU115, m_CU20, "", "", "", m_CU16, _
                            m_CU18, "", m_CU158, m_CU22, m_CU80, m_CU30, m_CU31, _
                            m_CU159, m_CU12, m_CU13, , m_CU168, , , , , , m_CU172) = True Then
            'Add By Sindy 2016/11/1
            strA49 = GetAcc49(adoquery.Fields("T15"), m_CU01 & m_CU02, m_A4904, m_A4905)
            If m_CU159 = "" Then
               m_CU159 = strA49
            Else
               m_CU159 = m_CU159 & IIf(strA49 <> "", vbCrLf & strA49, "")
            End If
            '2016/11/1 END
            
            'Add By Sindy 2016/11/9 瑞婷:有會計師E-Mail時,財務信箱要加會計師信箱寄發
            If m_A4905 <> "" Then
               If m_CU115 = "" Then
                  m_CU115 = m_A4905
               Else
                  m_CU115 = m_CU115 & ";" & m_A4905
               End If
            End If
            '2016/11/9 END
            
'            'Add By Sindy 2016/11/17 瑞婷:有會計師傳真時,傳真要加會計師傳真寄發
'            If m_A4904 <> "" Then
'               If m_CU18 = "" Then
'                  m_CU18 = m_A4904
'               Else
'                  m_CU18 = m_CU18 & ";" & m_A4904
'               End If
'            End If
'            '2016/11/17 END
            
            'Modify By Sindy 2016/11/15 執行Excel鍵時,T24增加註記若為每月提醒代填繳款書時上N,後面踢除T24=N的資料
            'Modify By Sindy 2017/3/16 T24增加註記若為不寄發扣繳核對資料時上N,後面踢除T24=N的資料
            'Modify By Sindy 2019/3/25 + T29記錄收據抬頭讀到的客戶編號
            'Modify By Sindy 2025/3/3 IIf(m_CU168 = "Y", "N", "Y") => IIf(m_CU168 <> "", "N", "Y")
            strSql = "update ACCTMP44q0" & _
                     " set T16=" & CNULL(IIf(m_CU115 = "", m_CU20, m_CU115)) & _
                     ",T17=" & CNULL(IIf(m_CU16 = "", m_CU22, m_CU16)) & _
                     ",T18=" & CNULL(m_CU18) & _
                     ",T19=" & CNULL(m_CU15) & _
                     ",T20=" & CNULL(IIf(m_CU80 = "", m_CU30 & m_CU31, m_CU80)) & _
                     ",T21=" & CNULL(ChgSQL(m_CU159)) & _
                     ",T22=" & CNULL(m_CU12) & _
                     ",T23=" & CNULL(m_CU13) & _
                     ",T24='" & IIf(m_CU172 = "N", "N", IIf(bolExcel = True, IIf(m_CU168 <> "", "N", "Y"), "Y")) & "'" & _
                     ",T25=" & CNULL(m_CU158) & _
                     ",T28=" & CNULL(m_A4904) & _
                     ",T29=" & CNULL(m_CU01 & m_CU02)
            'Modify By Sindy 2016/6/4
            If txtType = "2" And m_CU01 & m_CU02 <> "" Then '複合
               strSql = strSql & ",T26=decode(T02,'" & m_CU01 & m_CU02 & "',null,T02),T02=decode(T02,'" & m_CU01 & m_CU02 & "',T02,'" & m_CU01 & m_CU02 & "')"
            End If
            '2016/6/4 END
            'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
            strSql = strSql & " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                     " and T15='" & ChgSQL(adoquery.Fields("T15")) & "'"
            cnnConnection.Execute strSql, intI
         End If
         adoquery.MoveNext
      Loop
   End If
   adoquery.Close
   '先以收據抬頭+客戶編號抓資料
   strExc(0) = "select T15,T02 from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T24 is null order by T15,T02 asc"
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      adoquery.MoveFirst
      Do While Not adoquery.EOF
         'Modify By Sindy 2016/6/4 + m_CU01, m_CU02
         'Modify By Sindy 2017/3/16 + m_CU172
         If GetTitleCustData(adoquery.Fields("T15"), adoquery.Fields("T02"), "", m_CU01, m_CU02, _
                            m_CU15, m_CU115, m_CU20, "", "", "", m_CU16, _
                            m_CU18, "", m_CU158, m_CU22, m_CU80, m_CU30, m_CU31, _
                            m_CU159, m_CU12, m_CU13, False, m_CU168, , , , , , m_CU172) = True Then
            'Add By Sindy 2016/11/1
            strA49 = GetAcc49(adoquery.Fields("T15"), m_CU01 & m_CU02, m_A4904, m_A4905)
            If m_CU159 = "" Then
               m_CU159 = strA49
            Else
               m_CU159 = m_CU159 & IIf(strA49 <> "", vbCrLf & strA49, "")
            End If
            '2016/11/1 END
            
            'Add By Sindy 2016/11/9 瑞婷:有會計師E-Mail時,財務信箱要加會計師信箱寄發
            If m_A4905 <> "" Then
               If m_CU115 = "" Then
                  m_CU115 = m_A4905
               Else
                  m_CU115 = m_CU115 & ";" & m_A4905
               End If
            End If
            '2016/11/9 END
            
'            'Add By Sindy 2016/11/17 瑞婷:有會計師傳真時,傳真要加會計師傳真寄發
'            If m_A4904 <> "" Then
'               If m_CU18 = "" Then
'                  m_CU18 = m_A4904
'               Else
'                  m_CU18 = m_CU18 & ";" & m_A4904
'               End If
'            End If
'            '2016/11/17 END
            
            'Modify By Sindy 2016/11/15 執行Excel鍵時,T24增加註記若為每月提醒代填繳款書時上N,後面踢除T24=N的資料
            'Modify By Sindy 2017/3/16 T24增加註記若為不寄發扣繳核對資料時上N,後面踢除T24=N的資料
            'Modify By Sindy 2019/3/25 + T29記錄收據抬頭讀到的客戶編號
            strSql = "update ACCTMP44q0" & _
                     " set T16=" & CNULL(IIf(m_CU115 = "", m_CU20, m_CU115)) & _
                     ",T17=" & CNULL(IIf(m_CU16 = "", m_CU22, m_CU16)) & _
                     ",T18=" & CNULL(m_CU18) & _
                     ",T19=" & CNULL(m_CU15) & _
                     ",T20=" & CNULL(IIf(m_CU80 = "", m_CU30 & m_CU31, m_CU80)) & _
                     ",T21=" & CNULL(ChgSQL(m_CU159)) & _
                     ",T22=" & CNULL(m_CU12) & _
                     ",T23=" & CNULL(m_CU13) & _
                     ",T24='" & IIf(m_CU172 = "N", "N", IIf(bolExcel = True, IIf(m_CU168 = "Y", "N", "Y"), "Y")) & "'" & _
                     ",T25=" & CNULL(m_CU158) & _
                     ",T28=" & CNULL(m_A4904) & _
                     ",T29=" & CNULL(m_CU01 & m_CU02)
            'Modify By Sindy 2016/6/4
            If txtType = "2" And m_CU01 & m_CU02 <> "" Then '複合
               If adoquery.Fields("T02") <> m_CU01 & m_CU02 Then
                  strSql = strSql & ",T26=T02,T02='" & m_CU01 & m_CU02 & "'"
               End If
            End If
            '2016/6/4 END
            'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
            strSql = strSql & " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                     " and T15='" & ChgSQL(adoquery.Fields("T15")) & "'" & _
                     " and T02='" & adoquery.Fields("T02") & "'"
            cnnConnection.Execute strSql, intI
         End If
         adoquery.MoveNext
      Loop
   End If
   adoquery.Close
   '2016/1/7 END
   
   'Add By Sindy 2016/11/15 執行Excel鍵時,踢除每月提醒代填繳款書(T24=N)
   'Modify By Sindy 2017/3/16 T24增加註記若為不寄發扣繳核對資料時上N,後面踢除T24=N的資料
   'If bolExcel = True Then
   '2017/3/16 END
      strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
               " and T24='N'"
      cnnConnection.Execute strSql, intI
   'End If
   '2016/11/15 END
   'Add By Sindy 2015/12/9 踢除個人收據且未扣繳的資料
   strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
            " and T01 in(" & _
            "select T01 from ACCTMP44q0,acc1v0,acc0k0" & _
            " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and substr(T01,1,1)='E'" & _
            " and T01=a0k01(+) and a0k05='1' and a0k01=a1v02(+) and a1v06=0 and a1v02 is not null" & _
            ")"
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2015/12/9 踢除境外公司且未扣繳的資料
   strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
            " and T01 in(" & _
            "select T01 from ACCTMP44q0,acc1v0" & _
            " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T25='Y'" & _
            " and T01=a1v02(+) and a1v06=0 and a1v02 is not null" & _
            ")"
   cnnConnection.Execute strSql, intI
   '2015/12/9 END
   'Add By Sindy 2016/2/19 只要部份收款(and a1v05='Y')都沒有扣繳的問題
   '                       扣繳只會全扣不然就全不扣, 因為是部份收款所以全部不扣
   'Modify By Sindy 2019/12/30 E10802240:部份收款=Y
   strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
            " and T01 in(" & _
            "select T01 from ACCTMP44q0,acc1v0" & _
            " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
            " and T01=a1v02(+) and a1v05='Y'" & _
            ")"
   cnnConnection.Execute strSql, intI
   '2016/2/19 END
   
   'Add By Sindy 2016/11/3 更新收款單號
   strSql = "UPDATE ACCTMP44q0 set T27=(select min(a0m01) from acc0m0 where a0m02=T01)" & _
            "where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   '2016/11/3 END
   
   
GoToRun:
   '注意1.收款金額含銷退 2.服務費=10*(已扣+未扣) -->參考44t0扣繳憑單查詢及列印
   'Fee0:收款金額,Fee1:服務費,Fee2:可扣稅額,Fee3:收款扣繳額,Fee4:補扣繳額,Fee5:未扣繳額,Fee6:已扣繳額,Fee7:調整稅額,Fee8:收據未扣繳額
   'Modified by Morgan 2011/11/10 考慮拆收據情形
   'Modified by Morgan 2011/11/15 +ACCTMP44q0
   'Modified by Morgan 2011/12/21 +a0k33,a0j22,a0j25
   'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
   'Modify By Sindy 2014/10/21 +,T16,T17,T18
'********************
' 列印格式
'********************
   'Modify By Sindy 2015/12/11 + T15
   'Modify By Sindy 2015/12/24 + and T06='X'
   'Modify By Sindy 2016/6/4 + ,T26
   'Modify By Sindy 2016/6/4 CuNo ==> decode(T26,null,CuNo,T02) CuNo
   'Modify By Sindy 2016/11/3 + ,T27
   'Modify By Sindy 2019/3/25 + ,T29
   strMainSql = "select decode(T26,null,CuNo,T02) CuNo,a0k04,a0k11,a0l02,a0k01,a0k02" & _
                ",substrb(getcp10desc(cp01,cp10,a0j04),1,12) cp10N,substrb(na03,1,8) na03" & _
                ",Fee0,10*nvl(a1v04,0) Fee1,nvl(a1v04,0) Fee2,decode(a1v18,'1',nvl(a1v06,0),0) Fee3" & _
                ",decode(a1v18,'1',0,nvl(a1v06,0)) Fee4,nvl(a1v07,0) Fee5, nvl(decode(a1v15,null,0,a1v06),0) Fee6" & _
                ",nvl(a1v10,0) Fee7,a1p12,Fee8,T03,a0k33,a0j22,a0j25,cp10,T16,T17,T18,T19,T20,T21,T22,T23,T15,T26,T27,T28,T29" & _
                " from ( select CuNo,a0m02,max(a0l02) a0l02,min(a1p12) a1p12 from (" & stVTB & ") V1,acc0m0, acc0l0,acc1p0" & _
                " where a0m02=a0k01 and (a0m03 is null or substr(a0m03, 1, 1) = 'E')" & _
                " and a0l01=a0m01" & st0l0Con & " and a1p04(+)=a0l01 and a1p05(+)='113001'" & _
                " group by CuNo,a0m02 ) X,acc0k0, acc0j0,caseprogress,acc1v0" & _
                ",( select a1u03,sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)) Fee0" & _
                ", sum(nvl(a1u04, 0)-nvl(a1u08, 0)) as Fee1,a1u02" & _
                " From (" & stVTB & ") V1, acc1u0 where a1u02(+)=a0k01 group by a1u03,a1u02 ) Y" & _
                ",( select a1v02 Z1,nvl(sum(a1v07),0) Fee8" & _
                " From (" & stVTB & ") V1, acc1v0 where a1v02(+)=a0k01 group by a1v02 ) Z,ACCTMP44q0,nation" & _
                " Where a0k01(+)= a0m02 and a0j13(+)=a0k01 and cp09(+)=a0j01 and a1u03(+)=a0j01 and a1u02(+)=a0j13" & _
                " and a1v01(+)=a0j01 and a1v02(+)=a0j13 and a1v04>0 and Z1(+)=a0j13 and T01=a0k01 and T05='" & Me.Name & "' and T14='" & strUserNum & "' and T06='X' and na01(+)=a0j04"
   'Add By Sindy 2015/11/20 + 執行單一時, 才需要加入Acc1k0國外請款資料
   If txtType = "1" Then '單一
      'Modify By Sindy 2016/6/4 CuNo ==> decode(T26,null,CuNo,T02) CuNo
      strMainSql = strMainSql & " union " & _
               "select decode(T26,null,CuNo,T02) CuNo,a1k35 a0k04,a1v03 a0k11,a0y02 a0l02,a1k01 a0k01,a1k02,substrb(GETCP10DESCCaseNO(cp10,cp01,cp02,cp03,cp04),1,12) cp10N" & _
               ",substrb(GETNA03DESCCaseNO(cp01,cp02,cp03,cp04),1,8) na03,nvl(a1k30,0) Fee0,Fee1,Fee2,Fee3,Fee4,Fee5,Fee6,Fee7,a1p12,Fee8" & _
               ",T03,'' a0k33,'' a0j22,1 a0j25,cp10,T16,T17,T18,T19,T20,T21,T22,T23,T15,T26,T27,T28,T29" & _
               " from (select CuNo,a0z02,max(a0y02) a0y02,min(a1p12) a1p12 from (" & stVTB1K & ") V1,acc0z0,acc0y0,acc1p0" & _
               " where a0z02=a1k01 and a0y01=a0z01" & st0y0Con & " and a1p04(+)=a0y01 and a1p05(+)='113001' group by CuNo,a0z02) X" & _
               ",(select a1v02 Z1,sum(10*nvl(a1v04,0)) Fee1,nvl(sum(a1v04),0) Fee2,sum(decode(a1v18,'1',nvl(a1v06,0),0)) Fee3" & _
               ",sum(decode(a1v18,'1',0,nvl(a1v06,0))) Fee4,nvl(sum(a1v07),0) Fee5,sum(nvl(decode(a1v15,null,0,a1v06),0)) Fee6" & _
               ",nvl(sum(a1v10),0) Fee7,nvl(sum(a1v07),0) Fee8 From (" & stVTB1K & ") V1,acc1v0 where a1v02(+)=a1k01 group by a1v02) Z" & _
               ",ACCTMP44q0,acc1k0,acc1v0,caseprogress" & _
               " Where a1k01(+)=a0z02" & _
               " and a1k01=a1v02(+)" & _
               " and a1v01=cp09(+)" & _
               " and a1v04>0" & _
               " and Z1(+)=a1k01" & _
               " and T01=a1k01" & _
               " and T05='" & Me.Name & "'" & _
               " and T14='" & strUserNum & "' and T06='X'"
   End If
   '2015/11/20 END
   If BolContinueReSend = True Then GoTo ContinueReSend 'Add By Sindy 2016/12/15
   
'   'Add By Sindy 2015/12/23 暫時先踢除國外請款單資料
'   strSql = "delete from ACCTMP44q0 where substr(t01,1,1)='X'"
'   cnnConnection.Execute strSql, intI
'   '2015/12/23 END
   
   'Add By Sindy 2014/12/23 有1,2項過濾條件時,依收據抬頭為基準過濾資料
   '                        要整個抬頭收據資料都不符合條件,才可不顯示其抬頭收據資料
   'Modify By Sindy 2016/2/19 改為有勾選第2項過濾條件時,才要依收據抬頭為基準過濾資料,要整個抬頭收據資料都不符合條件,才可不顯示其抬頭收據資料
   '                          ex.祖嘉企業有限公司:E10428255, E10427456
   'Modify By Sindy 2016/2/24 + Or Check6.Value = 1
   If Check1.Value = 1 Or Check6.Value = 1 Or Check2.Value = 1 Then
      'Modify By Sindy 2016/3/29 目的是要 group by a0k01(收據編號) ex:E10409426 活全機器股份有限公司
      strExc(0) = "select T15,CuNo,A0K04,a0k01,sum(Fee0) Fee0,sum(Fee1) Fee1,sum(Fee5) Fee5,sum(Fee6) Fee6,T16,T27 from(" & strMainSql & ") " & _
                  " group by T15,A0K04,a0k01,T16,T27,CuNo"
      '2016/3/29 END
      'Add By Sindy 2016/11/3
      If Check6.Value = 1 Then '欲判斷單筆 2,001
         strExc(0) = strExc(0) & " order by T15 asc,a0k04 asc,CuNo asc,Fee5 desc"
      Else
         strExc(0) = strExc(0) & " order by T15,a0k04,CuNo"
      End If
      '2016/11/3 END
      intI = 1
      Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         adoquery.MoveFirst
         m_A0K04 = "": dblTot = 0: bolPrintCheck = False
         m_A0m01 = "" 'Add By Sindy 2016/11/3
         Do While Not adoquery.EOF
'            If InStr(adoquery.Fields("A0K04"), "金坤") > 0 Then
'               MsgBox adoquery.Fields("A0K04")
'            End If
            If m_A0K04 <> "" And adoquery.Fields("A0K04") <> m_A0K04 Then
               'Add By Sindy 2016/2/19
               If Check2.Value = 1 Then
               '2016/2/19 END
                  '檢查要不要刪除此抬頭
                  If bolPrintCheck = False Then
                     'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
                     strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                              " and T15='" & ChgSQL(m_A0K04) & "'"
                     cnnConnection.Execute strSql, intI
                  End If
                  'm_A0K04 = "": dblTot = 0: bolPrintCheck = False 'Modify By Sindy 2020/12/29 Mark
               End If
               m_A0K04 = "": dblTot = 0: bolPrintCheck = False 'Modify By Sindy 2020/12/29 換收據抬頭時,改回預設值
            End If
            
            'Modify By Sindy 2016/2/24 + Or Check6.Value = 1
            If Check1.Value = 1 Or Check6.Value = 1 Then '只印有應稅未扣的 (扣繳 2001 以上)
               'Add By Sindy 2016/11/3
               If m_A0m01 <> "" And "" & adoquery.Fields("T27") <> m_A0m01 Then
                  m_A0m01 = ""
                  bolPrintCheck = False
               End If
               '2016/11/3 END
               If bolPrintCheck = False Then
                  'Add By Sindy 2016/2/24 + if
                  If Check1.Value = 1 Then '稅額達 2,001 以上但未扣繳 (一收款單號合計達 2,001 含同收款單之所有收據)
                  '2016/2/24 END
                     'Add By Sindy 2016/1/27 單一收款單號合計達 2,001
                     'modify by sonia 2016/2/16  E10428147因グ分次收款會有錯,故改max(a0m01)
                     'strExc(0) = "select nvl(sum(a1v04),0) from acc0m0,acc1v0 where a0m01=(select a0m01 from acc0m0 where a0m02='" & adoquery.Fields("a0k01") & "') and a0m06=0" & _
                                 " and a0m02=a1v02"
                     'Modify By Sindy 2016/2/17 判斷 a0m06=0 只代表收款當時已扣繳稅額
                     '                          但有可能事後做補扣繳, 因此要檢查未扣繳需增加判斷 a1v07>0 : 有未扣繳額
                     strExc(0) = "select nvl(sum(a1v04),0),a0m01 from acc0m0,acc1v0,acc0k0 where a0m01=(select max(a0m01) from acc0m0,acc0k0 where a0m02='" & adoquery.Fields("a0k01") & "' and a0m06=0 and a0m02=a0k01 and a0k04='" & adoquery.Fields("a0k04") & "')" & _
                                 " and a0m02=a1v02 and a1v07>0 and a0m02=a0k01 and a0k04='" & adoquery.Fields("a0k04") & "' group by a0m01"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If RsTemp.Fields(0) >= 2001 Then
                           m_A0m01 = RsTemp.Fields("a0m01") '收款單號 Add By Sindy 2016/11/3
                           bolPrintCheck = True
                        End If
                     End If
                     '2016/1/27 END
                  Else '稅額達 2,001 以上但未扣繳 (單筆 2,001 含同收款單之所有收據)
                     If adoquery.Fields("Fee5") >= 2001 Then
                        'Add By Sindy 2016/11/3 抓取收款單號
                        m_A0m01 = "" & adoquery.Fields("T27") '收款單號
                        '2016/11/3 END
                        bolPrintCheck = True
                     End If
                  End If
               End If
               
               '不符合條件的資料就過濾掉
               If bolPrintCheck = False Then
                  strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                           " and T01='" & adoquery.Fields("a0k01") & "'"
                  cnnConnection.Execute strSql, intI
'               'Add By Sindy 2016/10/19 含同收款單之所有收據
'               Else
'                  If Check1.Value = 1 Then '一收款單號合計達 2,001 含同收款單之所有收據
'                     strExc(0) = "select a0m01,a0m02,a0k03,a0k04 from acc0m0,acc0k0 where a0m01=(select max(a0m01) from acc0m0,acc0k0" & _
'                                 " where a0m02='" & adoquery.Fields("a0k01") & "' and a0m02=a0k01" & IIf(Trim(cboTitle.Text) <> "", " and a0k04='" & adoquery.Fields("a0k04") & "'", "") & ")" & _
'                                 " and a0m02 not in (select T01 from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "')" & _
'                                 " and a0m02=a0k01(+)"
'                  Else '= Check6.Value = 1 : 單筆 2,001 含同收款單之所有收據
'                     strExc(0) = "select a0m01,a0m02,a0k03,a0k04 from acc0m0,acc0k0" & _
'                                 " where a0m02='" & adoquery.Fields("a0k01") & "' and a0m02=a0k01(+)" & _
'                                 IIf(Trim(cboTitle.Text) <> "", " and a0k04='" & adoquery.Fields("a0k04") & "'", "") & _
'                                 " and a0m02 not in (select T01 from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "')"
'                  End If
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     RsTemp.MoveFirst
'                     Do While Not RsTemp.EOF
'                        strSql = "insert into ACCTMP44q0(T01,T02,T05,T14,T15)" & _
'                                 " values('" & RsTemp.Fields("a0m02") & "','" & RsTemp.Fields("a0k03") & "','" & Me.Name & "','" & strUserNum & "','" & RsTemp.Fields("a0k04") & "')"
'                        cnnConnection.Execute strSql, intI
'                        RsTemp.MoveNext
'                     Loop
'                  End If
'               '2016/10/19 END
               End If
            ElseIf Check2.Value = 1 Then '只印全年服務費超過 20,000 或 服務費 未超過 20,000 但有扣繳的 或 當年有收款且有建信箱者
               dblTot = dblTot + Val("" & adoquery.Fields("Fee1"))
               If bolPrintCheck = False Then
                  If adoquery.Fields("Fee6") > 0 Then
                     bolPrintCheck = True
                  ElseIf dblTot > 20000 Then
                     bolPrintCheck = True
                  'Modify By Sindy 2015/2/17 +或 當年有收款且有建信箱者
                  ElseIf adoquery.Fields("Fee0") > 0 And "" & adoquery.Fields("T16") <> "" Then
                     bolPrintCheck = True
                  '2015/2/17 END
                  End If
               End If
            End If
            m_A0K04 = adoquery.Fields("A0K04")
            
            adoquery.MoveNext
         Loop
         'Add By Sindy 2016/2/19
         If Check2.Value = 1 Then
         '2016/2/19 END
            '檢查要不要刪除此抬頭
            If bolPrintCheck = False Then
               'Modify by Amy 2023/06/02 收據抬頭 有單引號會錯
               strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                        " and T15='" & ChgSQL(m_A0K04) & "'"
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
   End If
   '2014/12/23 END
   
   'Add By Sindy 2019/3/21 排除有扣單編號的資料
   'Modify By Sindy 2025/3/3 操作cmdMail按鍵,都要排除有扣單編號的資料
   'If Check2.Value = 1 And UCase(cmdMail.Tag) = UCase("EMail") Then
   If UCase(cmdMail.Tag) = UCase("EMail") Then
   '2025/3/3 END
      'Add By Sindy 2025/3/4 過濾 "無" 已扣繳額
      strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
               " and exists (select * from acc1v0 where a1v02=T01 and nvl(a1v06,0)=0)"
      cnnConnection.Execute strSql, intI
      '2025/3/4 END
      '過濾 "有" 扣單編號
      strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
               " and exists (select * from acc1v0 where a1v02=T01 and a1v15 is not null)"
      cnnConnection.Execute strSql, intI
      
      Call PrintA4_Excel_MailNO 'Add By Sindy 2025/9/5 信箱為NO的要另列清單
      'Add By Sindy 2025/3/3 因要發mail,所以僅留有建信箱者
      'Modify By Sindy 2025/9/5 增加刪除財務信箱是存NO者
      strSql = "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
               " and (T16 is null or Upper(T16)='NO')"
      cnnConnection.Execute strSql, intI
   End If
      
   '更新單次收款總應扣繳額(分次收款則取大的)
'   strSql = "update ACCTMP44q0 A set T03=(select max(X2) from (" & _
'      " select B.T01 X1,sum(nvl(u2.a1u04,0)+nvl(decode(a0j07,'Y',u2.a1u05),0)) X2,u1.a1u01" & _
'      " from ACCTMP44q0 B,acc1u0 u1,acc1u0 u2,acc0j0 where B.T05='" & Me.Name & "' and B.T14='" & strUserNum & "'" & _
'      " and u1.a1u02(+)=B.T01 and u2.a1u01(+)=u1.a1u01 and a0j01(+)=u2.a1u03 and a0j13(+)=u2.a1u02" & _
'      " group by B.T01,u1.a1u01) X where X1=A.T01)" & _
'      " where t05='" & Me.Name & "' and T14='" & strUserNum & "'"
   'Modify By Sindy 2013/12/6 ex.E10425001 分次收款,上列程式若是一收據多文號就會重覆計算
   strSql = "update ACCTMP44q0 A set T03=(select max(X2) from (" & _
      " select B.T01 X1,sum(nvl(u2.a1u04,0)+nvl(decode(a0j07,'Y',u2.a1u05),0)) X2,u2.a1u01" & _
      " from ACCTMP44q0 B,acc1u0 u2,acc0j0 where B.T05='" & Me.Name & "' and B.T14='" & strUserNum & "'" & _
      " and u2.a1u02(+)=B.T01 and a0j01(+)=u2.a1u03 and a0j13(+)=u2.a1u02" & _
      " group by B.T01,u2.a1u01) X where X1=A.T01)" & _
      " where t05='" & Me.Name & "' and T14='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI
   'end 2011/11/15
   
ContinueReSend:
   'Add By Sindy 2016/12/13 先判斷有無資料,後面再判斷較複雜
   strExc(0) = "select count(*) from(" & strMainSql & ")"
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If adoquery.Fields(0) = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2020/12/30
      MsgBox "無資料可列印！"
      adoquery.Close
      Set adoquery = Nothing
      Exit Sub
   Else
      InsertQueryLog (adoquery.Fields(0)) 'Add By Sindy 2020/12/30
   End If
   '2016/12/13 END
   
   txtNote.Top = 3720: txtNote.Visible = True 'Add By Sindy 2022/4/20
   If bolExcel = False Then
      bolMailSendErr = False 'Add By Sindy 2016/12/12 預設值
      '********************
      ' E-Mail
      '********************
      '有E-Mail時,Print A4紙張,寄Mail
      If txtType = "1" Then '單一:才要寄EMail
         strExc(0) = "select * from(" & strMainSql & ") where " & _
                     " T16 is not null" & _
                     " order by CuNo,a0k04,a0k11,a0l02, a0k01,a0j25,cp10"
         intI = 1
         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
         longRow_Mail = 0
         If intI = 1 Then
            longRow_Mail = adoquery.RecordCount
            bolPrintA4 = True
            'Modify by Amy 2024/05/17 財務2個特殊設定拆成3個 原:Pub_GetSpecMan("財務處總帳人員")
            strExc(0) = "select st02,ed01" & _
                        " from staff,ExtensionData" & _
                        " where ST01=ED02(+)" & _
                        " and st01='" & stTxtPerson & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strEmp = RsTemp.Fields("st02")
               strEMP_Tel = "" & RsTemp.Fields("ed01")
            End If
            
            '檢查是否有安裝PDFCreator
            PrinterIndex = -1
            For i = 0 To Printers.Count - 1
               If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
                  PrinterIndex = i
                  Exit For
               End If
            Next i
            If PrinterIndex < 0 Then
               MsgBox "請通知電腦中心安裝PDFCreator !!!"
               Exit Sub
            End If
            '*******************************
            PUB_SetOsDefaultPrinter Combo1 'Add By Sindy 2022/4/20
            PUB_RestorePrinter Printers(PrinterIndex).DeviceName
            '*******************************
            strTempFolder = App.path & "\" & "$$TempFolder"
            If Dir(strTempFolder, vbDirectory) = "" Then
               MkDir strTempFolder
            End If
            
'            Printer.Orientation = 1 '1.直印 2.橫印
'            Printer.PaperSize = 9
'            Call PrintA4(True, False, False)
            Call PrintA4_Excel(True, False, False)
         End If
         
         'Add By Sindy 2019/3/20
         '單純只是發提醒客戶摧扣繳憑單的EMail,發完後,後面的資料不用再Run
         If UCase(cmdMail.Tag) = UCase("EMail") Then
            GoTo RunAllEnd
         End If
         
         '********************
         ' 傳真
         '********************
         '無E-Mail有傳真時,Print A4紙張
         strExc(0) = "select * from(" & strMainSql & ") where " & _
                     " T16 is null and T18 is not null" & _
                     " order by CuNo,a0k04,a0k11,a0l02, a0k01,a0j25,cp10"
         intI = 1
         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
         longRow_Fax = 0
         If intI = 1 Then
            longRow_Fax = adoquery.RecordCount
            bolPrintA4 = True
            
            'Modify By Sindy 2015/12/21
'            MsgBox "請放置A4紙張，待放置完成，再按下確定鍵。", vbInformation
'            PUB_RestorePrinter Combo1
            '檢查是否有安裝PDFCreator
            PrinterIndex = -1
            For i = 0 To Printers.Count - 1
               If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
                  PrinterIndex = i
                  Exit For
               End If
            Next i
            If PrinterIndex < 0 Then
               MsgBox "請通知電腦中心安裝PDFCreator !!!"
               Exit Sub
            End If
            '*******************************
            PUB_SetOsDefaultPrinter Combo1 'Add By Sindy 2022/4/20
            PUB_RestorePrinter Printers(PrinterIndex).DeviceName
            '*******************************
            strTempFolder = PUB_Getdesktop & "\" & Text5 & "年單一傳真"
            If Dir(strTempFolder, vbDirectory) = "" Then
               MkDir strTempFolder
            End If
            '2015/12/21 END
            
'            Printer.Orientation = 1 '1.直印 2.橫印
'            Printer.PaperSize = 9
'            Call PrintA4(False, True, False)
            Call PrintA4_Excel(False, True, False)
         End If
         
         '********************
         ' 其他
         '********************
         '非上列者,Print A4紙張
         strExc(0) = "select * from(" & strMainSql & ") where " & _
                     " not (T16 is not null or (T16 is null and T18 is not null))" & _
                     " order by CuNo,a0k04,a0k11,a0l02,a0k01,a0j25,cp10"
         intI = 1
         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
         longRow_Other = 0
         If intI = 1 Then
            longRow_Other = adoquery.RecordCount
            bolPrintA4 = True
            
            '檢查是否有安裝PDFCreator
            PrinterIndex = -1
            For i = 0 To Printers.Count - 1
               If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
                  PrinterIndex = i
                  Exit For
               End If
            Next i
            If PrinterIndex < 0 Then
               MsgBox "請通知電腦中心安裝PDFCreator !!!"
               Exit Sub
            End If
            '*******************************
            PUB_SetOsDefaultPrinter Combo1 'Add By Sindy 2022/4/20
            PUB_RestorePrinter Printers(PrinterIndex).DeviceName
            '*******************************
            strTempFolder = PUB_Getdesktop & "\" & Text5 & "年其他"
            If Dir(strTempFolder, vbDirectory) = "" Then
               MkDir strTempFolder
            End If
            '2015/12/21 END
            
'            Printer.Orientation = 1 '1.直印 2.橫印
'            Printer.PaperSize = 9
'            Call PrintA4(False, False, True)
            Call PrintA4_Excel(False, False, True)
         End If
         
      Else '複合 維持列印A4,傳真
         MsgBox "請放置A4紙張，待放置完成，再按下確定鍵。", vbInformation
         '*******************************
         PUB_SetOsDefaultPrinter Combo1 'Add By Sindy 2022/4/20
         PUB_RestorePrinter Combo1
         '*******************************
         strExc(0) = "select * from(" & strMainSql & ") where " & _
                     " T18 is not null" & _
                     " order by CuNo,a0k04,a0k11,a0l02,a0k01,a0j25,cp10"
         intI = 1
         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
         longRow_Fax = 0
         If intI = 1 Then
            longRow_Fax = adoquery.RecordCount
            bolPrintA4 = True
            
'            Printer.Orientation = 1 '1.直印 2.橫印
'            Printer.PaperSize = 9
'            Call PrintA4(False, False, False)
            Call PrintA4_Excel(False, False, False)
         End If
         'Modify By Sindy 不管有無傳真一律A4列印出來
         strExc(0) = "select * from(" & strMainSql & ") where " & _
                     " T18 is null" & _
                     " order by CuNo,a0k04,a0k11,a0l02,a0k01,a0j25,cp10"
         intI = 1
         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
         longRow_Other = 0
         If intI = 1 Then
            longRow_Other = adoquery.RecordCount
            bolPrintA4 = True
            
'            Printer.Orientation = 1 '1.直印 2.橫印
'            Printer.PaperSize = 9
'            Call PrintA4(False, False, False)
            Call PrintA4_Excel(False, False, False)
         End If
      End If
'      '********************
'      ' 點陣中一刀
'      '********************
'      '無E-Mail無傳真時,列印點陣中一刀格式
'      If txtType = "1" Then '單一
'         strExc(0) = "select * from(" & strMainSql & ") where " & _
'                     " T16 is null and T18 is null" & _
'                     " order by CuNo,a0k04,a0k11,a0l02, a0k01,a0j25,cp10"
'      Else '複合
'         strExc(0) = "select * from(" & strMainSql & ") where " & _
'                     " T18 is null" & _
'                     " order by CuNo,a0k04,a0k11,a0l02, a0k01,a0j25,cp10"
'      End If
      
RunAllEnd:
      If longRow_Mail + longRow_Fax + longRow_Other = 0 Then
         MsgBox "無資料可列印！"
      Else
         If TempFileName <> "" Then
'            'Add By Sindy 2014/12/23
'            Print #ff, " E-Mail: " & dblTotMail & " 筆;　傳真: " & dblTotFax & " 筆;　其他: " & dblTotOther & " 筆"
'            '2014/12/23 END
'            Close ff
            'Modify By Sindy 2015/12/21
            '存檔
            intTxtCounter = intTxtCounter + 2
            wksAnnuity.Range("A" & intTxtCounter).Value = " E-Mail: " & dblTotMail & " 筆;　傳真: " & dblTotFax & " 筆;　其他: " & dblTotOther & " 筆"
            'Modify by Amy 2016/06/23 +判斷版本
            If Val(xlsAnnuity.Version) < 12 Then
                xlsAnnuity.Workbooks(1).SaveAs FileName:=TempFileName, FileFormat:=-4143
            Else
                xlsAnnuity.Workbooks(1).SaveAs FileName:=TempFileName, FileFormat:=56
            End If
            'end 2016/06/23
            xlsAnnuity.Workbooks.Close
            xlsAnnuity.Quit
            Set xlsAnnuity = Nothing
            Set wksAnnuity = Nothing
            '2015/12/21 END
         End If
         
         'Add by Sindy 2019/12/24 完成後,寄信通知操作人員
         If Check7.Value = 1 Then '寄發，作業執行完畢通知信
            'Add by Sindy 2021/2/22
            If UCase(cmdMail.Tag) = UCase("EMail") Then '單純只是發提醒客戶摧扣繳憑單的EMail
               'Modify By Sindy 2025/1/21 寄件人 strUserNum 改為 m_AccMail
               PUB_SendMail m_AccMail, strUserNum, "", Me.Caption & "（發催扣繳憑單Mail）,執行完畢！(操作的主機名稱=" & PUB_ReadHostName & ")", "作業已執行完畢" & vbCrLf & "請電腦中心寄桌面上的Excel檔(..客戶扣繳寄送狀況清單.xls)給您。", , , , , , , m_AccMail, m_AccMail, , , , "97038"
            '2021/2/22 END
            ElseIf txtType = "1" Then '單一:才要寄EMail
               'Modify By Sindy 2025/1/21 寄件人 strUserNum 改為 m_AccMail
               PUB_SendMail m_AccMail, strUserNum, "", Me.Caption & "（寄發扣繳信）,執行完畢！(操作的主機名稱=" & PUB_ReadHostName & ")", "作業已執行完畢" & vbCrLf & "請電腦中心寄桌面上的Excel檔(..客戶扣繳寄送狀況清單.xls)和資料夾(" & Left(strSrvDate(2), 3) & "年其他、" & Left(strSrvDate(2), 3) & "年單一傳真)給您。", , , , , , , m_AccMail, m_AccMail, , , , "97038"
            End If
         End If
         '2019/12/24 END
         
         'Modify By Sindy 2016/12/12 是否有E-Mail傳送失敗,增加顯示訊息的內容
         'MsgBox "列印完成！" & IIf(TempFileName <> "", "（清單已存至桌面：" & TempFileName & "）", "")
         MsgBox "完成！" & IIf(TempFileName <> "", "（清單已存至桌面：" & TempFileName & "）", "") & vbCrLf & _
                IIf(bolMailSendErr = True, "系統發信失敗：" & App.path & "\Log_" & App.EXEName & ".txt", "")
      End If
      
   Else
      'Excel檔要多產生,中南高電子檔,不分萬號存檔
      For Excel_kk = 1 To 4
         bolWriteData = False 'Add By Sindy 2016/11/17
         If Excel_kk = 1 Then '全部
            'Modify By Sindy 2016/11/30
            'strExc(0) = "select * from(" & strMainSql & ") "
            'Modify By Sindy 2018/1/5 and instr(T22,'中')=0 and T22<>'台南所' and T22<>'高雄所' ==> and ((instr(T22,'中')=0 and T22<>'台南所' and T22<>'高雄所') or t22 is null)
            strExc(0) = "select * from(" & Replace(strMainSql, "and T14='" & strUserNum & "'", "and T14='" & strUserNum & "' and ((instr(T22,'中')=0 and T22<>'台南所' and T22<>'高雄所') or t22 is null)") & ") "
            '2016/11/30 END
            'Modified by Morgan 2011/12/21
            'strExc(0) = strExc(0) & " order by CuNo,a0k04,a0k11,a0l02, a0k01, cp10"
            'Modify By Sindy 2016/6/4
            If txtType = "1" Then '單一
               strExc(0) = strExc(0) & " order by CuNo,a0k04,a0k11,a0l02, a0k01,a0j25,cp10"
            Else '複合
               '改以收據抬頭+客戶編號排序
               strExc(0) = strExc(0) & " order by a0k04,CuNo,a0k11,a0l02, a0k01,a0j25,cp10"
            End If
            '2016/6/4 END
         ElseIf Excel_kk = 2 Then '中所
            strExc(0) = "select * from(" & Replace(strMainSql, "and T14='" & strUserNum & "'", "and T14='" & strUserNum & "' and instr(T22,'中')>0") & ") "
            If txtType = "1" Then '單一
               strExc(0) = strExc(0) & " order by CuNo,a0k04,a0k11,a0l02, a0k01,a0j25,cp10"
            Else '複合
               '以收據抬頭+客戶編號排序
               strExc(0) = strExc(0) & " order by a0k04,CuNo,a0k11,a0l02, a0k01,a0j25,cp10"
            End If
         ElseIf Excel_kk = 3 Then '南所
            strExc(0) = "select * from(" & Replace(strMainSql, "and T14='" & strUserNum & "'", "and T14='" & strUserNum & "' and T22='台南所'") & ") "
            If txtType = "1" Then '單一
               strExc(0) = strExc(0) & " order by CuNo,a0k04,a0k11,a0l02, a0k01,a0j25,cp10"
            Else '複合
               '以收據抬頭+客戶編號排序
               strExc(0) = strExc(0) & " order by a0k04,CuNo,a0k11,a0l02, a0k01,a0j25,cp10"
            End If
         ElseIf Excel_kk = 4 Then '高所
            strExc(0) = "select * from(" & Replace(strMainSql, "and T14='" & strUserNum & "'", "and T14='" & strUserNum & "' and T22='高雄所'") & ") "
            If txtType = "1" Then '單一
               strExc(0) = strExc(0) & " order by CuNo,a0k04,a0k11,a0l02, a0k01,a0j25,cp10"
            Else '複合
               '以收據抬頭+客戶編號排序
               strExc(0) = strExc(0) & " order by a0k04,CuNo,a0k11,a0l02, a0k01,a0j25,cp10"
            End If
         End If
         intI = 1
         Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            lngPageNo = 0: stCustNo = "": bol1stPage = True: bolStarSign = False
            stCustName = "": stCustAddr = "": stUniNo = "": stCompName = "": stCompNo = "": stCustNo_OK = ""
            Erase lngSum
            With adoquery
               .MoveFirst
               Set xlsAnnuity_A4 = New Excel.Application
               Call SetExcelWorksheets
               Do While Not .EOF
                  '收據抬頭不同或客戶不同
                  If stCustName <> .Fields("a0k04") Or stCustNo <> .Fields("CuNo") Then
                     If lngPageNo > 0 Then
                        PrintSum_Excel lngSum, intCounter
                        Erase lngSum
                        lngPageNo = 0
                     End If
                     stCustNo = .Fields("CuNo")
                     stCustName = "" & .Fields("a0k04")
                     stMail = "" & .Fields("T16")
                     stTel = "" & .Fields("T17")
                     stFax = "" & .Fields("T18")
                     'Add By Sindy 2015/1/15
                     stCU15 = "" & .Fields("T19")
                     stCustAddr = "" & .Fields("T20")
                     stAccNote = "" & .Fields("T21")
                     stSaleArea = "" & .Fields("T22")
                     stSales = "" & .Fields("T23")
                     '2015/1/15 END
      '                  If PrintCheck(.Fields("CuNo")) = True Then
                        'Add By Sindy 2013/12/24 以客戶編號第2碼則存一個檔案
                        If stCustNo_OK <> "" Then
                           'Modify By Sindy 2016/6/4 + And txtType = "1" 單一才切檔案做存檔
                           'If Mid(.Fields("CuNo"), 2, 1) <> Mid(stCustNo_OK, 2, 1) And txtType = "1" Then
                           'Modify By Sindy 2016/11/17 單一且為全部資料時才切萬號
                           If Mid(.Fields("CuNo"), 2, 1) <> Mid(stCustNo_OK, 2, 1) And txtType = "1" And Excel_kk = 1 Then
                           '2016/11/17 END
                              '存檔
                              'Modify by Amy 2016/06/23 +判斷版本
                              If Val(xlsAnnuity_A4.Version) < 12 Then
                                   xlsAnnuity_A4.Workbooks(1).SaveAs FileName:=strExcelPath & Left(stCustNo_OK, 2) & "_" & Text5 & "年度扣繳明細核對表" & ACDate(strSrvDate(1)) & ServerTime & MsgText(43), FileFormat:=-4143
                              Else
                                   xlsAnnuity_A4.Workbooks(1).SaveAs FileName:=strExcelPath & Left(stCustNo_OK, 2) & "_" & Text5 & "年度扣繳明細核對表" & ACDate(strSrvDate(1)) & ServerTime & MsgText(43), FileFormat:=56
                              End If
                              'end 2016/06/23
                              xlsAnnuity_A4.Workbooks.Close
                              'xlsAnnuity_A4.Quit
                              Call SetExcelWorksheets
                           Else
                              '換頁
                              intCounter = intCounter + 1
                              wksAnnuity_A4.Range("A" & intCounter).Select
                              wksAnnuity_A4.HPageBreaks.add Before:=wksAnnuity_A4.Application.ActiveCell
                           End If
                        End If
                        stCustNo_OK = stCustNo
                        'stCustAddr = GetAddress(stCustNo)
                        If stCompNo <> "" & .Fields("a0k11") Then
                           stCompNo = "" & .Fields("a0k11")
                           GetCompInfo stCompNo, stCompName, stUniNo
                        End If
                        stTitle = "" & .Fields("a0k04")
                        PrintHead_Excel "E", stCustNo, Val(Text5), stUniNo, stCompName, stTitle, intCounter, stTel, stFax, stMail
      '                  End If
                  ElseIf lngPageNo > 0 Then
                        '公司不同
                        If stCompNo <> "" & .Fields("a0k11") Then
                           stCompNo = "" & .Fields("a0k11")
                           GetCompInfo stCompNo, stCompName, stUniNo
                           PrintSum_Excel lngSum, intCounter
                           Erase lngSum
                           stTitle = "" & .Fields("a0k04")
                           '換頁
                           intCounter = intCounter + 1
                           wksAnnuity_A4.Range("A" & intCounter).Select
                           wksAnnuity_A4.HPageBreaks.add Before:=wksAnnuity_A4.Application.ActiveCell
                           PrintHead_Excel "E", stCustNo, Val(Text5), stUniNo, stCompName, stTitle, intCounter, stTel, stFax, stMail
                        End If
                  End If
                  If lngPageNo > 0 Then
                     PrintData_Excel adoquery, intCounter
                     For intI = 1 To 5
                        lngSum(intI) = lngSum(intI) + .Fields("Fee" & intI)
                     Next
                     'dblSkipPageRow = dblSkipPageRow + 1
                  End If
                  .MoveNext
               Loop
                              
               'Add By Sindy 2014/12/9 檢查是否有寫入資料至Excel
               If bolWriteData = True Then
                  PrintSum_Excel lngSum, intCounter
                  Erase lngSum
'                  'Add By Sindy 2022/4/14
'                  With xlsAnnuity_A4.ActiveSheet.PageSetup
'                     .Zoom = False
'                     '.FitToPagesTall = 1 '縮放成一頁高
'                     .FitToPagesWide = 1 '縮放成一頁寬
'                     .FitToPagesTall = 1000 'Added by Morgan 2022/4/8 預設為1,筆數多時會縮小
'                  End With
'                  '2022/4/14 END
                  '存檔
                  'Modify by Amy 2016/06/23 +判斷版本
                  If Val(xlsAnnuity_A4.Version) < 12 Then
                     'If txtType = "1" Then '單一
                     If txtType = "1" And Excel_kk = 1 Then '單一且為全部資料時才切萬號 Modify By Sindy 2016/11/17
                        xlsAnnuity_A4.Workbooks(1).SaveAs FileName:=strExcelPath & Left(stCustNo, 2) & "_" & Text5 & "年度扣繳明細核對表" & ACDate(strSrvDate(1)) & ServerTime & MsgText(43), FileFormat:=-4143
                     Else '複合
                        xlsAnnuity_A4.Workbooks(1).SaveAs FileName:=strExcelPath & _
                           IIf(Excel_kk = 2, "台中", IIf(Excel_kk = 3, "台南", IIf(Excel_kk = 4, "高雄", ""))) & _
                           Text5 & "年度扣繳明細核對表" & ACDate(strSrvDate(1)) & ServerTime & MsgText(43), FileFormat:=-4143
                     End If
                  Else
                     'If txtType = "1" Then '單一
                     If txtType = "1" And Excel_kk = 1 Then '單一且為全部資料時才切萬號 Modify By Sindy 2016/11/17
                        xlsAnnuity_A4.Workbooks(1).SaveAs FileName:=strExcelPath & Left(stCustNo, 2) & "_" & Text5 & "年度扣繳明細核對表" & ACDate(strSrvDate(1)) & ServerTime & MsgText(43), FileFormat:=56
                     Else '複合
                        xlsAnnuity_A4.Workbooks(1).SaveAs FileName:=strExcelPath & _
                           IIf(Excel_kk = 2, "台中", IIf(Excel_kk = 3, "台南", IIf(Excel_kk = 4, "高雄", ""))) & _
                           Text5 & "年度扣繳明細核對表" & ACDate(strSrvDate(1)) & ServerTime & MsgText(43), FileFormat:=56
                     End If
                  End If
                  'end 2016/06/23
                  xlsAnnuity_A4.Workbooks.Close
                  xlsAnnuity_A4.Quit
                  'xlsAnnuity_A4.Visible = True
                  'xlsAnnuity_A4.WindowState = wdWindowStateMaximize
                  Set xlsAnnuity_A4 = Nothing
                  Set wksAnnuity_A4 = Nothing
               End If
            End With
         End If
      Next Excel_kk
      MsgBox "檔案已儲存至" & strExcelPath
      'Add By Sindy 2022/4/14
      '直接開啟視窗
      ShellExecute hLocalFile, "explore", strExcelPath, vbNullString, vbNullString, 1
      '2022/4/14 END
   End If
   
   adoquery.Close
   Set adoquery = Nothing
   txtNote.Visible = False 'Add By Sindy 2022/4/20
   
   'Add By Sindy 2016/1/6 檢查尚無處理到的資料,將T06欄位上註記為A
   strSql = "update ACCTMP44q0 set T06='A'" & _
            " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
            " and T06='X'"
   cnnConnection.Execute strSql, intI
   '2016/1/6 END
   
   '*******************************
   PUB_SetOsDefaultPrinter strPrinter 'Add By Sindy 2022/4/20
   PUB_RestorePrinter strPrinter 'Add By Sindy 2014/10/21 復原預設印表機
   '*******************************
   Exit Sub
   
ErrHnd:
   txtNote.Visible = False 'Add By Sindy 2022/4/20
   Set adoquery = Nothing
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   Set xlsAnnuity_A4 = Nothing
   Set wksAnnuity_A4 = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add By Sindy 2022/4/18 改用Excel列印; A4格式(E-Mail 或 傳真 或 其他)
Private Sub PrintA4_Excel(bolEmail As Boolean, bolFax As Boolean, bolOther As Boolean)
Dim strFilePathName As String, stCustName As String, stCompNo As String
Dim stFaxAcc As String
Dim strGetCustNo As String 'Add By Sindy 2019/3/25
   
   'If TempFileName = "" And Not (Left(txtCustNo(0), 6) = Left(txtCustNo(1), 6)) Then
   If TempFileName = "" And txtType = "1" Then
      'Modify By Sindy 2015/12/21
'      TempFileName = PUB_Getdesktop & "\" & Text5 & IIf(Val(FCDate(MaskEdBox1.Text)) = 0 And Val(FCDate(MaskEdBox2.Text)) = 0, "", "-" & Val(FCDate(MaskEdBox1.Text)) & "-" & Val(FCDate(MaskEdBox2.Text))) & "-" & txtCustNo(0) & "-" & txtCustNo(1) & "客戶扣繳寄送狀況清單.txt"
'      If ff > 0 Then Close #ff
'      ff = FreeFile
'      Open TempFileName For Output As ff
'      Print #ff, "客戶編號  客戶名稱                                 寄送狀況"
'      Print #ff, "========= ======================================== ========================================"
      Set xlsAnnuity = New Excel.Application
      TempFileName = PUB_Getdesktop & "\" & Text5 & IIf(Val(FCDate(MaskEdBox1.Text)) = 0 And Val(FCDate(MaskEdBox2.Text)) = 0, "", "-" & Val(FCDate(MaskEdBox1.Text)) & "-" & Val(FCDate(MaskEdBox2.Text))) & "-" & txtCustNo(0) & "-" & txtCustNo(1) & "客戶扣繳寄送狀況清單.xls"
      Call SetExcelWorksheetsTxt
      Call PrintHeadTxt_Excel(intTxtCounter)
      '2015/12/21 END
   End If
   
'   Set adocheck = adoquery.Clone
   lngPageNo = 0: stCustNo = "": bol1stPage = True: bolStarSign = False
   stCustName = "": stUniNo = "": stCompName = "": stCompNo = ""
   strGetCustNo = "" 'Add By Sindy 2019/3/25
   Erase lngSum
   With adoquery
      .MoveFirst
'      Printer.FontName = "標楷體"
'      Printer.FontSize = 12
'      iChrPix = Printer.TextWidth("A")
'      iRowPix = Printer.TextHeight("A") + 50
'      GetPleft
'      lngXo = 400: lngYo = 500
      Set xlsAnnuity_A4 = New Excel.Application
      Call SetExcelWorksheets
      Do While Not .EOF
            '客戶不同
            'If stCustNo <> .Fields("CuNo") Then
            If stCustName <> .Fields("a0k04") Or stCustNo <> .Fields("CuNo") Then
               If lngPageNo > 0 Then
                  PrintSum_Excel lngSum, intCounter
                  Call PrintNote_Excel(stCompNo, bolEmail, intCounter)
                  Erase lngSum
                  lngPageNo = 0
                  '最後一頁加空白頁
                  'Printer.NewPage
                  
                  stCustNo = convForm(CheckStr(stCustNo), 9)
                  stCustName = convForm(CheckStr(stCustName), 40)
                  
                  'Modify By Sindy 2015/12/21
                  'If bolEmail = True Then
                  If bolEmail = True Or bolFax = True Or bolOther = True Then
                  '2015/12/21 END
                     If bolEmail = True Then
                        bolMailSendOk = False
                        Call RunIsMail(stCustName, stCustNo, strGetCustNo, stCompNo, True)
                     ElseIf bolFax = True Then
                        Call RunIsFaxPDF_Excel(stCustName, stCustNo, stFax, stFaxAcc)
                     Else
                        Call RunIsFaxPDF_Excel(stCustName, stCustNo, "", "")
                     End If
                     If TempFileName <> "" Then
                        If bolEmail = True Then
                           'Modify By Sindy 2015/12/21
                           'Print #ff, stCustNo & " " & stCustName & " EMail：" & stMail & "（" & IIf(bolMailSendOk = False, "N", "Y") & "）"
                           Call PrintDataTxt_Excel(intTxtCounter, stCustName, " EMail：" & stMail & "（" & IIf(bolMailSendOk = False, "N", "Y") & "）")
                           '2015/12/21 END
                           dblTotMail = dblTotMail + 1 'Add By Sindy 2014/12/23
                        ElseIf bolFax = True Then
                           'Modify By Sindy 2015/12/21
                           'Print #ff, stCustNo & " " & stCustName & " FAX：" & stFax
                           Call PrintDataTxt_Excel(intTxtCounter, stCustName, " FAX：" & stFax)
                           '2015/12/21 END
                           dblTotFax = dblTotFax + 1 'Add By Sindy 2014/12/23
                        Else
                           Call PrintDataTxt_Excel(intTxtCounter, stCustName, "")
                           dblTotOther = dblTotOther + 1
                        End If
                     End If
                     
'                     Printer.Orientation = 1 '1.直印 2.橫印
'                     Printer.PaperSize = 9
                     lngPageNo = 0: stCustNo = "": bol1stPage = True: bolStarSign = False
                     stCustName = "": stUniNo = "": stCompName = "": stCompNo = ""
                     strGetCustNo = "" 'Add By Sindy 2019/3/25
'                     Printer.FontName = "標楷體"
'                     Printer.FontSize = 12
'                     iChrPix = Printer.TextWidth("A")
'                     iRowPix = Printer.TextHeight("A") + 50
'                     GetPleft
'                     lngXo = 400: lngYo = 500
                     
                  Else
                     xlsAnnuity_A4.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
                     '列印標題
                     xlsAnnuity_A4.ActiveSheet.PageSetup.PrintTitleRows = "$1:$6"
                     With xlsAnnuity_A4.ActiveSheet.PageSetup
                        .Zoom = False
                        '.FitToPagesTall = 1 '縮放成一頁高
                        .FitToPagesWide = 1 '縮放成一頁寬
                        .FitToPagesTall = 1000 'Added by Morgan 2022/4/8 預設為1,筆數多時會縮小
                     End With
                     '列印出來
                     xlsAnnuity_A4.Workbooks(1).PrintOut
                     xlsAnnuity_A4.Workbooks.Close
                  End If
                  Call SetExcelWorksheets
               End If
               
               stCustNo = .Fields("CuNo")
               stCustName = "" & .Fields("a0k04")
               strGetCustNo = "" & .Fields("T29") 'Add By Sindy 2019/3/25
               stMail = "" & .Fields("T16")
               stTel = "" & .Fields("T17")
               stFax = "" & .Fields("T18")
               stFaxAcc = "" & .Fields("T28") 'Add By Sindy 2016/11/18 會計師傳真
               'Add By Sindy 2015/1/15
               stCU15 = "" & .Fields("T19")
               stAccNote = "" & .Fields("T21")
               stSaleArea = "" & .Fields("T22")
               stSales = "" & .Fields("T23")
               '2015/1/15 END
'               If PrintCheck(stCustNo) = True Then
                  If stCompNo <> "" & .Fields("a0k11") Then
                     stCompNo = "" & .Fields("a0k11")
                     GetCompInfo stCompNo, stCompName, stUniNo
                  End If
                  stTitle = "" & .Fields("a0k04")
                  'PrintHead stCustNo, Val(Text5), stUniNo, stCompName, stTitle, stTel, stFax
                  PrintHead_Excel "R", stCustNo, Val(Text5), stUniNo, stCompName, stTitle, intCounter, stTel, stFax, stMail
'               End If
            ElseIf lngPageNo > 0 Then
                  '公司不同
                  If stCompNo <> "" & .Fields("a0k11") Then
                     PrintSum_Excel lngSum, intCounter
                     Call PrintNote_Excel(stCompNo, bolEmail, intCounter) 'Add By Sindy 2020/4/24
                     Erase lngSum
                     
                     xlsAnnuity_A4.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
                     '列印標題
                     xlsAnnuity_A4.ActiveSheet.PageSetup.PrintTitleRows = "$1:$6"
                     With xlsAnnuity_A4.ActiveSheet.PageSetup
                        .Zoom = False
                        '.FitToPagesTall = 1 '縮放成一頁高
                        .FitToPagesWide = 1 '縮放成一頁寬
                        .FitToPagesTall = 1000 'Added by Morgan 2022/4/8 預設為1,筆數多時會縮小
                     End With
                     
                     'Add By Sindy 2020/4/24
                     If bolEmail = True Then
                        frmPDF.Show
                        strFilePathName = stCustNo & "-" & Text5 & "年-客戶扣繳明細核對表(" & IIf(stCompNo = "1", "商標", IIf(stCompNo = "2", "智慧所", "法律所")) & ")"
                        'frmPDF.StartProcess strTempFolder, strFilePathName
                        'Printer.EndDoc
                        'frmPDF.EndtProcess
                        '存檔:
                        xlsAnnuity_A4.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strTempFolder & "\" & strFilePathName & ".pdf", Quality:=0, _
                              IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
                        xlsAnnuity_A4.Workbooks.Close
                        Unload frmPDF
                        If strEmailAttch <> "" Then strEmailAttch = strEmailAttch & "*"
                        strEmailAttch = strEmailAttch & strTempFolder & "\" & strFilePathName & ".pdf"
                        
                     Else
                        '列印出來
                        xlsAnnuity_A4.Workbooks(1).PrintOut
                        xlsAnnuity_A4.Workbooks.Close
                     End If
                     Call SetExcelWorksheets
                     '2020/4/24 END
                     
                     lngPageNo = 0: bol1stPage = True
                     stCompNo = "" & .Fields("a0k11")
                     stTitle = "" & .Fields("a0k04")
                     GetCompInfo stCompNo, stCompName, stUniNo
                     'PrintHead stCustNo, Val(Text5), stUniNo, stCompName, stTitle, stTel, stFax
                     PrintHead_Excel "R", stCustNo, Val(Text5), stUniNo, stCompName, stTitle, intCounter, stTel, stFax, stMail
                  End If
            End If

            If lngPageNo > 0 Then
               PrintData_Excel adoquery, intCounter, IIf(bolEmail = True, "E", IIf(bolFax = True, "F", "P"))
               For intI = 1 To 5
                  lngSum(intI) = lngSum(intI) + .Fields("Fee" & intI)
               Next
            End If
         .MoveNext
      Loop
      
      If lngPageNo > 0 Then
         PrintSum_Excel lngSum, intCounter
         Call PrintNote_Excel(stCompNo, bolEmail, intCounter)
         Erase lngSum
         
         stCustNo = convForm(CheckStr(stCustNo), 9)
         stCustName = convForm(CheckStr(stCustName), 40)
         
         'Modify By Sindy 2015/12/21
         If bolEmail = True Or bolFax = True Or bolOther = True Then
            If bolEmail = True Then
               bolMailSendOk = False
               Call RunIsMail(stCustName, stCustNo, strGetCustNo, stCompNo, True)
            ElseIf bolFax = True Then
               Call RunIsFaxPDF_Excel(stCustName, stCustNo, stFax, stFaxAcc)
            Else
               Call RunIsFaxPDF_Excel(stCustName, stCustNo, "", "")
            End If
            '2015/12/21 END
            If TempFileName <> "" Then
               If bolEmail = True Then
                  'Modify By Sindy 2015/12/21
                  'Print #ff, stCustNo & " " & stCustName & " EMail：" & stMail & "（" & IIf(bolMailSendOk = False, "N", "Y") & "）"
                  Call PrintDataTxt_Excel(intTxtCounter, stCustName, " EMail：" & stMail & "（" & IIf(bolMailSendOk = False, "N", "Y") & "）")
                  '2015/12/21 END
                  dblTotMail = dblTotMail + 1 'Add By Sindy 2014/12/23
               ElseIf bolFax = True Then
                  'Modify By Sindy 2015/12/21
                  'Print #ff, stCustNo & " " & stCustName & " FAX：" & stFax
                  Call PrintDataTxt_Excel(intTxtCounter, stCustName, " FAX：" & stFax)
                  '2015/12/21 END
                  dblTotFax = dblTotFax + 1 'Add By Sindy 2014/12/23
               Else
                  Call PrintDataTxt_Excel(intTxtCounter, stCustName, "")
                  dblTotOther = dblTotOther + 1
               End If
            End If
         Else
            xlsAnnuity_A4.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
            '列印標題
            xlsAnnuity_A4.ActiveSheet.PageSetup.PrintTitleRows = "$1:$6"
            With xlsAnnuity_A4.ActiveSheet.PageSetup
               .Zoom = False
               '.FitToPagesTall = 1 '縮放成一頁高
               .FitToPagesWide = 1 '縮放成一頁寬
               .FitToPagesTall = 1000 'Added by Morgan 2022/4/8 預設為1,筆數多時會縮小
            End With
            '列印出來
            xlsAnnuity_A4.Workbooks(1).PrintOut
            xlsAnnuity_A4.Workbooks.Close
         End If
      End If
      xlsAnnuity_A4.Quit
      Set xlsAnnuity_A4 = Nothing
      Set wksAnnuity_A4 = Nothing
   End With
End Sub

'Add By Sindy 2014/10/21 A4格式(E-Mail 或 傳真 或 其他)
Private Sub PrintA4(bolEmail As Boolean, bolFax As Boolean, bolOther As Boolean)
Dim strFilePathName As String, stCustName As String, stCompNo As String
Dim stFaxAcc As String
Dim strGetCustNo As String 'Add By Sindy 2019/3/25
   
   'If TempFileName = "" And Not (Left(txtCustNo(0), 6) = Left(txtCustNo(1), 6)) Then
   If TempFileName = "" And txtType = "1" Then
      'Modify By Sindy 2015/12/21
'      TempFileName = PUB_Getdesktop & "\" & Text5 & IIf(Val(FCDate(MaskEdBox1.Text)) = 0 And Val(FCDate(MaskEdBox2.Text)) = 0, "", "-" & Val(FCDate(MaskEdBox1.Text)) & "-" & Val(FCDate(MaskEdBox2.Text))) & "-" & txtCustNo(0) & "-" & txtCustNo(1) & "客戶扣繳寄送狀況清單.txt"
'      If ff > 0 Then Close #ff
'      ff = FreeFile
'      Open TempFileName For Output As ff
'      Print #ff, "客戶編號  客戶名稱                                 寄送狀況"
'      Print #ff, "========= ======================================== ========================================"
      Set xlsAnnuity = New Excel.Application
      TempFileName = PUB_Getdesktop & "\" & Text5 & IIf(Val(FCDate(MaskEdBox1.Text)) = 0 And Val(FCDate(MaskEdBox2.Text)) = 0, "", "-" & Val(FCDate(MaskEdBox1.Text)) & "-" & Val(FCDate(MaskEdBox2.Text))) & "-" & txtCustNo(0) & "-" & txtCustNo(1) & "客戶扣繳寄送狀況清單.xls"
      Call SetExcelWorksheetsTxt
      Call PrintHeadTxt_Excel(intTxtCounter)
      '2015/12/21 END
   End If
   
'   Set adocheck = adoquery.Clone
   lngPageNo = 0: stCustNo = "": bol1stPage = True: bolStarSign = False
   stCustName = "": stUniNo = "": stCompName = "": stCompNo = ""
   strGetCustNo = "" 'Add By Sindy 2019/3/25
   Erase lngSum
   With adoquery
      .MoveFirst
      Printer.FontName = "標楷體"
      Printer.FontSize = 12
      iChrPix = Printer.TextWidth("A")
      iRowPix = Printer.TextHeight("A") + 50
      GetPleft
      lngXo = 400: lngYo = 500
      Do While Not .EOF
            '客戶不同
            'If stCustNo <> .Fields("CuNo") Then
            If stCustName <> .Fields("a0k04") Or stCustNo <> .Fields("CuNo") Then
               If lngPageNo > 0 Then
                  PrintSum lngSum
                  Call PrintNote(stCompNo, bolEmail)
                  Erase lngSum
                  lngPageNo = 0
                  '最後一頁加空白頁
                  'Printer.NewPage
                  
                  stCustNo = convForm(CheckStr(stCustNo), 9)
                  stCustName = convForm(CheckStr(stCustName), 40)
                  
                  'Modify By Sindy 2015/12/21
                  'If bolEmail = True Then
                  If bolEmail = True Or bolFax = True Or bolOther = True Then
                  '2015/12/21 END
                     If bolEmail = True Then
                        bolMailSendOk = False
                        Call RunIsMail(stCustName, stCustNo, strGetCustNo, stCompNo, False)
                     ElseIf bolFax = True Then
                        Call RunIsFaxPDF(stCustName, stCustNo, stFax, stFaxAcc)
                     Else
                        Call RunIsFaxPDF(stCustName, stCustNo, "", "")
                     End If
                     If TempFileName <> "" Then
                        If bolEmail = True Then
                           'Modify By Sindy 2015/12/21
                           'Print #ff, stCustNo & " " & stCustName & " EMail：" & stMail & "（" & IIf(bolMailSendOk = False, "N", "Y") & "）"
                           Call PrintDataTxt_Excel(intTxtCounter, stCustName, " EMail：" & stMail & "（" & IIf(bolMailSendOk = False, "N", "Y") & "）")
                           '2015/12/21 END
                           dblTotMail = dblTotMail + 1 'Add By Sindy 2014/12/23
                        ElseIf bolFax = True Then
                           'Modify By Sindy 2015/12/21
                           'Print #ff, stCustNo & " " & stCustName & " FAX：" & stFax
                           Call PrintDataTxt_Excel(intTxtCounter, stCustName, " FAX：" & stFax)
                           '2015/12/21 END
                           dblTotFax = dblTotFax + 1 'Add By Sindy 2014/12/23
                        Else
                           Call PrintDataTxt_Excel(intTxtCounter, stCustName, "")
                           dblTotOther = dblTotOther + 1
                        End If
                     End If
                     
                     Printer.Orientation = 1 '1.直印 2.橫印
                     Printer.PaperSize = 9
                     lngPageNo = 0: stCustNo = "": bol1stPage = True: bolStarSign = False
                     stCustName = "": stUniNo = "": stCompName = "": stCompNo = ""
                     strGetCustNo = "" 'Add By Sindy 2019/3/25
                     Printer.FontName = "標楷體"
                     Printer.FontSize = 12
                     iChrPix = Printer.TextWidth("A")
                     iRowPix = Printer.TextHeight("A") + 50
                     GetPleft
                     lngXo = 400: lngYo = 500
                  End If
               End If
               
               stCustNo = .Fields("CuNo")
               stCustName = "" & .Fields("a0k04")
               strGetCustNo = "" & .Fields("T29") 'Add By Sindy 2019/3/25
               stMail = "" & .Fields("T16")
               stTel = "" & .Fields("T17")
               stFax = "" & .Fields("T18")
               stFaxAcc = "" & .Fields("T28") 'Add By Sindy 2016/11/18 會計師傳真
               'Add By Sindy 2015/1/15
               stCU15 = "" & .Fields("T19")
               stAccNote = "" & .Fields("T21")
               stSaleArea = "" & .Fields("T22")
               stSales = "" & .Fields("T23")
               '2015/1/15 END
'               If PrintCheck(stCustNo) = True Then
                  If stCompNo <> "" & .Fields("a0k11") Then
                     stCompNo = "" & .Fields("a0k11")
                     GetCompInfo stCompNo, stCompName, stUniNo
                  End If
                  stTitle = "" & .Fields("a0k04")
                  PrintHead stCustNo, Val(Text5), stUniNo, stCompName, stTitle, stTel, stFax
'               End If
            ElseIf lngPageNo > 0 Then
                  '公司不同
                  If stCompNo <> "" & .Fields("a0k11") Then
                     PrintSum lngSum
                     Call PrintNote(stCompNo, bolEmail) 'Add By Sindy 2020/4/24
                     Erase lngSum
                     'Add By Sindy 2020/4/24
                     If bolEmail = True Then
                        frmPDF.Show
                        strFilePathName = stCustNo & "-" & Text5 & "年-客戶扣繳明細核對表(" & IIf(stCompNo = "1", "商標", IIf(stCompNo = "2", "智慧所", "法律所")) & ")"
                        frmPDF.StartProcess strTempFolder, strFilePathName
                        Printer.EndDoc
                        frmPDF.EndtProcess
                        Unload frmPDF
                        If strEmailAttch <> "" Then strEmailAttch = strEmailAttch & "*"
                        strEmailAttch = strEmailAttch & strTempFolder & "\" & strFilePathName & ".pdf"
                     Else
                        Printer.EndDoc
                     End If
                     '2020/4/24 END
                     
                     lngPageNo = 0: bol1stPage = True
                     stCompNo = "" & .Fields("a0k11")
                     stTitle = "" & .Fields("a0k04")
                     GetCompInfo stCompNo, stCompName, stUniNo
                     PrintHead stCustNo, Val(Text5), stUniNo, stCompName, stTitle, stTel, stFax
                  End If
            End If

            If lngPageNo > 0 Then
               NewLine
               PrintData adoquery, IIf(bolEmail = True, "E", IIf(bolFax = True, "F", "P"))
               For intI = 1 To 5
                  lngSum(intI) = lngSum(intI) + .Fields("Fee" & intI)
               Next
            End If
         .MoveNext
      Loop
      If lngPageNo > 0 Then
         PrintSum lngSum
         Call PrintNote(stCompNo, bolEmail)
         Erase lngSum
         
         stCustNo = convForm(CheckStr(stCustNo), 9)
         stCustName = convForm(CheckStr(stCustName), 40)
         
         'Modify By Sindy 2015/12/21
         If bolEmail = True Or bolFax = True Or bolOther = True Then
            If bolEmail = True Then
               bolMailSendOk = False
               Call RunIsMail(stCustName, stCustNo, strGetCustNo, stCompNo, False)
            ElseIf bolFax = True Then
               Call RunIsFaxPDF(stCustName, stCustNo, stFax, stFaxAcc)
            Else
               Call RunIsFaxPDF(stCustName, stCustNo, "", "")
            End If
            '2015/12/21 END
            If TempFileName <> "" Then
               If bolEmail = True Then
                  'Modify By Sindy 2015/12/21
                  'Print #ff, stCustNo & " " & stCustName & " EMail：" & stMail & "（" & IIf(bolMailSendOk = False, "N", "Y") & "）"
                  Call PrintDataTxt_Excel(intTxtCounter, stCustName, " EMail：" & stMail & "（" & IIf(bolMailSendOk = False, "N", "Y") & "）")
                  '2015/12/21 END
                  dblTotMail = dblTotMail + 1 'Add By Sindy 2014/12/23
               ElseIf bolFax = True Then
                  'Modify By Sindy 2015/12/21
                  'Print #ff, stCustNo & " " & stCustName & " FAX：" & stFax
                  Call PrintDataTxt_Excel(intTxtCounter, stCustName, " FAX：" & stFax)
                  '2015/12/21 END
                  dblTotFax = dblTotFax + 1 'Add By Sindy 2014/12/23
               Else
                  Call PrintDataTxt_Excel(intTxtCounter, stCustName, "")
                  dblTotOther = dblTotOther + 1
               End If
            End If
         End If
      End If
      Printer.EndDoc
   End With
End Sub

'Add By Sindy 2014/10/23
'Modify By Sindy 2019/3/25 + strGetCustNo As String
'Modify By Sindy 2022/4/19 + , bolExcel As Boolean
Private Sub RunIsMail(stCustName As String, stCustNo As String, strGetCustNo As String, stCompNo As String, bolExcel As Boolean)
Dim strFilePathName As String, strSubject As String, strContent As String
   
   frmPDF.Show
   strFilePathName = stCustNo & "-" & Text5 & "年-客戶扣繳明細核對表(" & IIf(stCompNo = "1", "商標", IIf(stCompNo = "2", "智慧所", "法律所")) & ")"
   If bolExcel = False Then
      frmPDF.StartProcess strTempFolder, strFilePathName
      Printer.EndDoc
      frmPDF.EndtProcess
   'Add By Sindy 2022/4/19
   Else
      xlsAnnuity_A4.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
      '列印標題
      xlsAnnuity_A4.ActiveSheet.PageSetup.PrintTitleRows = "$1:$6"
      With xlsAnnuity_A4.ActiveSheet.PageSetup
         .Zoom = False
         '.FitToPagesTall = 1 '縮放成一頁高
         .FitToPagesWide = 1 '縮放成一頁寬
         .FitToPagesTall = 1000 'Added by Morgan 2022/4/8 預設為1,筆數多時會縮小
      End With
      
      '存檔:
      xlsAnnuity_A4.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strTempFolder & "\" & strFilePathName & ".pdf", Quality:=0, _
            IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
      xlsAnnuity_A4.Workbooks.Close
   End If
   '2022/4/19 END
   Unload frmPDF
   
   'Add By Sindy 2021/12/27 X54468030新力美貿易有110年度扣繳明細表請核對！
   '客戶和辜反應:信箱中沒有附加檔案,所以沒有收到110年度扣繳資料
   If Dir(strTempFolder & "\" & strFilePathName & ".pdf") = MsgText(601) Then
      Sleep 5000
   End If
   '2021/12/27 END
   
   'Add By Sindy 2020/4/24
   If strEmailAttch <> "" Then strEmailAttch = strEmailAttch & "*"
   strEmailAttch = strEmailAttch & strTempFolder & "\" & strFilePathName & ".pdf"
   '2020/4/24 END
   
   'Add By Sindy 2019/3/20
   If strGetCustNo = "" And stCustNo <> "" Then strGetCustNo = stCustNo
   If Right(strGetCustNo, 3) = "000" Then
      strGetCustNo = Left(strGetCustNo, 6)
   End If
   '提醒客戶的摧扣繳憑單EMail,無附件
   If UCase(cmdMail.Tag) = UCase("EMail") Then
      strSubject = IIf(strGetCustNo <> "", strGetCustNo, "") & Left(stCustName, 6) & "【重要提醒】請提供" & Text5 & "年所得扣繳憑單"
      'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個 原:Pub_GetSpecMan("財務處總帳人員")
      strContent = "E-Mail:" & IIf(Pub_StrUserSt03 = "M51" Or Check5.Value = 1, strUserNum, stMail) & vbCrLf & vbCrLf & _
                  "親愛的先生/小姐 您好，" & vbCrLf & vbCrLf & vbCrLf & _
                  "<B>提醒您</B>，" & vbCrLf & vbCrLf & _
                  "因即將申報" & Text5 & "年度所得，<B>目前尚未收到貴公司" & Text5 & "年扣繳憑單</B>，再麻煩您協助詢問財務。" & vbCrLf & vbCrLf & _
                  "請將扣繳憑單影本EMail( " & m_AccMail & " )、傳真( 02-25011666 )。" & vbCrLf & vbCrLf & _
                  "若有任何疑問，也歡迎您隨時與我們聯繫。" & vbCrLf & vbCrLf & _
                  "<B>若已提供扣繳憑單予本所，請略過本通知</B>" & vbCrLf & vbCrLf & _
                  "謝謝！" & vbCrLf & vbCrLf & vbCrLf
      'Modify By Sindy 2024/12/18
      strContent = strContent & _
                  "台北所 02-25061023 分機 542-547" & vbCrLf & _
                  "台中所 04-23270288 分機 63-65" & vbCrLf & _
                  "台南所 06-2743866  分機 10" & vbCrLf & _
                  "高雄所 07-2363602  分機 20"
      '2024/12/18 END
                  '& _
                  "財務處　" & strEmp & IIf(strEMP_Tel <> "", "（分機：" & strEMP_Tel & "）", "") & vbCrLf '& _
                  "台一國際專利法律事務所" & vbCrLf & _
                  "台北市長安東路２段１１２號９樓" & vbCrLf & _
                  "電話：０２－２５０６１０２３" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
                  "傳真：０２－２５０１１６６６"
      'Modify By Sindy 2020/4/24 iSignatureID=1
      'Modify By Sindy 2020/12/23 iSignatureID=5
      'Modify By Sindy 2025/1/21 寄件人 strUserNum 改為 m_AccMail
      PUB_SendMail m_AccMail, IIf(Pub_StrUserSt03 = "M51" Or Check5.Value = 1, strUserNum, stMail), "", strSubject, strContent, , , , True, , , m_AccMail, m_AccMail, , True, , , , False, , , , , , 5
      If bolMailSendOk = False Then
         bolMailSendErr = True
      End If
   Else
   '2019/3/20 END
      'Modify By Sindy 2019/3/25
      'strSubject = Left(stCustName, 4) & Text5 & "年度扣繳明細表請核對！"
      strSubject = IIf(strGetCustNo <> "", strGetCustNo, "") & Left(stCustName, 6) & Text5 & "年度扣繳明細表請核對！"
      '2019/3/25 END
      
      'Modify By Sindy 2018/12/21 + 顯示E-Mail
'      strContent = "E-Mail:" & IIf(Pub_StrUserSt03 = "M51" Or Check5.Value = 1, strUserNum, stMail) & vbCrLf & vbCrLf & _
'                   "您好，" & vbCrLf & _
'                   "附件為　" & stCU15 & Text5 & "年度扣繳資料，煩請核對資料" & vbCrLf
'      If Check1.Value = 1 Then '稅額達 2,001 以上但未扣繳 (一收款單號合計達 2,001 含同收款單之所有收據)
'         'strContent = strContent & "（" & Text5 & "年截至目前　" & stCU15 & "有單筆或一次付款稅額超過2000元但尚未扣繳，請您與本所連絡）" & vbCrLf
'         strContent = strContent & "（" & Text5 & "年截至目前　" & stCU15 & "一次付款稅額超過2000元但尚未扣繳，請您與本所連絡）" & vbCrLf
'      'Add By Sindy 2016/2/24 + Check6.Value = 1
'      ElseIf Check6.Value = 1 Then '稅額達 2,001 以上但未扣繳 (單筆 2,001 含同收款單之所有收據)
'         strContent = strContent & "（" & Text5 & "年截至目前　" & stCU15 & "有單筆稅額超過2000元但尚未扣繳，請您與本所連絡）" & vbCrLf
'      'Add By Sindy 2019/12/11
'      ElseIf Check2.Value = 1 Then '只印全年服務費超過 20,000 或 服務費未超過 20,000 但有扣繳的 或 當年有收款且有建信箱者
'         If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(Text5) Then
'            strContent = strContent & "* 以上為" & Text5 & "年12月31日止　" & stCU15 & "之扣繳資料，請核對；若12月31日後有增加屬於當年度之應扣繳款項，請自行加入並合計。" & vbCrLf
'         Else
'            strContent = strContent & "* 以上為" & Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日止　" & stCU15 & "之扣繳資料，請核對；若12月31日後有增加屬於當年度之應扣繳款項，請自行加入並合計。" & vbCrLf
'         End If
'         '2019/12/11 END
'      End If
'      '2016/2/24 END
'      'Modify By Sindy 2019/12/24 + *（若金額正確請勿回電或回覆信件，提供本所扣繳憑單即可）
'      strContent = strContent & "若扣繳資料有任何問題，請務必儘速聯絡，謝謝您的合作！" & vbCrLf & "*（若金額正確請勿回電或回覆信件，提供本所扣繳憑單即可）" & vbCrLf & vbCrLf & _
'                   "本所有分【專利商標】及【專利法律】 二個扣繳單位：" & vbCrLf & _
'                   "台一國際專利商標事務所(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
'                   "台一國際專利法律事務所(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf & vbCrLf & _
'                   "●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:71006@taie.com.tw，感謝您的配合！" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
'                   "財務處　" & strEmp & vbCrLf & _
'                   "台一國際專利法律事務所" & vbCrLf & _
'                   "台北市長安東路２段１１２號９樓" & vbCrLf & _
'                   "電話：０２－２５０６１０２３" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
'                   "傳真：０２－２５０１１６６６"
      'Modify By Sindy 2020/4/24
      'Modify By Sindy 2020/12/30 + 扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。
      'Modify By Sindy 2022/2/14 "提醒您，承辦人員已更換為 辜小姐 E-MAIL : 71005@taie.com.tw" & vbCrLf & vbCrLf Mark
      strContent = stMail & vbCrLf & vbCrLf & _
                  "您好，" & vbCrLf & _
                  "附件為　" & stCU15 & Text5 & "年度扣繳資料，若扣繳明細經核對無誤，請無須回覆，" & vbCrLf & _
                  "並請依附件內容開立扣繳憑單，扣繳憑單開立後請傳送至 " & m_AccMail & " 即可。" & vbCrLf & vbCrLf & _
                  "但若扣繳明細有誤差，請依以下本所所別，連絡貴公司當地對應的本所財務窗口，資訊如下：" & vbCrLf & vbCrLf
      'Modify By Sindy 2024/12/18
      strContent = strContent & _
                  "台北所 02-25061023 分機 542-547" & vbCrLf & _
                  "台中所 04-23270288 分機 63-65" & vbCrLf & _
                  "台南所 06-2743866  分機 10" & vbCrLf & _
                  "高雄所 07-2363602  分機 20" & vbCrLf & vbCrLf & _
                  "上述再請協助確認，謝謝。"
      '2024/12/18 END
                   '& _
                   "財務處　" & strEmp & IIf(strEMP_Tel <> "", "（分機：" & strEMP_Tel & "）", "") & vbCrLf
      '2020/4/24 END
      
      '當電腦中心人員或勾測試信箱時，均寄給操作人員
      'PUB_SendMail strUserNum, IIf(Pub_StrUserSt03 = "M51" Or Check5.Value = 1, strUserNum, stMail), "", strSubject, strContent, , strTempFolder & "\" & strFilePathName & ".pdf", , , , , , , , True
      'Modify By Sindy 2016/12/12 不顯示傳送失敗的Msg
      'Modify By Sindy 2016/12/15 IsByMsaIsp:對外信箱要設 True
      'Modify By Sindy 2020/4/24 iSignatureID=1
      'strTempFolder & "\" & strFilePathName & ".pdf"
      'Modify By Sindy 2020/12/23 iSignatureID=5
      'Modify By Sindy 2025/1/21 寄件人 strUserNum 改為 m_AccMail
      PUB_SendMail m_AccMail, IIf(Pub_StrUserSt03 = "M51" Or Check5.Value = 1, strUserNum, stMail), "", strSubject, strContent, , strEmailAttch, , True, , , m_AccMail, m_AccMail, , True, , , , False, , , , , , 5
      If bolMailSendOk = False Then
         bolMailSendErr = True
      End If
      '2016/12/12 END
      strEmailAttch = "" 'Add By Sindy 2020/4/24
   End If
   
   'Add By Sindy 2015/12/23 記錄執行日期時間
   strSql = "update ACCTMP44q0 set T08=" & strSrvDate(2) & ",T09=" & Right("000000" & ServerTime, 6) & " where t07='" & stCustNo & "' and t05='" & Me.Name & "' and T14='" & strUserNum & "' and T06='E'"
   cnnConnection.Execute strSql
   '2015/12/23 END
End Sub

'Add By Sindy 2015/12/21
'Modify By Sindy 2016/11/18 + stFaxAcc As String
Private Sub RunIsFaxPDF(stCustName As String, stCustNo As String, strFax As String, stFaxAcc As String)
Dim strFilePathName As String
   
   frmPDF.Show
   'Modify By Sindy 2016/11/18 + IIf(stFaxAcc <> "", ";" & stFaxAcc, "")
   strFilePathName = stCustNo & "-" & Left(stCustName, 4) & "-" & strFax & IIf(stFaxAcc <> "", ";" & stFaxAcc, "")
   frmPDF.StartProcess strTempFolder, Replace(strFilePathName, Chr(13) & Chr(10), "")
   Printer.EndDoc
   frmPDF.EndtProcess
   Unload frmPDF
   
   'Add By Sindy 2015/12/23 記錄執行日期時間
   strSql = "update ACCTMP44q0 set T08=" & strSrvDate(2) & ",T09=" & Right("000000" & ServerTime, 6) & " where t07='" & stCustNo & "' and t05='" & Me.Name & "' and T14='" & strUserNum & "' and T06='" & IIf(strFax = "" And stFaxAcc = "", "P", "F") & "'"
   cnnConnection.Execute strSql
   '2015/12/23 END
End Sub

'Add By Sindy 2022/4/18
Private Sub RunIsFaxPDF_Excel(stCustName As String, stCustNo As String, strFax As String, stFaxAcc As String)
Dim strFilePathName As String
   
   frmPDF.Show
   'Modify By Sindy 2016/11/18 + IIf(stFaxAcc <> "", ";" & stFaxAcc, "")
   strFilePathName = stCustNo & "-" & Left(stCustName, 4) & "-" & strFax & IIf(stFaxAcc <> "", ";" & stFaxAcc, "")
   
   xlsAnnuity_A4.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
   '列印標題
   xlsAnnuity_A4.ActiveSheet.PageSetup.PrintTitleRows = "$1:$6"
   With xlsAnnuity_A4.ActiveSheet.PageSetup
      .Zoom = False
      '.FitToPagesTall = 1 '縮放成一頁高
      .FitToPagesWide = 1 '縮放成一頁寬
      .FitToPagesTall = 1000 'Added by Morgan 2022/4/8 預設為1,筆數多時會縮小
   End With
   
   '存檔:
   xlsAnnuity_A4.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strTempFolder & "\" & Replace(strFilePathName, Chr(13) & Chr(10), "") & ".pdf", Quality:=0, _
         IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
   xlsAnnuity_A4.Workbooks.Close
   Unload frmPDF
   
   'Add By Sindy 2015/12/23 記錄執行日期時間
   strSql = "update ACCTMP44q0 set T08=" & strSrvDate(2) & ",T09=" & Right("000000" & ServerTime, 6) & " where t07='" & stCustNo & "' and t05='" & Me.Name & "' and T14='" & strUserNum & "' and T06='" & IIf(strFax = "" And stFaxAcc = "", "P", "F") & "'"
   cnnConnection.Execute strSql
   '2015/12/23 END
End Sub

'Add By Sindy 2022/4/19
Private Sub PrintNote_Excel(strCmp As String, Optional bolEmail As Boolean = False, Optional ByRef iRow As Integer)
   Dim strNoteDate As String
   
   If bolStarSign = True Then
      iRow = iRow + 1
      strExc(1) = "註記有 * 號的是本所執行業務報酬服務費超過20010則依法規定需扣繳，　貴公司該"
      strExc(2) = "筆付款時未扣稅款，若要補扣繳建議方式如下："
      strExc(3) = "１．請自行填寫各類所得繳款書並至銀行繳交該筆稅款，再將該繳款書傳真至本所，"
      'Modify by 2024/05/17 財務2個特殊設定拆成3個 原:Pub_GetSpecMan("財務處總帳人員")
      strExc(4) = "　　改成本所當即退還　貴公司該筆稅款(給付所得額10%)。"
      strExc(5) = "　　傳真 02-25011666 " & Pub_EMPTelNumbrandName(stTxtPerson, "") & "收" '& IIf(Pub_GetSpecMan("財務處總帳人員") = "71006", "楊小姐", "辜小姐") & "收"
      strExc(6) = "２．請自行填寫各類所得繳款書並寄至本所，本所繳交稅款後會"
      strExc(7) = "　　再將繳款書寄還　貴公司。"
      strExc(8) = "３．若以上說明有不詳之處，歡迎來電。 TEL 02-25061023" & Pub_EMPTelNumbrandName(stTxtPerson, "#") 'Modify By Sindy 2021/6/9 IIf(Pub_GetSpecMan("財務處總帳人員") = "71006", "#545 楊小姐", "#543 辜小姐")
      'end 2024/05/15
      For intI = 1 To 8 '6
         iRow = iRow + 1
         wksAnnuity_A4.Range("B" & iRow).Value = strExc(intI)
      Next
      
      'Added by Morgan 2012/12/26
      '都要印--辜
      'Add By Sindy 2014/10/28
      If bolEmail = False Then
      '2014/10/28 END
         iRow = iRow + 2
         wksAnnuity_A4.Range("B" & iRow).Value = "* 為能更有效率的核校資料，擬以電子郵件方式寄發扣繳核對表，"
         iRow = iRow + 1
         'Modify by Amy 2024/05/17 財務2個特殊設定拆成3個 原:Pub_GetSpecMan("財務處總帳人員")
         wksAnnuity_A4.Range("B" & iRow).Value = "  請提供您的e-mail address至" & m_AccMail & "，謝謝合作！"
      End If
      '2012/12/26
      'Add By Sindy 2014/10/21
      iRow = iRow + 1
      'Modify By Sindy 2020/12/28 備註日期調整,前後日期是要相同
      If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(Text5) Then
         strNoteDate = Text5 & "年12月31日"
      Else
         strNoteDate = Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日"
      End If
      wksAnnuity_A4.Range("B" & iRow).Value = "* 以上為" & strNoteDate & "止　" & stCU15 & "之扣繳資料，請核對；"
      iRow = iRow + 1
      wksAnnuity_A4.Range("B" & iRow).Value = "  若" & Right(strNoteDate, 6) & "後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
      iRow = iRow + 1
      wksAnnuity_A4.Range("B" & iRow).Value = "扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。"  'Add By Sindy 2020/12/30
      '2020/12/28 END
      'Modify By Sindy 2016/10/24
      iRow = iRow + 2
      
      bolStarSign = False
        
   Else
      iRow = iRow + 2
      'Modify by Amy 2024/05/17 財務2個特殊設定拆成3個, 分機資訊不寫於 MsgText(99)
      wksAnnuity_A4.Range("B" & iRow).Value = MsgText(99) & Pub_EMPTelNumbrandName(stTxtPerson, "#")
      '2010/12/14 add by sonia
      'Add By Sindy 2014/10/28
      If bolEmail = False Then
      '2014/10/28 END
         iRow = iRow + 1
         wksAnnuity_A4.Range("B" & iRow).Value = "* 為能更有效率的核校資料，擬以電子郵件方式寄發扣繳核對表，"
         iRow = iRow + 1
         'Modify by Amy 2024/05/17 財務2個特殊設定拆成3個 原:Pub_GetSpecMan("財務處總帳人員")
         wksAnnuity_A4.Range("B" & iRow).Value = "  請提供您的e-mail address至" & m_AccMail & "，謝謝合作！"
      End If
      '2010/12/14 end
      'Add By Sindy 2014/10/21
      iRow = iRow + 1
      'Modify By Sindy 2020/12/28 備註日期調整,前後日期是要相同
      If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(Text5) Then
         strNoteDate = Text5 & "年12月31日"
      Else
         strNoteDate = Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日"
      End If
      wksAnnuity_A4.Range("B" & iRow).Value = "* 以上為" & strNoteDate & "止　" & stCU15 & "之扣繳資料，請核對；"
      iRow = iRow + 1
      wksAnnuity_A4.Range("B" & iRow).Value = "  若" & Right(strNoteDate, 6) & "後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
      iRow = iRow + 1
      wksAnnuity_A4.Range("B" & iRow).Value = "扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。"  'Add By Sindy 2020/12/30
      '2020/12/28 END
      'Add By Sindy 2016/10/24
      iRow = iRow + 2
      
   End If
   'NewPage 'Remove by Morgan 2010/12/23 空白頁改印第二頁
   
   'Modify By Sindy 2021/12/2
   Dim strName As String, strNo As String, strType As String, strAddr As String
   If strCmp = "L" Then '台一國際法律事務所
      strName = CompNameQuery(strCmp): strNo = "77211833": strType = "9A-10": strAddr = "台北巿中山區朱園里7鄰長安東路2段110號4樓"
   Else '台一國際智慧財產事務所
      strName = CompNameQuery("2"): strNo = "04146457": strType = "9A-93 或 9A-91": strAddr = "台北巿中山區朱園里7鄰長安東路2段112號9樓"
   End If
   wksAnnuity_A4.Range("A" & iRow).Value = "本所扣繳資訊如下，請開立扣繳憑單"
   iRow = iRow + 1
   wksAnnuity_A4.Range("A" & iRow).Value = "1.所 得 人：" & CompNameQuery(strCmp)
   iRow = iRow + 1
   wksAnnuity_A4.Range("A" & iRow).Value = "2.統一編號：" & strNo
   iRow = iRow + 1
   wksAnnuity_A4.Range("A" & iRow).Value = "3.所得類別：" & strType
   iRow = iRow + 1
   wksAnnuity_A4.Range("A" & iRow).Value = "4.地　　址：" & strAddr
   'Add By Sindy 2021/12/16
   iRow = iRow + 1
   'Modify by Amy 2024/05/17 原:71005
   wksAnnuity_A4.Range("A" & iRow).Value = "5.扣繳憑單請  E-MAIL：" & m_AccMail
   '2021/12/16 END
   iRow = iRow + 2
   wksAnnuity_A4.Range("A" & iRow).Value = "感謝您的配合與支持！"
   Exit Sub
   '2021/12/2 END
End Sub

Private Sub PrintNote(strCmp As String, Optional bolEmail As Boolean = False)
   Dim strNoteDate As String
   
'   'Add By Sindy 2014/10/21
'   m_CU15 = ""
'   strExc(0) = "select DECODE(CU15,'0','台端','1','貴公司','貴單位') cu15 from customer" & _
'               " where cu01='" & Left(stCustNo, 8) & "' AND cu02='" & Mid(stCustNo, 9, 1) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      m_CU15 = "" & RsTemp.Fields("cu15")
'   End If
'   '2014/10/21 END
   
   If bolStarSign = True Then
      NewLine 10
      NewLine 9
      Printer.FontSize = 12
      strExc(1) = "註記有 * 號的是本所執行業務報酬服務費超過20010則依法規定需扣繳，　貴公司該"
      strExc(2) = "筆付款時未扣稅款，若要補扣繳建議方式如下："
      strExc(3) = "１．請自行填寫各類所得繳款書並至銀行繳交該筆稅款，再將該繳款書傳真至本所，"
      'Modify by Amy 2024/05/17 財務2個特殊設定拆成3個 原:Pub_GetSpecMan("財務處總帳人員")
      strExc(4) = "　　改成本所當即退還　貴公司該筆稅款(給付所得額10%)。傳真 02-25011666 " & Pub_EMPTelNumbrandName(stTxtPerson, "") & "收" '& IIf(Pub_GetSpecMan("財務處總帳人員") = "71006", "楊小姐", "辜小姐") & "收"
      strExc(5) = "２．請自行填寫各類所得繳款書並寄至本所，本所繳交稅款後會再將繳款書寄還　貴公司。"
      strExc(6) = "３．若以上說明有不詳之處，歡迎來電。 TEL 02-25061023" & Pub_EMPTelNumbrandName(stTxtPerson, "#") 'Modify By Sindy 2021/6/9 IIf(Pub_GetSpecMan("財務處總帳人員") = "71006", "#545 楊小姐", "#543 辜小姐")
      'end 2024/05/17
      For intI = 1 To 6
         Printer.CurrentX = lngXo + PLeft(2)
         Printer.CurrentY = lngY
         Printer.Print strExc(intI)
         lngY = lngY + 300
      Next
      
      'Added by Morgan 2012/12/26
      '都要印--辜
      'Add By Sindy 2014/10/28
      If bolEmail = False Then
      '2014/10/28 END
         NewLine
         Printer.CurrentX = lngXo + PLeft(2)
         Printer.CurrentY = lngY
         Printer.Print "* 為能更有效率的核校資料，擬以電子郵件方式寄發扣繳核對表，"
         NewLine
         Printer.CurrentX = lngXo + PLeft(2)
         Printer.CurrentY = lngY
         'Modify by Amy 2024/05/17 Pub_GetSpecMan("財務處總帳人員")
         Printer.Print "  請提供您的e-mail address至" & m_AccMail & "，謝謝合作！"
      End If
      '2012/12/26
      'Add By Sindy 2014/10/21
      NewLine
      Printer.CurrentX = lngXo + PLeft(2)
      Printer.CurrentY = lngY
      'Modify By Sindy 2020/12/28 備註日期調整,前後日期是要相同
      If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(Text5) Then
         strNoteDate = Text5 & "年12月31日"
      Else
         strNoteDate = Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日"
      End If
'      'Modify By Sindy 2015/1/21
'      'Printer.Print "* 以上為" & Text5 & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日止　" & stCU15 & "之扣繳資料，請核對；"
'      If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(Text5) Then
'         Printer.Print "* 以上為" & Text5 & "年12月31日止　" & stCU15 & "之扣繳資料，請核對；"
'      Else
'         Printer.Print "* 以上為" & Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日止　" & stCU15 & "之扣繳資料，請核對；"
'      End If
'      '2015/1/21 END
      Printer.Print "* 以上為" & strNoteDate & "止　" & stCU15 & "之扣繳資料，請核對；"
      NewLine
      Printer.CurrentX = lngXo + PLeft(2)
      Printer.CurrentY = lngY
      Printer.Print "  若" & Right(strNoteDate, 6) & "後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
      NewLine
      Printer.CurrentX = lngXo + PLeft(2)
      Printer.CurrentY = lngY
      Printer.Print "扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。" 'Add By Sindy 2020/12/30
      '2020/12/28 END
      'Modify By Sindy 2016/10/24
      NewLine
      NewLine
      
      bolStarSign = False
      
'      Printer.CurrentX = lngXo
'      Printer.CurrentY = lngY + 600
'      'Modify By Sindy 2020/4/24
'      If strCmp = "1" Then
'         Printer.Print "　　　　　　　　　　　　　　　(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf
'      ElseIf strCmp = "2" Then
'         Printer.Print "　　　　　　　　　　　　　　　(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf
'      Else 'L
'         Printer.Print "　　　　　　　　　　　　　　　(統編 77211833) 9A 代號10律師  台北巿長安東路2段110號4樓" & vbCrLf
'      End If
''      Printer.Print "本所有分【專利商標】及【專利法律】 二個扣繳單位：" & vbCrLf & _
''         "　　　　　　　　　　　　　(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
''         "　　　　　　　　　　　　　(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓"
'      '2020/4/24 END
'      '2016/10/24 END
'      '2014/10/21 END
'      'Add By Sindy 2016/12/12 改粗體字
'      Printer.FontBold = True
'      Printer.CurrentX = lngXo
'      Printer.CurrentY = lngY + 600 '825
'      'Modify By Sindy 2020/4/24
'      Printer.Print CompNameQuery(strCmp)
'      '2020/4/24 END
''      Printer.Print "台一國際專利商標事務所"
''      Printer.CurrentX = lngXo
''      Printer.CurrentY = lngY + 1075
''      Printer.Print "台一國際專利法律事務所"
'      Printer.CurrentX = lngXo
'      Printer.CurrentY = lngY + 900 '1325
'      Printer.Print "●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:" & Pub_GetSpecMan("財務處總帳人員") & "@taie.com.tw，感謝您的配合！"
'      Printer.FontBold = False
'      '2016/12/12 END
      
   Else
      NewLine
      NewLine
      Printer.FontSize = 12
      Printer.CurrentX = lngXo + PLeft(2)
      Printer.CurrentY = lngY
      'Modify by Amy 2024/05/17 財務2個特殊設定拆成3個, 分機資訊不寫於 MsgText(99)
      Printer.Print MsgText(99) & Pub_EMPTelNumbrandName(stTxtPerson, "#")
      '2010/12/14 add by sonia
      'Add By Sindy 2014/10/28
      If bolEmail = False Then
      '2014/10/28 END
         NewLine
         Printer.CurrentX = lngXo + PLeft(2)
         Printer.CurrentY = lngY
         Printer.Print "* 為能更有效率的核校資料，擬以電子郵件方式寄發扣繳核對表，"
         NewLine
         Printer.CurrentX = lngXo + PLeft(2)
         Printer.CurrentY = lngY
         'Modify by Amy 2024/05/17 Pub_GetSpecMan("財務處總帳人員")
         Printer.Print "  請提供您的e-mail address至" & m_AccMail & "，謝謝合作！"
      End If
      '2010/12/14 end
      'Add By Sindy 2014/10/21
      NewLine
      Printer.CurrentX = lngXo + PLeft(2)
      Printer.CurrentY = lngY
      'Modify By Sindy 2020/12/28 備註日期調整,前後日期是要相同
      If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(Text5) Then
         strNoteDate = Text5 & "年12月31日"
      Else
         strNoteDate = Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日"
      End If
'      'Modify By Sindy 2015/1/21
'      'Printer.Print "* 以上為" & Text5 & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日止　" & stCU15 & "之扣繳資料，請核對；"
'      If Val(Left(strSrvDate(2), Len(strSrvDate(2)) - 4)) > Val(Text5) Then
'         Printer.Print "* 以上為" & Text5 & "年12月31日止　" & stCU15 & "之扣繳資料，請核對；"
'      Else
'         Printer.Print "* 以上為" & Left(strSrvDate(2), Len(strSrvDate(2)) - 4) & "年" & Left(Right(strSrvDate(2), 4), 2) & "月" & Right(strSrvDate(2), 2) & "日止　" & stCU15 & "之扣繳資料，請核對；"
'      End If
'      '2015/1/21 END
      Printer.Print "* 以上為" & strNoteDate & "止　" & stCU15 & "之扣繳資料，請核對；"
      NewLine
      Printer.CurrentX = lngXo + PLeft(2)
      Printer.CurrentY = lngY
      Printer.Print "  若" & Right(strNoteDate, 6) & "後有增加屬於當年度之應扣繳款項，請自行加入並合計。"
      NewLine
      Printer.CurrentX = lngXo + PLeft(2)
      Printer.CurrentY = lngY
      Printer.Print "扣繳核對明細正確請不用回覆，扣單開立完成直接以MAIL提供即可。" 'Add By Sindy 2020/12/30
      '2020/12/28 END
      'Add By Sindy 2016/10/24
      NewLine
      NewLine
      
'      Printer.CurrentX = lngXo
'      Printer.CurrentY = lngY + 600
'      'Modify By Sindy 2020/4/24
'      If strCmp = "1" Then
'         Printer.Print "　　　　　　　　　　　　　　　(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf
'      ElseIf strCmp = "2" Then
'         Printer.Print "　　　　　　　　　　　　　　　(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓" & vbCrLf
'      Else 'L
'         Printer.Print "　　　　　　　　　　　　　　　(統編 77211833) 9A 代號10律師  台北巿長安東路2段110號4樓" & vbCrLf
'      End If
''      Printer.Print "本所有分【專利商標】及【專利法律】 二個扣繳單位：" & vbCrLf & _
''         "　　　　　　　　　　　　　(統編 04150022) 9A 代號91商標代理  台北巿長安東路2段112號10樓" & vbCrLf & _
''         "　　　　　　　　　　　　　(統編 04146457) 9A 代號93專利代理  台北巿長安東路2段112號9樓"
'      '2020/4/24 END
'      '2016/10/24 END
'      '2014/10/21 END
'      'Add By Sindy 2016/12/12 改粗體字
'      Printer.FontBold = True
'      Printer.CurrentX = lngXo
'      Printer.CurrentY = lngY + 600 '825
'      'Modify By Sindy 2020/4/24
'      Printer.Print CompNameQuery(strCmp)
'      '2020/4/24 END
''      Printer.Print "台一國際專利商標事務所"
''      Printer.CurrentX = lngXo
''      Printer.CurrentY = lngY + 1075
''      Printer.Print "台一國際專利法律事務所"
'      Printer.CurrentX = lngXo
'      Printer.CurrentY = lngY + 900 '1325
'      Printer.Print "●扣繳憑單開立完成，煩請郵寄至本所或(傳真02-25011666)或Mail:" & Pub_GetSpecMan("財務處總帳人員") & "@taie.com.tw，感謝您的配合！"
'      Printer.FontBold = False
'      '2016/12/12 END

   End If
   'NewPage 'Remove by Morgan 2010/12/23 空白頁改印第二頁
   
   'Modify By Sindy 2021/12/2
   Dim strName As String, strNo As String, strType As String, strAddr As String
   If strCmp = "L" Then '台一國際法律事務所
      strName = CompNameQuery(strCmp): strNo = "77211833": strType = "9A-10": strAddr = "台北巿中山區朱園里7鄰長安東路2段110號4樓"
   Else '台一國際智慧財產事務所
      strName = CompNameQuery("2"): strNo = "04146457": strType = "9A-93 或 9A-91": strAddr = "台北巿中山區朱園里7鄰長安東路2段112號9樓"
   End If
   Printer.CurrentX = lngXo
   Printer.CurrentY = lngY
   Printer.Print "本所扣繳資訊如下，請開立扣繳憑單"
   Call NewLine
   Printer.CurrentX = lngXo
   Printer.CurrentY = lngY
   Printer.Print "1.所 得 人：" & CompNameQuery(strCmp)
   Call NewLine
   Printer.CurrentX = lngXo
   Printer.CurrentY = lngY
   Printer.Print "2.統一編號：" & strNo
   Call NewLine
   Printer.CurrentX = lngXo
   Printer.CurrentY = lngY
   Printer.Print "3.所得類別：" & strType
   Call NewLine
   Printer.CurrentX = lngXo
   Printer.CurrentY = lngY
   Printer.Print "4.地　　址：" & strAddr
   'Add By Sindy 2021/12/16
   Call NewLine
   Printer.CurrentX = lngXo
   Printer.CurrentY = lngY
   'Modify by Amy 2024/05/17 原:71005
   Printer.Print "5.扣繳憑單請  E-MAIL：" & m_AccMail
   '2021/12/16 END
   Call NewLine
   Call NewLine
   Printer.CurrentX = lngXo
   Printer.CurrentY = lngY
   Printer.Print "感謝您的配合與支持！"
   Exit Sub
   '2021/12/2 END
   
End Sub

'strType（E:Email F:Fax P:紙本）
Private Sub PrintData(p_Rst As ADODB.Recordset, strType As String)
   'Added by Morgan 2011/12/21
   Dim stLstItem As String, stLstRecNo As String
   Dim dblAddAmt(4) As Double
   
   With p_Rst
      '"收款日期"
      Printer.CurrentX = lngXo + PLeft(1)
      Printer.CurrentY = lngY
      Printer.Print Format("" & .Fields("a0l02"), "###/##/##")
      '"收據號碼"
      Printer.CurrentX = lngXo + PLeft(2)
      Printer.CurrentY = lngY
      Printer.Print "" & .Fields("a0k01")
      '"收據日期"
      Printer.CurrentX = lngXo + PLeft(3)
      Printer.CurrentY = lngY
      Printer.Print Format("" & .Fields("a0k02"), "###/##/##")
      
      '"案件性質"
      Printer.CurrentX = lngXo + PLeft(8)
      Printer.CurrentY = lngY
      'Added by Morgan 2011/12/21
      If .Fields("a0k33") = "Y" Then
         Printer.Print convForm(CheckStr("" & .Fields("a0j22")), 12) '"" & .Fields("a0j22")
      Else
      'end 2011/12/21
         'Modified by Morgan 2011/12/27 取消 a0j20
         Printer.Print convForm(CheckStr("" & .Fields("cp10N")), 12) '"" & .Fields("cp10N")
      End If
      '"申請國家"
      Printer.CurrentX = lngXo + PLeft(9)
      Printer.CurrentY = lngY
      'Modified by Morgan 2011/12/30 取消 a0j21
      Printer.Print convForm(CheckStr("" & .Fields("na03")), 12) '"" & .Fields("na03")
      '票期
      Printer.CurrentX = lngXo + PLeft(10)
      Printer.CurrentY = lngY
      Printer.Print Format("" & .Fields("a1p12"), "###/##/##")
      
      'Added by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
      Erase dblAddAmt
      
      dblAddAmt(1) = Val("" & .Fields("Fee1"))
      dblAddAmt(2) = Val("" & .Fields("Fee3")) + Val("" & .Fields("Fee4"))
      dblAddAmt(3) = Val("" & .Fields("Fee5"))
      dblAddAmt(4) = Val("" & .Fields("Fee8"))
            
      stLstRecNo = "" & .Fields("a0k01")
      stLstItem = "" & .Fields("a0j22")
      If .Fields("a0k33") = "Y" Then
         .MoveNext
         Do While Not .EOF
            If stLstRecNo = .Fields("a0k01") And .Fields("a0k33") = "Y" And stLstItem = .Fields("a0j22") Then
               
               .MovePrevious
               For intI = 1 To 5
                  lngSum(intI) = lngSum(intI) + .Fields("Fee" & intI)
               Next
               .MoveNext
               
               dblAddAmt(1) = dblAddAmt(1) + Val("" & .Fields("Fee1"))
               dblAddAmt(2) = dblAddAmt(2) + Val("" & .Fields("Fee3")) + Val("" & .Fields("Fee4"))
               dblAddAmt(3) = dblAddAmt(3) + Val("" & .Fields("Fee5"))
               dblAddAmt(4) = dblAddAmt(4) + Val("" & .Fields("Fee8"))
            Else
               Exit Do
            End If
            .MoveNext
         Loop
         .MovePrevious
      End If
      
      '數字靠右
      '"給付所得額"
      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
      'strExc(1) = Format("" & .Fields("Fee1"), DDollar)
      strExc(1) = Format(dblAddAmt(1), DDollar)
      'Modify by Sindy 2016/12/13  由後往前移
      If (dblAddAmt(4) > 2000 Or Val("" & .Fields("T03")) > 20000) And dblAddAmt(3) > 0 Then
         strExc(1) = "*" & strExc(1)
         bolStarSign = True
      End If
      '2016/12/13 END
      Printer.CurrentX = lngXo + PLeft(5) - 1 * iChrPix - Printer.TextWidth(strExc(1))
      Printer.CurrentY = lngY
      Printer.Print strExc(1)
      
      'Remove by Morgan 2007/9/29
      ''"可扣繳稅額"
      'strExc(1) = Format("" & .Fields("Fee2"), DDollar)
      'Printer.CurrentX = lngXo + PLeft(6) - 1 * iChrPix - Printer.TextWidth(strExc(1))
      'Printer.CurrentY = lngY
      'Printer.Print strExc(1)
      
      '"已扣繳稅額"==>扣繳稅額
      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
      'strExc(1) = Format(Val("" & .Fields("Fee3")) + Val("" & .Fields("Fee4")), DDollar)
      strExc(1) = Format(dblAddAmt(2), DDollar)
      Printer.CurrentX = lngXo + PLeft(7) - 1 * iChrPix - Printer.TextWidth(strExc(1))
      Printer.CurrentY = lngY
      Printer.Print strExc(1)
      
      '"未扣繳稅額"
      'Add by Morgan 2007/9/29 收據未扣繳額大於2000的要加*號
      
'      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'      'strExc(1) = Format("" & .Fields("Fee5"), DDollar)
'      strExc(1) = Format(dblAddAmt(3), DDollar)
'
'      'Modify by Morgan 2010/12/23 不含2000--辜
'      'If Val("" & .Fields("Fee8")) >= 2000 Then
'      'Modified by Morgan 2011/11/15 單次收款總應扣繳額大於 20000 也要
'      'If Val("" & .Fields("Fee8")) > 2000 Then
'      'Modified by Morgan 2011/12/21 若收據有變更帳款類別則相同的依照列印順序合併
'      'If (Val("" & .Fields("Fee8")) > 2000 Or Val("" & .Fields("T03")) > 20000) And Val("" & .Fields("Fee5")) > 0 Then
'      If (dblAddAmt(4) > 2000 Or Val("" & .Fields("T03")) > 20000) And dblAddAmt(3) > 0 Then
'         strExc(1) = "*" & strExc(1)
'         bolStarSign = True
'      End If
'      'end 2007/9/29
'      Printer.CurrentX = lngXo + pLeft(8) - 1 * iChrPix - Printer.TextWidth(strExc(1))
'      Printer.CurrentY = lngY
'      Printer.Print strExc(1)
      
      'Add By Sindy 2015/12/23 記錄執行客戶編號
      strSql = "update ACCTMP44q0 set T06='" & strType & "',T07='" & .Fields("CuNo") & "' where t01='" & .Fields("a0k01") & "' and t05='" & Me.Name & "' and T14='" & strUserNum & "'"
      cnnConnection.Execute strSql
      '2015/12/23 END
   End With
End Sub

Private Sub PrintData_Excel(p_Rst As ADODB.Recordset, ByRef iRow As Integer, Optional ByVal strType As String = "")
   Dim stLstItem As String, stLstRecNo As String
   Dim dblAddAmt(4) As Double
   
   bolWriteData = True 'Add By Sindy 2014/12/9 有寫入資料至Excel
   iRow = iRow + 1
   With p_Rst
      '"收款日期"
      If IsNull("" & .Fields("a0l02").Value) = False Then
         wksAnnuity_A4.Range("A" & iRow).Value = Format("" & .Fields("a0l02"), "###/##/##")
      End If
      '"收據號碼"
      If IsNull("" & .Fields("a0k01").Value) = False Then
         wksAnnuity_A4.Range("B" & iRow).Value = "" & .Fields("a0k01")
      End If
      '"收據日期"
      If IsNull("" & .Fields("a0k02").Value) = False Then
         wksAnnuity_A4.Range("C" & iRow).Value = Format("" & .Fields("a0k02"), "###/##/##")
      End If
      '"案件性質"
      If .Fields("a0k33") = "Y" Then
         wksAnnuity_A4.Range("F" & iRow).Value = convForm(CheckStr("" & .Fields("a0j22")), 12)
      Else
         wksAnnuity_A4.Range("F" & iRow).Value = convForm(CheckStr("" & .Fields("cp10N")), 12)
      End If
      '"申請國家"
      wksAnnuity_A4.Range("G" & iRow).Value = "" & .Fields("na03")
      '票期
      wksAnnuity_A4.Range("H" & iRow).Value = Format("" & .Fields("a1p12"), "###/##/##")
      'Add By Sindy 2016/6/4 + T26
      If "" & .Fields("T26") <> "" And strType = "" Then
         wksAnnuity_A4.Range("H" & iRow).Value = "客戶編號：" & .Fields("T26")
      End If
      '2016/6/4 END
      
      '若收據有變更帳款類別則相同的依照列印順序合併
      Erase dblAddAmt
      
      dblAddAmt(1) = Val("" & .Fields("Fee1"))
      dblAddAmt(2) = Val("" & .Fields("Fee3")) + Val("" & .Fields("Fee4"))
      dblAddAmt(3) = Val("" & .Fields("Fee5"))
      dblAddAmt(4) = Val("" & .Fields("Fee8"))
      
      stLstRecNo = "" & .Fields("a0k01")
      stLstItem = "" & .Fields("a0j22")
      If .Fields("a0k33") = "Y" Then
         .MoveNext
         Do While Not .EOF
            If stLstRecNo = .Fields("a0k01") And .Fields("a0k33") = "Y" And stLstItem = .Fields("a0j22") Then
               
               .MovePrevious
               For intI = 1 To 5
                  lngSum(intI) = lngSum(intI) + .Fields("Fee" & intI)
               Next
               .MoveNext
               
               dblAddAmt(1) = dblAddAmt(1) + Val("" & .Fields("Fee1"))
               dblAddAmt(2) = dblAddAmt(2) + Val("" & .Fields("Fee3")) + Val("" & .Fields("Fee4"))
               dblAddAmt(3) = dblAddAmt(3) + Val("" & .Fields("Fee5"))
               dblAddAmt(4) = dblAddAmt(4) + Val("" & .Fields("Fee8"))
            Else
               Exit Do
            End If
            .MoveNext
         Loop
         .MovePrevious
      End If
      
      '"給付所得額"
      strExc(1) = Format(dblAddAmt(1), DDollar)
      If (dblAddAmt(4) > 2000 Or Val("" & .Fields("T03")) > 20000) And dblAddAmt(3) > 0 Then
         strExc(1) = "*" & strExc(1)
         bolStarSign = True
      End If
      wksAnnuity_A4.Range("D" & iRow).Value = strExc(1)
      '"已扣繳稅額"==>扣繳稅額
      strExc(1) = Format(dblAddAmt(2), DDollar)
      wksAnnuity_A4.Range("E" & iRow).Value = strExc(1)
'      '"未扣繳稅額"
'      strExc(1) = Format(dblAddAmt(3), DDollar)
'      If (dblAddAmt(4) > 2000 Or Val("" & .Fields("T03")) > 20000) And dblAddAmt(3) > 0 Then
'         strExc(1) = "*" & strExc(1)
'         bolStarSign = True
'      End If
'      wksAnnuity_A4.Range("F" & iRow).Value = strExc(1)
'      'strTemp = "F" & iRow & ":F" & iRow
'      wksAnnuity_A4.Range("F" & iRow & ":F" & iRow).Select
      
      'Add By Sindy 2022/4/18 記錄執行客戶編號
      If strType <> "" Then
         strSql = "update ACCTMP44q0 set T06='" & strType & "',T07='" & .Fields("CuNo") & "' where t01='" & .Fields("a0k01") & "' and t05='" & Me.Name & "' and T14='" & strUserNum & "'"
         cnnConnection.Execute strSql
      End If
      '2022/4/18 END
   End With
End Sub

Private Sub GetCompInfo(ByVal p_CompNo As String, ByRef p_CompName As String, ByRef p_UniNo As String)
   strExc(0) = "select a0802,a0807 from acc080 where a0801='" & p_CompNo & "'"
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      p_CompName = "" & RsTemp.Fields("a0802")
      p_UniNo = "" & RsTemp.Fields("a0807")
   End If
   'Modify By Sindy 2020/4/24
   p_CompName = CompNameQuery(p_CompNo)
   '2020/4/24 END
End Sub

Private Sub PrintLine()
   lngY = lngY + iRowPix
   Printer.DrawStyle = vbSolid
   Printer.Line (lngXo, lngY)-(lngXo + 19.5 * 567, lngY)
   lngY = lngY - iRowPix + 60
End Sub

Private Sub PrintSum(p_Sum() As Long)
   PrintLine
   NewLine
   Printer.FontSize = 11
   Printer.FontBold = True
   Printer.CurrentX = lngXo + PLeft(3)
   Printer.CurrentY = lngY
   Printer.Print "合計："
   '"給付所得額"
   strExc(1) = Format(p_Sum(1), DDollar)
   Printer.CurrentX = lngXo + PLeft(5) - 1 * iChrPix - Printer.TextWidth(strExc(1))
   Printer.CurrentY = lngY
   Printer.Print strExc(1)
   
   'Remove by Morgan 2007/9/29
   '"可扣繳稅額"
   'strExc(1) = Format(p_Sum(2), DDollar)
   'Printer.CurrentX = lngXo + PLeft(6) - 1 * iChrPix - Printer.TextWidth(strExc(1))
   'Printer.CurrentY = lngY
   'Printer.Print strExc(1)
   
   '"已扣繳稅額" '扣繳稅額
   strExc(1) = Format(p_Sum(3) + p_Sum(4), DDollar)
   Printer.CurrentX = lngXo + PLeft(7) - 1 * iChrPix - Printer.TextWidth(strExc(1))
   Printer.CurrentY = lngY
   Printer.Print strExc(1)
'   '"未扣繳稅額"
'   strExc(1) = Format(p_Sum(5), DDollar)
'   Printer.CurrentX = lngXo + pLeft(8) - 1 * iChrPix - Printer.TextWidth(strExc(1))
'   Printer.CurrentY = lngY
'   Printer.Print strExc(1)

   Printer.FontBold = False
End Sub

Private Sub PrintSum_Excel(p_Sum() As Long, ByRef iRow As Integer)
   wksAnnuity_A4.Range("A" & iRow & ":H" & iRow).Select
   With wksAnnuity_A4.Application.Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .Weight = xlThin
      .ColorIndex = xlAutomatic
   End With
   iRow = iRow + 1
   wksAnnuity_A4.Range("C" & iRow).Value = "合計："
   '"給付所得額"
   strExc(1) = Format(p_Sum(1), DDollar)
   wksAnnuity_A4.Range("D" & iRow).Value = strExc(1)
   '"已扣繳稅額" '扣繳稅額
   strExc(1) = Format(p_Sum(3) + p_Sum(4), DDollar)
   wksAnnuity_A4.Range("E" & iRow).Value = strExc(1)
'   '"未扣繳稅額"
'   strExc(1) = Format(p_Sum(5), DDollar)
'   wksAnnuity_A4.Range("F" & iRow).Value = strExc(1)
   wksAnnuity_A4.Range("A" & iRow & ":H" & iRow).Select
   With wksAnnuity_A4.Application.Selection.Font
      .Bold = True '粗體
   End With
   wksAnnuity_A4.Range("D6" & ":E" & iRow).Select
   With wksAnnuity_A4.Application.Selection
      .HorizontalAlignment = xlHAlignRight
   End With
End Sub

''Modify By Sindy 2014/10/15 +strAccNote
'Private Function GetAddress(p_CustNo As String, Optional ByRef strAccNote As String) As String
'   ''Modify by Morgan 2007/1/22 客戶地址若客戶狀態有資料時優先抓
'   'Modify By Sindy 2014/10/15 +,CU159
'   strExc(0) = "select NVL(CU80,cu30||cu31),CU159 from customer where cu01='" & Left(p_CustNo, 8) & "' and cu02='" & Mid(p_CustNo, 9) & "'"
'   intI = 1
'   'edit by nickc 2007/02/07 不用 dll 了
'   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      GetAddress = Trim("" & RsTemp.Fields(0))
'      strAccNote = Trim("" & RsTemp.Fields(1)) 'Add By Sindy 2014/10/15
'   End If
'End Function

'信封
Private Sub PrintCover(p_CustAddr As String, p_CustName As String)
   Dim lngX1 As Long, lngY1 As Long, lngX2 As Long, lngY2 As Long
   
   If bol1stPage = False Then
      'Printer.NewPage
      NewPage
   Else
      bol1stPage = False
   End If
   
   lngYo = lngYo - 500 'Add by Morgan 2007/10/1 向上調整，印完調回來，否則封面印下半頁時會超出可印範圍
   
   Printer.FontSize = 14
   Printer.FontName = "標楷體"
   Printer.CurrentX = 750
   Printer.CurrentY = lngYo + 600
   Printer.Print "寄件人：台一國際智慧財產事務所"
   Printer.CurrentX = 750
   Printer.CurrentY = lngYo + 950
   Printer.Print "地　址：台北市長安東路二段１１２號九樓"
   Printer.CurrentX = 750
   Printer.CurrentY = lngYo + 1300
   Printer.Print "電　話：０２－２５０６１０２３"
   
   Printer.FontSize = 18
   Printer.CurrentX = 9800
   Printer.CurrentY = lngYo + 600
   strExc(1) = "平信"
   Printer.Print strExc(1)
   
   lngX1 = 9750
   lngX2 = lngX1 + Printer.TextWidth(strExc(1)) + 100
   lngY1 = lngYo + 550
   lngY2 = lngY1 + Printer.TextHeight(strExc(1)) + 100
   Printer.Line (lngX1, lngY1)-(lngX2, lngY1)
   Printer.Line (lngX1, lngY1)-(lngX1, lngY2)
   Printer.Line (lngX2, lngY1)-(lngX2, lngY2)
   Printer.Line (lngX1, lngY2)-(lngX2, lngY2)
   
   Printer.FontSize = 14
   Printer.CurrentX = 1850
   Printer.CurrentY = lngYo + 3700
   strExc(1) = "收件人："
   Printer.Print strExc(1)
   
   lngX1 = 1850 + Printer.TextWidth(strExc(1))
   lngY1 = lngYo + 3700
   Pub_SmartPrint p_CustAddr, lngX1, lngY1, 125, 350
   
   lngY1 = lngY1 + 400
   Printer.FontSize = 20
   Pub_SmartPrint p_CustName, lngX1, lngY1, 125, 400
      
   Printer.CurrentX = 5000
   Printer.CurrentY = lngYo + 5800
   Printer.Print "會計部"
   
   Printer.FontSize = 14
   Printer.CurrentX = 1800
   Printer.CurrentY = lngYo + 6600
   strExc(1) = " 附註：內附扣繳明細（僅提供核對參考，不作為繳稅之依據）"
   Printer.Print strExc(1)
   
   lngX1 = 1750
   lngX2 = lngX1 + Printer.TextWidth(strExc(1)) + 100
   lngY1 = lngYo + 6550
   lngY2 = lngY1 + Printer.TextHeight(strExc(1)) + 100
   Printer.Line (lngX1, lngY1)-(lngX2, lngY1)
   Printer.Line (lngX1, lngY1)-(lngX1, lngY2)
   Printer.Line (lngX2, lngY1)-(lngX2, lngY2)
   Printer.Line (lngX1, lngY2)-(lngX2, lngY2)
   
   lngYo = lngYo + 500 'Add by Morgan 2007/10/1 向上調整，印完調回來，否則封面印下半頁時會超出可印範圍
   
   'Add by Morgan 2010/12/23 加印空白頁
   NewPage
End Sub

'表頭
Private Sub PrintHead(p_CuNo As String, p_Year As String, p_UniNo As String, p_CompName As String, p_Title As String, _
                      p_Tel As String, p_Fax As String)
   'Modify by Morgan 2006/11/20 改Letter(8.5x11)紙張
   'If lngYo > 0 Then
   '   Printer.NewPage
   '   lngYo = 0
   'Else
   '   lngYo = Printer.Height / 2
   'End If
   If bol1stPage = False Then
      'Modify By Sindy 2014/10/23
      If bolPrintA4 = True Then
      '2014/10/23 END
         Printer.NewPage
         lngYo = 500
      Else
         NewPage
      End If
   Else
      bol1stPage = False
   End If
   'end 2006/11/20
   lngPageNo = lngPageNo + 1
   Printer.FontSize = 16
   
   'Add By Sindy 2014/10/28 請交會計
   If bolPrintA4 = True Then
      Printer.Line (500, 400)-(2200, 900), , B
      Printer.CurrentX = 700
      Printer.CurrentY = 500
      Printer.Print "請交會計"
   End If
   '2014/10/28 END
   
   strExc(1) = ReportTitle(422)
   Printer.CurrentX = lngXo + 2900
   Printer.CurrentY = lngYo
   Printer.Print strExc(1)
   '條件
   Printer.FontSize = 12
   
   'Add By Sindy 2014/10/22
   Printer.CurrentX = lngXo + 9000 'PLeft(9) + 500 '=>8960
   Printer.CurrentY = lngYo
   Printer.Print "電話：" & p_Tel
   Printer.CurrentX = lngXo + 9000 'PLeft(9) + 500
   Printer.CurrentY = lngYo + 250
   Printer.Print "傳真：" & p_Fax
   '2014/10/22 END
   
   strExc(1) = "扣繳編號：" & p_UniNo
   Printer.CurrentX = lngXo + 0
   Printer.CurrentY = lngYo + 500
   Printer.Print strExc(1)
   
   'Modify By Sindy 2021/12/3 Mark
'   Printer.FontBold = True 'Add By Sindy 2016/12/13 粗體
'   Printer.FontSize = 14
'   'strExc(1) = "公司：" & p_CompName
'   strExc(1) = p_CompName
'   Printer.CurrentX = lngXo + 0
'   Printer.CurrentY = lngYo + 750
'   Printer.Print strExc(1)
'   Printer.FontBold = False 'Add By Sindy 2016/12/13
'   Printer.FontSize = 12 'Add By Sindy 2016/12/13
   
   Printer.CurrentX = lngXo + 4300
   Printer.CurrentY = lngYo + 500
   Printer.Print "扣繳年度：" & p_Year
   
   Printer.CurrentX = lngXo + 4300
   Printer.CurrentY = lngYo + 750
   Printer.Print "客戶編號：" & p_CuNo
   
   Printer.CurrentX = lngXo + 4300
   Printer.CurrentY = lngYo + 1000
   Printer.Print "收據抬頭：" & p_Title
   
   strExc(1) = "日期：" & Format(strSrvDate(2), "###/##/##")
   Printer.CurrentX = lngXo + 9000
   Printer.CurrentY = lngYo + 500
   Printer.Print strExc(1)
   
   strExc(1) = "頁次：" & lngPageNo
   Printer.CurrentX = lngXo + 9000
   Printer.CurrentY = lngYo + 750
   Printer.Print strExc(1)
   
   lngY = lngYo + 1350
   Printer.CurrentX = lngXo + PLeft(1)
   Printer.CurrentY = lngY
   Printer.Print "收款日期"
   Printer.CurrentX = lngXo + PLeft(2)
   Printer.CurrentY = lngY
   Printer.Print "收據號碼"
   Printer.CurrentX = lngXo + PLeft(3)
   Printer.CurrentY = lngY
   Printer.Print "收據日期"
   Printer.CurrentX = lngXo + PLeft(5) - 1 * iChrPix - Printer.TextWidth("給付所得額")
   Printer.CurrentY = lngY
   Printer.Print "給付所得額"
   'Remove by Morgan 2007/9/29
   'Printer.CurrentX = lngXo + PLeft(6) - 1 * iChrPix - Printer.TextWidth("可扣繳稅額")
   'Printer.CurrentY = lngY
   'Printer.Print "可扣繳稅額"
   Printer.CurrentX = lngXo + PLeft(7) - 1 * iChrPix - Printer.TextWidth("扣繳稅額")
   Printer.CurrentY = lngY
   Printer.Print "扣繳稅額" '已扣繳稅額
'   Printer.CurrentX = lngXo + pLeft(8) - 1 * iChrPix - Printer.TextWidth("未扣繳稅額")
'   Printer.CurrentY = lngY
'   Printer.CurrentY = lngY
'   Printer.Print "未扣繳稅額"
   Printer.CurrentX = lngXo + PLeft(8)
   Printer.CurrentY = lngY
   Printer.Print "案件性質"
   Printer.CurrentX = lngXo + PLeft(9)
   Printer.CurrentY = lngY
   Printer.Print "申請國家"
   Printer.CurrentX = lngXo + PLeft(10)
   Printer.CurrentY = lngY
   Printer.Print "票期"
   
   PrintLine
End Sub

'表頭
'p_RptType= R.給客戶的報表 E.原就是產生Excel檔
Private Sub PrintHead_Excel(ByVal p_RptType As String, ByVal p_CuNo As String, ByVal p_Year As String, ByVal p_UniNo As String, _
                            ByVal p_CompName As String, ByVal p_Title As String, ByRef iRow As Integer, _
                            ByVal p_Tel As String, ByVal p_Fax As String, ByVal p_Mail As String)
Dim i As Integer, strTemp As String
'Dim stCustAddr As String, strAccNote As String 'Add By Sindy 2014/10/15
   
   'stCustAddr = GetAddress(p_CuNo, strAccNote) 'Add By Sindy 2014/10/15
   
   lngPageNo = lngPageNo + 1
   With wksAnnuity_A4
      If p_RptType = "E" Then
         For i = 1 To 4
            If i = 1 Then
               .Range("E" & iRow).Value = ReportTitle(422)
            ElseIf i = 2 Then
               iRow = iRow + 2
               .Range("A" & iRow).Value = "扣繳編號：" & p_UniNo
               .Range("D" & iRow).Value = "扣繳年度：" & p_Year
               .Range("G" & iRow).Value = "日期：" & Format(strSrvDate(2), "###/##/##")
               .Range("I" & iRow).Value = IIf(p_Mail <> "", "E-Mail：" & p_Mail & vbCrLf, "") & "電話：" & p_Tel & vbCrLf & "傳真：" & p_Fax & vbCrLf & "會計備註：" & IIf(stAccNote <> "", vbCrLf & stAccNote, "") 'Add By Sindy 2014/10/15
            ElseIf i = 3 Then
               iRow = iRow + 1
               '.Range("A" & iRow).Value = "公司：" & p_CompName
               .Range("A" & iRow).Value = p_CompName
               .Range("D" & iRow).Value = "客戶編號：" & p_CuNo
   '            .Range("H" & iRow).Value = "頁次：" & lngPageNo
   '            'Add By Sindy 2014/12/12 EXCEL表中帶出業務區及智權人員(加在客戶編號右邊)
   '            strExc(0) = "select cu01,cu02,cu12,cu13,st02,a0902 from customer,staff,acc090 where cu01=substr('" & p_CuNo & "',1,8) and cu02=substr('" & p_CuNo & "',9)" & _
   '                        " and cu13=st01(+)" & _
   '                        " and cu12=a0901(+)"
   '            intI = 1
   '            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '            If intI = 1 Then
   '               .Range("G" & iRow).Value = "" & RsTemp.Fields("a0902")
   '               .Range("H" & iRow).Value = "" & RsTemp.Fields("st02")
   '            End If
   '            '2014/12/12 END
               'Modify By Sindy 2015/1/15 帶出業務區及智權人員(加在客戶編號右邊)
               .Range("G" & iRow).Value = stSaleArea
               .Range("H" & iRow).Value = stSales
               '2015/1/15 END
            ElseIf i = 4 Then
               iRow = iRow + 1
               .Range("D" & iRow).Value = "收據抬頭：" & p_Title
            End If
            strTemp = "A" & iRow & ":H" & iRow
            .Range(strTemp).Select
            If i = 1 Then
               With .Application.Selection
                  .HorizontalAlignment = xlCenter
                  .Font.Size = 16
               End With
            End If
         Next i
      Else
         For i = 1 To 4
            If i = 1 Then
               .Range("A" & iRow).Value = "請交會計"
               .Range("A1").Select
               With .Application.Selection.Borders(xlEdgeLeft)
                   .LineStyle = xlContinuous
                   .ColorIndex = xlAutomatic
                   .tintandshade = 0
                   .Weight = xlThin
               End With
               With .Application.Selection.Borders(xlEdgeTop)
                   .LineStyle = xlContinuous
                   .ColorIndex = xlAutomatic
                   .tintandshade = 0
                   .Weight = xlThin
               End With
               With .Application.Selection.Borders(xlEdgeBottom)
                   .LineStyle = xlContinuous
                   .ColorIndex = xlAutomatic
                   .tintandshade = 0
                   .Weight = xlThin
               End With
               With .Application.Selection.Borders(xlEdgeRight)
                   .LineStyle = xlContinuous
                   .ColorIndex = xlAutomatic
                   .tintandshade = 0
                   .Weight = xlThin
               End With
               .Range("A" & iRow).Select
               With .Application.Selection
                  .HorizontalAlignment = xlCenter
                  .Font.Size = 13
               End With
               
               .Range("E" & iRow).Value = ReportTitle(422)
               .Range("B" & iRow & ":H" & iRow).Select
               With .Application.Selection
                  .HorizontalAlignment = xlCenter
                  .Font.Size = 16
               End With
               
            ElseIf i = 2 Then
               iRow = iRow + 2
               .Range("A" & iRow).Value = "扣繳編號：" & p_UniNo
               .Range("D" & iRow).Value = "扣繳年度：" & p_Year
               .Range("G" & iRow).Value = "　　電話：" & p_Tel
               
               .Range("G" & iRow & ":H" & iRow).Select
               With .Application.Selection
                  .HorizontalAlignment = xlLeft
                  .MergeCells = True
               End With
            ElseIf i = 3 Then
               iRow = iRow + 1
               .Range("D" & iRow).Value = "客戶編號：" & p_CuNo
               .Range("G" & iRow).Value = "　　傳真：" & p_Fax
               
               .Range("G" & iRow & ":H" & iRow).Select
               With .Application.Selection
                  .HorizontalAlignment = xlLeft
                  .MergeCells = True
               End With
            ElseIf i = 4 Then
               iRow = iRow + 1
               .Range("A" & iRow).Value = "收據抬頭：" & p_Title
               .Range("G" & iRow).Value = "　　日期：" & Format(strSrvDate(2), "###/##/##")
               
               .Range("G" & iRow & ":H" & iRow).Select
               With .Application.Selection
                  .HorizontalAlignment = xlLeft
                  .MergeCells = True
               End With
            End If
         Next i
      End If
      
      '標題
      iRow = iRow + 1
      .Range("A" & iRow).Value = "收款日期"
      .Range("B" & iRow).Value = "收據號碼"
      .Range("C" & iRow).Value = "收據日期"
      .Range("D" & iRow).Value = "給付所得額"
      .Range("E" & iRow).Value = "扣繳稅額" '已扣繳稅額
'      .Range("F" & iRow).Value = "未扣繳稅額"
      .Range("F" & iRow).Value = "案件性質"
      .Range("G" & iRow).Value = "申請國家"
      .Range("H" & iRow).Value = "票期"
      strTemp = "A" & iRow & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
      End With
      .Range("D" & iRow & ":E" & iRow).Select
      With .Application.Selection
         .HorizontalAlignment = xlHAlignRight
      End With
   End With
End Sub

Private Sub NewLine(Optional p_iMarginRows As Integer = 2)
   Dim lngHeight As Long
   
   'Add By Sindy 2014/10/23
   If bolPrintA4 = True Then
      'lngHeight = Printer.ScaleHeight
      lngHeight = 15000
      If lngY + iRowPix > lngHeight Then
         PrintHead stCustNo, Val(Text5), stUniNo, stCompName, stTitle, stTel, stFax
      End If
   Else
   '2014/10/23 END
      If lngYo > 500 Then
         lngHeight = Printer.ScaleHeight
      Else
         lngHeight = Printer.ScaleHeight / 2
      End If
      'Modify by Morgan 2006/11/20 改Letter(8.5x11)紙張
      'If lngY + iRowPix > Printer.ScaleHeight - (Printer.Height / 2 - lngYo) - 2 * iRowPix Then
      If lngY + iRowPix > lngHeight - p_iMarginRows * iRowPix Then
      'end 2006/11/20
         If p_iMarginRows = 2 Then
            PrintLine
         End If
         PrintHead stCustNo, Val(Text5), stUniNo, stCompName, stTitle, stTel, stFax
      End If
   End If
   lngY = lngY + iRowPix
End Sub

Private Sub GetPleft()
   ReDim PLeft(11)
   
   PLeft(0) = 0
   '1 收款日期
   PLeft(1) = 0
   '收據號碼
   PLeft(2) = PLeft(1) + 10 * iChrPix
   '收據日期
   PLeft(3) = PLeft(2) + 10 * iChrPix
   '給付所得額
   PLeft(4) = PLeft(3) + 10 * iChrPix
   '可扣繳稅額
   PLeft(5) = PLeft(4) + 11 * iChrPix 'Modify by Morgan 2007/9/29
   '已扣繳稅額
   PLeft(6) = PLeft(4) + 11 * iChrPix
   '未扣繳稅額
   PLeft(7) = PLeft(6) + 11 * iChrPix
   '案件性質
   PLeft(8) = PLeft(7) + 11 * iChrPix - 1000
   '申請國家
   PLeft(9) = PLeft(8) + 12.5 * iChrPix '13
   '票期
   PLeft(10) = PLeft(9) + 9 * iChrPix '9
End Sub

'Private Function PrintCheck(ByVal p_CustNo As String) As Boolean
'   Dim dblTot As Double
'   If Check1.Value = 1 Then '只印有應稅未扣的 (扣繳 2001 以上)
'      adocheck.MoveFirst
'      adocheck.Find "CuNo='" & p_CustNo & "'", , adSearchForward
'      Do While Not adocheck.EOF
'         If adocheck.Fields("CuNo") <> p_CustNo Then
'            Exit Do
'         Else
'            'Modify By Sindy 2014/9/26
'            'If adocheck.Fields("Fee5") >= 2000 Then
'            If adocheck.Fields("Fee5") >= 2001 Then
'            '2014/9/26 END
'               PrintCheck = True
'               Exit Do
'            End If
'            adocheck.MoveNext
'         End If
'      Loop
'   'Add by Morgan 2009/12/14
'   ElseIf Check2.Value = 1 Then '只印全年服務費超過 20,000 或 服務費 未超過 20,000 但有扣繳的
'      adocheck.MoveFirst
'      adocheck.Find "CuNo='" & p_CustNo & "'", , adSearchForward
'      Do While Not adocheck.EOF
'         If adocheck.Fields("CuNo") <> p_CustNo Then
'            Exit Do
'         Else
'            dblTot = dblTot + Val("" & adocheck.Fields("Fee1"))
'            If adocheck.Fields("Fee6") > 0 Then
'               PrintCheck = True
'               Exit Do
'            'Modified by Morgan 2012/9/14
'            'ElseIf dblTot >= 20000 Then
'            ElseIf dblTot > 20000 Then
'               PrintCheck = True
'               Exit Do
'            End If
'            adocheck.MoveNext
'         End If
'      Loop
'   Else
'      PrintCheck = True
'   End If
'End Function

Private Sub txtType_GotFocus()
   TextInverse txtType
   CloseIme
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub NewPage()
   If lngYo > 500 Then
      Printer.NewPage
      lngYo = 500
   Else
      lngYo = 500 + Printer.Height / 2
   End If
End Sub
