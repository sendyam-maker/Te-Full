VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc11i0 
   AutoRedraw      =   -1  'True
   Caption         =   "退費收訖憑單維護"
   ClientHeight    =   5650
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   7090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5650
   ScaleWidth      =   7090
   Begin VB.TextBox txtPrint 
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
      Left            =   1392
      MaxLength       =   1
      TabIndex        =   10
      Top             =   4850
      Width           =   405
   End
   Begin VB.CheckBox Check2 
      Caption         =   "單張列印"
      Height          =   195
      Left            =   5610
      TabIndex        =   12
      Top             =   5270
      Width           =   1100
   End
   Begin VB.ComboBox cmbPrinter 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1260
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   5210
      Width           =   4110
   End
   Begin VB.TextBox txtA2519 
      Enabled         =   0   'False
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
      Left            =   1620
      MaxLength       =   25
      TabIndex        =   9
      Top             =   3600
      Width           =   5230
   End
   Begin VB.TextBox txtA2518 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4140
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Width           =   1065
   End
   Begin VB.TextBox txtA2512 
      Enabled         =   0   'False
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
      Left            =   6435
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1950
      Width           =   405
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5850
      TabIndex        =   39
      Top             =   90
      Width           =   930
   End
   Begin VB.TextBox txtA2504 
      Enabled         =   0   'False
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
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1950
      Width           =   1605
   End
   Begin VB.TextBox txtA2503 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1290
      Width           =   1605
   End
   Begin VB.TextBox txtA2501 
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
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   0
      Top             =   60
      Width           =   1605
   End
   Begin VB.TextBox txtA2502 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   390
      Width           =   435
   End
   Begin VB.TextBox txtA2505 
      Enabled         =   0   'False
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
      Left            =   1260
      MaxLength       =   15
      TabIndex        =   4
      Top             =   960
      Width           =   1605
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   300
      Left            =   3015
      Picture         =   "Frmacc11i0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   350
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "類別3:銷退付款之款項說明，列印時系統會自動帶出。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   220
      Left            =   1320
      TabIndex        =   50
      Top             =   3300
      Width           =   5470
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1.列印 2.Word檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   1900
      TabIndex        =   49
      Top             =   4850
      Width           =   1600
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "輸出方式："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   230
      TabIndex        =   48
      Top             =   4850
      Width           =   1140
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   230
      TabIndex        =   47
      Top             =   5230
      Width           =   970
   End
   Begin MSForms.TextBox txtA2514 
      Height          =   960
      Left            =   1260
      TabIndex        =   8
      Top             =   2280
      Width           =   5610
      VariousPropertyBits=   -1466941415
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "9895;1693"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA2513 
      Height          =   330
      Left            =   1260
      TabIndex        =   13
      Top             =   1620
      Width           =   5610
      VariousPropertyBits=   671105049
      MaxLength       =   100
      Size            =   "9895;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "   5:翻譯費 6:銷帳轉暫收轉帳 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1710
      TabIndex        =   46
      Top             =   720
      Width           =   2805
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "( 1:暫收款 3:銷退付款 4:銷帳轉暫收退費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1755
      TabIndex        =   45
      Top             =   420
      Width           =   3690
   End
   Begin VB.Label lblAlert 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "只有銷退的暫收款可新增！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   3195
      TabIndex        =   44
      Top             =   1320
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label lblNotice 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "J單號回執單之列印日期將為系統日！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   3195
      TabIndex        =   43
      Top             =   990
      Width           =   3690
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "無法回收說明"
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
      Left            =   230
      TabIndex        =   42
      Top             =   3630
      Width           =   1350
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "日期"
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
      Left            =   3555
      TabIndex        =   41
      Top             =   90
      Width           =   450
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "繳款書份數"
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
      Left            =   5085
      TabIndex        =   40
      Top             =   2010
      Width           =   1140
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "修改人員："
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
      Left            =   230
      TabIndex        =   38
      Top             =   4260
      Width           =   1130
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "修改日期："
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
      Left            =   2610
      TabIndex        =   37
      Top             =   4260
      Width           =   1130
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "修改時間："
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
      Left            =   4820
      TabIndex        =   36
      Top             =   4260
      Width           =   1130
   End
   Begin VB.Label lblA2515 
      BackStyle       =   0  '透明
      Caption         =   "xxx"
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
      Left            =   1440
      TabIndex        =   35
      Top             =   4260
      Width           =   1110
   End
   Begin VB.Label lblA2516 
      BackStyle       =   0  '透明
      Caption         =   "xxx"
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
      Left            =   3690
      TabIndex        =   34
      Top             =   4260
      Width           =   1110
   End
   Begin VB.Label lblA2517 
      BackStyle       =   0  '透明
      Caption         =   "xxx"
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
      Left            =   5940
      TabIndex        =   33
      Top             =   4260
      Width           =   1110
   End
   Begin VB.Label lblA2511 
      BackStyle       =   0  '透明
      Caption         =   "xxx"
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
      Left            =   5940
      TabIndex        =   32
      Top             =   4530
      Width           =   1110
   End
   Begin VB.Label lblA2510 
      BackStyle       =   0  '透明
      Caption         =   "xxx"
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
      Left            =   3690
      TabIndex        =   31
      Top             =   4530
      Width           =   1110
   End
   Begin VB.Label lblA2508 
      BackStyle       =   0  '透明
      Caption         =   "xxx"
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
      Left            =   5940
      TabIndex        =   30
      Top             =   3990
      Width           =   1110
   End
   Begin VB.Label lblA2507 
      BackStyle       =   0  '透明
      Caption         =   "xxx"
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
      Left            =   3690
      TabIndex        =   29
      Top             =   3990
      Width           =   1110
   End
   Begin VB.Label lblA2509 
      BackStyle       =   0  '透明
      Caption         =   "xxx"
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
      Left            =   1440
      TabIndex        =   28
      Top             =   4530
      Width           =   1110
   End
   Begin VB.Label lblA2506 
      BackStyle       =   0  '透明
      Caption         =   "xxx"
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
      Left            =   1440
      TabIndex        =   27
      Top             =   3990
      Width           =   1110
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "建立時間："
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
      Left            =   4820
      TabIndex        =   26
      Top             =   3990
      Width           =   1130
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收回時間："
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
      Left            =   4820
      TabIndex        =   25
      Top             =   4530
      Width           =   1130
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "建立日期："
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
      Left            =   2610
      TabIndex        =   24
      Top             =   3990
      Width           =   1130
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收回日期："
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
      Left            =   2610
      TabIndex        =   23
      Top             =   4530
      Width           =   1130
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "建立人員："
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
      Left            =   230
      TabIndex        =   22
      Top             =   3990
      Width           =   1130
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收回人員："
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
      Left            =   230
      TabIndex        =   21
      Top             =   4530
      Width           =   1130
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "金額"
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
      Left            =   225
      TabIndex        =   20
      Top             =   1950
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "款項說明"
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
      Left            =   225
      TabIndex        =   19
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶抬頭"
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
      Left            =   225
      TabIndex        =   18
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "回執單號"
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
      Left            =   225
      TabIndex        =   17
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      Left            =   225
      TabIndex        =   16
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "回執類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   225
      TabIndex        =   15
      Top             =   420
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "單據編號"
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
      Left            =   225
      TabIndex        =   14
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc11i0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改; 報表2022/4/12 已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Add by Morgan 2007/4/12 只有銷退的暫收款可新增
Option Explicit

Dim strPrinter As String 'Add By Sindy 2022/4/11

Public Function FormSave() As Boolean
   Dim strRetNo As String
   If SaveCheck = False Then
      Exit Function
   End If
On Error GoTo ErrHnd
   adoTaie.BeginTrans
   If strSaveConfirm = MsgText(3) Then
      strRetNo = AutoNo("H", 5, 1)
      strSql = "INSERT INTO ACC250(A2501,A2502,A2503,A2504,A2505,A2506,A2512,A2513,A2514,A2518,A2519)" & _
         " VALUES('" & strRetNo & "','" & txtA2502 & "','" & txtA2503 & "'," & Val(Format(txtA2504)) & ",'" & txtA2505 & "'" & _
         ",'" & strUserNum & "'," & Val(Format(txtA2512)) & ",'" & ChgSQL(txtA2513) & "','" & ChgSQL(txtA2514) & "'" & _
         "," & DBDATE(txtA2518) & ",'" & ChgSQL(txtA2519) & "')"
      adoTaie.Execute strSql
   Else
      strRetNo = txtA2501
      strSql = "UPDATE ACC250 SET A2502='" & txtA2502 & "',A2504=" & Val(Format(txtA2504)) & ",A2505='" & txtA2505 & "'" & _
         ",A2512=" & Val(Format(txtA2512)) & ",A2513='" & ChgSQL(txtA2513) & "',A2514='" & ChgSQL(txtA2514) & "'" & _
         ",A2515='" & strUserNum & "',A2516=" & strSrvDate(1) & ",A2517=TO_CHAR(SYSDATE,'HH24MISS')" & _
         ",A2519='" & ChgSQL(txtA2519) & "' WHERE A2501='" & strRetNo & "'"
      adoTaie.Execute strSql
   End If
   adoTaie.CommitTrans
   FormSave = True
   txtA2501 = strRetNo
   Call QueryData
   Exit Function
   
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Function SaveCheck() As Boolean
   '新增,修改
   If strSaveConfirm = MsgText(3) Then
      'Add by Morgan 2009/5/26
      If txtA2502 <> "4" And txtA2502 <> "6" Then
         MsgBox "只可新增 4 或 6 的回執類別!!"
         txtA2502.SetFocus
         Exit Function
      End If
      
      If txtA2505.Tag <> txtA2505 Then
         MsgBox "暫收單號已更動請重新檢查資料！", vbExclamation
         txtA2505.SetFocus
         Exit Function
      End If
      If Left(txtA2505, 1) <> "J" Then
         MsgBox "請輸入暫收單號！", vbExclamation
         txtA2505.SetFocus
         Exit Function
      ElseIf QueryData1(False) = False Then
         MsgBox "暫收款單號錯誤！", vbExclamation
         Exit Function
      End If
   End If
   If txtA2504 = "" Then
      MsgBox "請輸入金額！", vbExclamation
      txtA2504.SetFocus
      Exit Function
   End If
   
   'Add by Sindy 2021/12/14 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Function
   End If
   
   SaveCheck = True
End Function

'p_iDir: 0=本筆 1=下筆 2=末筆 -1=上筆 -2=首筆
Private Function QueryData(Optional p_iDir As Integer) As Boolean
   
   Select Case p_iDir
      Case 0
         strExc(0) = "select * from acc250 where a2501='" & txtA2501 & "'"
      Case 1
         strExc(0) = "select * from acc250 a where a.a2501=(select min(b.a2501) from acc250 b where b.a2501>'" & txtA2501 & "')"
      Case 2
         strExc(0) = "select * from acc250 a where a.a2501=(select max(b.a2501) from acc250 b)"
      Case -1
         strExc(0) = "select * from acc250 a where a.a2501=(select max(b.a2501) from acc250 b where b.a2501<'" & txtA2501 & "')"
      Case -2
         strExc(0) = "select * from acc250 a where a.a2501=(select min(b.a2501) from acc250 b)"
   End Select
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '已有回執紀錄
   If intI = 1 Then
      With RsTemp
      txtA2501.Text = "" & .Fields("a2501")
      txtA2502.Text = "" & .Fields("a2502")
      'Add by Morgan 2008/1/25
      If txtA2502 = "4" Then
         lblNotice.Visible = True
      End If
      txtA2503.Text = "" & .Fields("a2503")
      txtA2504.Text = Format("" & .Fields("a2504"), "#,##0")
      txtA2505.Text = "" & .Fields("a2505")
      txtA2518.Text = ChangeWStringToTDateString("" & .Fields("a2518"))
      lblA2506.Caption = "" & .Fields("a2506")
      lblA2507.Caption = ChangeWStringToTDateString("" & .Fields("a2507"))
      lblA2508.Caption = Format("" & .Fields("a2508"), Tformat)
      lblA2509.Caption = "" & .Fields("a2509")
      lblA2510.Caption = ChangeWStringToTDateString("" & .Fields("a2510"))
      lblA2511.Caption = Format("" & .Fields("a2511"), Tformat)
      txtA2512.Text = "" & .Fields("a2512")
      txtA2513.Text = "" & .Fields("a2513")
      txtA2514.Text = "" & .Fields("a2514")
      txtA2519.Text = "" & .Fields("a2519")
      lblA2515.Caption = "" & .Fields("a2515")
      lblA2516.Caption = ChangeWStringToTDateString("" & .Fields("a2516"))
      lblA2517.Caption = Format("" & .Fields("a2517"), Tformat)
      End With
      txtA2501.Tag = txtA2501
      QueryData = True
   End If
End Function

Private Function QueryData1(Optional p_bolSetData As Boolean = True) As Boolean
   strExc(0) = "select * from acc0t0,acc0s0,acc0k0,acc250 where a0t01='" & txtA2505 & "' and a0s01(+)=a0t07 and a0k01(+)=a0s02 and a2505(+)=a0t01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      If strSaveConfirm = MsgText(3) And Not IsNull(.Fields("a2501")) Then
         MsgBox "該暫收款單號已有回執單不可再新增！"
         Exit Function
      End If
      If p_bolSetData = True Then
         'Remove by Morgan 2009/5/26 改可人工輸入
         'txtA2502.Text = "4"
         txtA2503.Text = "" & .Fields("a0k03")
         txtA2504.Text = Val("" & .Fields("a0s05"))
         txtA2513.Text = "" & .Fields("a0k04")
         '新增且有收據編號
         If Not IsNull(.Fields("a0k01")) Then
            txtA2514.Text = "退回 " & PUB_GetCaseInfo("" & .Fields("a0k01")) & " 款項"
            'Add by Morgan 2009/5/26
            If txtA2502 = "6" Then
               txtA2514.Text = txtA2514.Text & ",轉辦他案"
            End If
         End If
         txtA2518.Text = ChangeTStringToTDateString("" & .Fields("a0t03"))
      End If
      txtA2505.Tag = txtA2505
      End With
      QueryData1 = True
   End If
End Function

Private Sub cmdPrint_Click()
   If txtA2501.Text = "" Then
      MsgBox "請輸入回執單號！", vbExclamation
   ElseIf txtA2501.Text <> txtA2501.Tag Then
      MsgBox "單號已改請重新查詢！", vbExclamation
   'Added by Lydia 2024/03/12
   ElseIf txtPrint <> "1" And txtPrint <> "2" Then
      MsgBox "請輸入方式1, 2 ！", vbExclamation
      txtPrint.SetFocus
      txtPrint_GotFocus
   'end 2024/03/12
   Else
      PrintSheet
   End If
End Sub

Public Sub PrintSheet()
Dim stChoice As String
Dim p_adoquery1 As New ADODB.Recordset
Dim p_lngYo As Long, p_lngPageNo As Long
'Added by Lydia 2024/03/12
Dim strFileName As String
Dim hLocalFile As Long
   
   strExc(0) = "select * from acc250 where a2501='" & txtA2501.Text & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Added by Lydia 2024/03/12
      If txtPrint = "2" Then '產生Word檔並且開啟
         Call Pub_ChkExcelPath(strFileName)
         strFileName = strFileName & "\" & Trim(txtA2501) & "_退費收訖憑單.doc"
         If Dir(strFileName) <> "" Then
           If PUB_ChkFileOpening(strFileName, True) = True Then
              Exit Sub
           End If
           If PUB_DelPCOrgFile(strFileName) = False Then
              Exit Sub
           End If
         End If
      Else
         strFileName = ""
      End If
      'end 2024/03/12
      Printer.EndDoc
      PUB_SetOsDefaultPrinter cmbPrinter 'Add by Sindy 2022/4/12
      'Modified by Lydia 2024/03/12 +strFileName存檔位置
      PUB_PrintReceipt RsTemp, 0, 0, , , IIf(Check2.Value = 1, True, False), strFileName
      Printer.EndDoc
      PUB_SetOsDefaultPrinter strPrinter 'Add by Sindy 2022/4/12
      
      'Added by Lydia 2024/03/12
      If Len(strFileName) > 2 Then
         ShellExecute hLocalFile, "open", strFileName, vbNullString, vbNullString, 1
      End If
      'end 2024/03/12
   End If
End Sub

Public Sub cmdSearch_Click()
   If txtA2501.Text = "" Then
      MsgBox "請輸入回執單號！", vbExclamation
   Else
      FormClear True
      If QueryData = False Then
         MsgBox "回執單號不存在！", vbExclamation
         Exit Sub
      End If
   End If
End Sub

Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
   '表單初始化
   'Modify by Amy 2023/10/06 原H:5390
   'Modified by Lydia 2024/03/12 H:5600 >> 5760
   'modify by sonia 2024/6/13 H:5760->6210
   PUB_InitForm Me, 7215, 6210, strBackPicPath1
   
   '畫面初值設定
   FormClear
   MoveLast
   
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter 'Add By Sindy 2022/4/11
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2022/4/11
   If cmbPrinter.Text <> cmbPrinter.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   '2022/4/11 END
   
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc11i0 = Nothing
End Sub

Public Sub FormClear(Optional p_bolSearch As Boolean = False)
   If p_bolSearch = False Then
      txtA2501.Text = ""
   End If
   txtA2502.Text = ""
   txtA2503.Text = ""
   txtA2504.Text = ""
   txtA2505.Text = ""
   txtA2513.Text = ""
   txtA2514.Text = ""
   txtA2518.Text = ""
   txtA2519.Text = ""
   lblA2506.Caption = ""
   lblA2507.Caption = ""
   lblA2508.Caption = ""
   lblA2509.Caption = ""
   lblA2510.Caption = ""
   lblA2511.Caption = ""
   lblA2515.Caption = ""
   lblA2516.Caption = ""
   lblA2517.Caption = ""
   lblNotice.Visible = False 'Add by Morgan 2008/1/25
   txtPrint = "" 'Added by Lydia 2024/03/12
End Sub

Public Sub FormEnable()
   Dim bolEnable As Boolean
   
   If strSaveConfirm = MsgText(3) Then
      lblAlert.Visible = True
   Else
      lblAlert.Visible = False
   End If
   
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      bolEnable = True
   Else
      bolEnable = False
   End If
   cmdSearch.Enabled = Not bolEnable
   txtA2501.Enabled = Not bolEnable
   txtA2504.Enabled = bolEnable
   txtA2512.Enabled = bolEnable
   txtA2502.Enabled = bolEnable
   If strSaveConfirm = MsgText(3) Then
      txtA2505.Enabled = bolEnable
   Else
      txtA2505.Enabled = False
   End If
   txtA2513.Enabled = bolEnable
   txtA2514.Enabled = bolEnable
   If strSaveConfirm = MsgText(3) Then
      txtA2502.SetFocus
   ElseIf strSaveConfirm = MsgText(4) Then
      txtA2504.SetFocus
   Else
      txtA2501.SetFocus
   End If
   txtA2519.Enabled = bolEnable 'Add By Sindy 2021/12/14
End Sub

Public Function FormCheck() As Boolean
   If txtA2501.Text = "" Then
      FormCheck = False
      MsgBox "請先查詢回執單號！", vbCritical
      Exit Function
   ElseIf txtA2501.Tag <> txtA2501.Text Then
      FormCheck = False
      MsgBox "回執單號已更動，請重新查詢！", vbCritical
      Exit Function
   End If
   If lblA2510 <> "" Then
      FormCheck = False
      MsgBox "已收回不可修改或刪除！", vbCritical
      Exit Function
   End If
   FormCheck = True
End Function

Public Function FormDelete() As Boolean
   If FormCheck = False Then
      Exit Function
   End If
On Error GoTo ErrHnd
   strSql = "delete from acc250 where a2501='" & txtA2501 & "'"
   adoTaie.Execute strSql
   FormDelete = True
   
   If QueryData(1) = False Then
      If QueryData(2) = False Then
         FormClear
         MsgBox "資料庫已無資料！", vbInformation
      End If
   End If
   
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub txtA2501_GotFocus()
   TextInverse txtA2501
   CloseIme
End Sub

'Added by Lydia 2024/03/12
Private Sub txtA2501_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA2502_GotFocus()
   TextInverse txtA2502
   CloseIme
End Sub

Private Sub txtA2502_KeyPress(KeyAscii As Integer)
   If strSaveConfirm = MsgText(3) Then
      If KeyAscii <> 8 And KeyAscii <> Asc("4") And KeyAscii <> Asc("6") Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtA2503_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA2504_GotFocus()
   TextInverse txtA2504
   CloseIme
End Sub

Private Sub txtA2505_GotFocus()
   If strSaveConfirm = MsgText(3) Then
      If Left(txtA2505, 1) = "J" Then
         txtA2505.SelStart = 1
         txtA2505.SelLength = Len(txtA2505) - 1
      End If
   End If
   CloseIme
End Sub

Private Sub txtA2505_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtA2505_Validate(Cancel As Boolean)
   If strSaveConfirm = MsgText(3) Then
      '暫收款
      If Left(txtA2505.Text, 1) = "J" And txtA2505 <> "J" Then
         If QueryData1 = False Then
            Cancel = True
         End If
       End If
   End If
End Sub

Private Sub txtA2512_GotFocus()
   TextInverse txtA2512
   CloseIme
End Sub

Private Sub txtA2513_GotFocus()
   TextInverse txtA2513
   OpenIme
End Sub

Private Sub txtA2513_Validate(Cancel As Boolean)
   Cancel = Not CheckLengthIsOK(txtA2513, 200)
End Sub

Private Sub txtA2514_GotFocus()
   TextInverse txtA2514
   OpenIme
End Sub

Public Sub FormRequery()
   If txtA2501.Tag <> "" Then
      txtA2501.Text = txtA2501.Tag
      cmdSearch_Click
   End If
End Sub

Public Sub MoveNext()
   If txtA2501.Tag = "" Then
      MsgBox "尚未有查詢紀錄，無法搜尋上下筆！"
   Else
      txtA2501.Text = txtA2501.Tag
      If QueryData(1) = False Then
         MsgBox "已經是最後一筆！"
      End If
   End If
End Sub

Public Sub MovePrevious()
   If txtA2501.Tag = "" Then
      MsgBox "尚未有查詢紀錄，無法搜尋上下筆！"
   Else
      txtA2501.Text = txtA2501.Tag
      If QueryData(-1) = False Then
         MsgBox "已經是第一筆！"
      End If
   End If
End Sub

Public Sub MoveFirst()
   If QueryData(-2) = False Then
      MsgBox "查無資料！"
   End If
End Sub

Public Sub MoveLast()
   If QueryData(2) = False Then
      MsgBox "查無資料！"
   End If
End Sub

Private Sub txtA2514_Validate(Cancel As Boolean)
   Cancel = Not CheckLengthIsOK(txtA2514, 200)
End Sub

Private Sub txtA2519_GotFocus()
   TextInverse txtA2519
   OpenIme
End Sub

'Added by Lydia 2024/03/12
Private Sub txtPrint_GotFocus()
   TextInverse txtPrint
End Sub

'Added by Lydia 2024/03/12
Private Sub txtPrint_Validate(Cancel As Boolean)
   If txtPrint <> "" Then
      If txtPrint <> "1" And txtPrint <> "2" Then
         MsgBox "請輸入方式1,2 ！", vbExclamation
         Cancel = True
         txtPrint.SetFocus
         txtPrint_GotFocus
         Exit Sub
      End If
   End If
End Sub
