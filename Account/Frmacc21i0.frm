VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc21i0 
   AutoRedraw      =   -1  'True
   Caption         =   "折讓輸入"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   8760
   Begin VB.TextBox Text23 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3576
      TabIndex        =   40
      Top             =   2760
      Width           =   1572
   End
   Begin VB.TextBox Text22 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2550
      TabIndex        =   38
      Top             =   3534
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.TextBox Text18 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2550
      TabIndex        =   36
      Top             =   3900
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2550
      TabIndex        =   34
      Top             =   3168
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.TextBox Text19 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      TabIndex        =   32
      Top             =   3534
      Width           =   1572
   End
   Begin VB.TextBox Text21 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      TabIndex        =   30
      Top             =   3168
      Width           =   1572
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      TabIndex        =   4
      Top             =   2760
      Width           =   1572
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      TabIndex        =   27
      Top             =   1320
      Width           =   1572
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4080
      TabIndex        =   25
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   2640
      Picture         =   "Frmacc21i0.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   255
      Width           =   350
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      TabIndex        =   24
      Top             =   1680
      Width           =   372
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   23
      Top             =   1680
      Width           =   252
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   22
      Top             =   1680
      Width           =   852
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   21
      Top             =   2040
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      MaxLength       =   14
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   9
      Top             =   1680
      Width           =   492
   End
   Begin VB.TextBox Text20 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6840
      TabIndex        =   8
      Top             =   960
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   4080
      TabIndex        =   5
      Top             =   240
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.TextBox Text11 
      Height          =   330
      Left            =   2910
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2040
      Width           =   5535
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "9763;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text8 
      Height          =   330
      Left            =   3270
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1680
      Width           =   5175
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "9128;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text3 
      Height          =   330
      Left            =   4080
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   600
      Width           =   4335
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "7646;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text4 
      Height          =   330
      Left            =   6840
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   240
      Width           =   1575
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "2778;582"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "規費折讓金額(台幣)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   1512
      TabIndex        =   41
      Top             =   2817
      Width           =   1992
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "折讓金額(美金)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   996
      TabIndex        =   39
      Top             =   3573
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "折讓後金額(美金)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   756
      TabIndex        =   37
      Top             =   3939
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "請款金額(美金)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   996
      TabIndex        =   35
      Top             =   3207
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   4920
      X2              =   8400
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "折讓後金額(台幣)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   33
      Top             =   3572
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "折讓後金額(外幣)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   31
      Top             =   3206
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "折讓金額(台幣)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   29
      Top             =   2798
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "請款金額(台幣)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   28
      Top             =   1358
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "匯率"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   998
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "申請人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   2078
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   4260
      Left            =   210
      Top             =   90
      Width           =   8370
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "折讓日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   2438
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "折讓金額(外幣)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   2438
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   1718
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "請款金額(外幣)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   998
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "幣別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   998
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "姓名(日)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   278
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "姓名(英)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3120
      TabIndex        =   13
      Top             =   639
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   12
      Top             =   639
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "請款日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3120
      TabIndex        =   11
      Top             =   279
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   10
      Top             =   279
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc21i0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/02 改成Form2.0 ; Text3、Text4、Text8、Text11
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc1k0 As New ADODB.Recordset
Dim NTRate As Double   '2009/4/24 ADD BY SONIA
Dim USRate As Double   '2009/4/24 ADD BY SONIA

Private Sub Command5_Click()
   Acc1k0Refresh
   If adoacc1k0.RecordCount <> 0 Then
      FormShow
      RecordShow
   End If
End Sub

Private Sub Command5_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command5_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
'edit by nickc 2007/02/08
'   '93.3.16 ADD BY SONIA
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
'   '93.3.16 END
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
   strFormName = Name
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   If adoacc1k0.RecordCount <> 0 Then
      adoacc1k0.MoveFirst
   End If
   'adoacc1k0.Find "a1k01 = '" & strItemNo & "'", 0, adSearchForward, 1
   'If adoacc1k0.EOF = False Then
   '   FormShow
   '   RecordShow
   'End If
   Text1 = strItemNo
   Acc1k0Refresh
   If adoacc1k0.RecordCount <> 0 Then
      FormShow
      RecordShow
   End If
   strItemNo = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 4900  'Modified by Lydia 2021/12/02 Height 4700=>4900
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath1)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   If adoacc1k0.RecordCount <> 0 Then
      adoacc1k0.MoveLast
      adoacc1k0.MoveFirst
      RecordShow
   End If
   MaskEdBox2.Mask = DFormat
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc21i0 = Nothing
End Sub

Private Sub MaskEdBox2_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label10 & MsgText(52), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label10 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text10_Change()
   If Text10 = MsgText(601) Then
      Exit Sub
   End If
   Text11 = CustomerQuery(Text10, 1)
   If Text11 = MsgText(601) Then
      Text11 = CustomerQuery(Text10, 2)
      If Text11 = MsgText(601) Then
         Text11 = CustomerQuery(Text10, 3)
      End If
   End If
End Sub

Private Sub Text12_Change()
   CaseQuery
End Sub

Private Sub Text13_Change()
   CaseQuery
End Sub

Private Sub Text14_Change()
   CaseQuery
End Sub

Private Sub Text16_Change()
   '2009/4/26 cancel BY SONIA
   'Text17 = Val(Text9) * Val(Text15)
   'Text18 = Val(Text6) - Val(Text9)
   'Text19 = Val(Text16) - Val(Text17)
   '2009/4/26 END
End Sub

Private Sub Text2_Change()
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   Text3 = FagentQuery(Text2, 2)
   Text4 = FagentQuery(Text2, 1)
End Sub

Private Sub Text7_Change()
   CaseQuery
End Sub

'Add By Sindy 2009/07/15
Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

'Add By Sindy 2009/07/15
Public Sub Text7_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(Text7) = False Then
      ' 檢查系統類別
      If IsCorrectSysKind(Text7) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text7_GotFocus
         GoTo EXITSUB
      End If
      ' 檢查使用者權限
      If IsUserHasRightOfSystem(strUserNum, Text7) = False Then
         '2009/8/21 ADD BY SONIA FCP程序可輸入P及CFP請款單
         If Not (GetStaffDepartment(strUserNum) = "F22" And (Text7 = "P" Or Text7 = "PS" Or Text7 = "CFP" Or Text7 = "CPS")) Then
         '2009/8/21 END
            Cancel = True
            strTit = "資料檢核"
            strMsg = "您沒有使用該系統類別的權限"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Text7_GotFocus
            GoTo EXITSUB
         End If
      End If
   Else
'      Cancel = True
'      strTit = "資料檢核"
'      strMsg = "本所案號中的系統別不可空白"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      Text7_GotFocus
'      GoTo EXITSUB
   End If
EXITSUB:
End Sub

'Private Sub Text9_Change()
Private Sub Text9_LostFocus()
   '2009/4/24 MODIFY BY SONIA
   'Text17 = Val(Text9) * Val(Text15)
   'Text18 = Val(Text6) - Val(Text9)
   Text17 = Round(Val(Text9) * Val(NTRate), 0) 'Modify By Sindy 2012/11/29 改可以四捨五入
   Text21 = Val(Text20) - Val(Text9)
'   Text22 = Format(Val(Text9) * USRate, FAmount) 'Modify By Sindy 2012/11/29 Mark
'   Text18 = Val(Text6) - Val(Text22) 'Modify By Sindy 2012/11/29 Mark
   '2009/4/24 END
   Text19 = Val(Text16) - Val(Text17)
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'Add By Sindy 2012/11/29
Private Sub Text17_LostFocus()
   Text9 = Format(((Val(Text17) * 100) / NTRate) / 100, FAmount)
   Text21 = Val(Text20) - Val(Text9)
   Text19 = Val(Text16) - Val(Text17)
End Sub

'Add By Sindy 2012/11/29
Private Sub Text17_GotFocus()
   TextInverse Text17
End Sub

'Add By Sindy 2012/11/29
Private Sub Text17_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.MaxRecords = intMax
   adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & Text1 & "' order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
Dim intFee As Integer 'Add By Sindy 2010/8/25
   Text1 = adoacc1k0.Fields("a1k01").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc1k0.Fields("a1k02").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc1k0.Fields("a1k02").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc1k0.Fields("a1k03").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc1k0.Fields("a1k03").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k18").Value) Then
      Text5 = MsgText(601)
   Else
      Text5 = adoacc1k0.Fields("a1k18").Value
   End If
   If IsNull(adoacc1k0.Fields("a1k10").Value) Then
      Text15 = MsgText(601)
   Else
      Text15 = adoacc1k0.Fields("a1k10").Value
   End If
   'Modify By Sindy 2012/11/29
'   If IsNull(adoacc1k0.Fields("a1k08").Value) Then
'      Text6 = MsgText(601)
'   Else
'      Text6 = adoacc1k0.Fields("a1k08").Value
'   End If
   If IsNull(adoacc1k0.Fields("a1k08").Value) Then
      Text20 = MsgText(601)
   Else
      Text20 = adoacc1k0.Fields("a1k08").Value
   End If
   '2012/11/29 End
   If IsNull(adoacc1k0.Fields("a1k11").Value) Then
      Text16 = MsgText(601)
   Else
      Text16 = adoacc1k0.Fields("a1k11").Value
   End If
   Text7 = adoacc1k0.Fields("a1k13").Value
   Text12 = adoacc1k0.Fields("a1k14").Value
   Text13 = adoacc1k0.Fields("a1k15").Value
   Text14 = adoacc1k0.Fields("a1k16").Value
   Text10 = CaseCustShow(Text7, Text12, Text13, Text14, 1)
   'Modify By Sindy 2012/11/29
'   If IsNull(adoacc1k0.Fields("a1k06").Value) Then
'      '2009/4/24 MODIFY BY SONIA
'      'Text9 = MsgText(601)
'      Text22 = MsgText(601)
'   Else
'      '2009/4/24 MODIFY BY SONIA
'      'Text9 = adoacc1k0.Fields("a1k06").Value
'      Text22 = adoacc1k0.Fields("a1k06").Value
'   End If
   If IsNull(adoacc1k0.Fields("a1k06").Value) Then
      Text17 = MsgText(601)
   Else
      Text17 = adoacc1k0.Fields("a1k06").Value
   End If
   '2012/11/29 End
   MaskEdBox2.Mask = MsgText(601)
   If IsNull(adoacc1k0.Fields("a1k07").Value) Then
      MaskEdBox2.Text = MsgText(601)
   Else
      MaskEdBox2.Text = CFDate(adoacc1k0.Fields("a1k07").Value)
   End If
   MaskEdBox2.Mask = DFormat
   '2009/4/24 ADD BY SONIA
   If IsNull(adoacc1k0.Fields("a1k31").Value) Then
      Text9 = MsgText(601)
   Else
      Text9 = adoacc1k0.Fields("a1k31").Value
   End If
   Text23 = "" & adoacc1k0.Fields("a1k36").Value 'Added by Morgan 2018/3/20
   NTRate = PUB_GetUSXRate_1(Replace(MaskEdBox1.Text, "/", ""), Text5.Text)
   '2010/3/3 MODIFY BY SONIA X09803810溢位
   'Text20 = Format(((Text16 * 100 * 100) \ (NTRate * 100)) / 100, FAmount)
   '2010/12/7 MODIFY BY SONIA X09812517溢位
   'intFee = ((Val(Text16) * 100) / NTRate) 'Modify By Sindy 2010/8/25 ((Text16 * 100) \ NTRate)
   'Text20 = Format(intFee / 100, FAmount)
   'Text20 = Format(((Val(Text16) * 100) / NTRate) / 100, FAmount) 'Modify By Sindy 2012/11/29 Mark
   '2010/12/7 END
   '2010/3/3 END
   'If Text5 = "USD" Then Text20 = Text6 'Modify By Sindy 2012/11/29 Mark
   '抓請款幣別對美金匯率
   USRate = PUB_GetDNRate(Replace(MaskEdBox1.Text, "/", ""), Text5.Text)
   'Text17 = Val(Text9) * Val(NTRate) 'Modify By Sindy 2012/11/29 Mark
   Text21 = Val(Text20) - Val(Text9)
'   Text22 = Format(Val(Text9) * USRate, FAmount) 'Modify By Sindy 2012/11/29 Mark
'   Text18 = Val(Text6) - Val(Text22) 'Modify By Sindy 2012/11/29 Mark
   Text19 = Val(Text16) - Val(Text17)
   '2009/4/24 END
   '2010/2/23 ADD BY SONIA 已收款不可輸折讓 X09803813
   If adoacc1k0.Fields("a1k29").Value = "Y" Then
      tool15_enabled
   Else
      tool8_enabled
   End If
   '2010/2/23 END
End Sub

'*************************************************
'  顯示查詢資料(案件基本資料)
'
'*************************************************
Private Sub CaseQuery()
   If Text7 = MsgText(601) Or Text12 = MsgText(601) Or Text13 = MsgText(601) Or Text14 = MsgText(601) Then
      Exit Sub
   End If
   Text8 = CaseNameShow(Text7, Text12, Text13, Text14, 1)
   If Text8 = MsgText(601) Then
      Text8 = CaseNameShow(Text7, Text12, Text13, Text14, 2)
      If Text8 = MsgText(601) Then
         Text8 = CaseNameShow(Text7, Text12, Text13, Text14, 3)
      End If
   End If
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   If adoacc1k0.RecordCount = 0 Then
      Exit Sub
   End If
   CountShow adoacc1k0.Bookmark, adoacc1k0.RecordCount
End Sub

'*************************************************
'  重新整理國外請款資料
'
'*************************************************
Public Sub Acc1k0Refresh()
On Error GoTo Checking
   If adoacc1k0.State = adStateOpen Then
      adoacc1k0.Close
   End If
   adoacc1k0.CursorLocation = adUseClient
   adoacc1k0.MaxRecords = intMax
   adoacc1k0.Open "select * from acc1k0 where a1k01 >= '" & Text1 & "' order by a1k01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

