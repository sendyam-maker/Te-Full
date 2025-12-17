VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1224 
   AutoRedraw      =   -1  'True
   Caption         =   "銷退資料查詢"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   9405
   Begin VB.TextBox Text1 
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
      Height          =   300
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   51
      Top             =   240
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
      Height          =   300
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   21
      Top             =   720
      Width           =   1572
   End
   Begin VB.TextBox Text3 
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
      Height          =   300
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   20
      Top             =   1080
      Width           =   612
   End
   Begin VB.TextBox Text4 
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
      Height          =   300
      Left            =   1560
      TabIndex        =   19
      Top             =   1920
      Width           =   1572
   End
   Begin MSForms.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False     
      Height          =   300
      Left            =   4560
      TabIndex        =   18
      Top             =   1920
      Width           =   4692
      VariousPropertyBits=   679493661
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   1560
      TabIndex        =   17
      Top             =   2280
      Width           =   1572
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   4560
      TabIndex        =   16
      Top             =   2280
      Width           =   1572
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   1560
      TabIndex        =   15
      Top             =   2640
      Width           =   1572
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   4560
      TabIndex        =   14
      Top             =   2640
      Width           =   1572
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   1560
      TabIndex        =   13
      Top             =   3000
      Width           =   1572
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   4590
      TabIndex        =   12
      Top             =   3000
      Width           =   1572
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   7680
      TabIndex        =   11
      Top             =   2640
      Width           =   1572
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   7680
      TabIndex        =   10
      Top             =   3000
      Width           =   1572
   End
   Begin VB.TextBox Text14 
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
      Height          =   300
      Left            =   7680
      MaxLength       =   14
      TabIndex        =   9
      Top             =   3360
      Width           =   1572
   End
   Begin VB.TextBox Text15 
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
      Height          =   300
      Left            =   1560
      MaxLength       =   14
      TabIndex        =   8
      Top             =   3720
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
      Height          =   300
      Left            =   4560
      MaxLength       =   14
      TabIndex        =   7
      Top             =   3720
      Width           =   1572
   End
   Begin VB.TextBox Text17 
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
      Height          =   300
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1440
      Width           =   612
   End
   Begin VB.TextBox Text21 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   7680
      TabIndex        =   5
      Top             =   4560
      Width           =   1572
   End
   Begin VB.TextBox Text22 
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
      Height          =   300
      Left            =   7680
      TabIndex        =   4
      Top             =   720
      Width           =   1572
   End
   Begin VB.TextBox Text24 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
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
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   3360
      Width           =   1572
   End
   Begin VB.TextBox Text25 
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
      Height          =   300
      Left            =   4560
      TabIndex        =   2
      Top             =   3360
      Width           =   852
   End
   Begin VB.TextBox Text26 
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
      Height          =   300
      Left            =   7680
      MaxLength       =   14
      TabIndex        =   1
      Top             =   3720
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4560
      TabIndex        =   22
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
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
      Height          =   300
      Left            =   7680
      TabIndex        =   23
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   7680
      TabIndex        =   24
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
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
   Begin MSForms.TextBox Text27 
      Height          =   300
      Left            =   6240
      TabIndex        =   0
      Top             =   1080
      Width           =   3012
      VariousPropertyBits=   -1467989987
      BackColor       =   14737632
      MaxLength       =   35
      ScrollBars      =   2
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收據單號"
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
      Left            =   240
      TabIndex        =   52
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "銷退單號"
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
      Left            =   240
      TabIndex        =   50
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銷退日期"
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
      Left            =   3480
      TabIndex        =   49
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "類別"
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
      Left            =   240
      TabIndex        =   48
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "(1.銷帳 2.退費 3.銷帳+退費)"
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
      Left            =   2280
      TabIndex        =   47
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   46
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "收文日期"
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
      Left            =   6720
      TabIndex        =   45
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "案件名稱"
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
      Left            =   3480
      TabIndex        =   44
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "應收服務費"
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
      Left            =   240
      TabIndex        =   43
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "應收規費"
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
      Left            =   3480
      TabIndex        =   42
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "已收服務費"
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
      Left            =   240
      TabIndex        =   41
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "已收規費"
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
      Left            =   3480
      TabIndex        =   40
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "已退服務費"
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
      Left            =   240
      TabIndex        =   39
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "已退規費"
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
      Left            =   3480
      TabIndex        =   38
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "已銷金額"
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
      Left            =   6720
      TabIndex        =   37
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "銷帳金額"
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
      Left            =   6720
      TabIndex        =   36
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "未收金額"
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
      Left            =   6720
      TabIndex        =   35
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "退費服務費"
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
      Left            =   240
      TabIndex        =   34
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "退費規費"
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
      Left            =   3480
      TabIndex        =   33
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "退費方式"
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
      Left            =   240
      TabIndex        =   32
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "(1.轉應付款 2.轉暫收款)"
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
      Left            =   2280
      TabIndex        =   31
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "欲處理日期"
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
      Left            =   6480
      TabIndex        =   30
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1215
      Left            =   120
      Top             =   600
      Width           =   9255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label29 
      BackStyle       =   0  '透明
      Caption         =   "轉出單據編號"
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
      Left            =   6240
      TabIndex        =   29
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label31 
      BackStyle       =   0  '透明
      Caption         =   "扣繳金額"
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
      Left            =   240
      TabIndex        =   28
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label32 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
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
      Left            =   3480
      TabIndex        =   27
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label33 
      BackStyle       =   0  '透明
      Caption         =   "稅款退費金額"
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
      Left            =   6240
      TabIndex        =   26
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label23 
      BackStyle       =   0  '透明
      Caption         =   "備註"
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
      Left            =   5760
      TabIndex        =   25
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "Frmacc1224"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/08 Form2.0已修改 Text5/Text27
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit
Public adoacc0k0 As New ADODB.Recordset
Public adoacc0s0 As New ADODB.Recordset
Public adopatent As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset

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
   Me.Width = 9500
   Me.Height = 4700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   FormShowE
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strCon1 = ""
   StatusClear
   tool3_enabled
   Select Case strFormLink
      Case "Frmacc1220"
         Frmacc1220.Enabled = True
      Case "Frmacc1230"
         Frmacc1230.Enabled = True
      Case "Frmacc1240"
         Frmacc1240.Enabled = True
      'Add By Sindy 2016/6/8
      Case "Frmacc12d0"
         Frmacc12d0.Enabled = True
      '2016/6/8 END
   End Select
   Set Frmacc1224 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0s0.CursorLocation = adUseClient
   adoacc0s0.Open "select * from acc0s0 where a0s01 = '" & strCon1 & "'", adoTaie
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'************************************************
'  顯示資料表(國內銷帳退費資料)
'
'************************************************
Public Sub FormShowE()
   If IsNull(adoacc0s0.Fields("a0s02").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0s0.Fields("a0s02").Value
   End If
   Text2 = adoacc0s0.Fields("a0s01").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0s0.Fields("a0s03").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0s0.Fields("a0s03").Value)
   End If
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc0s0.Fields("a0s10").Value) Then
      Text22 = MsgText(601)
   Else
      Text22 = adoacc0s0.Fields("a0s10").Value
   End If
   If IsNull(adoacc0s0.Fields("a0s04").Value) Then
      Text3 = MsgText(601)
   Else
      Text3 = adoacc0s0.Fields("a0s04").Value
   End If
   If IsNull(adoacc0s0.Fields("a0s18").Value) Then
      Text27 = MsgText(601)
   Else
      Text27 = adoacc0s0.Fields("a0s18").Value
   End If
   CaseQuery
   Acc0k0Query
   'SumShow 'Removed by Morgan 2011/11/1 併入Acc0k0Query
   'Text13 = Val(Text6) + Val(Text7) - Val(Text8) - Val(Text9) - Val(Text12) 'Removed by Morgan 2011/11/1 併入Acc0k0Query
   If IsNull(adoacc0s0.Fields("a0s05").Value) Then
      Text14 = MsgText(601)
   Else
      Text14 = adoacc0s0.Fields("a0s05").Value
   End If
   If IsNull(adoacc0s0.Fields("a0s06").Value) Then
      Text15 = MsgText(601)
   Else
      Text15 = adoacc0s0.Fields("a0s06").Value
   End If
   If IsNull(adoacc0s0.Fields("a0s07").Value) Then
      Text16 = MsgText(601)
   Else
      Text16 = adoacc0s0.Fields("a0s07").Value
   End If
   If IsNull(adoacc0s0.Fields("a0s08").Value) Then
      Text17 = MsgText(601)
   Else
      Text17 = adoacc0s0.Fields("a0s08").Value
   End If
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(adoacc0s0.Fields("a0s09").Value) Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = CFDate(adoacc0s0.Fields("a0s09").Value)
   End If
   MaskEdBox3.Mask = DFormat
   If IsNull(adoacc0s0.Fields("a0s17").Value) Then
      Text26 = MsgText(601)
   Else
      Text26 = adoacc0s0.Fields("a0s17").Value
   End If
End Sub

'*************************************************
'  案件資料查詢
'
'*************************************************
Public Sub CaseQuery()
   adoacc0k0.CursorLocation = adUseClient
   adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & adoacc0s0.Fields("a0s02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   adopatent.CursorLocation = adUseClient
   'Modified by Morgan 2011/11/1 nvl(cp05 - 19110000, 0)
   'adopatent.Open "select cp01, cp02, cp03, cp04, cp05 from caseprogress where cp60 = '" & adoacc0s0.Fields("a0s02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   adopatent.Open "select cp01, cp02, cp03, cp04, nvl(cp05 - 19110000, 0) cp05 from caseprogress where cp09 in (select a1u03 from acc1u0 where a1u01 = '" & adoacc0s0.Fields("a0s01").Value & "') order by cp05", adoTaie, adOpenStatic, adLockReadOnly
End Sub

'*************************************************
'  查詢顯示(國內收據資料)
'
'*************************************************
Private Sub Acc0k0Query()
    'Add By Cheng 2003/05/15
    Dim StrSQLa As String
   
   If adoacc0k0.RecordCount = 0 Then
      Exit Sub
   End If
   If adopatent.RecordCount <> 0 Then
      If IsNull(adopatent.Fields("cp01").Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = adopatent.Fields("cp01").Value
         If IsNull(adopatent.Fields("cp02").Value) = False Then
            Text4 = Text4 & adopatent.Fields("cp02").Value
         End If
         If IsNull(adopatent.Fields("cp03").Value) = False Then
            Text4 = Text4 & adopatent.Fields("cp03").Value
         End If
         If IsNull(adopatent.Fields("cp04").Value) = False Then
            Text4 = Text4 & adopatent.Fields("cp04").Value
         End If
      End If
      MaskEdBox2.Mask = MsgText(601)
      If IsNull(adopatent.Fields("cp05").Value) Then
         MaskEdBox2.Text = MsgText(601)
      Else
         MaskEdBox2.Text = CFDate(adopatent.Fields("cp05").Value)
      End If
      MaskEdBox2.Mask = DFormat
      adoquery.CursorLocation = adUseClient
        'Modify By Cheng 2003/05/15
'      strSQLA = "select nvl(pa07, nvl(pa06, pa05) from patent where pa01 = '" & adopatent.Fields(0).Value & "' and pa02 = '" & adopatent.Fields(1).Value & "' and pa03 = '" & adopatent.Fields(2).Value & "' and pa04 = '" & adopatent.Fields(3).Value & "' " & _
'                    "select nvl(tm07, nvl(tm06, tm05) from trademark where tm01 = '" & adopatent.Fields(0).Value & "' and tm02 = '" & adopatent.Fields(1).Value & "' and tm03 = '" & adopatent.Fields(2).Value & "' and tm04 = '" & adopatent.Fields(3).Value & "' " & _
'                    "select nvl(lc07, nvl(lc06, lc05) from lawcase where lc01 = '" & adopatent.Fields(0).Value & "' and lc02 = '" & adopatent.Fields(1).Value & "' and lc03 = '" & adopatent.Fields(2).Value & "' and lc04 = '" & adopatent.Fields(3).Value & "' " & _
'                    "select hc06 from hirecase where hc01 = '" & adopatent.Fields(0).Value & "' and hc02 = '" & adopatent.Fields(1).Value & "' and hc03 = '" & adopatent.Fields(2).Value & "' and hc04 = '" & adopatent.Fields(3).Value & "' " & _
'                    "select nvl(sp07, nvl(sp06, sp05)) from servicepractice where sp01 = '" & adopatent.Fields(0).Value & "' and sp02 = '" & adopatent.Fields(1).Value & "' and sp03 = '" & adopatent.Fields(2).Value & "' and sp04 = '" & adopatent.Fields(3).Value & "'"
      StrSQLa = "select nvl(pa07, nvl(pa06, pa05)) from patent where pa01 = '" & adopatent.Fields(0).Value & "' and pa02 = '" & adopatent.Fields(1).Value & "' and pa03 = '" & adopatent.Fields(2).Value & "' and pa04 = '" & adopatent.Fields(3).Value & "' " & _
                    " Union select nvl(tm07, nvl(tm06, tm05)) from trademark where tm01 = '" & adopatent.Fields(0).Value & "' and tm02 = '" & adopatent.Fields(1).Value & "' and tm03 = '" & adopatent.Fields(2).Value & "' and tm04 = '" & adopatent.Fields(3).Value & "' " & _
                    " Union select nvl(lc07, nvl(lc06, lc05)) from lawcase where lc01 = '" & adopatent.Fields(0).Value & "' and lc02 = '" & adopatent.Fields(1).Value & "' and lc03 = '" & adopatent.Fields(2).Value & "' and lc04 = '" & adopatent.Fields(3).Value & "' " & _
                    " Union select hc06 from hirecase where hc01 = '" & adopatent.Fields(0).Value & "' and hc02 = '" & adopatent.Fields(1).Value & "' and hc03 = '" & adopatent.Fields(2).Value & "' and hc04 = '" & adopatent.Fields(3).Value & "' " & _
                    " Union select nvl(sp07, nvl(sp06, sp05)) from servicepractice where sp01 = '" & adopatent.Fields(0).Value & "' and sp02 = '" & adopatent.Fields(1).Value & "' and sp03 = '" & adopatent.Fields(2).Value & "' and sp04 = '" & adopatent.Fields(3).Value & "'"
      adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount <> 0 Then
         If IsNull(adoquery.Fields(0).Value) Then
            Text5 = MsgText(601)
         Else
            Text5 = adoquery.Fields(0).Value
         End If
      Else
         Text5 = ""
      End If
      adoquery.Close
   Else
     Text4 = ""
     MaskEdBox2.Mask = ""
     MaskEdBox2.Text = ""
     MaskEdBox2.Mask = DFormat
   End If
   adopatent.Close
   
   'Modified by Morgan 2011/11/1
   'If IsNull(adoacc0k0.Fields("a0k06").Value) Then
   '   Text6 = MsgText(601)
   'Else
   '   Text6 = adoacc0k0.Fields("a0k06").Value
   'End If
   'If IsNull(adoacc0k0.Fields("a0k07").Value) Then
   '   Text7 = MsgText(601)
   'Else
   '   Text7 = adoacc0k0.Fields("a0k07").Value
   'End If
   'If IsNull(adoacc0k0.Fields("a0k17").Value) Then
   '   Text8 = MsgText(601)
   'Else
   '   Text8 = adoacc0k0.Fields("a0k17").Value
   'End If
   'If IsNull(adoacc0k0.Fields("a0k18").Value) Then
   '   Text9 = MsgText(601)
   'Else
   '   Text9 = adoacc0k0.Fields("a0k18").Value
   'End If
   strExc(0) = "select * from (select a0j02 A1,sum(a0j09) a0j09,sum(a0j10) a0j10 from acc1u0,acc0j0" & _
      " where a1u01 = '" & adoacc0s0.Fields("a0s01").Value & "' and a0j01(+)=a1u03 and a0j13(+)=a1u02" & _
      " group by a0j02) A,(select a0j02 B1,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u06) a1u06" & _
      ",sum(a1u07) a1u07,sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10 from acc1u0,acc0j0" & _
      " where a1u02='" & adoacc0s0.Fields("a0s02").Value & "' and a1u01<>'" & adoacc0s0.Fields("a0s01").Value & "'" & _
      " and a0j01(+)=a1u03 and a0j13(+)=a1u02 group by a0j02) B where B1(+)=A1"
   intI = 1
   Set adoacc0k0 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoacc0k0
      Text6 = Val("" & .Fields("a0j09"))
      Text7 = Val("" & .Fields("a0j10"))
      Text8 = Val("" & .Fields("a1u04"))
      Text9 = Val("" & .Fields("a1u05"))
      Text12 = Val("" & .Fields("a1u07")) + Val("" & .Fields("a1u09"))
      Text10 = Val("" & .Fields("a1u08"))
      Text11 = Val("" & .Fields("a1u10"))
      Text13 = Val(Text6) + Val(Text7) - Val(Text8) - Val(Text9) + Val(Text10) + Val(Text11) - Val(Text12)
      Text24 = Val("" & .Fields("a1u06"))
      End With
   Else
      Text6 = MsgText(601)
      Text7 = MsgText(601)
      Text8 = MsgText(601)
      Text9 = MsgText(601)
      Text10 = MsgText(601)
      Text11 = MsgText(601)
      Text12 = MsgText(601)
      Text13 = MsgText(601)
      Text24 = MsgText(601)
   End If
   'end 2011/11/1
   adoacc0k0.Close
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Private Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0s05), sum(a0s06), sum(a0s07) from acc0s0 where a0s01 = '" & strCon1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text12 = MsgText(601)
      Else
         Text12 = adoaccsum.Fields(0).Value
      End If
      If IsNull(adoaccsum.Fields(1).Value) Then
         Text10 = MsgText(601)
      Else
         Text10 = adoaccsum.Fields(1).Value
      End If
      If IsNull(adoaccsum.Fields(2).Value) Then
         Text11 = MsgText(601)
      Else
         Text11 = adoaccsum.Fields(2).Value
      End If
   Else
      Text12 = MsgText(601)
      Text10 = MsgText(601)
      Text11 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

