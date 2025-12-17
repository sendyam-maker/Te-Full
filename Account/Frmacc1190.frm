VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1190 
   AutoRedraw      =   -1  'True
   Caption         =   "銷帳退費作業"
   ClientHeight    =   5520
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9528
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   9528
   Begin VB.CommandButton cmdCaseNo 
      Caption         =   "本所案號　"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   74
      Top             =   2220
      Width           =   1290
   End
   Begin VB.CommandButton cmdDot 
      Caption         =   "工作點數"
      Height          =   300
      Left            =   8205
      TabIndex        =   72
      Top             =   2257
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "已收回發票"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4590
      TabIndex        =   71
      Top             =   1748
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2925
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   1733
      Width           =   1440
   End
   Begin VB.TextBox Text30 
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6525
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   150
      Width           =   1035
   End
   Begin VB.TextBox Text31 
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8190
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   137
      Width           =   1080
   End
   Begin VB.ComboBox cboCaseNo 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1485
      Style           =   2  '單純下拉式
      TabIndex        =   64
      Top             =   2250
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "退公開費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   63
      Top             =   1748
      Width           =   1365
   End
   Begin VB.TextBox Text28 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7605
      TabIndex        =   7
      Top             =   1740
      Width           =   1572
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   4260
      TabIndex        =   60
      Top             =   167
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Left            =   2790
      TabIndex        =   59
      Top             =   167
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   2280
      Picture         =   "Frmacc1190.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   144
      Width           =   350
   End
   Begin VB.TextBox Text23 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3048
      TabIndex        =   57
      Top             =   4905
      Width           =   3012
   End
   Begin VB.TextBox Text26 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7608
      MaxLength       =   14
      TabIndex        =   11
      Top             =   4050
      Width           =   1572
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4488
      TabIndex        =   55
      Top             =   3690
      Width           =   852
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1488
      TabIndex        =   53
      Top             =   3690
      Width           =   1572
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7608
      TabIndex        =   50
      Top             =   660
      Width           =   1572
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7632
      TabIndex        =   48
      Top             =   4545
      Width           =   1572
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1488
      TabIndex        =   46
      Top             =   4905
      Width           =   1572
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1488
      TabIndex        =   44
      Top             =   4545
      Width           =   1572
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1488
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1380
      Width           =   612
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4488
      MaxLength       =   14
      TabIndex        =   10
      Top             =   4050
      Width           =   1572
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1488
      MaxLength       =   14
      TabIndex        =   9
      Top             =   4050
      Width           =   1572
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7608
      MaxLength       =   14
      TabIndex        =   8
      Top             =   3690
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7608
      TabIndex        =   37
      Top             =   3330
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7608
      TabIndex        =   34
      Top             =   2970
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4488
      TabIndex        =   32
      Top             =   3330
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1488
      TabIndex        =   30
      Top             =   3330
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4488
      TabIndex        =   28
      Top             =   2970
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1488
      TabIndex        =   26
      Top             =   2970
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4488
      TabIndex        =   24
      Top             =   2610
      Width           =   1572
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
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1488
      TabIndex        =   22
      Top             =   2610
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1488
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1020
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1470
      MaxLength       =   15
      TabIndex        =   12
      Top             =   660
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1125
      MaxLength       =   15
      TabIndex        =   0
      Top             =   144
      Width           =   1170
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   4485
      TabIndex        =   2
      Top             =   660
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
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
      Left            =   7605
      TabIndex        =   19
      Top             =   2610
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
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
      Left            =   7605
      TabIndex        =   6
      Top             =   1380
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSForms.TextBox Text27 
      Height          =   315
      Left            =   6165
      TabIndex        =   4
      Top             =   1020
      Width           =   3015
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "5318;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text18 
      Height          =   300
      Left            =   3048
      TabIndex        =   77
      Top             =   4545
      Width           =   3015
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "5318;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   315
      Left            =   4488
      TabIndex        =   76
      Top             =   2250
      Width           =   3660
      VariousPropertyBits=   671105049
      BackColor       =   14737632
      Size            =   "6456;556"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   75
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "已列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   216
      Left            =   5544
      TabIndex        =   73
      Top             =   1428
      Visible         =   0   'False
      Width           =   672
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "發票號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1935
      TabIndex        =   70
      Top             =   1785
      Width           =   900
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "票號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6030
      TabIndex        =   67
      Top             =   195
      Width           =   450
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "票期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7695
      TabIndex        =   66
      Top             =   195
      Width           =   450
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "轉出單號2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6435
      TabIndex        =   62
      Top             =   1785
      Width           =   1035
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "暫收款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4620
      TabIndex        =   61
      Top             =   195
      Width           =   1125
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5685
      TabIndex        =   58
      Top             =   1065
      Width           =   450
   End
   Begin VB.Label Label33 
      BackStyle       =   0  '透明
      Caption         =   "稅款退費金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6165
      TabIndex        =   56
      Top             =   4050
      Width           =   1455
   End
   Begin VB.Label Label32 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3405
      TabIndex        =   54
      Top             =   3690
      Width           =   1215
   End
   Begin VB.Label Label31 
      BackStyle       =   0  '透明
      Caption         =   "扣繳金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   52
      Top             =   3690
      Width           =   1215
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3150
      TabIndex        =   51
      Top             =   189
      Width           =   900
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "轉出單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6435
      TabIndex        =   49
      Top             =   705
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   -72
      Top             =   4944
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   855
      Left            =   45
      Top             =   4425
      Width           =   9255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1605
      Left            =   45
      Top             =   540
      Width           =   9255
   End
   Begin VB.Label Label26 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "暫收款總金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6255
      TabIndex        =   47
      Top             =   4545
      Width           =   1320
   End
   Begin VB.Label Label25 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   45
      Top             =   4905
      Width           =   975
   End
   Begin VB.Label Label24 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   43
      Top             =   4545
      Width           =   1455
   End
   Begin VB.Label Label22 
      BackStyle       =   0  '透明
      Caption         =   "欲處理日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6435
      TabIndex        =   42
      Top             =   1403
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackStyle       =   0  '透明
      Caption         =   "(1.轉應付款 2.轉暫收款 3.匯款)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2208
      TabIndex        =   41
      Top             =   1404
      Width           =   3252
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "退費方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   168
      TabIndex        =   40
      Top             =   1403
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackStyle       =   0  '透明
      Caption         =   "退費規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3405
      TabIndex        =   39
      Top             =   4050
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "退費服務費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   38
      Top             =   4050
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "未收金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   36
      Top             =   3330
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "銷帳金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   35
      Top             =   3690
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "已銷金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   33
      Top             =   2970
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "已退規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3405
      TabIndex        =   31
      Top             =   3330
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "已退服務費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   29
      Top             =   3330
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "已收規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3405
      TabIndex        =   27
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "已收服務費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   25
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "應收規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3405
      TabIndex        =   23
      Top             =   2610
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "應收服務費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   21
      Top             =   2610
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "案件名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3405
      TabIndex        =   20
      Top             =   2250
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "收文日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   18
      Top             =   2610
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "(1.銷帳 2.退費 3.銷帳+退費)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2205
      TabIndex        =   17
      Top             =   1043
      Width           =   2940
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   168
      TabIndex        =   16
      Top             =   1043
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "銷退日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3405
      TabIndex        =   15
      Top             =   705
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "銷退單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   165
      TabIndex        =   14
      Top             =   705
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "單據編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   168
      TabIndex        =   13
      Top             =   168
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc1190"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/19 Form2.0已修改; Text5、Text18、Text27
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc0k0 As New ADODB.Recordset
Public adoacc0s0 As New ADODB.Recordset
Public adoacc0t0 As New ADODB.Recordset
Public adoacc0t0o As New ADODB.Recordset
Public adopatent As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public strDocNo As String
'Add by Morgan 2004/4/5
Public bolTaxed As Boolean '是否已收扣單
'Add by Morgan 2004/12/28
Public m_stDeliver As String '銷退方式
'Add by Morgan 2011/4/13
Dim m_EditList '尚未維護收文分配資料的收文號
Dim m_Assigning As Boolean
Public m_CheckAssign As Boolean
Public m_AssignNo As String
Public m_IsOpen As Boolean 'Add by Morgan 2011/10/14
Public m_KeepItem As String 'Added by Morgan 2015/8/11 保留項次
Dim m_Transfered As String 'Added by Morgan 2014/1/2 未收款沖帳傳票是否已過帳 Y/N
Dim m_A0T18 As String 'Added by Morgan 2014/8/5 暫收款公司別
Dim m_A0J07 As String 'Added by Morgan 2015/4/10 是否合併
Dim m_CaseNo As String, m_KeyNo As String, m_BillNo As String  'Added by Lydia 2021/08/24 記錄操作的本所案號m_CaseNo、銷退單號m_KeyNo、單據編號m_BillNo
Public m_LOS02 As String 'Added by Morgan 2022/6/21

Private Sub cboCaseNo_Click()
   SetCaseData
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then
      If m_Transfered = "N" Then
         MsgBox "有未收款沖帳傳票且未過帳，請先做發票作廢作業！", vbExclamation
         Check2.Value = 0
      Else
         strExc(0) = "select * from acc0m0 where a0m02='" & Text1 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "請款單已收款，請先做發票作廢作業！", vbExclamation
            Check2.Value = 0
         End If
      End If
   End If
End Sub

Private Sub cmdCaseNo_Click()
   tool3_enabled
   Me.Tag = ""
   Frmacc1240.Show
   Set Frmacc1240.frmCall = Me
   Frmacc1240.cmdCrtRct.Visible = False
   strSaveConfirm = ""
   Me.Enabled = False
End Sub

Private Sub Command3_Click()
   If m_IsOpen = False Then OpenTable 'Add by Morgan 2011/10/14
   MaskEdBox1.Enabled = True 'Add by Amy  2014/10/28
   If Option1.Value Then
      If adoacc0s0.RecordCount = 0 Or Text1 = MsgText(601) Then
         Exit Sub
      End If
      adoacc0s0.Find "a0s02 = '" & Text1 & "'", 0, adSearchForward, 1
      If adoacc0s0.EOF = False Then
         'Add by Morgan 2004/4/5
         Call Check1V0(bolTaxed)
         FormShowE
         Frmacc0000.StatusBar1.Panels(2).Text = adoacc0s0.Bookmark & MsgText(35) & adoacc0s0.RecordCount
      Else
         MsgBox MsgText(33), , MsgText(5)
         adoacc0s0.MoveFirst
      End If
   Else
      If adoacc0t0.RecordCount = 0 Or Text1 = MsgText(601) Then
         Exit Sub
      End If
      adoacc0t0.Find "a0s02 = '" & Text1 & "'", 0, adSearchForward, 1
      If adoacc0t0.EOF = False Then
         FormShowJ
         Frmacc0000.StatusBar1.Panels(2).Text = adoacc0t0.Bookmark & MsgText(35) & adoacc0t0.RecordCount
      Else
         MsgBox MsgText(33), , MsgText(5)
         adoacc0t0.MoveFirst
      End If
   End If
End Sub

Private Sub Command3_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         Command3_Click
         Exit Sub
   End Select
   KeyEnter KeyCode
End Sub

Private Sub Form_Activate()
   strFormName = Name
   
   'Added by Morgan 2015/12/8
   If strFormLink = "Frmacc1240" Then
      '全域變數要清除,否則可能殘留查詢畫面的值而導致存檔時發生錯誤
      strCon1 = ""
      strCon2 = ""
      strCon3 = ""
      strCon4 = ""
      strFormLink = ""
      If Me.Tag <> "" Then
         Text1 = Me.Tag
         Me.Tag = ""
         Text1_Validate False
         Text2.SetFocus
      End If
      Exit Sub
   End If
   'end 2015/12/8
   
   strFormLink = ""
   
   'Add by Morgan 2011/5/30
   CheckAssign
   
   If m_Assigning Then
      strExc(0) = GetNextNo
      If strExc(0) <> "" Then
         Frmacc11l0.m_sAssignNo = strExc(0)
         Frmacc11l0.m_sCallType = "E"
         Set Frmacc11l0.m_fCallForm = Me
         Frmacc11l0.Show
         Me.Visible = False
      Else
         m_Assigning = False
         tool1_enabled
         MenuDisabled
      End If
      Exit Sub
   End If
   'end 2011/5/30
   
   If strItemNo = MsgText(601) Then
      Exit Sub
   End If
   
   Text2 = strItemNo
   Select Case Mid(strCon1, 1, 1)
      Case "E"
         Option1.Value = True
      Case "J"
         Option2.Value = True
   End Select
   FormRefresh
   If Option1.Value Then
      If adoacc0s0.RecordCount <> 0 Then
         adoacc0s0.MoveFirst
         adoacc0s0.Find "a0s01 = '" & Text2 & "'", 0, adSearchForward, 1
         If adoacc0s0.EOF = False Then
            FormShowE
            Frmacc0000.StatusBar1.Panels(2).Text = adoacc0s0.Bookmark & MsgText(35) & adoacc0s0.RecordCount
         End If
      End If
   Else
      If adoacc0t0.RecordCount <> 0 Then
         adoacc0t0.MoveFirst
         adoacc0t0.Find "a0s01 = '" & Text2 & "'", 0, adSearchForward, 1
         If adoacc0t0.EOF = False Then
            FormShowJ
            Frmacc0000.StatusBar1.Panels(2).Text = adoacc0t0.Bookmark & MsgText(35) & adoacc0t0.RecordCount
         End If
      End If
   End If
   strItemNo = MsgText(601)
End Sub
'Added by Lydia 2021/08/24
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  'Form2.0 記錄鍵盤傳入順序
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   'Add by Morgan 2004/9/29
   If KeyCode = vbKeyF3 Then
      If EditCheck = False Then Exit Sub
   End If
   '2004/9/29 end
   'Modified by Lydia 2021/08/24
   'KeyEnter KeyCode
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
    
   'Mark by Lydia 2021/08/18 法務系統的工作點數分配功能先上線
   'cmdDot.Visible = False 'Added Lydia 2020/04/10
   'Modify by Amy 2023/10/06 原:9500, 5790
   PUB_InitForm Me, 9650, 5980, strBackPicPath1
   strItemNo = MsgText(601)
   MaskEdBox1.Mask = DFormat
   'Modified by Morgan 2011/10/14
   '先不抓資料否則進畫面會很久
   'OpenTable
   'If adoacc0s0.RecordCount <> 0 Then
   '   Frmacc0000.StatusBar1.Panels(2).Text = adoacc0s0.Bookmark & MsgText(35) & adoacc0s0.RecordCount
   'End If
   m_IsOpen = False

   SetCmdCaseNo 'Added by Morgan 2015/12/8
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   strTrackMode = "" 'Added by Lydia 2021/08/24 Form2.0 記錄鍵盤傳入順序(清除)
   MenuEnabled
   Set Frmacc1190 = Nothing
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      Exit Sub
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) Then
      Exit Sub
   End If
   If Mid(MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
      Text2 = UpdateNo("acc0s0", "a0s01", 5, MaskEdBox1.Text, MsgText(805))
   Else
      'Text2 = AutoNo(MsgText(805), 5)
      Text2 = strDocNo
   End If
End Sub

Private Sub MaskEdBox3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Option1_Click()
   FormRefresh
   'Modify by Morgan 2005/6/2 加控制收據不可只退費
   If Text3 = "2" Then
      MsgBox "選收據時類別不可為2", vbExclamation
      Text3 = "1"
   End If
End Sub

Private Sub Option2_Click()
   FormRefresh
End Sub

Private Sub Text1_Change()
   Text30 = "": Text31 = "": Text4 = "": Check2.Value = 0: Check2.Visible = False
   m_LOS02 = "" 'Added by Morgan 2022/6/21
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

Private Sub Text1_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   
   If Option1.Value Then
   
      'Added by Mogan 2025/5/22
      If PUB_Chk440(Text1.Text, "1") Then
         Text1_GotFocus
         Cancel = True
         Exit Sub
      End If
      'end 2025/5/22
      
      ObjectEnabled_E
      CaseQuery
      Acc0k0Query
      'Remove by Morgan 2011/10/17 Acc0k0Query做了
      'SumShow
      'Text13 = Val(Text6) + Val(Text7) - Val(Text8) - Val(Text9) - Val(Text12)
      'end2011/10/17
      
      'Add by Morgan 2004/4/5
      Call Check1V0(bolTaxed)
   Else
      ObjectEnabled_J
      Acc0t0Query
      Text17 = "1"
   End If
End Sub

Private Function Check1V0(ByRef bolTaxed As Boolean) As Boolean

On Error GoTo flgErr

   Dim stSQL As String, adQuery As New ADODB.Recordset
   stSQL = "Select * From acc1v0 where a1v02='" & Text1 & "' and a1v15 is not null"
   adQuery.CursorLocation = adUseClient
   adQuery.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
   If adQuery.RecordCount > 0 Then
      MsgBox "本筆收據已有扣單資料，稅款金額不可修改！", vbExclamation
      bolTaxed = True
      Text26.Enabled = False
   Else
      bolTaxed = False
   End If
   Check1V0 = True
flgErr:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub Text14_GotFocus()
   TextInverse Text14
End Sub

Private Sub Text14_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
   If Val(Text14) > Val(Text6) + Val(Text7) Then
      MsgBox MsgText(108), , MsgText(5)
      Cancel = True
      Text14.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text15_GotFocus()
   TextInverse Text15
End Sub

Private Sub Text15_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text15_Validate(Cancel As Boolean)
   If Val(Text15) > Val(Text8) Then
      MsgBox MsgText(109), , MsgText(5)
      Cancel = True
      Text15.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text16_GotFocus()
   TextInverse Text16
End Sub

Private Sub Text16_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   If Val(Text16) > Val(Text9) Then
      MsgBox MsgText(109), , MsgText(5)
      Cancel = True
      Text16.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Text17_GotFocus()
   TextInverse Text17
End Sub

Private Sub Text17_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
   'KeyEnter KeyCode 'Remove by Lydia 2021/08/24
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   'Modify by Amy 2014/11/07 解未key單據號碼,直接輸類別1,因acc0k0沒資料所以會error
   If Text17 = MsgText(601) Or Text1 = MsgText(601) Or Text1 = "E" Then
      Exit Sub
   End If
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If strCon1 <> MsgText(601) Then
      Exit Sub
   End If
   'Add by Amy 2014/09/30 +判斷acc0k0沒資料清空此欄位(為可觸發acc1191)
   If adoacc0k0.EOF = True Then '按了查詢後再新增
        MsgBox "無國內收據資料!", , MsgText(5)
        Text17 = ""
        Exit Sub
   End If
   'end 2014/09/30
   Select Case Text17
      Case "1"
         strCon5 = Text2
         tool3_enabled
         Frmacc1191.Show
         Me.Enabled = False
         'Add by Morgan 2004/12/27 新增時若選收據編號、轉應付款(1)且備註欄為空白時預設為收據抬頭
         If Option1.Value = True And Text27.Text = "" Then
            Text27.Text = "" & adoacc0k0.Fields("a0k04")
         End If
         '2004/12/27 end
      '2007/6/27 add by sonia 新增時若選收據編號、轉暫收款(2)且備註欄為空白時預設為收據抬頭
      Case "2"
         If Option1.Value = True And Text27.Text = "" Then
            Text27.Text = "" & adoacc0k0.Fields("a0k04")
         End If
      '2007/6/27 end
   End Select
End Sub

Private Sub Text19_Change()
    If Text19 = MsgText(601) Then
        'Add By Cheng 2004/03/31
        '清空智權人員名稱
        Text18 = ""
        'End
        Exit Sub
    End If
    Text18 = StaffQuery(Text19)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text20_Change()
   If Text20 = MsgText(601) Then
      Exit Sub
   End If
   Text23 = CustomerQuery(Text20, 1)
End Sub

Private Sub Text26_GotFocus()
   TextInverse Text26
End Sub

Private Sub Text26_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Text26_Validate(Cancel As Boolean)
'   If Val(Text26) > Val(Text24) Then
'      MsgBox MsgText(109), , MsgText(5)
'      Cancel = True
'      Text26.SetFocus
'      Exit Sub
'   End If
End Sub

Private Sub Text27_GotFocus()
   TextInverse Text27
   'edit by nickc 2007/06/11  切換輸入法改用API
   OpenIme
End Sub

'Modified by Lydia 2021/08/19 改成Form 2.0
'Private Sub Text27_KeyUp(KeyCode As Integer, Shift As Integer)
'     KeyEnter KeyCode
Private Sub Text27_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     KeyEnter Val(KeyCode)
'end 2021/08/19
End Sub

Private Sub Text27_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   'Modify by Morgan 2005/6/2 加控制收據不可只退費
   If KeyAscii = Asc("2") And Option1.Value = True Then
      KeyAscii = 0
      MsgBox "收據不可選2", vbExclamation
   End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  物件可使用狀態設定1
'
'*************************************************
Public Sub ObjectEnabled_E()
   Text2.Enabled = True
   MaskEdBox1.Enabled = True
   Text3.Enabled = True
   Text14.Enabled = True
   Text15.Enabled = True
   Text16.Enabled = True
   Text17.Enabled = True
   MaskEdBox3.Enabled = True
   Text26.Enabled = True
   
   'Add by Morgan 2004/4/5
   '已收扣單不可改稅額
   If bolTaxed Then Text26.Enabled = False
End Sub

'*************************************************
'  物件可使用狀態設定2
'
'*************************************************
Public Sub ObjectEnabled_J()
   Text2.Enabled = False
   'MaskEdBox1.Enabled = False
   Text3.Enabled = False
   Text14.Enabled = False
   Text15.Enabled = False
   Text16.Enabled = False
   Text17.Enabled = False
   'MaskEdBox3.Enabled = False
   Text26.Enabled = False
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Public Sub OpenTable()
On Error GoTo Checking
   'Modify by Morgan 2011/10/14
   'adoacc0k0.CursorLocation = adUseClient
   'adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0s0.State = adStateOpen Then adoacc0s0.Close
   adoacc0s0.CursorLocation = adUseClient
   If Option1.Value Then
       adoacc0s0.Open "select * from acc0s0 where substr(a0s02, 1, 1) = 'E' order by a0s01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   Else
       adoacc0s0.Open "select * from acc0s0 where substr(a0s02, 1, 1) = 'J' order by a0s01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End If
   
   'Added by Morgan 2024/11/12
   If Text2 <> "" And adoacc0s0.RecordCount > 0 Then
      adoacc0s0.Find "a0s01 = '" & Text2 & "'", 0, adSearchForward, 1
   End If
   'end 2024/11/12
   
   If adoacc0t0.State = adStateOpen Then adoacc0t0.Close 'Added by Morgan 2024/11/12
   adoacc0t0.CursorLocation = adUseClient
   adoacc0t0.Open "select * from acc0t0, acc0s0 where a0t01 = a0s02 and (a0t09 is not null and a0t09 <> 0) order by a0t01 asc", adoTaie, adOpenStatic, adLockReadOnly
   m_IsOpen = True 'Add by Morgan 2011/10/14
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
      
   'Added by Morgan 2015/12/8
   Label34.Visible = False
   SetCmdCaseNo
   'end 2015/12/8
   
   If Me.m_IsOpen = False Then Exit Sub
   If IsNull(adoacc0s0.Fields("a0s02").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0s0.Fields("a0s02").Value
   End If
   'CaseQuery 'Remove by Morgan 2011/10/17 下面會做不必重複
   Text2 = adoacc0s0.Fields("a0s01").Value
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0s0.Fields("a0s03").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0s0.Fields("a0s03").Value)
   End If
   MaskEdBox1.Tag = "" & adoacc0s0.Fields("a0s03").Value 'Add by Amy 2015/01/21
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
   'Remove by Morgan 2011/10/17 Acc0k0Query做了
   'SumShow
   'Text13 = Val(Text6) + Val(Text7) - Val(Text8) - Val(Text9) - Val(Text12) 'Remove by Morgan 2011/10/17
   'end 2011/10/17
   
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
   
   'Add by Morgan 2009/7/1
   If adoacc0s0.Fields("a0s24") = "Y" Then
      Check1.Value = 1
   Else
      Check1.Value = 0
   End If
   
   'Added by Morgan 2014/1/17
   Text4 = "" & adoacc0s0.Fields("A0S26")
   If Text4 <> "" Then
      Check2.Visible = True
      If adoacc0s0.Fields("A0S27") = "N" Then
         Check2.Value = 0
      Else
         Check2.Value = 1
      End If
   End If
   'end 2014/1/17
   'Add by Amy 2014/10/28 +a1p22有值不可修改銷退日
   If Text3 = "3" And CheckExistA1p22(IIf(adoacc0k0.Fields("a0k11") = "J", "J", "1"), "Z", Text2) = True Then
        MaskEdBox1.Enabled = False
    End If
    'end 2014/10/28
End Sub

'*************************************************
'  案件資料查詢
'
'*************************************************
Public Sub CaseQuery()
   If adoacc0k0.State = adStateOpen Then adoacc0k0.Close
   adoacc0k0.CursorLocation = adUseClient
   'Modified by Morgan 2014/1/2 +抓發票資料
   'adoacc0k0.Open "select * from acc0k0 where a0k01 = '" & Text1 & "' and (a0k09 is null or a0k09 = 0)", adoTaie, adOpenStatic, adLockReadOnly
   adoacc0k0.Open "select * from acc0k0,acc431,acc430 where a0k01 = '" & Text1 & "' and (a0k09 is null or a0k09 = 0) and axc02(+)=a0k01 and a4301(+)=axc01", adoTaie, adOpenStatic, adLockReadOnly
   
   If adopatent.State = adStateOpen Then adopatent.Close
   adopatent.CursorLocation = adUseClient
   'Modify by Morgan 2011/10/17 考慮多案一收據情形
   'adopatent.Open "select cp01, cp02, cp03, cp04, cp05 from caseprogress where cp60 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   'Modified by Morgan 2015/4/10 +a0j07
   strExc(0) = "select cp01, cp02, cp03, cp04,min(cp01||cp02||cp03||cp04) CaseNo,min(cp05) cp05" & _
      ",nvl(sum(a0j09),0) a0j09,nvl(sum(a0j10),0) a0j10,nvl(sum(U1.a1u04),0) a1u04" & _
      ",nvl(sum(U1.a1u05),0) a1u05,nvl(sum(U1.a1u06),0) a1u06" & _
      ",nvl(sum(U1.a1u07),0) a1u07,nvl(sum(U1.a1u08),0) a1u08" & _
      ",nvl(sum(U1.a1u09),0) a1u09,nvl(sum(U1.a1u10),0) a1u10" & _
      ",min(U2.a1u01) a1u01,max(a0j07) a0j07" & _
      " from acc0j0,caseprogress,(select A1U03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u06) a1u06,sum(a1u07) a1u07" & _
      ",sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10 from acc1u0" & _
      " where a1u02='" & Text1 & "' and a1u01<>'" & Text2 & "' GROUP BY A1U03) U1" & _
      ",(SELECT A1U01,A1U03 FROM acc1u0 U2 WHERE A1U01='" & Text2 & "') U2" & _
      " where a0j13='" & Text1 & "' and cp09(+)=a0j01 and U1.a1u03(+)=a0j01" & _
      " AND U1.A1U03(+)=A0J01 AND U2.A1U03(+)=A0J01" & _
      " group by cp01,cp02,cp03,cp04"
   adopatent.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Private Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select sum(a0s05), sum(a0s06), sum(a0s07) from acc0s0 where a0s02 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
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

'*************************************************
'  查詢顯示(國內收據資料)
'
'*************************************************
Public Sub Acc0k0Query()
'Add by Morgan 2011/10/17 考慮多對多收據(舊程式放下面比較容易看)
ClearQuery
If adoacc0k0.RecordCount > 0 Then
   If Val("" & adoacc0k0.Fields("a0k19").Value) > 0 Then Label34.Visible = True 'Added by Morgan 2015/12/7
   
   Text25 = "" & adoacc0k0.Fields("a0k16").Value
   'Added by Morgan 2014/1/2
   '發票號碼
   Text4 = "" & adoacc0k0("axc01")
   m_Transfered = ""
   If Text4 <> "" Then
      Check2.Visible = True
      Check2.Enabled = True
      '若發票有未收款沖帳傳票(A4317)且已過帳時不可勾選已收回發票
      If Not IsNull(adoacc0k0("A4317")) Then
         strExc(0) = "select * from acc021 where ax201='J' and ax202='" & adoacc0k0("A4317") & "' and ax210>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Check2.Enabled = False
            m_Transfered = "Y"
         Else
            m_Transfered = "N"
         End If
      End If
      '若發票已申報(ACC410.A4111有值)
      If Check2.Enabled = True Then
         strExc(1) = adoacc0k0("A4302") \ 100
         '2014/12/2 modify by sonia 因二個月一期,故以發票月份-1個月去抓發票號碼檔,否則雙月讀不到
         'strExc(0) = "select * from acc410 where a4101>=" & strExc(1) & " and a4102<=" & strExc(1) & " and a4111>0"
         strExc(0) = "select * from acc410 where a4101>=" & Val(strExc(1)) - 1 & " and a4102<=" & strExc(1) & " and a4111>0"
         '2014/12/2 end
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Check2.Enabled = False
         End If
      End If
   End If
   
   '票號,票期(最大票期)
   strExc(0) = "select * from acc1u0,acc1p0 where a1u02='" & Text1 & "' and a1p04(+)=a1u01 and a1p09 is not null and a1p12>0  order by a1p12 desc,a1p09 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Text30 = RsTemp("a1p09")
      Text31 = ChangeTStringToTDateString(RsTemp("a1p12"))
   End If
   'end 2014/1/2
End If

With adopatent
If .RecordCount > 0 Then
   .MoveFirst
   Do While Not .EOF
      cboCaseNo.AddItem .Fields("CaseNo")
      If .Fields("a1u01") = Text2.Text Then
         'Memo by Lydia 2020/04/10 會觸發cboCaseNo_Click
         cboCaseNo.ListIndex = .AbsolutePosition - 1
      End If
      .MoveNext
   Loop
   If cboCaseNo.ListIndex < 0 Then
      cboCaseNo.ListIndex = 0
   End If
   If strSaveConfirm = "A" Or strSaveConfirm = "E" Then
      Me.cboCaseNo.Enabled = True
      SetLOS02 'Added by Morgan 2022/6/21
   Else
      Me.cboCaseNo.Enabled = False
   End If
End If
End With
'Remove by Morgan 2011/10/17
'--舊程式已刪除--
End Sub

'Add by Morgan 2011/10/11
Private Sub SetCaseData()
   MaskEdBox2.Mask = MsgText(601)
   MaskEdBox2.Text = MsgText(601)
   With adopatent
   .MoveFirst
   .Find " CaseNo='" & cboCaseNo.Text & "'"
   If Not .EOF Then
      MaskEdBox2.Text = CFDate(ACDate(.Fields("cp05")))
      Text5 = GetCaseName(cboCaseNo.Text)
      Text6 = .Fields("a0j09")
      Text7 = .Fields("a0j10")
      Text8 = .Fields("a1u04")
      Text9 = .Fields("a1u05")
      Text12 = .Fields("a1u07") + .Fields("a1u09")
      Text10 = .Fields("a1u08")
      Text11 = .Fields("a1u10")
      Text13 = Val(Text6) + Val(Text7) - Val(Text8) - Val(Text9) + Val(Text10) + Val(Text11) - Val(Text12)
      Text24 = .Fields("a1u06")
      m_A0J07 = "" & .Fields("a0j07") 'Added by Morgan 2015/4/10
      'Added by Lydia 2020/04/10
      'Mark by Lydia 2020/04/20 先隱藏
      'Remove Mark by Lydia 2021/08/18 法務系統的工作點數分配功能先上線
      If InStr(UCase(cboCaseNo.Text), "L") > 0 Then  '法務案: 顯示工作點數按鈕
          cmdDot.Visible = True
      Else
          cmdDot.Visible = False
      End If
      'end 2020/04/20
      'end 2020/04/10
   End If
   End With
   MaskEdBox2.Mask = DFormat
End Sub

'Add by Morgan 2004/4/6
'取得已扣繳稅額
Private Function GetA1V06() As String
   Dim stSQL As String, rsQuery As New ADODB.Recordset
   
On Error GoTo flgErr

   GetA1V06 = "0"
   With rsQuery
      If .State = adStateOpen Then .Close
      .CursorLocation = adUseClient
      .Open "select sum(a1v06) from acc1v0 where a1v02 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
          GetA1V06 = Format(Val("" & .Fields(0)))
      End If
   End With
   
flgErr:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   Set rsQuery = Nothing

End Function

'*************************************************
'  清除查詢顯示
'
'*************************************************
Private Sub ClearQuery()
   'Modify by Morgan 2011/10/11
   'Text4 = ""
   cboCaseNo.Clear
   
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text12 = ""
   Text10 = ""
   Text11 = ""
   Text13 = ""
   Text19 = ""
   Text18 = ""
   Text20 = ""
   Text23 = ""
   Text21 = ""
End Sub

'*************************************************
'  顯示資料表(國內暫收款資料(主檔))
'
'*************************************************
Public Sub FormShowJ()
   
   Label34.Visible = False 'Added by Morgan 2015/12/7
   If IsNull(adoacc0t0.Fields("a0t01").Value) Then
      Text1 = MsgText(601)
   Else
      Text1 = adoacc0t0.Fields("a0t01").Value
   End If
   If IsNull(adoacc0t0.Fields("a0t05").Value) Then
      Text19 = MsgText(601)
   Else
      Text19 = adoacc0t0.Fields("a0t05").Value
   End If
   If IsNull(adoacc0t0.Fields("a0t06").Value) Then
      Text20 = MsgText(601)
   Else
      Text20 = adoacc0t0.Fields("a0t06").Value
   End If
   If IsNull(adoacc0t0.Fields("a0t08").Value) Then
      Text21 = MsgText(601)
   Else
      Text21 = adoacc0t0.Fields("a0t08").Value
   End If
   If IsNull(adoacc0t0.Fields("a0s01").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0t0.Fields("a0s01").Value
   End If
   MaskEdBox1.Mask = MsgText(601)
   If IsNull(adoacc0t0.Fields("a0s03").Value) Then
      MaskEdBox1.Text = MsgText(601)
   Else
      MaskEdBox1.Text = CFDate(adoacc0t0.Fields("a0s03").Value)
   End If
   MaskEdBox1.Tag = adoacc0t0.Fields("a0s03").Value 'Add by Amy 2015/01/21
   MaskEdBox1.Mask = DFormat
   If IsNull(adoacc0t0.Fields("a0s18").Value) Then
      Text27 = MsgText(601)
   Else
      Text27 = adoacc0t0.Fields("a0s18").Value
   End If
   MaskEdBox3.Mask = MsgText(601)
   If IsNull(adoacc0t0.Fields("a0s09").Value) Then
      MaskEdBox3.Text = MsgText(601)
   Else
      MaskEdBox3.Text = CFDate(adoacc0t0.Fields("a0s09").Value)
   End If
   MaskEdBox3.Mask = DFormat
   If IsNull(adoacc0t0.Fields("a0s10").Value) Then
      Text22 = MsgText(601)
   Else
      Text22 = adoacc0t0.Fields("a0s10").Value
   End If
   'Add by Morgan 2006/12/21
   Text28 = "" & adoacc0t0.Fields("a0s23").Value
   
    'Add by Amy 2014/10/28 +a1p22有值不可修改銷退日
    If CheckExistA1p22(adoacc0t0.Fields("a0t18"), "Z", Text2) = True Then
        MaskEdBox1.Enabled = False
    End If
    'end 2014/10/28
End Sub

'*************************************************
'  查詢顯示(國內暫收款資料(主檔))
'
'*************************************************
Private Sub Acc0t0Query()
   adoacc0t0o.CursorLocation = adUseClient
   adoacc0t0o.Open "select * from acc0t0 where a0t01 = '" & Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0t0o.RecordCount = 0 Then
      ClearQuery
      adoacc0t0o.Close
      Exit Sub
   'Added by Lydia 2024/11/28 未收文客戶暫收款管制
   Else
      If "" & adoacc0t0o.Fields("a0t06") = "X03072010" Then
         MsgBox "此暫收款之客戶為" & adoacc0t0o.Fields("a0t06") & CustomerQuery(adoacc0t0o.Fields("a0t06"), 1) & "，不可沖帳 !", vbExclamation
         ClearQuery
         adoacc0t0o.Close
         Exit Sub
      End If
   'end 2024/11/28
   End If
   '2005/7/1 ADD BY SONIA
   adoquery.CursorLocation = adUseClient
   'Modify by Morgan 2007/4/27 加判斷科目為2401的否則收據銷退轉暫收的會無法轉應付
   'adoquery.Open "select * from acc1P0 where a1P23 = '" & Text1 & "' AND A1P07>0", adoTaie, adOpenStatic, adLockReadOnly
   adoquery.Open "select * from acc1P0 where a1p05='2401' and a1P23 = '" & Text1 & "' AND A1P07>0", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount > 0 Then
      Select Case adoquery.Fields("A1P02")
         Case "Z"
            MsgBox "已退費, 無法再退費...", , MsgText(5)
            ClearQuery
            adoquery.Close
            adoacc0t0o.Close
            Exit Sub
         Case "A"
            MsgBox "已轉收款, 無法再退費...", , MsgText(5)
            ClearQuery
            adoquery.Close
            adoacc0t0o.Close
            Exit Sub
         Case Else
      End Select
   Else
      adoquery.Close
   End If
   '2005/7/1 END
   If IsNull(adoacc0t0o.Fields("a0t05").Value) Then
      Text19 = MsgText(601)
   Else
      Text19 = adoacc0t0o.Fields("a0t05").Value
   End If
   If IsNull(adoacc0t0o.Fields("a0t06").Value) Then
      Text20 = MsgText(601)
   Else
      Text20 = adoacc0t0o.Fields("a0t06").Value
   End If
   If IsNull(adoacc0t0o.Fields("a0t08").Value) Then
      Text21 = MsgText(601)
   Else
      Text21 = adoacc0t0o.Fields("a0t08").Value
   End If
   '2007/6/27 add by sonia
   If IsNull(adoacc0t0o.Fields("a0t17").Value) Then
      'Modify by Morgan 2007/8/16 新增暫收款銷退時備註預設為客戶名稱
      'Text27 = MsgText(601)
      If Text27 = "" Then
         Text27 = Text23
      End If
      'end 2007/8/16
   Else
      Text27 = adoacc0t0o.Fields("a0t17").Value
   End If
   '2007/6/27 end
   m_A0T18 = "" & adoacc0t0o.Fields("a0t18").Value 'Added by Morgan 2014/8/5
   adoacc0t0o.Close
End Sub

'*************************************************
'  更新資料表
'
'*************************************************
Public Sub FormRefresh()
'   Frmacc1190_Clear
   If Option1.Value Then
      If adoacc0s0.State = adStateOpen Then
         adoacc0s0.Close
      End If
      adoacc0s0.CursorLocation = adUseClient
      adoacc0s0.Open "select * from acc0s0 where substr(a0s02, 1, 1) = 'E' order by a0s01 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      If adoacc0s0.RecordCount <> 0 Then
         Frmacc0000.StatusBar1.Panels(2).Text = adoacc0s0.Bookmark & MsgText(35) & adoacc0s0.RecordCount
      Else
         StatusClear
      End If
   Else
      If adoacc0t0.State = adStateOpen Then
         adoacc0t0.Close
      End If
      adoacc0t0.CursorLocation = adUseClient
      adoacc0t0.Open "select * from acc0t0, acc0s0 where a0t01 = a0s02 and (a0t09 is not null and a0t09 <> 0) order by a0t01 asc", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0t0.RecordCount <> 0 Then
         Frmacc0000.StatusBar1.Panels(2).Text = adoacc0t0.Bookmark & MsgText(35) & adoacc0t0.RecordCount
      Else
         StatusClear
      End If
   End If
   m_IsOpen = True 'Add by Morgan 2011/10/14

   If Option1.Value Then
      ObjectEnabled_E
   Else
      ObjectEnabled_J
   End If
   
   SetCmdCaseNo 'Added by Morgan 2015/12/8
End Sub

'Added by Morgan 2015/12/8
Public Sub SetCmdCaseNo()
   If strSaveConfirm = MsgText(3) And Option1.Value Then
      cmdCaseNo.Visible = True
   Else
      cmdCaseNo.Visible = False
   End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 = MsgText(601) Then
      MsgBox MsgText(52), , MsgText(5)
      Cancel = True
      Text3.SetFocus
      Exit Sub
   End If
   If Text3 = "2" Or Text3 = "3" Then
      If adoacc0k0.RecordCount <> 0 Then
         If (IsNull(adoacc0k0.Fields("a0k18").Value) Or Val(adoacc0k0.Fields("a0k18").Value) = 0) And (IsNull(adoacc0k0.Fields("a0k17").Value) Or Val(adoacc0k0.Fields("a0k17").Value) = 0) Then
            MsgBox MsgText(110), , MsgText(5)
            Cancel = True
            Text3.SetFocus
            Exit Sub
         End If
      End If
   End If
   Select Case Text3
      Case "1"
         Text14.Enabled = True
         Text17.Enabled = False
         MaskEdBox3.Enabled = False
         Text15.Enabled = False
         Text16.Enabled = False
         Text26.Enabled = False
         Text15 = 0 'Add by Morgan 2011/10/21 避免點退費且輸入金額後改點銷帳導致資料殘留
         Text16 = 0 'Add by Morgan 2011/10/21 避免點退費且輸入金額後改點銷帳導致資料殘留
         Text26 = 0 'Add by Morgan 2011/10/21 避免點退費且輸入金額後改點銷帳導致資料殘留
      Case Else
         Text14 = 0 'Add by Morgan 2011/10/21 避免點銷帳且輸入金額後改點退費導致資料殘留
         Text14.Enabled = False
         Text17.Enabled = True
         MaskEdBox3.Enabled = True
         Text15.Enabled = True
         Text16.Enabled = True
         Text26.Enabled = True
         'Add by Morgan 2004/4/5
         '已收扣單不可改稅額
         If bolTaxed Then Text26.Enabled = False
   End Select
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Call PUB_SaveTrackMode(1, KeyCode)  'Added by Lydia 2021/08/24 Form2.0 記錄鍵盤傳入順序
   
   Select Case KeyCode
      Case vbKeyF12
         If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
            Exit Sub
         End If
         Select Case Text17
            Case "1"
               strCon5 = Text2
               tool3_enabled
               Frmacc1191.Show
               Me.Enabled = False
         End Select
   End Select
   KeyEnter KeyCode 'Added by Lydia 2021/08/24
End Sub

Public Sub Mail2Sales()
   Dim stSales As String
   Dim stSubject As String
   Dim stContent As String
   Dim ii As Integer
   Dim strRecvNo As String 'Add By Sindy 2023/3/27
   
   stContent = "銷退單號：" & Text2
   '收據
   If Option1.Value = True Then
      stSales = adoacc0k0.Fields("a0k20")
      'Modify by Morgan 2011/10/11
      'stSubject = "銷退通知 -> " & Text4 & " " & Text5
      stSubject = "銷退通知 -> " & cboCaseNo.Text & " " & Text5
      If cboCaseNo.ListCount > 1 Then
         stSubject = stSubject & " 等..."
      End If
      'end 2011/10/11
      
      stContent = stContent & vbCrLf & "收據編號：" & Text1
      stContent = stContent & vbCrLf & "客戶名稱：" & adoacc0k0.Fields("a0k03") & " " & CustomerQuery("" & adoacc0k0.Fields("a0k03"), 1)
      
      'Modify by Morgan 2011/10/11
      'stContent = stContent & vbCrLf & "本所案號：" & Text4
      stContent = stContent & vbCrLf & "本所案號：" & cboCaseNo.Text
      stContent = stContent & vbCrLf & "案件名稱：" & Text5
      stContent = stContent & vbCrLf & "申請國家：" & adoacc0k0.Fields("a0k23") & " " & GetPrjNationName("" & adoacc0k0.Fields("a0k23"))
      stContent = stContent & vbCrLf & "銷退日期：" & MaskEdBox1.Text
      stContent = stContent & vbCrLf & "類　　別：" & IIf(Text3 = "1", "銷帳", IIf(Text3 = "2", "退費", "銷+退"))
      If Text3 = "1" Then
         stContent = stContent & vbCrLf & "銷帳金額：" & Format(Val(Format(Text14)), "#,##0")
      Else
         'Modify by Morgan 2007/2/8
         'stContent = stContent & vbCrLf & "退費方式：" & IIf(Text3 = "1", "轉應付", "轉暫收")
         'Modified by Morgan 2023/11/8
         'stContent = stContent & vbCrLf & "退費方式：" & IIf(Text17 = "1", "轉應付", "轉暫收")
         stContent = stContent & vbCrLf & "退費方式：" & IIf(Text17 = "1", "轉應付", IIf(Text17 = "2", "轉暫收", "匯款"))
         'End 2007/2/8
         stContent = stContent & vbCrLf & "退費服務費：" & Format(Val(Format(Text15)), "#,##0")
         stContent = stContent & vbCrLf & "退費規費：" & Format(Val(Format(Text16)), "#,##0")
         stContent = stContent & vbCrLf & "稅款退費金額：" & Format(Val(Format(Text26)), "#,##0")
      End If
      
      'Add By Sindy 2023/3/27 抓總收文號
      strExc(0) = "select a0j13,a0j01 from acc0j0,acc0k0 where a0k01='" & Text1 & "' and a0k01=a0j13"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strRecvNo = "" & RsTemp.Fields("a0j01")
      End If
      '2023/3/27 END
   '暫收款
   Else
      stSales = Text19
      stSubject = "銷退通知 -> " & Text20 & " " & Text23
      stContent = stContent & vbCrLf & "暫收單號：" & Text1
      stContent = stContent & vbCrLf & "客戶名稱：" & Text20 & " " & Text23
      stContent = stContent & vbCrLf & "銷退日期：" & MaskEdBox1.Text
      stContent = stContent & vbCrLf & "暫收款金額：" & Format(Val(Format(Text21)), "#,##0")
   End If
   
   'Modify By Sindy 2023/3/27 + strRecvNo
   PUB_SendMail strUserNum, stSales, strRecvNo, stSubject, stContent
End Sub
'Add by Morgan 2006/10/31 檢查寄送方式
Public Function CheckSendOpt(p_Opt As Integer) As Boolean
   Dim stSalesNo As String, stOffice As String
   
   If Option1.Value = True Then
      stSalesNo = "" & adoacc0k0("a0k20")
   Else
      'Modify by Moran 2009/2/10
      'stSalesNo = "" & adoacc0t0("a0t05")
      stSalesNo = Text19
      'end 2009/2/10
   End If
   stOffice = PUB_GetST06(stSalesNo)
   If p_Opt = 1 And stOffice = "1" Then
      MsgBox "北所智權人員不可選寄分所！"
   ElseIf p_Opt = 2 And stOffice <> "1" Then
      MsgBox "非北所智權人員不可選交智權人員！"
   Else
      CheckSendOpt = True
   End If
End Function

'Add by Morgan 2011/4/13
'收文分配查詢
Public Sub CheckAssign()
   If m_CheckAssign = True Then
      m_CheckAssign = False
      If m_AssignNo <> "" Then
         'Modify by Morgan 2011/10/12 考慮拆收據改抓 0j0
         'strExc(0) = "select cp09 from caseprogress where cp60 in ('" & m_AssignNo & "')" & _
            " and exists(select * from acc0n0 where a0n01=cp09)"
         strExc(0) = "select a0j01 from acc0j0 where a0j13 in ('" & m_AssignNo & "')" & _
            " and exists(select * from acc0n0 where a0n01=a0j01)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_Assigning = True
            RsTemp.MoveFirst
            m_EditList = Split(RsTemp.GetString(, , , ","), ",")
            Frmacc11l0.m_sAssignNo = GetNextNo
            Frmacc11l0.m_sCallType = "E"
            Set Frmacc11l0.m_fCallForm = Me
            Frmacc11l0.Show
            Me.Visible = False
         End If
      End If
   End If
End Sub
'讀取尚未維護收文分配資料的收文號
Private Function GetNextNo() As String
   Dim ii As Integer
   For ii = LBound(m_EditList) To UBound(m_EditList)
      If m_EditList(ii) <> "" Then
         GetNextNo = m_EditList(ii)
         m_EditList(ii) = ""
         Exit For
      End If
   Next
End Function

'Add by Morgan 2011/10/11
'改寫為函數便於重複使用並加申請國家參數
Private Function GetCaseName(pCaseNo As String, Optional ByRef pNation As String) As String
   
   If pCaseNo = "" Then Exit Function
   If adoquery.State = adStateOpen Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select nvl(pa07, nvl(pa06, pa05)),pa09 from patent where " & ChgPatent(pCaseNo) & _
                 "union select nvl(tm07, nvl(tm06, tm05)),tm10 from trademark where " & ChgTradeMark(pCaseNo) & _
                 "union select nvl(lc07, nvl(lc06, lc05)),lc15 from lawcase where " & ChgLawcase(pCaseNo) & _
                 "union select hc06,'000' pa09 from hirecase where " & ChgHirecase(pCaseNo) & _
                 "union select nvl(sp07, nvl(sp06, sp05)),sp09 from servicepractice where " & ChgService(pCaseNo), adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         GetCaseName = MsgText(601)
      Else
         GetCaseName = adoquery.Fields(0).Value
      End If
      pNation = "" & adoquery(1)
   Else
      GetCaseName = ""
   End If
   adoquery.Close
End Function

'Add by Morgan 2011/10/20
Public Sub RestoreData()
   strSql = "update caseprogress set (cp77,cp78)=(select nvl(sum(a1u07),0)+nvl(sum(a1u09),0)" & _
         ",nvl(sum(a1u08),0)+nvl(sum(a1u10),0) from acc1u0 where a1u03=cp09 and a1u01<>'" & Text2 & "')" & _
         " where cp09 in (select a1u03 from acc1u0 where a1u01='" & Text2 & "')"
   adoTaie.Execute strSql, intI
   
   strSql = "update caseprogress set cp79 = nvl(cp16, 0) - nvl(cp75, 0) - nvl(cp77, 0) + nvl(cp78, 0)" & _
      " where cp09 in (select a1u03 from acc1u0 where a1u01='" & Text2 & "')"
   adoTaie.Execute strSql, intI
   
   'Added by Morgan 2012/4/19 更新進度檔已扣繳金額
   strSql = "update caseprogress set cp76=(select nvl(sum(a1u06),0) from acc1u0 where a1u03=cp09) " & _
      " where cp09 in (select a1u03 from acc1u0 where a1u01='" & Text2 & "')"
   adoTaie.Execute strSql, intI
   'end 2012/4/19
   
   
If adoacc0k0.Fields("a0k11") <> "J" Then 'Added by Morgan 2014/1/3 排除J公司
   '還原 acc1v0 資料
   If adoquery.State = adStateOpen Then adoquery.Close
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select a1u02,a1u03,a1v01,a1v02 from acc1u0,acc1v0 where a1u01='" & Text2 & "' and a1v01(+)=a1u03 and a1v02(+)=a1u02", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoquery.EOF = False
      If IsNull(adoquery.Fields("a1v01")) Then
         'Modified by Morgan 2011/12/27 取消 a0j20,a0j21
         strSql = "insert into acc1v0(a1v01,a1v02,a1v03,a1v04,a1v05,a1v06,a1v07,a1v09,a1v12,a1v13,a1v17,a1v18)" & _
            " select a0j01 a1v01,a0j13 a1v02,a0k11 a1v03" & _
            ",0.1*(nvl(a0j09,0)-nvl(x2,0)+decode(a0j07,'Y',nvl(a0j10,0)-nvl(x3,0),0)) a1v04" & _
            ",nvl(a0k13,'N') a1v05,nvl(x1,0) a1v06" & _
            ",0.1*(nvl(a0j09,0)-nvl(x2,0)+decode(a0j07,'Y',nvl(a0j10,0)-nvl(x3,0),0))-nvl(x1,0) a1v07" & _
            ",a0k16 a1v09,getcp10desc(cp01,cp10,a0j04) a1v12,na03 a1v13,y1 a1v17,decode(sign(x1),1,'1') a1v18" & _
            " From acc0j0,acc0k0,(select a1u02,a1u03,sum(a1u06) x1,sum(a1u07) x2,sum(a1u09) x3 from acc1u0" & _
            " where a1u02='" & adoquery("a1u02") & "' and a1u03='" & adoquery("a1u03") & "' and a1u01<>'" & Text2 & "'" & _
            " group by a1u02,a1u03) x,(select a0m02,max(a0m03) y1 from acc0m0 where a0m02='" & Text1 & "'" & _
            " group by a0m02) y,caseprogress,nation" & _
            " where  a0j01='" & adoquery("a1u03") & "' and a0j13='" & adoquery("a1u02") & "'" & _
            " and a0k01(+)=a0j13 and a1u03(+)=a0j01 and a1u02(+)=a0j13 and a0m02(+)=a0j13" & _
            " and exists(select * from acc1u0 where a1u02(+)=a0j13 and a1u03=a0j01 and substr(a1u01,1,1)='F')" & _
            " and cp09(+)=a0j01 and na01(+)=a0j04 "
            
      Else
         strSql = "update acc1v0 set (a1v04,a1v06,a1v07)=(" & _
            " select 0.1*(nvl(max(a0j09),0)-nvl(sum(a1u07),0)+decode(max(a0j07),'Y',nvl(max(a0j10),0)-nvl(sum(a1u09),0),0)) a1v04" & _
            ",nvl(sum(a1u06),0) a1v06" & _
            ",0.1*(nvl(max(a0j09),0)-nvl(sum(a1u07),0)+decode(max(a0j07),'Y',nvl(max(a0j10),0)-nvl(sum(a1u09),0),0))-nvl(sum(a1u06),0) a1v07" & _
            " from acc0j0,acc1u0" & _
            " where a0j01=a1v01 and a0j13=a1v02 and a1u02(+)=a1v02 and a1u03(+)=a1v01 and a1u01(+)<>'" & Text2 & "')" & _
            " where a1v01='" & adoquery("a1u03") & "' and a1v02='" & adoquery("a1u02") & "'"
            
      End If
      adoTaie.Execute strSql, intI
      adoquery.MoveNext
   Loop
   adoquery.Close
End If 'Added by Morgan 2014/1/3
   
   adoTaie.Execute "delete from acc1u0 where a1u01 = '" & Text2 & "'", intI
End Sub

'Add by Morgan 2004/9/29
Public Function EditCheck() As Boolean
   'Added by Morgan 2014/1/3
   '已開發票檢查銷退折讓是否已申報
   If Text4 <> "" Then
      strExc(0) = "select * from acc460 where a4601='" & Text2 & "' and a4606>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "銷退折讓已申報不可修改或刪除！", vbExclamation
         Exit Function
      End If
   End If
   'end 2014/1/3
   
   If Text17.Text = "1" And Text22.Text <> "" Then
      strExc(0) = "SELECT A0O11 FROM ACC0O0 WHERE A0O01='" & Text22.Text & "' union SELECT A0O11 FROM ACC0t0,ACC0O0 WHERE A0t01='" & Text28.Text & "' and a0o01=a0t07 order by 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      With RsTemp
      If .RecordCount = 0 Then
         MsgBox "無法讀取轉出單據編號的應付款資料，不可修改或刪除！"
         Exit Function
      ElseIf "" & .Fields("A0O11").Value <> "" Then
         MsgBox "此筆資料已付款，不可修改或刪除！"
         Exit Function
      ElseIf .RecordCount > 1 Then
         MsgBox "轉出單號2已轉應付，不可修改或刪除！"
         Exit Function
      End If
      End With
   End If
   SetLOS02 'Added by Morgan 2022/6/21
   EditCheck = True
End Function

Public Sub Frmacc1190_Delete()
   Dim douFullTax As Double
   Dim bolTrans As Boolean

   
On Error GoTo Checking
   
   With Frmacc1190
      'Added by Lydia 2021/08/24 記錄操作的本所案號m_CaseNo、銷退單號m_KeyNo、單據編號m_BillNo
      m_CaseNo = "": m_BillNo = "": m_KeyNo = ""
      If cmdDot.Visible = True Then
          m_CaseNo = .cboCaseNo.Text
          m_KeyNo = .Text2.Text
          m_BillNo = .Text1.Text
      End If
      'end 2021/08/24
      '收據
      If .Option1.Value Then
         If DeleteCheck("select a0s01 from acc0s0 where a0s01 = '" & .Text2 & "'") = MsgText(603) Then
            Exit Sub
         End If
         'Add by Morgan 2011/10/20
         adoTaie.BeginTrans
         bolTrans = True
         
         .RestoreData '考慮拆收據改批次還原 cp 資料,要原 1u0 資料需先執行
         'end 2011/10/20
         
         'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
         adoTaie.Execute "delete from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'"
         adoTaie.Execute "delete from acc0s0 where a0s01 = '" & .Text2 & "'"
         adoTaie.Execute "delete from acc1u0 where a1u01 = '" & .Text2 & "'"
         
         adoTaie.Execute "update acc0k0 set a0k10 = null where a0k01 = '" & .Text1 & "'"
         
         'Added by Morgan 2014/1/3
         '已開發票
         If .Text4 <> "" Then
            '刪除 銷退折讓明細檔
            adoTaie.Execute "delete from acc460 where a4601 = '" & .Text2 & "'", intI
            '發票沒有任何銷退紀錄時清除A4309
            adoTaie.Execute "update acc430 set a4309=null where a4301 = '" & .Text4 & "' and not exists(select * from acc0s0 where a0s26=a4301)", intI
         End If
         PUB_UpdateReceiptStatus .Text1 'Added by Morgan 2014/1/17 更新收據結清與介紹獎金發放日期
         'end 2014/1/3
         
         'Add by Morgan 2011/10/20
         adoTaie.CommitTrans
         bolTrans = False
         'end 2011/10/20
         
         .m_CheckAssign = True 'Add by Morgan 2011/5/30
         
'Remove by Morgan 2011/10/18 改由上面程式以批次方式執行
'--舊程式已刪除--
         
         .adoacc0s0.Requery
         If .adoacc0s0.RecordCount <> 0 Then
            .adoacc0s0.MoveFirst
            Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0s0.Bookmark & MsgText(35) & .adoacc0s0.RecordCount
         Else
            StatusClear
         End If
         
      '暫收款
      Else
         If DeleteCheck("select a0t01 from acc0t0 where a0t01 = '" & .Text1 & "'") = MsgText(603) Then
            Exit Sub
         End If
         
         'Add by Morgan 2011/10/20
         adoTaie.BeginTrans
         bolTrans = True
         'end 2011/10/20
         
         'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
         adoTaie.Execute "delete from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'"
         adoTaie.Execute "delete from acc0s0 where a0s01 = '" & .Text2 & "'"
         adoTaie.Execute "delete from acc1u0 where a1u01 = '" & .Text2 & "'"
         
         '2005/7/13 MODIFY BY SONIA
         'adoTaie.Execute "update acc0t0 set a0t10 = '' where a0t01 = '" & .Text1 & "'"
         adoTaie.Execute "update acc0t0 set a0t10 = '',A0T09=NULL where a0t01 = '" & .Text1 & "'"
         '2005/7/13 END
         
         PUB_UpdateReceiptStatus .Text1 'Added by Morgan 2015/5/15
         
         'Add by Morgan 2011/10/20
         adoTaie.CommitTrans
         bolTrans = False
         'end 2011/10/20

         .adoacc0t0.Requery
         If .adoacc0t0.RecordCount <> 0 Then
            .adoacc0t0.MoveFirst
            Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0t0.Bookmark & MsgText(35) & .adoacc0t0.RecordCount
         Else
            StatusClear
         End If
      End If
   End With
   
   'Added by Lydia 2021/08/24 法務案有分配工作點數，檢查剩餘點數與工作點數不符時進入工作點數分配畫面
   If bolTrans = False And m_KeyNo <> "" And m_BillNo <> "" And m_CaseNo <> "" Then
       If ChkA1n05PBalence(m_CaseNo, m_BillNo) = False Then
           Call ProcShowPoint(m_CaseNo, m_KeyNo, m_BillNo)
       End If
   End If
   'end 2021/08/24
   
   Exit Sub
   
Checking:
   If bolTrans = True Then adoTaie.RollbackTrans 'Add by Morgan 2011/10/20
   MsgBox Err.Description, , MsgText(5)
   
End Sub

Public Sub Frmacc1190_Clear()
   With Frmacc1190
      If .Option1.Value Then
         .Text1 = "E"
      Else
         .Text1 = "J"
      End If
      TextInverse .Text1
      .Text2 = ""
      MaskEdBox1.Tag = "" 'Add by Amy 2015/01/21
      If .MaskEdBox1.Text = MsgText(29) Or .MaskEdBox1.Text = MsgText(601) Then
         .MaskEdBox1.Mask = ""
         .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
         .MaskEdBox1.Mask = DFormat
      End If
      .Text22 = ""
      .Text3 = ""
      .Text27 = ""
      'Modify by Morgan 2011/10/11
      '.Text4 = ""
      .cboCaseNo.Clear
      .MaskEdBox2.Mask = ""
      .MaskEdBox2.Text = ""
      .MaskEdBox2.Mask = DFormat
      .Text5 = ""
      .Text6 = ""
      .Text7 = ""
      .Text8 = ""
      .Text9 = ""
      .Text12 = ""
      .Text10 = ""
      .Text11 = ""
      .Text13 = ""
      .Text14 = ""
      .Text15 = ""
      .Text16 = ""
      .Text17 = ""
      .MaskEdBox3.Mask = ""
      .MaskEdBox3.Text = CFDate(ACDate(ServerDate))
      .MaskEdBox3.Mask = DFormat
      .Text19 = ""
      .Text20 = ""
      .Text21 = ""
      .Text26 = ""
      .Text24 = ""
      .Text18 = ""
      .Text23 = ""
      .Text25 = ""
      .Text28 = "" 'Add by Morgan 2007/1/5
      'Modified by Lydia 2021/08/24
      '.Text1.SetFocus
      If .Enabled = True Then .Text1.SetFocus
      
      'Added by Morgan 2015/12/8
      .Label34.Visible = False
      SetCmdCaseNo
      'end 2015/12/8
   End With
End Sub

'Modified by Morgan 2024/11/11 摘要加單引號控制,客戶名稱可能會有 Ex:X00403140
Public Sub Frmacc1190_Save()
Dim strYes As String
Dim strNo As String
Dim douService As Double
Dim douTax As Double
Dim strSerialNo As String
Dim douAService As Double
Dim douATax As Double
Dim lngTax As Double
Dim strMan As String
Dim strCust As String
Dim strRemark As String
Dim strRemark2112 As String   '2005/5/12 ADD BY SONIA
Dim douDfee As Double
Dim douDtax As Double
Dim strSalesNo As String
Dim strAccNo As String
Dim strDocuNo As String
Dim strDYes As String
Dim strDept As String
'Add by Morgan 2004/4/5
'退費扣繳金額
Dim stA1U06 As String
'轉應付借貸方摘要
Dim strRemark1 As String, strRemark2 As String
Dim douFee As Double
'Ken 92/12/22 新增扣繳明細資料
Dim douFullTax As Double
Dim douPayTax As Double
Dim douNonePayTax As Double
Dim strPayMethod As String
Dim strYear As String
Dim strCPN As String
Dim strANN As String
'Added by Morgan 2014/1/2
Dim strCompNo As String '公司別
Dim strMaxFee As String '最大規費
Dim strMaxFeeNo As String '最大規費科目項次
Dim bolShowForm As Boolean 'Added by Morgan 2014/8/７是否有彈出畫面
Dim strMsg As String 'Add by Amy 2014/10/28
      
      'Added by Lydia 2021/08/19 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
      If PUB_ChkUniText(Me, , True, "TextBox") = False Then
          Text27.SetFocus  '備註
          Text27_GotFocus
          strControlButton = MsgText(602)
          Exit Sub
      End If
      'end 2021/08/19
      
On Error GoTo Checking

   With Frmacc1190
      .m_KeepItem = "" 'Added by Morgan 2015/8/11
      'Added by Lydia 2021/08/24 記錄操作的本所案號m_CaseNo、銷退單號m_KeyNo、單據編號m_BillNo
      m_CaseNo = "": m_BillNo = "": m_KeyNo = ""
      If cmdDot.Visible = True Then
          m_CaseNo = .cboCaseNo.Text
          m_KeyNo = .Text2.Text
          m_BillNo = .Text1.Text
      End If
      'end 2021/08/24
      
      'Added by Morgan 2014/1/3
      If .Option1.Value Then
         strCompNo = "" & .adoacc0k0("a0k11")
      Else
         'Modified by Morgan 2014/8/5
         'strCompNo = "" & .adoacc0t0("a0t18")
         strCompNo = m_A0T18
         'end 2014/8/5
      End If
      'modify by sonia 2020/7/28 加L公司改寫法
      'If strCompNo <> "J" Then strCompNo = "1" 'Added by Morgan 2014/1/13 除J公司外其他存 1 --秀玲
      If strCompNo < "A" Then strCompNo = "1" 'Added by Morgan 2014/1/13 除J公司外其他存 1 --秀玲
      'end 2014/1/3
      
      'Modify by Amy 2014/10/28 +銷退日期必填及系統日檢查
      If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
            MsgBox .Label3 & MsgText(52), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
      End If
      If DateCheck(.MaskEdBox1.Text) = MsgText(603) Then
            MsgBox .Label3 & MsgText(63), , MsgText(5)
            strControlButton = MsgText(602)
            .MaskEdBox1.SetFocus
            Exit Sub
       End If
      
       '系統日檢查
       If MaskEdBox1.Enabled = True And ((.Option1.Value And Text3 <> "1") Or Option2.Value) Then
            '收據:類別為1.銷帳不會寫入acc1p0 , 故不用判斷/暫收:類別為鎖住不可輸,但a0s04固定存1
            'Modify by Amy 2015/01/21 +新增時及有修改才檢查
            If strSaveConfirm = "A" Or (strSaveConfirm = "E" And Val(FCDate(MaskEdBox1.Text)) <> Val(MaskEdBox1.Tag)) Then
                If ChkWorkData(strCompNo, DBDATE(MaskEdBox1), strMsg) = False Then
                    MsgBox Label3 & strMsg, , MsgText(5)
                    strControlButton = MsgText(602)
                    MaskEdBox1.SetFocus
                    Exit Sub
                End If
            End If
      End If
      'end 2014/10/28
      
      'Added by Morgan 2024/11/12 --瑞婷
      '稅款退費金額不應為負值
      If Val(Text26) < 0 Then
         strMsg = "不應為負值!!"
         MsgBox Label33 & strMsg, , MsgText(5)
         strControlButton = MsgText(602)
         MaskEdBox1.SetFocus
         Exit Sub
      End If
      'end 2024/11/12
      
      'Added by Morgan 2025/7/8
      If .Option1.Value And strCompNo = "J" And Left(cboCaseNo, 3) = "ACS" Then
         If Val(Text15) > 0 Then
            If Val(Text16) = 0 Then
               MsgBox "智權公司ACS案件在退費時,不能沒有輸退費規費(稅)!!", vbCritical
               strControlButton = MsgText(602)
               Text16.SetFocus
               Exit Sub
            Else
               intI = Round(Val(Text15) * 0.05)
               If Val(Text16) <> intI Then
                  If MsgBox("退費規費(稅)錯誤!!應該為 " & intI & " (" & Val(Text15) & "x0.05)。" & vbCrLf & vbCrLf & "是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                     strControlButton = MsgText(602)
                     Text16.SetFocus
                     Exit Sub
                  End If
               End If
            End If
         End If
      End If
      'end 2025/7/8

      '收據編號
      If .Option1.Value Then
         If .Text2 = MsgText(601) Then
            MsgBox .Label2 & MsgText(10), , MsgText(5)
            strControlButton = MsgText(602)
            .Text2.SetFocus
            Exit Sub
         Else
            'Modify by Amy 2014/10/28 原銷退日期判斷往上搬
            'Added by Morgan 2013/12/31
            '收款日期不可大於銷退日期
            strExc(0) = "select sqldatet(a0l02) from acc0m0,acc0l0 where a0m02='" & .Text1 & "' and a0l01(+)=a0m01 and a0l02>" & Val(FCDate(.MaskEdBox1.Text))
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "收據/請款單 " & .Text1 & " 的收款日期(" & RsTemp(0) & ") 晚於本次銷退日期(" & .MaskEdBox1 & ")，不可作業！", vbCritical
               strControlButton = MsgText(602)
               If MaskEdBox1.Enabled = True Then .MaskEdBox1.SetFocus 'Modify by Amy 2014/10/28 +if
               Exit Sub
            End If
            'end 2013/12/31
            
            'If strSaveConfirm = MsgText(3) Then 'Remove by Morgan 2011/10/17 修改也要判斷
               Select Case .Text3
                  Case "1"
                     If (Val(.Text14) + Val(.Text12)) > (Val(.Text6) + Val(.Text7)) Then
                        MsgBox MsgText(108), , MsgText(5)
                        strControlButton = MsgText(602)
                        .Text14.SetFocus
                        Exit Sub
                     End If
                     
                     'Add by Morgan 2011/10/20
                     If Val(.Text14) > Val(.Text13) Then
                        MsgBox "銷帳金額不可大於未收金額", , MsgText(5)
                        strControlButton = MsgText(602)
                        .Text14.SetFocus
                        Exit Sub
                     End If
                     If Val(.Text15) + Val(.Text16) > 0 Then
                        MsgBox "選銷帳不可輸入退費金額!!"
                        strControlButton = MsgText(602)
                        .Text15.SetFocus
                        Exit Sub
                     End If
                     'end 2011/10/20
                     
                  Case "2", "3"
                     'Add by Morgan 2011/10/20
                     If Val(.Text8) + Val(.Text9) = 0 Then
                        MsgBox "尚未收款無法退費!!"
                        strControlButton = MsgText(602)
                        .Text3.SetFocus
                        Exit Sub
                     End If
                     If Val(.Text15) + Val(.Text16) = 0 Then
                        MsgBox "請輸入退費金額!!"
                        strControlButton = MsgText(602)
                        .Text15.SetFocus
                        Exit Sub
                     End If
                     'end 2011/10/20
                                          
                     If Val(.Text15) > Val(.Text8) Then
                        MsgBox MsgText(109), , MsgText(5)
                        strControlButton = MsgText(602)
                        .Text15.SetFocus
                        Exit Sub
                     End If
                     If Val(.Text16) > Val(.Text9) Then
                        MsgBox MsgText(109), , MsgText(5)
                        strControlButton = MsgText(602)
                        .Text16.SetFocus
                        Exit Sub
                     End If
                     
                     'Added by Morgan 2023/3/21
                     If Val(.Text15) > Val(.Text8) - Val(.Text10) Then
                        MsgBox "退費服務費不可大於未退金額!!"
                        strControlButton = MsgText(602)
                        .Text15.SetFocus
                        Exit Sub
                     End If
                     If Val(.Text16) > Val(.Text9) - Val(.Text11) Then
                        MsgBox "退費規費不可大於未退金額!!"
                        strControlButton = MsgText(602)
                        .Text16.SetFocus
                        Exit Sub
                     End If
                     'end 2023/3/21
                     
                     'Add by Morgan 2004/4/6
                     '全額退費時，稅款只能為 0 或全退
                     If Val(.Text6) = Val(.Text15) And Val(.Text7) = Val(.Text16) Then
                        If Val(.Text26) <> 0 And Val(.Text24) <> Val(.Text26) Then
                           strControlButton = MsgText(602)
                           MsgBox "全額退費時，稅款只能為 0 或全退！", vbExclamation
                           Exit Sub
                        End If
                     'Added by Morgan 2014/1/6
                     ElseIf Val(.Text24) = 0 And Val(.Text26) > 0 Then
                        strControlButton = MsgText(602)
                        MsgBox "未扣繳不可退稅款！", vbExclamation
                        Exit Sub
                     End If
                     
                     'Added by Morgan 2015/4/10
                     'Modified by Morgan 2015/5/6 有扣單照舊 --瑞婷
                     If bolTaxed = False Then
                        If Val(.Text24) > 0 And Val(.Text26) = 0 And (Val(.Text15) > 0 Or (Val(.Text16) > 0 And m_A0J07 = "Y")) Then
                           'Modified by Morgan 2015/5/6 改只提醒不強制 --瑞婷
                           'MsgBox ("本收據已有扣繳事實,退稅金額請記得輸入！", vbExclamation
                           If MsgBox("本收據已有扣繳事實，但退稅金額未輸入，是否要繼續？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                              strControlButton = MsgText(602)
                              Exit Sub
                           End If
                        End If
                     
                     'Removed by Morgan 2015/10/27 改存檔時有扣單的不刪除
                     ''Added by Morgan 2015/8/10
                     ''有扣單時不可全銷退,否則刪除1v0後扣單號會不見且有扣繳沒收款亦不合理
                     'ElseIf Val(.Text6) = Val(.Text15) And Val(.Text7) = Val(.Text16) Then
                     '   MsgBox "本收據已有扣單，不可全額銷退！", vbExclamation
                     '   strControlButton = MsgText(602)
                     '   Exit Sub
                     ''end 2015/8/10
                     'end 2015/10/27
                     
                     End If
                     'end 2015/4/10
                     
                  Case Else
                     MsgBox MsgText(199), , MsgText(5)
                     strControlButton = MsgText(602)
                     .Text3.SetFocus
                     Exit Sub
               End Select
            'End If
            Select Case .Text17
               'Modified by Morgan 2023/11/8 +選項3
               Case "1", "2", "3"
               Case Else
                  If .Text3 <> "1" Then
                     'MsgBox MsgText(200), , MsgText(5)
                     MsgBox MsgText(199), , MsgText(5)
                     strControlButton = MsgText(602)
                     .Text17.SetFocus
                     Exit Sub
                  End If
            End Select
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            .adoquery.Open "select ax210 from acc1p0, acc021 where a1p01 = ax201 and a1p22 = ax202 and a1p03 = ax203 and ax210 is not null and a1p04 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
            If .adoquery.RecordCount <> 0 Then
               MsgBox MsgText(14), , MsgText(5)
               strControlButton = MsgText(602)
               .adoquery.Close
               Exit Sub
            End If
            .adoquery.Close
            .adoquery.CursorLocation = adUseClient
            .adoquery.Open "select a0o11 from acc0o0 where a0o01 = '" & .Text22 & "' and a0o11 is not null", adoTaie, adOpenStatic, adLockReadOnly
            If .adoquery.RecordCount <> 0 Then
               MsgBox "此應付款資料已付款,不可修改 !", , MsgText(5)
               strControlButton = MsgText(602)
               .adoquery.Close
               Exit Sub
            End If
            .adoquery.Close
         End If
         
         adoTaie.BeginTrans
         
         m_IsOpen = False 'Added by Morgan 2024/11/12 改一律重抓單筆資料,後面要 Resync 時才不會跑太久
         
         'Add by Morgan 2011/10/14
         If .m_IsOpen = False Then
            If .adoacc0s0.State = adStateOpen Then .adoacc0s0.Close
            .adoacc0s0.CursorLocation = adUseClient
            .adoacc0s0.Open "select * from acc0s0 where a0s01 = '" & .Text2 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
         End If
         'end 2011/10/14
         
         If strSaveConfirm = MsgText(3) Then
            If .adoacc0s0.RecordCount <> 0 Then
               .adoacc0s0.Find "a0s01 = '" & .Text2 & "'", 0, adSearchForward, 1
               If .adoacc0s0.EOF = False Then
                  MsgBox MsgText(9), , MsgText(5)
                  strControlButton = MsgText(602)
                  .Text2.SetFocus
                  adoTaie.RollbackTrans
                  Exit Sub
               End If
            End If
            .adoacc0s0.AddNew
         End If
         
         .adoacc0s0.Fields("a0s01").Value = .Text2
         If .Text1 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s02").Value = .Text1
         Else
            .adoacc0s0.Fields("a0s02").Value = Null
         End If
         If .MaskEdBox1.Text <> MsgText(601) And .MaskEdBox1.Text <> MsgText(29) Then
            .adoacc0s0.Fields("a0s03").Value = Val(FCDate(.MaskEdBox1.Text))
         Else
            .adoacc0s0.Fields("a0s03").Value = Null
         End If
         If .Text3 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s04").Value = .Text3
         Else
            .adoacc0s0.Fields("a0s04").Value = Null
         End If
         If .Text27 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s18").Value = .Text27
         Else
            .adoacc0s0.Fields("a0s18").Value = Null
         End If
         If .Text14 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s05").Value = Val(.Text14)
         Else
            .adoacc0s0.Fields("a0s05").Value = 0
         End If
         If .Text15 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s06").Value = Val(.Text15)
         Else
            .adoacc0s0.Fields("a0s06").Value = 0
         End If
         If .Text16 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s07").Value = Val(.Text16)
         Else
            .adoacc0s0.Fields("a0s07").Value = 0
         End If
         If .Text17 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s08").Value = .Text17
         Else
            .adoacc0s0.Fields("a0s08").Value = Null
         End If
         If .MaskEdBox3.Text <> MsgText(601) And .MaskEdBox3.Text <> MsgText(29) Then
            .adoacc0s0.Fields("a0s09").Value = Val(FCDate(.MaskEdBox3.Text))
         Else
            .adoacc0s0.Fields("a0s09").Value = Null
         End If
         If .Text26 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s17").Value = Val(.Text26)
         Else
            .adoacc0s0.Fields("a0s17").Value = 0
         End If
         If strCon1 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s19").Value = strCon1
            'strCon1 = MsgText(601)
         End If
         If strCon2 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s22").Value = strCon2
            'strCon2 = MsgText(601)
         End If
         If strCon3 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s20").Value = strCon3
            'strCon3 = MsgText(601)
         End If
         If strCon4 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s21").Value = strCon4
            'strCon4 = MsgText(601)
         End If
         
         'Add by Morgan 2009/7/1
         '是否退公開費
         If .Check1.Value = 1 Then
            .adoacc0s0.Fields("a0s24").Value = "Y"
         Else
            .adoacc0s0.Fields("a0s24").Value = Null
         End If
         
         If strSaveConfirm = MsgText(3) Then
            .adoacc0s0.Fields("a0s11").Value = Val(strSrvDate(2))
            .adoacc0s0.Fields("a0s12").Value = ServerTime
            .adoacc0s0.Fields("a0s13").Value = strUserNum
         Else
            .adoacc0s0.Fields("a0s14").Value = Val(strSrvDate(2))
            .adoacc0s0.Fields("a0s15").Value = ServerTime
            .adoacc0s0.Fields("a0s16").Value = strUserNum
         End If
         
         If .adoquery.State = 1 Then .adoquery.Close
         .adoquery.CursorLocation = adUseClient
         'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
         .adoquery.Open "select a1p22 from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount <> 0 Then
            If IsNull(.adoquery.Fields("a1p22").Value) Then
               strDocuNo = "null"
               strDYes = "null"
            Else
               strDocuNo = "'" & .adoquery.Fields("a1p22").Value & "'"
               strDYes = "'Y'"
            End If
         Else
            strDocuNo = "null"
            strDYes = "null"
         End If
         .adoquery.Close
         
         'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
         adoTaie.Execute "delete from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'"
         
         '銷+退
         If .Text3 = "3" Then
            '轉應付
            If .Text17 = "1" Then
               If Mid(.Text22, 1, 1) = "G" Then
                  strNo = .Text22
                  adoTaie.Execute "delete from acc0o0 where a0o01 = '" & strNo & "'"
               Else
                  strNo = AutoNo(MsgText(804), 5, 1)
               End If
               
            '轉暫收款
            ElseIf .Text17 = "2" Then
               If .Text22 <> "" And Mid(.Text22, 1, 1) = "J" Then
                  strNo = .Text22
                  adoTaie.Execute "delete from acc0t0 where a0t01 = '" & strNo & "'"
               Else
                  strNo = AutoNo(MsgText(806), 5, 1)
               End If
            End If
         End If
            
         'Added by Morgan 2014/1/2
         '有發票號碼(J公司)
         If .Text4 <> "" Then
            .adoacc0s0.Fields("a0s26").Value = .Text4
            .adoacc0s0.Fields("a0s27").Value = IIf(.Check2.Value = 0, "N", Null)
            'Add by Amy 2023/10/05 若 發票日期 在電子發票上線前,則不需上傳盟立(避免盟立回傳錯誤),直接上Tag
            If Val(adoacc0k0("A4302")) < TranInvoiceDate Then
               .adoacc0s0.Fields("a0s28").Value = "111111"
               .adoacc0s0.Fields("a0s29").Value = "240000"
            End If
            
            '更新發票檔
            strSql = "update acc430 set a4309='Y' where a4301='" & .Text4 & "'"
            adoTaie.Execute strSql, intI
            'modify by sonia 2018/9/13 I10700651只銷E10710001部分,故只抓有銷帳的
            strExc(0) = "select a0j02,na03,decode(sk02,'1','專利','2','商標','其他') Sys,nvl(a0j22,getcp10desc(cp01,cp10,a0j04)) cp10N from acc0j0,nation,caseprogress,systemkind where a0j13='" & .Text1 & "' and na01(+)=a0j04 and cp09(+)=a0j01 and sk01(+)=cp01 order by a0j25,a0j01"
            'strExc(0) = "select a0j02,na03,decode(sk02,'1','專利','2','商標','其他') Sys,nvl(a0j22,getcp10desc(cp01,cp10,a0j04)) cp10N from acc0j0,nation,caseprogress,systemkind,acc1u0 where a0j13='" & .Text1 & "' and na01(+)=a0j04 and cp09(+)=a0j01 and sk01(+)=cp01 and '" & .Text2 & "'=a1u01(+) and a0j01=a1u03 order by a0j25,a0j01"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = RsTemp("na03") & RsTemp("Sys") & "/" & RsTemp("cp10N")
               '銷
               If .Text3 = "1" Then
                  strExc(2) = Val(.Text14)
               '銷+退
               Else
                  strExc(2) = Val(.Text15) + Val(.Text16)
               End If
               '2014/12/2 modify by sonia 個人無稅額
               ''銷售額
               'strExc(3) = Round(Val(strExc(2)) / 1.05)
               ''稅額
               'strExc(4) = Val(strExc(2)) - Val(strExc(3))
               If Not IsNull(.adoacc0k0("A4303")) Then
                  '銷售額
                  strExc(3) = Round(Val(strExc(2)) / 1.05)
                  '稅額
                  strExc(4) = Val(strExc(2)) - Val(strExc(3))
               Else
                  '銷售額
                  strExc(3) = Val(strExc(2))
                  '稅額
                  strExc(4) = 0
               End If
               '2014/12/2 end
               '修改時要先刪除
               If strSaveConfirm = MsgText(4) Then
                  strSql = "delete acc460 where a4601='" & .Text2 & "' and nvl(a4606,0)=0"
                  adoTaie.Execute strSql, intI
               End If
               '新增銷退折讓明細檔
               strSql = "insert into acc460(A4601,A4602,A4603,A4604,A4605) values('" & .Text2 & "','" & strExc(1) & "','" & RsTemp("a0j02") & "'," & strExc(3) & "," & strExc(4) & ")"
               adoTaie.Execute strSql, intI
            End If
            

            
            'Added by Morgan 2014/1/20
            strSalesNo = "" & .adoacc0k0("a0k20")
            strMan = PUB_GetShortName(strSalesNo)
            strRemark = strMan & "/" & MidB("" & .adoacc0k0.Fields("a0k04").Value, 1, 16)
            
            'Modified by Morgan 2015/8/17
            '1.銷帳 2.銷退規費(未結清)
            If .Text3 = "1" Or (.Text3 = "3" And Val(.Text13) > 0) Then
               '有未收款沖帳傳票
               If m_Transfered <> "" Then
                  '非全銷只帶科目,金額讓User自行輸入
                  'Modified by Morgan 2025/7/8 改都要帶金額--瑞婷
                  'If Val(.Text14) = Val(.Text6) + Val(.Text7) Then
                     douService = Round(Val(Text14) / 1.05)
                     douTax = Val(Text14) - douService
                  'Else
                  '   douService = 0
                  '   douTax = 0
                  'End If
                  'end 2025/7/8
                  
                  '應收未收款 2141
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                  'Modified by Morgan 2015/8/4 +a1p17對沖代號(本所案號)
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                     " values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2141', 'TOT', " & douService & ", 0, null, null, null, null, null, '" & strRemark & "/" & .Text4 & "', '" & .adoacc0k0.Fields("a0k03") & "', '" & strSalesNo & "', '" & cboCaseNo & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                  m_KeepItem = m_KeepItem & strSerialNo & ";" 'Added by Morgan 2015/8/11
                  '銷項稅額 2119
                  'Modified by Morgan 2015/5/28 摘要加發票號
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                  'Modified by Morgan 2015/8/4 +a1p17對沖代號(本所案號)
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                     " values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2119', 'TOT', " & douTax & ", 0, null, null, null, null, null, '" & strRemark & "/" & .Text4 & "', '" & .adoacc0k0.Fields("a0k03") & "', '" & strSalesNo & "', '" & cboCaseNo & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                  m_KeepItem = m_KeepItem & strSerialNo & ";" 'Added by Morgan 2015/8/11
                  '應收帳款 1133 (2020/10/15 發現2017/4/5 已改用1141未入帳應收帳款但此處漏改)
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                  'Modified by Morgan 2015/8/4 +a1p17對沖代號(本所案號)
                  'modify by sonia 2017/3/16 未收款沖帳傳票已不用1135應收銷項稅額,所以此處改全額(douService+douTax)
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                     " values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '1141', 'TOT',0, " & douService + douTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "/" & .Text4 & "', '" & .adoacc0k0.Fields("a0k03") & "', '" & strSalesNo & "', '" & cboCaseNo & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                  m_KeepItem = m_KeepItem & strSerialNo & ";" 'Added by Morgan 2015/8/11
                  
'cancel by sonia 2017/3/16 未收款沖帳傳票已不用1135應收銷項稅額,所以此處取消
'                  '應收銷項稅額 1135
'                  'Modified by Morgan 2015/5/28 摘要加發票號
'                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
'                  'Modified by Morgan 2015/8/4 +a1p17對沖代號(本所案號)
'                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
'                     " values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '1135', 'TOT',0, " & douTax & ",null, null, null, null, null, '" & strRemark & "/" & .Text4 & "', '" & .adoacc0k0.Fields("a0k03") & "', '" & strSalesNo & "', '" & cboCaseNo & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
'                  m_KeepItem = m_KeepItem & strSerialNo & ";" 'Added by Morgan 2015/8/11

               End If
            End If
            'end 2015/8/17
            'end 2014/1/20
            
            'Added by Morgan 2015/8/17
            '未結清退規費
            If .Text3 = "3" And Val(.Text13) > 0 And Val(.Text16) > 0 Then
            
'Removed by Morgan 2015/11/10 改結清時借規費扣
'               '預收銷項稅額
'               strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
'               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
'                  " values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2405', 'TOT', 0, 0, null, null, null, null, null, '" & strRemark & "/" & .Text4 & "', '" & .adoacc0k0.Fields("a0k03") & "', '" & strSalesNo & "', '" & cboCaseNo & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
'               m_KeepItem = m_KeepItem & strSerialNo & ";"
'end 2015/11/10
               
            '未收全銷(已收部分)
            ElseIf .Text3 = "1" And Val(.Text8) + Val(.Text9) > 0 And Val(.Text14) = Val(.Text13) Then
               
               'Removed by Morgan 2015/11/10 改結清時借規費扣
               ''純服務費
               'If Val(.Text9) = 0 Then
                  strAccNo = "2201" '應付規費
               'Else
               '   strAccNo = "2405" '預收銷項稅額
               'End If
               'end 2015/11/10
               
               strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                  " values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & strAccNo & "', 'TOT', 0, 0, null, null, null, null, null, '" & ChgSQL(strRemark) & "/" & .Text4 & "', '" & .adoacc0k0.Fields("a0k03") & "', '" & strSalesNo & "', '" & cboCaseNo & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
               m_KeepItem = m_KeepItem & strSerialNo & ";"
               
               If m_Transfered <> "" Then
                  'modify by sonia 2017/3/16 未收款沖帳傳票已不用1135應收銷項稅額,所以此處改1133應收帳款
                  strAccNo = "1133" '應收帳款
               Else
                  strAccNo = "2119" '銷項稅額
               End If
               strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                  " values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & strAccNo & "', 'TOT', 0, 0, null, null, null, null, null, '" & ChgSQL(strRemark) & "/" & .Text4 & "', '" & .adoacc0k0.Fields("a0k03") & "', '" & strSalesNo & "', '" & cboCaseNo & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
               m_KeepItem = m_KeepItem & strSerialNo & ";"
            End If
            'end 2015/8/17
            
         Else
            .adoacc0s0.Fields("a0s26").Value = Null
            .adoacc0s0.Fields("a0s27").Value = Null
         End If
         'end 2014/1/2
         
         .m_CheckAssign = True '收文分配檢查 Add by Morgan 2011/5/30
         
         'Modified by Morgan 2014/1/21 轉應付或暫收款程式合併
         '銷+退
         If .Text3 = "3" Then
            
            If .adoquery.State = 1 Then .adoquery.Close
            .adoquery.CursorLocation = adUseClient
            'Modify by Morgan 2011/10/12 考慮拆收據情形(已收要改抓1u0)
            '.adoquery.Open "select * from caseprogress, acc0k0, acc0j0 where cp60 = a0k01 and cp09 = a0j01 (+) and a0k01 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
            strExc(0) = "select a.*,b.*,c.*,d.*,getcp10desc(cp01,cp10,a0j04) cp10N from acc0k0 a, acc0j0 b, caseprogress c,(select a1u02,a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u06)" & _
               " from acc1u0 where a1u02='" & .Text1 & "' and a1u01<>'" & .Text2 & "'" & _
               " group by a1u02,a1u03) d where a0k01='" & .Text1 & "' and a0j13(+)=a0k01 and a0j02='" & .cboCaseNo & "' and cp09(+)=a0j01 and a1u02(+)=a0j13 and a1u03(+)=a0j01"
            .adoquery.Open strExc(0), adoTaie, adOpenStatic, adLockReadOnly
            If .adoquery.RecordCount <> 0 Then
               '轉應付
               If .Text17 = "1" Then
                  '新增應付資料
                  'Modified by Morgan 2014/1/2 +a0o07(改指定欄位)
                  'adoTaie.Execute "insert into acc0o0 values ('" & strNo & "', '2', '" & IIf(IsNull(.adoQuery.Fields("a0k03").Value), "", .adoQuery.Fields("a0k03").Value) & "', null, " & Val(FCDate(.MaskEdBox3.Text)) & ", " & Val(FCDate(.MaskEdBox3.Text)) & ", '" & .Text2 & "', '" & strCon2 & "', null, '2', '" & strUserNum & "', " & strSrvDate(2) & ", to_char(sysdate,'hh24miss'), null, null, null)"
                  strSql = "insert into acc0o0(a0o01,a0o02,a0o03,a0o04,a0o05,a0o06,a0o09,a0o10,a0o11,a0o19,a0o15,a0o13,a0o14,a0o18,a0o16,a0o17,a0o07) values ('" & strNo & "', '2', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', null, " & Val(FCDate(.MaskEdBox3.Text)) & ", " & Val(FCDate(.MaskEdBox3.Text)) & ", '" & .Text2 & "', '" & strCon2 & "', null, '2', '" & strUserNum & "', " & strSrvDate(2) & ", to_char(sysdate,'hh24miss'), null, null, null,'" & strCompNo & "')"
                  adoTaie.Execute strSql, intI
                  'end 2014/1/2
               '轉暫收款
               ElseIf .Text17 = "2" Then
                  'Modified by Morgan 2014/1/3
                  'adoTaie.Execute "insert into acc0t0 values ('" & strNo & "', '3', " & Val(FCDate(.MaskEdBox1.Text)) & ", " & Val(FCDate(.MaskEdBox3.Text)) & ", '" & .adoacc0k0.Fields("a0k20").Value & "', '" & .adoacc0k0.Fields("a0k03").Value & "', '" & .Text2 & "', " & Val(.Text15) + Val(.Text16) - Val(.Text26) & ", " & Val(FCDate(.MaskEdBox1.Text)) & ", null, '" & .Text27 & "', '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", null, null, null)"
                  adoTaie.Execute "insert into acc0t0(A0T01,A0T02,A0T03,A0T04,A0T05,A0T06,A0T07,A0T08,A0T09,A0T10,A0T17,A0T13,A0T11,A0T12,A0T16,A0T14,A0T15,A0T18) values ('" & strNo & "', '3', " & Val(FCDate(.MaskEdBox1.Text)) & ", " & Val(FCDate(.MaskEdBox3.Text)) & ", '" & .adoacc0k0.Fields("a0k20").Value & "', '" & .adoacc0k0.Fields("a0k03").Value & "', '" & .Text2 & "', " & Val(.Text15) + Val(.Text16) - Val(.Text26) & ", " & Val(FCDate(.MaskEdBox1.Text)) & ", null, '" & .Text27 & "', '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", null, null, null,'" & strCompNo & "')"
               End If
               
               '智權人員代碼
               strSalesNo = "" & .adoquery.Fields("a0k20").Value
               '智權人員簡稱
               strMan = PUB_GetShortName(strSalesNo)
               '收據抬頭
               strCust = MidB("" & .adoquery.Fields("a0k04").Value, 1, 16)
               
               douService = 0
               douTax = 0
               Do While .adoquery.EOF = False
                  
                  '借方摘要
                  'Modified by Morgan 2011/12/27 取消 a0j20
                  If .Text17 = "1" Then
                     strRemark = strMan & "/" & Left(strCust, 4) & "/" & "" & .adoquery.Fields("cp10N").Value & "/" & .Text2
                  ElseIf .Text17 = "2" Then
                     strRemark = strMan & "/" & Left(strCust, 5) & "/" & "" & .adoquery.Fields("cp10N").Value & "/" & .Text2
                  'Added by Morgan 2023/11/8
                  ElseIf .Text17 = "3" Then
                     strRemark = strMan & "/" & Left(strCust, 5) & "/" & "" & .adoquery.Fields("cp10N").Value & "/" & .Text2
                  End If
                  
                  'If douAService <> Val(.Text15) Then
                     'Modified by Morgan 2013/12/19 會有J公司,取消 a1p01='1' 條件
                     strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                     .adoaccsum.CursorLocation = adUseClient
                     .adoaccsum.Open "select cpm11 from casepropertymap where cpm01 = '" & .adoquery.Fields("cp01").Value & "' and cpm02 = '" & .adoquery.Fields("cp10").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                     If .adoaccsum.RecordCount <> 0 Then
                        'Modify by Morgan 2011/10/17 考慮拆收據情形
                        'If (douAService + IIf(IsNull(.adoquery.Fields("cp73").Value), 0, .adoquery.Fields("cp73").Value)) <= Val(.Text15) Then
                        '   douService = IIf(IsNull(.adoquery.Fields("cp73").Value), 0, .adoquery.Fields("cp73").Value)
                        If (douAService + Val("" & .adoquery.Fields("a1u04"))) <= Val(.Text15) Then
                           douService = Val("" & .adoquery.Fields("a1u04"))
                        'end 2011/10/17
                        Else
                           douService = Val(.Text15) - douAService
                        End If
                        If IsNull(.adoaccsum.Fields("cpm11").Value) Then
                           strAccNo = "XXX"
                        Else
                           strAccNo = .adoaccsum.Fields("cpm11").Value
                        End If
                        'Modify by Morgan 2008/6/10 改判斷非台灣
                        'If .adoquery.Fields("a0j04").Value = "020" Then
                        If .adoquery.Fields("a0j04").Value <> "000" Then
                           If "" & .adoquery.Fields("cp01") = "P" Then
                              strAccNo = "411103"
                           ElseIf "" & .adoquery.Fields("cp01") = "T" Then
                              strAccNo = "410103"
                           End If
                        End If

                        If IsNull(.adoquery.Fields("cp01").Value) Then
                           strDept = "null"
                        Else
                           '93.11.25 MODIFY BY SONIA
                           'strDept = "'" & .adoquery.Fields("cp01").Value & "'"
                           'MODIFY BY SONIA 2016/1/5
                           'Select Case Mid(strAccNo, 1, 4)
                           '   Case "4101", "4151"
                           '      strDept = "T"
                           '   Case "4111"
                           '      strDept = "P"
                           '   Case "4121"
                           '      strDept = "CFT"
                           '   Case "4172"
                           '      If .adoaccsum.Fields("cpm11").Value = "417202" Then
                           '         strDept = "T"
                           '      Else
                           '         strDept = "FCT"
                           '      End If
                           '   Case "4131"
                           '      strDept = "CFP"
                           '   Case "4171"
                           '      strDept = "FCP"
                           '   Case "4141"
                           '      strDept = "L"
                           '   Case "4181"
                           '      strDept = "L"
                           '   Case "4161"
                           '      strDept = "FCL"
                           '   'add by sonia 2015/12/31
                           '   Case "4172"
                           '      strDept = "FCT"
                           '   'end 2015/12/31
                           '   Case Else
                           '      strDept = "TOT"
                           'End Select
                           If Left(strAccNo, 1) = "4" Then
                              strDept = PUB_GETAccNODept(strAccNo, strDept)
                           Else
                              strDept = "TOT"
                           End If
                           'END 2016/1/5
                           '93.11.25 END
                        End If
                        If douService <> 0 Then
                           '93.11.25 MODIFY BY SONIA
                           'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('1', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & strAccNo & "', " & strDept & ", " & douService & ", 0, null, null, null, null, null, '" & strRemark & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & strSalesNo & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                           'ADD BY SONIA 2016/1/5 105年起法務收入改其他部門收入(傳CP09以判斷案件性質及收文人員)
                           If Val(FCDate(.MaskEdBox1.Text)) >= 1050101 And (Left(strAccNo, 4) = "4141" Or Left(strAccNo, 4) = "4161" Or Left(strAccNo, 4) = "4181") Then
                              InsertLawACC1P0 strCompNo, "Z", strSerialNo, .Text2, strAccNo, strDept, Val(douService), 0, "", "", "", "", "", ChgSQL(strRemark), IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value), strSalesNo, .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value, Val(FCDate(.MaskEdBox1.Text)), "", "", 0, Replace(strDocuNo, "'", ""), strNo, "", 0, "", Replace(strDYes, "'", ""), "", "", .adoquery.Fields("a0j01")
                           Else
                           'END 2016/1/5
                              adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & strAccNo & "', '" & strDept & "', " & douService & ", 0, null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & strSalesNo & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                           End If  'add by sonia 2016/1/5
                           '93.11.25 END
                        End If
                     End If
                     douAService = douAService + douService
                     .adoaccsum.Close
                  'End If
                  
                  '退規費
                  'If douATax <> Val(.Text16) Then
                     'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
                     strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                     .adoaccsum.CursorLocation = adUseClient
                     .adoaccsum.Open "select cpm12 from casepropertymap where cpm01 = '" & .adoquery.Fields("cp01").Value & "' and cpm02 = '" & .adoquery.Fields("cp10").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                     If .adoaccsum.RecordCount <> 0 Then
                        'Modify by Morgan 2011/10/17 考慮拆收據情形
                        'If (douATax + IIf(IsNull(.adoquery.Fields("cp74").Value), 0, .adoquery.Fields("cp74").Value)) <= Val(.Text16) Then
                        '   douTax = IIf(IsNull(.adoquery.Fields("cp74").Value), 0, .adoquery.Fields("cp74").Value)
                        If (douATax + Val("" & .adoquery.Fields("a1u05").Value)) <= Val(.Text16) Then
                           douTax = Val("" & .adoquery.Fields("a1u05").Value)
                        'end 2011/10/17
                        Else
                           douTax = Val(.Text16) - douATax
                        End If
                        
                        If IsNull(.adoaccsum.Fields("cpm12").Value) Then
                           strAccNo = "XXX"
                        Else
                           strAccNo = .adoaccsum.Fields("cpm12").Value
                        End If
                        'Modify by Morgan 2008/6/10 改判斷非台灣
                        'If .adoquery.Fields("a0j04").Value = "020" Then
                        If .adoquery.Fields("a0j04").Value <> "000" Then
                           If "" & .adoquery.Fields("cp01") = "P" Then
                              strAccNo = "220112"
                           ElseIf "" & .adoquery.Fields("cp01") = "T" Then
                              strAccNo = "220111"
                           End If
                        End If
                        If douTax <> 0 Then
                           '93.11.25 MODIFY BY SONIA
                           'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('1', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & strAccNo & "', '" & MsgText(55) & "', " & douTax & ", 0, null, null, null, null, null, '" & strRemark & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & strAccNo & "', '" & IIf(strAccNo = "610103", strDept, MsgText(55)) & "', " & douTax & ", 0, null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                           '93.11.25 END
                           
                           'Added by Morgan 2014/1/13
                           If douTax > Val(strMaxFee) Then
                              strMaxFee = douTax
                              strMaxFeeNo = strSerialNo
                           End If
                           'end 2014/1/13
                           
                        End If
                     End If
                     douATax = douATax + douTax
                     .adoaccsum.Close
                  'End If
                  'Modify by Morgan 2011/8/17 更新cp -- 移到後面做(原程式已刪除)
                  .adoquery.MoveNext
               Loop
               .adoquery.MoveLast
               
               'Added by Morgan 2014/1/21
               'J公司
               'Modified by Morgan 2015/8/12
               '已結清
               'modify by sonia 2017/6/8 J公司只要銷退都要顯示Frmacc1194,不管有沒有開發票E10613585
               'If .Text4 <> "" And Val(.Text13) = 0 Then
               If .adoacc0k0.Fields("a0k11") = "J" And .Text3 = "3" Then
                  strRemark = strMan & "/" & Left(strCust, 10)
                  '部分銷退只帶科目,金額讓User自行輸入
                  'Modified by Morgan 2025/7/8 改都要帶金額--瑞婷
                  'If Val(.Text6) = Val(.Text15) And Val(.Text7) = Val(.Text16) Then
                     douService = Val(.Text15) + Val(.Text16)
                     douTax = douService - Round(douService / 1.05)
                     '大於稅額才扣,否則讓借貸不平由User自行調整
                     If Val(strMaxFee) > douTax Then
                        strSql = "update acc1p0 set a1p07=a1p07-" & douTax & " where a1p04='" & .Text2 & "' and a1p02='Z' and a1p03='" & strMaxFeeNo & "'"
                        adoTaie.Execute strSql, intI
                     End If
                  'Else
                  '   douService = 0
                  '   douTax = 0
                  'End If
                  'end 2025/7/8
                     
                  '銷項稅額 2119
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                  'Modified by Morgan 2015/5/28 摘要加發票號
                  'Modified by Morgan 2015/8/4 +a1p17對沖代號(本所案號)
                  'modify by sonia 2017/3/16 未收款沖帳傳票已不用1135應收銷項稅額,所以此處改1133應收帳款
                  'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                     " values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & IIf(m_Transfered = "Y", "1135", "2119") & "', 'TOT', " & douTax & ", 0, null, null, null, null, null, '" & strRemark & "/" & .Text4 & "', '" & .adoacc0k0.Fields("a0k03") & "', '" & strSalesNo & "', '" & cboCaseNo & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                  adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27)" & _
                     " values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & IIf(m_Transfered = "Y", "1133", "2119") & "', 'TOT', " & douTax & ", 0, null, null, null, null, null, '" & ChgSQL(strRemark) & "/" & .Text4 & "', '" & .adoacc0k0.Fields("a0k03") & "', '" & strSalesNo & "', '" & cboCaseNo & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                  m_KeepItem = m_KeepItem & strSerialNo & ";" 'Added by Morgan 2015/8/11
               End If
               'end 2015/8/12
               'end 2014/1/21
               
               'lngTax = Val(.Text15) * 0.1
               lngTax = Val(.Text26)
                  
               '轉應付
               If .Text17 = "1" Then
                  'Add by Morgan 2004/4/7
                  '貸方摘要
                  'Modify by Morgan 2004/12/28 改摘要
                  'strRemark = strMan & "/" & Left(strCust, 10) & "/" & .Text2 & "/" & strNo
                  strRemark = strMan & "/" & Left(strCust, 10)
                  Select Case strCon1
                     Case "1"
                        strRemark = strRemark & "/開票"
                        strRemark = strRemark & "/" & strCon2
                     Case "2"
                        strRemark = strRemark & "/退原票"
                        strRemark = strRemark & "/" & strCon3
                  End Select
                  'Add by Morgan 2004/12/28 改摘要
                  Select Case .m_stDeliver
                     Case "0"
                        strRemark = strRemark & "/寄出"
                     Case "1"
                        strRemark = strRemark & "/寄分所"
                     Case "2"
                        strRemark = strRemark & "/交智權人員"
                     'Add by Morgan 2006/10/31
                     Case "3"
                        strRemark = strRemark & "/寄出特別"
                  End Select
                  strRemark = strRemark & "/" & .Text2
                  '2004/12/28 end
                  
                  'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                  If Val(.Text26) = 0 Then
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2112', '" & MsgText(55) & "', 0, " & Val(.Text15) + Val(.Text16) & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                  Else
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2112', '" & MsgText(55) & "', 0, " & Val(.Text15) + Val(.Text16) - lngTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                     If Val(.Text26) > 0 Then
                        'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
                        strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                        'Modified by Morgan 2025/1/13 有扣單才改為2631科目
                        If Not bolTaxed Or Val(IIf(IsNull(.adoquery.Fields("a0k16").Value), 0, .adoquery.Fields("a0k16").Value)) = Val(Mid(CFDate(strSrvDate(2)), 1, 3)) Then
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '1203', '" & MsgText(55) & "', 0, " & lngTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                        Else
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2631', '" & MsgText(55) & "', 0, " & lngTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                        End If
                     End If
                  End If
                  
               '轉暫收款
               ElseIf .Text17 = "2" Then
                  'Add by Morgan 2004/4/7
                  '貸方摘要
                  strRemark = strMan & "/" & Left(strCust, 10) & "/" & .Text2 & "/" & strNo
                  'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                  'lngTax = Val(.Text15) * 0.1
                  lngTax = Val(.Text26)
                  'Modified by Morgan 2015/8/21 2401的其他對沖要放暫收款單號
                  If Val(.Text26) = 0 Then
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2401', '" & MsgText(55) & "', 0, " & Val(.Text15) + Val(.Text16) & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "/" & strNo & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ",'" & strNo & "')"
                  Else
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2401', '" & MsgText(55) & "', 0, " & Val(.Text15) + Val(.Text16) - lngTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "/" & strNo & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ",'" & strNo & "')"
                     If .adoquery.Fields("a0k16").Value > 0 Then
                        strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                        'Modified by Morgan 2025/1/13 有扣單才改為2631科目
                        If Not bolTaxed Or Val(IIf(IsNull(.adoquery.Fields("a0k16").Value), 0, .adoquery.Fields("a0k16").Value)) = Val(Mid(CFDate(ACDate(ServerDate)), 1, 3)) Then
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '1203', '" & MsgText(55) & "', 0, " & lngTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                        Else
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2631' , '" & MsgText(55) & "', 0, " & lngTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                        End If
                     End If
                  End If
               'Added by Morgan 2023/11/8
               ElseIf .Text17 = "3" Then
                  '貸方摘要
                  strRemark = strMan & "/" & Left(strCust, 10) & "/" & .Text2 & "/匯款"
                  If strCompNo = "J" Then
                     strExc(1) = "110303" '智權公司
                  ElseIf strCompNo = "J" Then
                     strExc(1) = "110502" '法律所
                  Else
                     strExc(1) = "110602" '智慧所
                  End If
                  strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                  If Val(.Text26) = 0 Then
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & strExc(1) & "', '" & MsgText(55) & "', 0, " & Val(.Text15) + Val(.Text16) & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                  Else
                     adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '" & strExc(1) & "', '" & MsgText(55) & "', 0, " & Val(.Text15) + Val(.Text16) - lngTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                     If Val(.Text26) > 0 Then
                        strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", 3)
                        'Modified by Morgan 2025/1/13 有扣單才改為2631科目
                        If Not bolTaxed Or Val(IIf(IsNull(.adoquery.Fields("a0k16").Value), 0, .adoquery.Fields("a0k16").Value)) = Val(Mid(CFDate(strSrvDate(2)), 1, 3)) Then
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '1203', '" & MsgText(55) & "', 0, " & lngTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                        Else
                           adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '" & strSerialNo & "', '" & .Text2 & "', '2631', '" & MsgText(55) & "', 0, " & lngTax & ", null, null, null, null, null, '" & ChgSQL(strRemark) & "', '" & IIf(IsNull(.adoquery.Fields("a0k03").Value), "", .adoquery.Fields("a0k03").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0k20").Value), "", .adoquery.Fields("a0k20").Value) & "', '" & .adoquery.Fields("cp01").Value & .adoquery.Fields("cp02").Value & .adoquery.Fields("cp03").Value & .adoquery.Fields("cp04").Value & "',  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strNo & "', null, 0, null, " & strDYes & ")"
                        End If
                     End If
                  End If
               'end 2023/11/8
               End If
            End If
            .adoquery.Close
            .Text22 = strNo
         End If
         
         If .Text22 <> MsgText(601) Then
            .adoacc0s0.Fields("a0s10").Value = .Text22
         Else
            .adoacc0s0.Fields("a0s10").Value = Null
         End If
         .adoacc0s0.UpdateBatch
         adoTaie.Execute "update acc0k0 set a0k10 = '" & .Text2 & "' where a0k01 = '" & .Text1 & "'"
         Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0s0.Bookmark & MsgText(35) & .adoacc0s0.RecordCount
         
         'Add by Morgan 2011/10/18 若為修改存檔多對多收據可能案號不同故資料要先還原
         If strSaveConfirm = MsgText(4) Then .RestoreData
                  
         ' 退費或銷帳足額時
         Select Case .Text3
            '銷帳
            Case "1"
               
               '全額銷帳
               If (Val(.Text6) + Val(.Text7)) = Val(.Text14) Then
                  
                  'Add by Morgan 2011/10/18 改用批次新增,另因未收款不會有 a1v0 所以不必再考慮
                  strSql = "insert into acc1u0(a1u01,a1u02,a1u03,a1u04,a1u05,a1u06,a1u07,a1u08,a1u09,a1u10)" & _
                     " select '" & .Text2 & "', '" & .Text1 & "', a0j01, 0, 0, 0,a0j09, 0,a0j10, 0" & _
                     " from acc0j0 where a0j13='" & .Text1 & "' and a0j02='" & .cboCaseNo & "'"
                  adoTaie.Execute strSql, intI
                  
                  strSql = "update caseprogress set (cp77,cp78)=(select nvl(sum(a1u07),0)+nvl(sum(a1u09),0)" & _
                     ",nvl(sum(a1u08),0)+nvl(sum(a1u10),0) from acc1u0 where a1u03=cp09)" & _
                     " where cp09 in (select a1u03 from acc1u0 where a1u01='" & .Text2 & "')"
                  adoTaie.Execute strSql, intI
                  
                  strSql = "update caseprogress set cp79 = nvl(cp16, 0) - nvl(cp75, 0) - nvl(cp77, 0) + nvl(cp78, 0)" & _
                     " where cp09 in (select a1u03 from acc1u0 where a1u01='" & .Text2 & "')"
                  adoTaie.Execute strSql, intI
                  'end 2011/10/18
                  
'Remove by Morgan 2011/10/18 改由上面程式以批次方式執行
'--舊程式已刪除--
                  
                  'Add By Sindy 2010/6/18
                  Dim dblsumA0K06 As Double, dblsumA0K07 As Double
                  Dim dblsumA1U07 As Double, dblsumA1U09 As Double
                  
                  dblsumA0K06 = 0: dblsumA0K07 = 0
                  dblsumA1U07 = 0: dblsumA1U09 = 0
                  .adoaccsum.CursorLocation = adUseClient
                  .adoaccsum.Open "select nvl(a0k06,0),nvl(a0k07,0) from acc0k0 where a0k01 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoaccsum.RecordCount <> 0 Then
                     If IsNull(.adoaccsum.Fields(0)) = False Then
                        dblsumA0K06 = .adoaccsum.Fields(0)
                     End If
                     If IsNull(.adoaccsum.Fields(1)) = False Then
                        dblsumA0K07 = .adoaccsum.Fields(1)
                     End If
                  End If
                  .adoaccsum.Close
                  .adoaccsum.CursorLocation = adUseClient
                  .adoaccsum.Open "select sum(nvl(a1u07,0)),sum(nvl(a1u09,0)) from acc1u0 where a1u02 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoaccsum.RecordCount <> 0 Then
                     If IsNull(.adoaccsum.Fields(0)) = False Then
                        dblsumA1U07 = .adoaccsum.Fields(0)
                     End If
                     If IsNull(.adoaccsum.Fields(1)) = False Then
                        dblsumA1U09 = .adoaccsum.Fields(1)
                     End If
                  End If
                  .adoaccsum.Close
                  If (dblsumA0K06 = dblsumA1U07) And (dblsumA0K07 = dblsumA1U09) Then
                     'Modified by Lydia 2023/12/12 排除Z.確定不印
                     'adoTaie.Execute "update acc0k0 set a0k32=null where a0k01 = '" & .Text1 & "'"
                     adoTaie.Execute "update acc0k0 set a0k32=null where a0k01 = '" & .Text1 & "' and nvl(a0k32,'Y') <>'Z' "
                  End If
                  '2010/6/18 End
                  
                  PUB_UpdateReceiptStatus .Text1 'Added by Morgan 2015/5/15
                  
                  adoTaie.CommitTrans
                  'Add by Morgan 2006/8/30
                  If strSaveConfirm = "A" Then
                     .Mail2Sales
                     cmdCaseNo.Visible = False 'Added by Morgan 2015/12/9
                  End If
                  
                  'Added by Lydia 2021/08/24 法務案有分配工作點數，檢查剩餘點數與工作點數不符時進入工作點數分配畫面
                  If m_KeyNo <> "" And m_BillNo <> "" And m_CaseNo <> "" Then
                      If ChkA1n05PBalence(m_CaseNo, m_BillNo) = False Then
                          Call ProcShowPoint(m_CaseNo, m_KeyNo, m_BillNo)
                      End If
                  End If
                  'end 2021/08/24
                  
                  Exit Sub
               
               'Modified by Morgan 2015/4/13 從 If 判斷式外面移來
               Else
               
                  '更新部分收款註記
                  If .adoquery.State = adStateOpen Then .adoquery.Close
                  .adoquery.CursorLocation = adUseClient
                  'Modify by Morgan 2011/10/12 考慮拆收據情形
                  '.adoquery.Open "select sum(nvl(cp16, 0)-nvl(cp75, 0)) as Namount from caseprogress where cp60 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
                  'Modified by Morgan 2013/12/31 +Ramount 已收金額(收款-退費)
                  .adoquery.Open "select nvl(a0k06,0)+nvl(a0k07,0)-nvl(a1u04,0)-nvl(a1u05,0)-nvl(a1u07,0)+nvl(a1u08,0)-nvl(a1u09,0)+nvl(a1u10,0) as Namount,nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0) Ramount from acc0k0,(select a1u02,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07,sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10 from acc1u0 where a1u02='" & .Text1 & "' and a1u01<>'" & .Text2 & "' group by a1u02) x where a0k01='" & .Text1 & "' and a1u02(+)=a0k01", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoquery.RecordCount <> 0 Then
                     If IsNull(.adoquery.Fields("Namount").Value) = False Then
                        If Val(.adoquery.Fields("Namount").Value) = Val(.Text14) Then
                           
                           'Added by Morgan 2015/4/13
                           '未收全銷不必分配
                           strSql = "insert into acc1u0(a1u01,a1u02,a1u03,a1u04,a1u05,a1u06,a1u07,a1u08,a1u09,a1u10)" & _
                              " select '" & .Text2 & "', '" & .Text1 & "', a0j01, 0, 0, 0,nvl(a0j09,0)-nvl(a1u04,0)-nvl(a1u07,0)+nvl(a1u08,0) a1u07, 0, nvl(a0j10,0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) a1u09, 0" & _
                              " from acc0j0,(select a1u02,a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07,sum(a1u08) a1u08,sum(a1u09) a1u09,sum(a1u10) a1u10 from acc1u0 where a1u02='" & .Text1 & "' and a1u01<>'" & .Text2 & "' group by a1u02,a1u03 ) x" & _
                              " where a0j13='" & .Text1 & "' and a0j02='" & .cboCaseNo & "' and a1u02(+)=a0j13 and a1u03(+)=a0j01" & _
                              " and nvl(a0j09,0)+nvl(a0j10,0)-nvl(a1u04,0)-nvl(a1u05,0)-nvl(a1u07,0)+nvl(a1u08,0)-nvl(a1u09,0)+nvl(a1u10,0)>0"
                           adoTaie.Execute strSql, intI
                           
                           strSql = "update caseprogress set (cp77,cp78)=(select nvl(sum(a1u07),0)+nvl(sum(a1u09),0)" & _
                              ",nvl(sum(a1u08),0)+nvl(sum(a1u10),0) from acc1u0 where a1u03=cp09)" & _
                              " where cp09 in (select a1u03 from acc1u0 where a1u01='" & .Text2 & "')"
                           adoTaie.Execute strSql, intI
                           
                           strSql = "update caseprogress set cp79 = nvl(cp16, 0) - nvl(cp75, 0) - nvl(cp77, 0) + nvl(cp78, 0)" & _
                              " where cp09 in (select a1u03 from acc1u0 where a1u01='" & .Text2 & "')"
                           adoTaie.Execute strSql, intI
                           'end 2015/4/13
                           
                           adoTaie.Execute "update acc0k0 SET a0k13 = 'N' where a0k01 = '" & .Text1 & "'"
                           
                           'Added by Morgan 2013/12/31
                           '有收款
                           If .adoquery.Fields("Ramount").Value > 0 Then
                           
                              'Added by Morgan 2015/4/13
                              strSql = "update acc1v0 set a1v05='N' where (a1v01,a1v02) in (select a0j01,a0j13 from acc0j0 where a0j13='" & .Text1 & "' and a0j02='" & .cboCaseNo & "')"
                              adoTaie.Execute strSql, intI
                              'end 2015/4/13
                              
                              'Added by Morgan 2015/8/10 扣繳金額也要重算
                              strSql = "update acc1v0 set (a1v04,a1v06,a1v07)=(" & _
                                 " select 0.1*(nvl(max(a0j09),0)-nvl(sum(a1u07),0)+decode(max(a0j07),'Y',nvl(max(a0j10),0)-nvl(sum(a1u09),0),0)) a1v04" & _
                                 ",nvl(sum(a1u06),0) a1v06" & _
                                 ",0.1*(nvl(max(a0j09),0)-nvl(sum(a1u07),0)+decode(max(a0j07),'Y',nvl(max(a0j10),0)-nvl(sum(a1u09),0),0))-nvl(sum(a1u06),0) a1v07" & _
                                 " from acc0j0,acc1u0" & _
                                 " where a0j01=a1v01 and a0j13=a1v02 and a1u02(+)=a1v02 and a1u03(+)=a1v01)" & _
                                 " where (a1v01,a1v02) in (select a0j01,a0j13 from acc0j0 where a0j13='" & .Text1 & "' and a0j02='" & .cboCaseNo & "')"
                              adoTaie.Execute strSql, intI
                              'end 2015/8/10
                              
                              PUB_InvProc .Text1
                              
                           'Removed by Morgan 2015/5/15 改在下面呼叫共用函數
                           '   '更新是否結清(a0k37),介紹獎金發放日期(a0k36)
                           '   adoTaie.Execute "update acc0k0 SET a0k36=decode(a0k34,null,null," & strSrvDate(2) & "),a0k37 = 'Y' where a0k01 = '" & .Text1 & "'", intI
                           ''沒收款
                           'Else
                           '   '更新是否結清(a0k37)
                           '   adoTaie.Execute "update acc0k0 SET a0k37 = 'N' where a0k01 = '" & .Text1 & "'", intI
                           End If
                           'end 2013/12/31
                           
                           PUB_UpdateReceiptStatus .Text1 'Added by Morgan 2015/5/15
                           
                           'Added by Morgan 2015/4/13
                           adoTaie.CommitTrans
                           
                           'Add by Morgan 2006/8/30
                           If strSaveConfirm = "A" Then
                              .Mail2Sales
                              cmdCaseNo.Visible = False 'Added by Morgan 2015/12/9
                           End If
                           
                           .adoquery.Close
                           
                           'Added by Lydia 2021/08/24 法務案有分配工作點數，檢查剩餘點數與工作點數不符時進入工作點數分配畫面
                           If m_KeyNo <> "" And m_BillNo <> "" And m_CaseNo <> "" Then
                               If ChkA1n05PBalence(m_CaseNo, m_BillNo) = False Then
                                   Call ProcShowPoint(m_CaseNo, m_KeyNo, m_BillNo)
                               End If
                           End If
                           'end 2021/08/24
                           
                           Exit Sub
                           'end 2015/4/13
                           
                        End If
                     End If
                  End If
                  .adoquery.Close
               
               'end 2015/4/13
               
               End If
            '退費
            Case "2"
              'Modify by Morgan 2011/10/17 2005年已加控制收據不可只退費,故刪除舊程式碼以免修改程式時浪費時間多考慮
              GoTo Checking
              
            '銷帳+退費
            Case "3"
               
               '全額銷退
               If Val(.Text6) = Val(.Text15) And Val(.Text7) = Val(.Text16) Then
                                
                  'Add by Morgan 2011/10/18 改用批次新增,另因為全額退所以 a1v0 一律刪除
                  'Modified by Morgan 2015/10/28 有扣單的扣繳為0
                  'modify by sonia 2018/12/25 -1*a1v06改為-1*nvl(a1v06,0)否則會變NULL(E10222483)
                  strSql = "insert into acc1u0(a1u01,a1u02,a1u03,a1u04,a1u05,a1u06,a1u07,a1u08,a1u09,a1u10)" & _
                     " select '" & .Text2 & "', '" & .Text1 & "', a0j01, 0, 0, " & IIf(bolTaxed, "0", "-1*nvl(a1v06,0)") & ",a0j09, a0j09,a0j10, a0j10" & _
                     " from acc0j0,acc1v0 where a0j13='" & .Text1 & "' and a0j02='" & .cboCaseNo & "' and a1v01(+)=a0j01 and a1v02(+)=a0j13"
                  adoTaie.Execute strSql, intI
                  
                  'Modified by Morgan 2015/10/28 有扣單的不刪
                  If bolTaxed Then
                     strSql = "update acc1v0 set a1v04=0,a1v07=-1*a1v04 where (a1v01,a1v02) in " & _
                        " (select a1u03,a1u02 from acc1u0 where a1u01='" & .Text2 & "')"
                     adoTaie.Execute strSql, intI
                  Else
                     strSql = "delete from acc1v0 where (a1v01,a1v02) in " & _
                        " (select a1u03,a1u02 from acc1u0 where a1u01='" & .Text2 & "') and a1v15 is null"
                     adoTaie.Execute strSql, intI
                  End If
                  'end 2015/10/28
                  
                  strSql = "update caseprogress set (cp77,cp78)=(select nvl(sum(a1u07),0)+nvl(sum(a1u09),0)" & _
                     ",nvl(sum(a1u08),0)+nvl(sum(a1u10),0) from acc1u0 where a1u03=cp09)" & _
                     " where cp09 in (select a1u03 from acc1u0 where a1u01='" & .Text2 & "')"
                  adoTaie.Execute strSql, intI
                  
                  strSql = "update caseprogress set cp79 = nvl(cp16, 0) - nvl(cp75, 0) - nvl(cp77, 0) + nvl(cp78, 0)" & _
                     " where cp09 in (select a1u03 from acc1u0 where a1u01='" & .Text2 & "')"
                  adoTaie.Execute strSql, intI
                  'end 2011/10/18
                  
                  'Added by Morgan 2012/4/19 更新進度檔已扣繳金額
                  strSql = "update caseprogress set cp76=(select nvl(sum(a1u06),0) from acc1u0 where a1u03=cp09) " & _
                     " where cp09 in (select a1u03 from acc1u0 where a1u01='" & .Text2 & "')"
                  adoTaie.Execute strSql, intI
                  'end 2012/4/19

                  PUB_UpdateReceiptStatus .Text1
'Remove by Morgan 2011/10/18 改由上面程式以批次方式執行
'--舊程式已刪除--

                  'Added by Morgan 2016/2/9
                  '發票退費彈出分錄畫面讓User可修改
                  'Modified by Morgan 2022/6/21 +案源也要
                  If bolShowForm = False And (.Text4 <> "" Or m_LOS02 <> "") Then
                     Frmacc1190.Enabled = False
                     Frmacc1194.Show
                  'add by sonia 2021/5/27 J公司只要銷退都要顯示Frmacc1194,不管有沒有開發票E11008325
                  ElseIf .adoacc0k0.Fields("a0k11") = "J" And .Text3 = "3" Then
                     Frmacc1190.Enabled = False
                     Frmacc1194.Show
                  'end 2021/5/27
                  End If
                  'end 2016/2/9
                  
                  adoTaie.CommitTrans
                  
                  'Add by Morgan 2006/8/30
                  If strSaveConfirm = "A" Then
                     .Mail2Sales
                     cmdCaseNo.Visible = False 'Added by Morgan 2015/12/9
                  End If
                  
                  'Added by Lydia 2021/08/24 法務案有分配工作點數，檢查剩餘點數與工作點數不符時進入工作點數分配畫面
                  If m_KeyNo <> "" And m_BillNo <> "" And m_CaseNo <> "" Then
                      If ChkA1n05PBalence(m_CaseNo, m_BillNo) = False Then
                          Call ProcShowPoint(m_CaseNo, m_KeyNo, m_BillNo)
                      End If
                  End If
                  'end 2021/08/24
                  
                  Exit Sub
               End If
         End Select
         
         ' 退費或銷帳不足額時
         If Val(.Text14) <> 0 Or Val(.Text15) <> 0 Or Val(.Text16) <> 0 Then
            If Val(.Text14) < (Val(.Text6) + Val(.Text7)) Or Val(.Text15) < Val(.Text8) Or Val(.Text16) < Val(.Text9) Then
              'Removed by Morgan 2024/11/12 此設定沒用且與上面的收票銀行重疊,若取消重存會錯
              'strCon4 = .cboCaseNo 'Add by Morgan 2011/10/18
              'end 2024/11/12
              strCon5 = .Text26
              strCon6 = .Text2
              strCon7 = .Text1
              strCon8 = .Text15
              strCon9 = .Text16
              strTitle = .MaskEdBox1.Text
              strItemNo = strDocuNo
              Select Case .Text3
                 Case "1"
                    strCon10 = .Text14
                 Case "3"
                    strCon10 = Val(.Text15) + Val(.Text16)
                 Case Else
                    strCon10 = ""
              End Select
              'tool3_enabled 'Removed by Morgan 2024/11/12
              'Frmacc1190.Enabled = False 'Removed by Morgan 2024/11/12
              '已收扣單時鎖住稅款退費金額
              Frmacc1192.Show
              Frmacc1192.DataGrid1.Columns(10).Locked = .bolTaxed
              
              'Added by Morgan 2024/11/12 改強制回應視窗並包在Transaction內
              Me.Tag = ""
              Set Frmacc1192.frmCall = Me
              Frmacc1192.Visible = False
              Frmacc1192.Show vbModal
              strFormName = Name
              If Me.Tag = "N" Then
                  adoTaie.RollbackTrans
                  If strSaveConfirm = MsgText(4) Then
                     adoacc0s0.Resync
                  End If
                  strControlButton = MsgText(602)
                  Exit Sub
               ElseIf Me.Tag = "F" Then
                  Frmacc1194.Show
               End If
              'end 2024/11/12
              
              bolShowForm = True  'Added by Morgan 2014/8/7
            End If
         End If
      
      'Add by Morgan 2005/1/18 選暫收款單號
      Else
         .m_stDeliver = ""
         Do While .m_stDeliver = ""
            'Modify by Morgan 2006/10/31 加 4:寄出特別
            .m_stDeliver = InputBox("請輸入銷退方式 1:寄出 2:寄分所 3:交智權人員 4:寄出特別")
            If .m_stDeliver <> "1" And .m_stDeliver <> "2" And .m_stDeliver <> "3" And .m_stDeliver <> "4" Then
               MsgBox "只可輸入 1,2,3,4", vbCritical
               .m_stDeliver = ""
            Else
               .m_stDeliver = .m_stDeliver - 1
               'Add by Morgan 2006/10/31
               If .CheckSendOpt(.m_stDeliver) = False Then
                  .m_stDeliver = ""
               End If
               'end 2006/10/31
            End If
         Loop
         '2004/12/27 end
      
         '暫收款退費
         If .adoquery.State = adStateOpen Then
            .adoquery.Close
         End If
         .adoquery.CursorLocation = adUseClient
         .adoquery.Open "select ax210 from acc1p0, acc021 where a1p01 = ax201 and a1p22 = ax202 and a1p03 = ax203 and ax210 is not null and a1p04 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount <> 0 Then
            MsgBox MsgText(14), , MsgText(5)
            strControlButton = MsgText(602)
            .adoquery.Close
            Exit Sub
         End If
         .adoquery.Close
         '2005/12/7 ADD BY SONIA
         .adoquery.CursorLocation = adUseClient
         .adoquery.Open "select a0o11 from acc0o0 where a0o01 = '" & .Text22 & "' and a0o11 is not null", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount <> 0 Then
            MsgBox "此應付款資料已付款,不可修改 !", , MsgText(5)
            strControlButton = MsgText(602)
            .adoquery.Close
            Exit Sub
         End If
         .adoquery.Close
        '2005/12/7 END
         'adoTaie.BeginTrans 'Remove by Morgan 2007/6/6 移到下面
         If strSaveConfirm = MsgText(3) Then
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            .adoquery.Open "select * from acc0s0 where a0s02 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
            If .adoquery.RecordCount <> 0 Then
               MsgBox MsgText(9), , MsgText(5)
               strControlButton = MsgText(602)
               .Text1.SetFocus
               Exit Sub
            End If
            .adoquery.Close
         End If
         If .Text1 = MsgText(601) Then
            MsgBox MsgText(10), , MsgText(5)
            'Modify by Morgan 2007/6/6
            'adoTaie.RollbackTrans
            strControlButton = MsgText(602)
            'end 2007/6/6
            Exit Sub
         End If

         adoTaie.BeginTrans 'Add by Morgan 2007/6/6
         .adoacc0t0o.CursorLocation = adUseClient
         .adoacc0t0o.Open "select * from acc0t0 where a0t01 = '" & .Text1 & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
         If .adoacc0t0o.RecordCount <> 0 Then
            .adoacc0t0o.Fields("a0t09").Value = Val(FCDate(.MaskEdBox1.Text))
            .adoacc0t0o.Fields("a0t14").Value = Val(strSrvDate(2))
            .adoacc0t0o.Fields("a0t15").Value = ServerTime
            .adoacc0t0o.Fields("a0t16").Value = strUserNum
            .adoacc0t0o.UpdateBatch
         End If
         .adoacc0t0o.Close
         
         'Modified by Morgan 2024/11/29
         '.adoacc0t0.Close
         If .adoacc0t0.State = adStateOpen Then .adoacc0t0.Close
         If .adoquery.State = adStateOpen Then .adoquery.Close
         'end 2024/11/29
         
         '2005/12/7 ADD BY SONIA
         .adoquery.CursorLocation = adUseClient
         'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
         .adoquery.Open "select a1p22 from acc1p0 where a1p02 = 'Z' and a1p04 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount <> 0 Then
            If IsNull(.adoquery.Fields("a1p22").Value) Then
               strDocuNo = "null"
               strDYes = "null"
            Else
               strDocuNo = "'" & .adoquery.Fields("a1p22").Value & "'"
               strDYes = "'Y'"
            End If
         Else
            strDocuNo = "null"
            strDYes = "null"
         End If
         .adoquery.Close
         '2005/12/7 END
         adoTaie.Execute "delete from acc0s0 where a0s01 = '" & .Text2 & "'"
         'Modify by Morgan 2007/1/5 分錄不用刪改可輸入
         'adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'Z' and a1p04 = '" & .Text2 & "'"
         adoTaie.Execute "delete from acc0o0 where a0o01 = '" & .Text22 & "'"
         adoTaie.Execute "insert into acc0s0(a0s01,a0s02,a0s03,a0s04,a0s05,a0s06,a0s07,a0s08,a0s09,a0s10,a0s17,a0s18,a0s19,a0s20,a0s21,a0s22,a0s13,a0s11,a0s12,a0s16,a0s14,a0s15,a0s23) values ('" & .Text2 & "', '" & .Text1 & "', " & Val(FCDate(.MaskEdBox1.Text)) & ", '1', " & Val(.Text21) & ", 0, 0, null, " & Val(FCDate(.MaskEdBox3.Text)) & ", null, null, '" & .Text27 & "', null, null, null, null, '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", null, null, null,'" & .Text28 & "')"
         .adoquery.CursorLocation = adUseClient
         .adoquery.Open "select * from acc0t0 where a0t01 = '" & .Text1 & "'", adoTaie, adOpenStatic, adLockReadOnly
         If .adoquery.RecordCount <> 0 And Val(.Text21) <> 0 Then
            If IsNull(.adoquery.Fields("a0t05").Value) = False Then
               .adoaccsum.CursorLocation = adUseClient
               .adoaccsum.Open "select sn01 from salesno where sn02 = '" & .adoquery.Fields("a0t05").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
               If .adoaccsum.RecordCount <> 0 Then
                  If IsNull(.adoaccsum.Fields("sn01").Value) Then
                     strMan = ""
                  Else
                     strMan = .adoaccsum.Fields("sn01").Value
                  End If
               End If
               .adoaccsum.Close
            Else
               strMan = ""
            End If
            'Add by Morgan 2007/8/16 預設備註資料--瑞婷
            If .Text27 <> "" Then
               strCust = .Text27
            Else
            'end 2007/8/16
               strCust = CustomerQuery(.adoquery.Fields("a0t06").Value, 1)
               If strCust = "" Then
                  strCust = CustomerQuery(.adoquery.Fields("a0t06").Value, 2)
               End If
            End If
            
            'Modify by Morgan 2005/1/18 改摘要
            'strRemark = strMan & "/" & strCust & "/" & .Text2
            strRemark = strMan & "/" & strCust
            '2005/5/12 ADD BY SONIA
            If .Text27 <> "" Then
               strRemark2112 = strMan & "/" & Mid(.Text27, 1, 6)
               strRemark = strMan & "/" & .Text27     '2007/6/27 add by sonia D096061743
            Else
               strRemark2112 = strMan & "/" & strCust
            End If
            '2005/5/12 END
            Select Case .m_stDeliver
               Case "0"
                  strRemark = strRemark & "/寄出"
                  strRemark2112 = strRemark2112 & "/寄出"       '2005/5/12 ADD BY SONIA
               Case "1"
                  strRemark = strRemark & "/寄分所"
                  strRemark2112 = strRemark2112 & "/寄分所"     '2005/5/12 ADD BY SONIA
               Case "2"
                  strRemark = strRemark & "/交智權人員"
                  strRemark2112 = strRemark2112 & "/交智權人員"     '2005/5/12 ADD BY SONIA
               'Add by Morgan 2006/10/31
               Case "3"
                  strRemark = strRemark & "/寄出特別"
                  strRemark2112 = strRemark2112 & "/寄出特別"
            End Select
            '2014/2/25 modify by sonia 辜說改抓原收據銷退時之銷帳單號
            'strRemark = strRemark & "/" & .Text2
            If "" & .adoquery.Fields("a0t07").Value <> "" Then
               strRemark = strRemark & "/原銷帳單號" & "" & .adoquery.Fields("a0t07").Value
            Else
               strRemark = strRemark & "/" & .Text2
            End If
            '2014/2/25 end
            strRemark2112 = strRemark2112 & "/" & .Text22     '2005/5/12 ADD BY SONIA
            '2005/1/18 end
            
            'Ken 93/02/03 修改時, 原付款單號不變
            If .Text22 <> MsgText(601) Then
               strSerialNo = .Text22
            Else
               strSerialNo = AutoNo(MsgText(804), 5, 1)
            End If
            '2005/12/7 MODIFY BY SONIA
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('1', 'Z', '001', '" & .Text2 & "', '2401', 'TOT', " & Val(.Text21) & ", 0, null, null, null, null, null, '" & strRemark & "/" & .Text1 & "', '" & IIf(IsNull(.adoquery.Fields("a0t06").Value), "", .adoquery.Fields("a0t06").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0t05").Value), "", .adoquery.Fields("a0t05").Value) & "', null,  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, null, '" & strSerialNo & "', null, 0, null, null)"
            'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('1', 'Z', '002', '" & .Text2 & "', '2112', 'TOT', 0, " & Val(.Text21) & ", null, null, null, null, null, '" & strRemark2112 & "', '" & IIf(IsNull(.adoquery.Fields("a0t06").Value), "", .adoquery.Fields("a0t06").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0t05").Value), "", .adoquery.Fields("a0t05").Value) & "', null,  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, null, '" & strSerialNo & "', null, 0, null, null)"
            'Modify by Morgan 2007/1/5 改新增時才要做
            If strSaveConfirm = MsgText(3) Then
               '2012/9/25 MODIFY BY SONIA 加存暫收款單號於A1P30
               'adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('1', 'Z', '001', '" & .Text2 & "', '2401', 'TOT', " & Val(.Text21) & ", 0, null, null, null, null, null, '" & strRemark & "/" & .Text1 & "', '" & IIf(IsNull(.adoquery.Fields("a0t06").Value), "", .adoquery.Fields("a0t06").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0t05").Value), "", .adoquery.Fields("a0t05").Value) & "', null,  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strSerialNo & "', null, 0, null, " & strDYes & ")"
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27, a1p30) values ('" & strCompNo & "', 'Z', '001', '" & .Text2 & "', '2401', 'TOT', " & Val(.Text21) & ", 0, null, null, null, null, null, '" & ChgSQL(strRemark) & "/" & .Text1 & "', '" & IIf(IsNull(.adoquery.Fields("a0t06").Value), "", .adoquery.Fields("a0t06").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0t05").Value), "", .adoquery.Fields("a0t05").Value) & "', null,  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strSerialNo & "', null, 0, null, " & strDYes & ",'" & .Text1 & "')"
               adoTaie.Execute "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27) values ('" & strCompNo & "', 'Z', '002', '" & .Text2 & "', '2112', 'TOT', 0, " & Val(.Text21) & ", null, null, null, null, null, '" & ChgSQL(strRemark2112) & "', '" & IIf(IsNull(.adoquery.Fields("a0t06").Value), "", .adoquery.Fields("a0t06").Value) & "', '" & IIf(IsNull(.adoquery.Fields("a0t05").Value), "", .adoquery.Fields("a0t05").Value) & "', null,  " & Val(FCDate(.MaskEdBox1.Text)) & ", null, null, 0, " & strDocuNo & ", '" & strSerialNo & "', null, 0, null, " & strDYes & ")"
            Else
               'Modified by Morgan 2014/1/20 會有J公司,取消 a1p01='1' 條件
               adoTaie.Execute "update acc1p0 set a1p14='" & ChgSQL(strRemark) & "/" & .Text1 & "' where a1p02='Z' and a1p04='" & .Text2 & "' and a1p05='2401'"
               adoTaie.Execute "update acc1p0 set a1p14='" & ChgSQL(strRemark2112) & "' where a1p02='Z' and a1p04='" & .Text2 & "' and a1p05='2112'"
            End If
            'end 2007/1/5
            '2005/12/7 END
            '2005/5/3 MODIFY BY SONIA
            'adoTaie.Execute "insert into acc0o0 (a0o01, a0o02, a0o03, a0o04, a0o05, a0o06, a0o09, a0o10, a0o11, a0o19, a0o15, a0o13, a0o14, a0o18, a0o16, a0o17) values ('" & strSerialNo & "', '2', '" & .Text20 & "', null, " & Val(FCDate(.MaskEdBox1.Text)) & ", " & Val(FCDate(.MaskEdBox1.Text)) & ", '" & .Text2 & "', null, null, '2', '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ", '" & strUserNum & "', " & Val(ACDate(ServerDate)) & ", " & ServerTime & ")"
            'Modified by Morgan 2014/1/17 +a0o07
            adoTaie.Execute "insert into acc0o0 (a0o01, a0o02, a0o03, a0o04, a0o05, a0o06, a0o09, a0o10, a0o11, a0o19, a0o15, a0o13, a0o14, a0o18, a0o16, a0o17,a0o07) values ('" & strSerialNo & "', '2', '" & .Text20 & "', null, " & Val(FCDate(.MaskEdBox1.Text)) & ", " & Val(FCDate(.MaskEdBox1.Text)) & ", '" & .Text2 & "', '" & .Text27 & "', null, '2', '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ", '" & strUserNum & "', " & Val(strSrvDate(2)) & ", " & ServerTime & ",'" & strCompNo & "')"
            '2005/5/3 END
            .Text22 = strSerialNo
            adoTaie.Execute "update acc0s0 set a0s10 = '" & strSerialNo & "' where a0s01 = '" & .Text2 & "'"
         End If
         .adoquery.Close
         .adoacc0t0.CursorLocation = adUseClient
         .adoacc0t0.Open "select * from acc0t0, acc0s0 where a0t01 = a0s02 and (a0t09 is not null AND A0T09 <> 0) order by a0t01 asc", adoTaie, adOpenStatic, adLockReadOnly
         .adoacc0t0.Find "a0t01 = '" & .Text1 & "'", 0, adSearchForward, 1
         Frmacc0000.StatusBar1.Panels(2).Text = .adoacc0t0.Bookmark & MsgText(35) & .adoacc0t0.RecordCount
         
         'Add by Morgan 2007/1/5
         Frmacc1190.Enabled = False
         Frmacc1194.Show
         ''end 2007/1/5
         
         bolShowForm = True  'Added by Morgan 2014/8/7
      End If
      
      'Added by Morgan 2014/8/7
      '發票退費彈出分錄畫面讓User可修改
      'Modified by Morgan 2022/6/21 +案源也要
      If bolShowForm = False And (.Text4 <> "" Or m_LOS02 <> "") Then
         Frmacc1190.Enabled = False
         Frmacc1194.Show
         
      'add by sonia 2021/5/27 J公司只要銷退都要顯示Frmacc1194,不管有沒有開發票E11008325
      'Modified by Morgan 2024/11/12
      'ElseIf .Option1.Value Then
      ElseIf bolShowForm = False And .Option1.Value Then
      'end 2024/11/12
         If .adoacc0k0.Fields("a0k11") = "J" And .Text3 = "3" Then
            Frmacc1190.Enabled = False
            Frmacc1194.Show
         End If
      'end 2021/5/27
      End If
      'end 2014/8/7
      
      adoTaie.CommitTrans
      'Add by Morgan 2006/8/30
      If strSaveConfirm = "A" Then
         .Mail2Sales
         cmdCaseNo.Visible = False 'Added by Morgan 2015/12/9
      End If
      
      'Added by Lydia 2021/08/24 法務案有分配工作點數，檢查剩餘點數與工作點數不符時進入工作點數分配畫面
      If m_KeyNo <> "" And m_BillNo <> "" And m_CaseNo <> "" Then
          If ChkA1n05PBalence(m_CaseNo, m_BillNo) = False Then
              Call ProcShowPoint(m_CaseNo, m_KeyNo, m_BillNo)
          End If
      End If
      'end 2021/08/24
      
      Exit Sub
      
Checking:
   adoTaie.RollbackTrans
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   strControlButton = MsgText(602) 'Add by Morgan 2011/10/17
   End With
End Sub

'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Added by Lydia 2020/04/10 法務工作點數分配
Private Sub CmdDot_Click()
   '判斷非編輯狀態
   If Frmacc0000.Toolbar1.Buttons.Item(1).Enabled = True And Frmacc0000.Toolbar1.Buttons.Item(4).Enabled = True And Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True Then
      '保留: 原本使用Frmacc21h5
      'If Val(m_CaseCP18) = 0 Then
      '    MsgBox "無點數可供分配!", vbExclamation
      '    Exit Sub
      'End If
'      If Left(Text1, 1) = "E" Or Text1 = "" Then
'         Set Frmacc21h5.m_PrevForm = Me
'         Frmacc21h5.txtA0K01.Text = Text1.Text
'         Frmacc21h5.m_bolPrev = True
'         Frmacc21h5.Show
'         Me.Visible = False
'      Else
'         MsgBox "點數輸入限國內收據單號!!", vbCritical
'      End If
      'end 保留
      If Text1 = "" And Text2 = "" Then
          MsgBox "請先輸入單據號碼!!", vbExclamation
          Exit Sub
      End If
      If Option2.Value = True Then
          MsgBox "限收據號碼!!", vbExclamation
          Exit Sub
      End If
      If cboCaseNo.Text = "" Then
          MsgBox "請先選擇本所案號!!", vbExclamation
          Exit Sub
      End If
      
      'Memo by Lydia 2021/08/24 改成模組
      Call ProcShowPoint(cboCaseNo.Text, Text2.Text, Text1.Text)
   End If
End Sub

'Added by Lydia 2021/08/24
Private Function ProcShowPoint(ByVal pSysNo As String, pKeyNo As String, pBillNo As String) As Boolean
Dim rsA9 As New ADODB.Recordset
Dim mCP(1 To 4) As String
Dim intA As Integer
Dim strA As String

'判斷非編輯狀態
'   If Frmacc0000.Toolbar1.Buttons.Item(1).Enabled = True And Frmacc0000.Toolbar1.Buttons.Item(4).Enabled = True And Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True Then
      If pKeyNo = "" And pBillNo = "" Then
          MsgBox "請先輸入單據號碼!!", vbExclamation
          Exit Function
      End If
      If pSysNo = "" Then
          MsgBox "請先選擇本所案號!!", vbExclamation
          Exit Function
      End If
      
      '抓收文號的資料
      Call ChgCaseNo(pSysNo, mCP)
      strA = "select cp09,cp18 from caseprogress where cp60='" & pBillNo & "' and nvl(cp18,0)>0 " & _
                "and cp01='" & mCP(1) & "' and cp02='" & mCP(2) & "' and cp03='" & mCP(3) & "' and cp04='" & mCP(4) & "' "
      intA = 1
      Set rsA9 = ClsLawReadRstMsg(intA, strA)
      If intA = 1 Then
         If rsA9.RecordCount = 1 Then
            Set frm071021.m_PrevForm = Me
            frm071021.m_bolPrev = True
            frm071021.m_KeyList = "" & rsA9.Fields("cp09")
            Frmacc1190.Enabled = False
            frm071021.Show
         Else
            Frmacc1190.Enabled = False
            Frmacc1195.m_A0J13 = pBillNo
            Frmacc1195.Show
         End If
      End If
      Set rsA9 = Nothing
      ProcShowPoint = True
'   Else
'      MsgBox "尚在輸入資料中!!", vbCritical
'      ProcShowPoint = False
'   End If
   
End Function

'Added by Lydia 2021/08/24 法務案有分配工作點數，檢查剩餘點數與工作點數不符時進入工作點數分配畫面
Private Function ChkA1n05PBalence(ByVal pSysNo As String, ByVal pNo As String) As Boolean
'pSysNo: 本所案號
'pNo: 收據號碼/ 收文號
Dim strQ1 As String, intQ As Integer
Dim rsQuery As New ADODB.Recordset


   ChkA1n05PBalence = True
   If InStr(UCase(pSysNo), "L") = 0 Then Exit Function
   
   If Left(pNo, 1) < "E" Then
       strQ1 = " cp09='" & pNo & "' "
   Else
       strQ1 = " cp60='" & pNo & "' "
   End If
   
   strQ1 = "select cp01,cp02,cp03,cp04,cp13,cp09,cp18,sum(a1u07) a1u07,sum(a1n05) a1n05 " & _
               "from ( select cp01,cp02,cp03,cp04,cp09,cp13,cp18,sum(nvl(a1u07,0)/1000) a1u07, 0 as a1n05 " & _
               "from caseprogress, acc1u0 where " & strQ1 & " and cp09=a1u03(+) and cp60=a1u02(+) " & _
               "group by cp01,cp02,cp03,cp04,cp09,cp13,cp18 " & _
               "union select cp01,cp02,cp03,cp04,cp09,cp13,cp18,0 as a1u07,sum(nvl(a1n05,0)) a1n05 " & _
               "from caseprogress,acc1n0 where " & strQ1 & " and cp09=a1n03(+) and a1n02='3' " & _
               "group by cp01,cp02,cp03,cp04,cp09,cp13,cp18 " & _
               ") group by cp01,cp02,cp03,cp04,cp09,cp13,cp18 "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
           '法務案有分配工作點數，檢查剩餘點數與工作點數不符時進入工作點數分配畫面
           If Val("" & rsQuery.Fields("a1n05")) <> 0 Then
               If Val("" & rsQuery.Fields("cp18")) - Val("" & rsQuery.Fields("a1u07")) <> Val("" & rsQuery.Fields("a1n05")) Then
                  ChkA1n05PBalence = False
               End If
           End If
           rsQuery.MoveNext
      Loop
   End If
End Function

'Added by Morgan 2022/6/21
'以收據號讀取案源類別
Private Sub SetLOS02()
   Dim intQ As Integer, stSQL As String
   Dim rsQuery As ADODB.Recordset
   
   m_LOS02 = ""
   stSQL = "select nvl(l2.los02,l1.los02) los02,a0k11" & _
      " from acc0k0,acc0j0,caseprogress,lawofficesource l1,lawofficesource l2" & _
      " where a0k01='" & Me.Text1 & "' and a0j13(+)=a0k01 and cp09(+)=a0j01" & _
      " and l1.los01(+)=cp09 and l2.los15(+)=cp162"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      m_LOS02 = "" & rsQuery("los02")
      If m_LOS02 <> "" And strSaveConfirm = "A" Then
         MsgBox "本案為【" & m_LOS02 & "】類案源，請確認是否有【" & IIf(rsQuery("a0k11") = "L", "1", "L") & "】公司收據也要銷退！", vbInformation
      End If
   End If
End Sub
