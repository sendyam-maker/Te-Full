VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc7151 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "分所出納之智權人員繳款確認"
   ClientHeight    =   3468
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3468
   ScaleWidth      =   9336
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   8100
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1590
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   7335
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1590
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   6570
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1590
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   5805
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1590
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   1260
      MaxLength       =   13
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3020
      Width           =   960
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   990
      Width           =   1050
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   990
      Width           =   1050
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   6390
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   990
      Width           =   915
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   8010
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   990
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4860
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1590
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4095
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1590
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1590
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2385
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1590
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1590
      Width           =   870
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   270
      MaxLength       =   8
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1590
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "存檔(&S)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   6885
      TabIndex        =   3
      Top             =   240
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&R)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   7860
      TabIndex        =   4
      Top             =   240
      Width           =   1380
   End
   Begin MSForms.Label lblA4431 
      Height          =   495
      Left            =   6300
      TabIndex        =   47
      Top             =   1920
      Width           =   2760
      VariousPropertyBits=   8388627
      Size            =   "4868;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblA4412 
      Height          =   500
      Left            =   1300
      TabIndex        =   22
      Top             =   1920
      Width           =   4500
      VariousPropertyBits=   8388627
      Size            =   "7937;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA4415 
      Height          =   495
      Left            =   1260
      TabIndex        =   1
      Top             =   2460
      Width           =   7800
      VariousPropertyBits=   -1463795685
      MaxLength       =   250
      ScrollBars      =   2
      Size            =   "13758;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCUser 
      Height          =   300
      Left            =   1215
      TabIndex        =   10
      Top             =   500
      Width           =   1080
      VariousPropertyBits=   19
      Size            =   "1905;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSales 
      Height          =   300
      Left            =   1215
      TabIndex        =   9
      Top             =   250
      Width           =   1080
      VariousPropertyBits=   19
      Size            =   "1905;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "其他備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   22
      Left            =   5850
      TabIndex        =   46
      Top             =   1950
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "其　他"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   5910
      TabIndex        =   45
      Top             =   1380
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "留分所"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   270
      TabIndex        =   43
      Top             =   3065
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   2565
      Left            =   90
      Top             =   870
      Width           =   9150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收據收款金額合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   315
      TabIndex        =   42
      Top             =   1035
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "服務費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   2295
      TabIndex        =   41
      Top             =   1035
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "規費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   4185
      TabIndex        =   40
      Top             =   1035
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "扣繳"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   5895
      TabIndex        =   39
      Top             =   1035
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "總計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   7515
      TabIndex        =   38
      Top             =   1035
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權人員備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   19
      Left            =   270
      TabIndex        =   37
      Top             =   1920
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "手續費"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   7395
      TabIndex        =   36
      Top             =   1380
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "溢收款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   6630
      TabIndex        =   35
      Top             =   1380
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "抵暫收款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   4920
      TabIndex        =   34
      Top             =   1380
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "現　金"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   4155
      TabIndex        =   33
      Top             =   1380
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "分所電匯"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   3255
      TabIndex        =   32
      Top             =   1380
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "台北電匯"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   2400
      TabIndex        =   31
      Top             =   1380
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "票據金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   1545
      TabIndex        =   30
      Top             =   1380
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "票　　號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   465
      TabIndex        =   29
      Top             =   1380
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "補扣繳/外幣"
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
      Index           =   20
      Left            =   8063
      TabIndex        =   28
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "出納人員備註"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   270
      TabIndex        =   27
      Top             =   2460
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRTime 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5940
      TabIndex        =   13
      Top             =   240
      Width           =   810
   End
   Begin VB.Label lblCDate 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   12
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label lblRDate 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   240
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "繳款日期時間："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2820
      TabIndex        =   8
      Top             =   240
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "智權人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   7
      Top             =   240
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "出納確認日期時間："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "出納人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   480
      Width           =   1050
   End
End
Attribute VB_Name = "Frmacc7151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/28 Form2.0已修改(lblSales,lblCUser,Text1(8)->lblA4412,Text1(13)->lblA4431,Text1(10)->txtA4415)
'Created by Morgan 2013/12/12
Option Explicit

Public m_A4401 As String
Public m_A4402 As String
Public m_A4403 As String

Private Sub cmdOK_Click(Index As Integer)
   'add by sonia 2022/1/2 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
      Exit Sub
   End If
   'end 2022/1/2
   
   Select Case Index
   Case 0
      If FormSave = True Then
         'Removed by Morgan 2023/11/16 取消以免誤發--秀玲/財務處
         'PUB_AccDeliverInform "2", , m_A4401, m_A4402, m_A4403 'Added by Morgan 2016/8/17
         Unload Me
      End If
   Case 1
      Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   'modify by sonia 2021/12/28
   'Text1(10).SetFocus
   txtA4415.SetFocus
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc7151 = Nothing
End Sub

Private Sub QueryData()
   strExc(0) = "select * " & _
      " from (select AXD01,AXD02,AXD03,sum(axd06) S1,sum(axd07) S2,sum(axd08) S3" & _
      " From acc441 where AXD01='" & m_A4401 & "' and AXD02=" & m_A4402 & " and AXD03=" & m_A4403 & _
      " group by AXD01,AXD02,AXD03),acc440" & _
      " where a4401(+)=AXD01 and a4402(+)=AXD02 and a4403(+)=AXD03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      txtTot(2) = Format(.Fields("S1"), "#,##0")
      txtTot(3) = Format(.Fields("S2"), "#,##0")
      txtTot(4) = Format(.Fields("S3"), "#,##0")
      txtTot(5) = Format(.Fields("S1") + .Fields("S2") - .Fields("S3"), "#,##0")
   '   Text1(0) = "" & .Fields("A4404")
      'Add by Lydia 2015/01/14 +票號A4428,留分所A4429
      Text1(0) = "" & .Fields("A4428")
      Text1(11) = Format(Val("" & .Fields("A4429")), "#,##0")
      
      Text1(1) = Format(Val("" & .Fields("A4405")), "#,##0")
      Text1(2) = Format(Val("" & .Fields("A4406")), "#,##0")
      Text1(3) = Format(Val("" & .Fields("A4407")), "#,##0")
      Text1(4) = Format(Val("" & .Fields("A4408")), "#,##0")
      Text1(5) = Format(Val("" & .Fields("A4409")), "#,##0")
      Text1(6) = Format(Val("" & .Fields("A4410")), "#,##0")
      Text1(7) = Format(Val("" & .Fields("A4411")), "#,##0")
      'modify by sonia 2021/12/28
      'Text1(8) = "" & .Fields("A4412")
      lblA4412 = "" & .Fields("A4412")
      Text1(9) = Format(Val("" & .Fields("A4422")), "#,##0")
      'modify by sonia 2021/12/28
      'Text1(10) = "" & .Fields("A4415")
      txtA4415 = "" & .Fields("A4415")
      'Added by Morgan 2015/7/15
      Text1(12) = Format(Val("" & .Fields("A4430")), "#,##0")
      'modify by sonia 2021/12/28
      'Text1(13) = "" & .Fields("A4431")
      lblA4431 = "" & .Fields("A4431")
      'end 2015/7/15
      End With
   End If
End Sub

Private Function FormSave() As Boolean
   
On Error GoTo ErrHnd
   'Add by Lydia 2015/01/14 +票號A4428,留分所A4429
   'modify by sonia 2021/12/28 Text1(10)->txtA4415
   strSql = "update acc440 set A4415='" & ChgSQL(txtA4415) & "'" & _
      ",A4413=" & strSrvDate(1) & ",A4414='" & strUserNum & "',A4423=to_char(sysdate,'hh24miss')" & _
      ",A4428=" & CNULL(ChgSQL(Text1(0))) & ",A4429=" & Val(Format(Text1(11), "##0")) & _
      " where a4401='" & m_A4401 & "' and a4402=" & m_A4402 & " And A4403 = " & m_A4403
   cnnConnection.Execute strSql, intI
   
   FormSave = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub Text1_GotFocus(Index As Integer)
   If Text1(Index).Locked = False Then
      TextInverse Text1(Index)
      If Index = 10 Then
         OpenIme
      Else 'Add by Lydia 2015/01/14
         CloseIme
      End If
   End If
End Sub

'Add by Lydia 2015/01/14
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 11 Then
      If KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If Index = 11 Then
   Text1(11) = Format(Val(Text1(11)), "#,##0")
End If
End Sub

'add by sonia 2021/12/28
Private Sub txtA4415_GotFocus()
   TextInverse txtA4415
   OpenIme
End Sub
'end 2021/12/28
