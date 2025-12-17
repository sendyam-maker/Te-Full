VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc1124 
   Caption         =   "收據開立作業-執行中"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   5100
   Begin VB.CommandButton Command1 
      Caption         =   "返回(&R)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3165
      TabIndex        =   24
      Top             =   3930
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "略過(&K)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1005
      TabIndex        =   23
      Top             =   3930
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "繼續(&C)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2085
      TabIndex        =   22
      Top             =   3930
      Width           =   960
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   10
      Left            =   1350
      TabIndex        =   21
      Top             =   3480
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   9
      Left            =   1350
      TabIndex        =   20
      Top             =   3150
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   8
      Left            =   1350
      TabIndex        =   19
      Top             =   2805
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   7
      Left            =   1350
      TabIndex        =   18
      Top             =   2475
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   6
      Left            =   1350
      TabIndex        =   17
      Top             =   2160
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   5
      Left            =   1350
      TabIndex        =   16
      Top             =   1830
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   4
      Left            =   1350
      TabIndex        =   15
      Top             =   1500
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   3
      Left            =   1350
      TabIndex        =   14
      Top             =   1152
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   2
      Left            =   1350
      TabIndex        =   13
      Top             =   828
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   1
      Left            =   1350
      TabIndex        =   12
      Top             =   504
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   0
      Left            =   1350
      TabIndex        =   11
      Top             =   180
      Width           =   3570
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6297;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   225
      TabIndex        =   10
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   225
      TabIndex        =   9
      Top             =   3150
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   225
      TabIndex        =   8
      Top             =   2805
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   225
      TabIndex        =   7
      Top             =   2475
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年費年度："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   225
      TabIndex        =   6
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規費："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   225
      TabIndex        =   5
      Top             =   1830
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服務費："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   225
      TabIndex        =   4
      Top             =   1500
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   3
      Top             =   1152
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   225
      TabIndex        =   2
      Top             =   828
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   1
      Top             =   504
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   1050
   End
End
Attribute VB_Name = "Frmacc1124"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Create by Morgan 2010/12/7
Option Explicit


Private Sub Command1_Click(Index As Integer)
   Frmacc1123.m_Rtn = Index
   Unload Me
End Sub

Private Sub Form_Load()
   'Modify by Amy 2023/10/06 原W:5130 /H:4770
   PUB_InitForm Me, 5220, 4950
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc1124 = Nothing
End Sub
