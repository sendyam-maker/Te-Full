VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170015 
   BorderStyle     =   1  '單線固定
   Caption         =   "所得稅率表"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6645
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   1
      Left            =   5400
      TabIndex        =   72
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "存檔(&S)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   4440
      TabIndex        =   71
      Top             =   30
      Width           =   915
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   6
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   5
      Text            =   "99999999"
      Top             =   795
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Index           =   5
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "99999999"
      Top             =   1935
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Index           =   4
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "99999999"
      Top             =   1605
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Index           =   3
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "99999999"
      Top             =   1275
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Index           =   2
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "99999999"
      Top             =   945
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Index           =   1
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "99999999"
      Top             =   630
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   7
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "99.99"
      Top             =   780
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   8
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "99999999"
      Top             =   765
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   10
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "99.99"
      Top             =   1035
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   11
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   10
      Text            =   "99999999"
      Top             =   1020
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   12
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   11
      Text            =   "99999999"
      Top             =   1305
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   13
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   12
      Text            =   "99.99"
      Top             =   1290
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   14
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   13
      Text            =   "99999999"
      Top             =   1275
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   15
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   14
      Text            =   "99999999"
      Top             =   1560
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   16
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   15
      Text            =   "99.99"
      Top             =   1545
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   17
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   16
      Text            =   "99999999"
      Top             =   1530
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   18
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   17
      Text            =   "99999999"
      Top             =   1815
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   19
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   18
      Text            =   "99.99"
      Top             =   1800
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   20
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   19
      Text            =   "99999999"
      Top             =   1785
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   21
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   20
      Text            =   "99999999"
      Top             =   2070
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   22
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   21
      Text            =   "99.99"
      Top             =   2055
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   23
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   22
      Text            =   "99999999"
      Top             =   2040
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   24
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   23
      Text            =   "99999999"
      Top             =   2325
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   25
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   24
      Text            =   "99.99"
      Top             =   2310
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   26
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   25
      Text            =   "99999999"
      Top             =   2295
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   27
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   26
      Text            =   "99999999"
      Top             =   2580
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   28
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   27
      Text            =   "99.99"
      Top             =   2565
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   29
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   28
      Text            =   "99999999"
      Top             =   2550
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   30
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   29
      Text            =   "99999999"
      Top             =   2835
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   31
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   30
      Text            =   "99.99"
      Top             =   2820
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   32
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   31
      Text            =   "99999999"
      Top             =   2805
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   33
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   32
      Text            =   "99999999"
      Top             =   3090
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   34
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   33
      Text            =   "99.99"
      Top             =   3075
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   35
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   34
      Text            =   "99999999"
      Top             =   3060
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   36
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   35
      Text            =   "99999999"
      Top             =   3345
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   37
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   36
      Text            =   "99.99"
      Top             =   3330
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   38
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   37
      Text            =   "99999999"
      Top             =   3315
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   39
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   39
      Text            =   "99999999"
      Top             =   3600
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   40
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   41
      Text            =   "99.99"
      Top             =   3585
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   41
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   43
      Text            =   "99999999"
      Top             =   3570
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   42
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   44
      Text            =   "99999999"
      Top             =   3855
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   43
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   45
      Text            =   "99.99"
      Top             =   3840
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   44
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   46
      Text            =   "99999999"
      Top             =   3825
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   45
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   47
      Text            =   "99999999"
      Top             =   4110
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   46
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   48
      Text            =   "99.99"
      Top             =   4095
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   47
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   49
      Text            =   "99999999"
      Top             =   4080
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   48
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   50
      Text            =   "99999999"
      Top             =   4365
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   49
      Left            =   4770
      MaxLength       =   5
      TabIndex        =   51
      Text            =   "99.99"
      Top             =   4350
      Width           =   520
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   50
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   52
      Text            =   "99999999"
      Top             =   4335
      Width           =   825
   End
   Begin VB.TextBox txtIT 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   9
      Left            =   3690
      MaxLength       =   8
      TabIndex        =   8
      Text            =   "99999999"
      Top             =   1050
      Width           =   825
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   120
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   4800
      Width           =   5700
      VariousPropertyBits=   671105055
      Size            =   "10054;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "２."
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
      Left            =   3420
      TabIndex        =   70
      Top             =   1095
      Width           =   255
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "３."
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
      Left            =   3420
      TabIndex        =   69
      Top             =   1350
      Width           =   255
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "４."
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
      Left            =   3420
      TabIndex        =   68
      Top             =   1605
      Width           =   255
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "５."
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
      Left            =   3420
      TabIndex        =   67
      Top             =   1860
      Width           =   255
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "６."
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
      Left            =   3420
      TabIndex        =   66
      Top             =   2115
      Width           =   255
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "７."
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
      Left            =   3420
      TabIndex        =   65
      Top             =   2370
      Width           =   255
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "８."
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
      Left            =   3420
      TabIndex        =   64
      Top             =   2625
      Width           =   255
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "９."
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
      Left            =   3420
      TabIndex        =   63
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "１０."
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
      Left            =   3240
      TabIndex        =   62
      Top             =   3135
      Width           =   450
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "１１."
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
      Left            =   3240
      TabIndex        =   61
      Top             =   3390
      Width           =   450
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "１２."
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
      Left            =   3240
      TabIndex        =   60
      Top             =   3645
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "１３."
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
      Left            =   3240
      TabIndex        =   59
      Top             =   3900
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "１４."
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
      Left            =   3240
      TabIndex        =   58
      Top             =   4155
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "１５."
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
      Left            =   3240
      TabIndex        =   57
      Top             =   4410
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "１."
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
      Left            =   3420
      TabIndex        =   56
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "薪資          稅率(%)    累進差額"
      Height          =   180
      Left            =   3930
      TabIndex        =   55
      Top             =   570
      Width           =   2325
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "薪   資   扣   除   額："
      Height          =   180
      Left            =   150
      TabIndex        =   54
      Top             =   1980
      Width           =   1620
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "標準扣除額－單身："
      Height          =   180
      Left            =   150
      TabIndex        =   53
      Top             =   1650
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "標準扣除額－夫妻："
      Height          =   180
      Left            =   150
      TabIndex        =   42
      Top             =   1320
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "扶養親屬寬減額    ："
      Height          =   180
      Left            =   150
      TabIndex        =   40
      Top             =   990
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "免          稅          額："
      Height          =   180
      Left            =   150
      TabIndex        =   38
      Top             =   660
      Width           =   1620
   End
End
Attribute VB_Name = "frm170015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/21 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/29 add by sonia
Option Explicit

Dim m_FieldList() As FIELDITEM
Dim TF_IT As Integer '欄位數
Dim oText As Object
Dim idx As Integer


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0  '存檔
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               InitialField
               MsgBox "更新所得稅率表資料完成！", vbInformation
            Else
               MsgBox "更新所得稅率表資料錯誤！", vbInformation
            End If
         End If
      Case 1  '結束
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   textCUID.BackColor = &H8000000F
   InitialField
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170015 = Nothing
End Sub

' 初始化欄位陣列及抓資料
Private Sub InitialField()
Dim CUID(1 To 6) As String
   
   strExc(0) = "select * from IncomeTax "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      ClearField
      With RsTemp
         TF_IT = .Fields.Count
         ReDim m_FieldList(TF_IT) As FIELDITEM
         For Each oText In txtIT
            idx = oText.Index
            m_FieldList(idx).fiName = "IT" & Format(idx, "00")
            'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
            'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
               m_FieldList(idx).fiType = 0
            'Else
            '   m_FieldList(idx).fiType = 1
            'End If
            'end 2017/06/29
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
            m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
            oText.Text = m_FieldList(idx).fiOldData
         Next
         CUID(1) = "" & .Fields("it51")
         CUID(2) = "" & .Fields("it52")
         CUID(3) = "" & .Fields("it53")
         CUID(4) = "" & .Fields("it54")
         CUID(5) = "" & .Fields("it55")
         CUID(6) = "" & .Fields("it56")
      End With
   End If
   UpdateCUID CUID, textCUID
   If Me.Visible = True Then
      txtIT(1).SetFocus
      txtIT_GotFocus 1
   End If
   
End Sub

Private Sub ClearField()
   For Each oText In txtIT
      oText.Text = Empty
   Next
   
   For intI = 1 To TF_IT
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""

End Sub

Private Sub UpdateFieldNewData()
   For Each oText In txtIT
      idx = oText.Index
      m_FieldList(idx).fiNewData = oText.Text
   Next
End Sub

Private Function ModRecord() As Boolean
Dim stSQL As String, stSet As String, stCols As String, stValues As String
Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE IncomeTax SET "
   stSet = ""
   For Each oText In txtIT
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where it10 is not null; end; "
      
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   For Each oText In txtIT
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtIT_Validate idx, bCancel
         If bCancel = True Then
            txtIT(idx).SetFocus
            txtIT_GotFocus idx
            Exit Function
         End If
      End If
   Next
   
   If txtIT(1) = "" Then
      ShowMsg "請輸入免稅額 !"
      txtIT(1).SetFocus
      txtIT_GotFocus 1
      Exit Function
   End If
   If txtIT(2) = "" Then
      ShowMsg "請輸入扶養親屬寬減額 !"
      txtIT(2).SetFocus
      txtIT_GotFocus 2
      Exit Function
   End If
   If txtIT(3) = "" Then
      ShowMsg "請輸入標準扣除額－夫妻 !"
      txtIT(3).SetFocus
      txtIT_GotFocus 3
      Exit Function
   End If
   If txtIT(4) = "" Then
      ShowMsg "請輸入標準扣除額－單身 !"
      txtIT(4).SetFocus
      txtIT_GotFocus 4
      Exit Function
   End If
   If txtIT(5) = "" Then
      ShowMsg "請輸入薪資扣除額 !"
      txtIT(5).SetFocus
      txtIT_GotFocus 5
      Exit Function
   End If
   If txtIT(6) = "" Then
      ShowMsg "請輸入薪資級距的第 1 項 !"
      txtIT(6).SetFocus
      txtIT_GotFocus 6
      Exit Function
   End If
   If txtIT(7) = "" Then
      ShowMsg "請輸入稅率級距的第 1 項 !"
      txtIT(7).SetFocus
      txtIT_GotFocus 7
      Exit Function
   End If
   
   TxtValidate = True
    
End Function

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub txtIT_GotFocus(Index As Integer)
   TextInverse txtIT(Index)
   CloseIme
End Sub

Private Sub txtIT_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub txtIT_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 39, 42, 45, 48
         If txtIT(Index) <> "" Then
            If Val(txtIT(Index)) <= Val(txtIT(Index - 3)) Then
               ShowMsg "薪資級距前後有錯誤, 請詳細檢查 !"
               Cancel = True
            End If
         End If
      Case 10, 13, 16, 19, 22, 25, 28, 31, 34, 37, 40, 43, 46, 49
         If txtIT(Index) = "" Then
            If txtIT(Index - 1) <> "" Then
               ShowMsg "薪資級距有輸入, 請輸入此級距的稅率 !"
               Cancel = True
            End If
         Else
            If Val(txtIT(Index)) <= Val(txtIT(Index - 3)) Then
               ShowMsg "稅率級距前後有錯誤, 請詳細檢查 !"
               Cancel = True
            End If
         End If
   End Select
   
   If Cancel = True Then TextInverse txtIT(Index)
End Sub
