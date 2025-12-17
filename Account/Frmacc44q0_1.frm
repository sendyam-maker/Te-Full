VERSION 5.00
Begin VB.Form Frmacc44q0_1 
   Caption         =   "作業說明"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   8120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   8120
   Begin VB.TextBox Text2 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
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
      Height          =   740
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Frmacc44q0_1.frx":0000
      Top             =   90
      Width           =   6020
   End
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6720
      TabIndex        =   1
      Top             =   180
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3440
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Frmacc44q0_1.frx":0036
      Top             =   900
      Width           =   8000
   End
End
Attribute VB_Name = "Frmacc44q0_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/3/24 Form2.0已檢查 (無需修改的物件)
Option Explicit


Private Sub Command3_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc44q0_1 = Nothing
End Sub
