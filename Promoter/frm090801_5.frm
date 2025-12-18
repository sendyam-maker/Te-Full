VERSION 5.00
Begin VB.Form frm090801_5 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "微個體身分規範"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7245
      TabIndex        =   2
      Top             =   60
      Width           =   930
   End
   Begin VB.Label Label2 
      Caption         =   $"frm090801_5.frx":0000
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   135
      TabIndex        =   1
      Top             =   540
      Width           =   7995
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "微個體身分規範"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   7995
   End
End
Attribute VB_Name = "frm090801_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/22 改成Form2.0 (無)
'Create by Morgan 2013/3/20
Option Explicit
Dim m_MousePointer As Integer

Private Sub cmdOK_Click(Index As Integer)
   Unload Me
   Screen.MousePointer = m_MousePointer
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   m_MousePointer = Screen.MousePointer
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090801_5 = Nothing
End Sub
