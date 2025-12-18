VERSION 5.00
Begin VB.Form frm040104_3_1 
   AutoRedraw      =   -1  'True
   Caption         =   "補存取碼或優先權證明書"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   2880
   Begin VB.OptionButton Option1 
      Caption         =   "補存取碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1700
   End
   Begin VB.OptionButton Option1 
      Caption         =   "優先權證明書"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1700
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請選擇"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frm040104_3_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (無)
'Create by Amy 2014/03/20
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
    strPublicTemp = MsgText(601)
    If Index = 0 Then
        If Option1(0).Value = False And Option1(1).Value = False Then
            MsgBox "請選擇補存取碼或優先權證明書！", vbCritical
            Exit Sub
        End If
        If Option1(0).Value = True Then strPublicTemp = "1"
        If Option1(1).Value = True Then strPublicTemp = "2"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040104_3_1 = Nothing
End Sub

