VERSION 5.00
Begin VB.Form frm060206_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "事件說明"
   ClientHeight    =   4140
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9132
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   9132
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8070
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   1020
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frm060206_1.frx":0000
      Top             =   972
      Width           =   9045
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frm060206_1.frx":012D
      Top             =   450
      Width           =   9045
   End
End
Attribute VB_Name = "frm060206_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2023/8/9
Option Explicit


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060206_1 = Nothing
End Sub
