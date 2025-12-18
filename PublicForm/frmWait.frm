VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   1  '單線固定
   Caption         =   "寄信倒數"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   1560
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Left            =   180
      Top             =   510
   End
   Begin VB.Label lblCountDown 
      Alignment       =   2  '置中對齊
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1440
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1350
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
Public iWaitSec As Integer

Private Sub Form_Load()
   lblCountDown.Caption = iWaitSec
   Me.Left = Screen.Width / 2 - Me.Width / 2
   Me.Top = Screen.Height / 2 - Me.Height / 2
   'Me.Width = 0
   'Me.Height = 0
End Sub

Private Sub Timer1_Timer()
   Timer1.Tag = Val(Timer1.Tag) + 1
   If Val(Timer1.Tag) > iWaitSec Then
      Unload Me
   Else
      lblCountDown.Caption = iWaitSec - Val(Timer1.Tag)
   End If
End Sub

