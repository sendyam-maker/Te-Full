VERSION 5.00
Begin VB.Form Frmacc0001 
   BorderStyle     =   0  '沒有框線
   Caption         =   "Form1"
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   300
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "Frmacc0001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/2/9 Form2.0不用改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/26 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/28 日期欄已修改
Option Explicit
Dim strCounter As String

Private Sub Form_Load()
   Me.Height = 300
   Me.Width = 300
   Me.Move 9230, 0
   strCounter = "1"
   Image1 = LoadPicture(strPicPath & "1.bmp")
   Timer1.Interval = 50
End Sub

Private Sub Timer1_Timer()
   strCounter = Val(strCounter) + 1
   Image1 = LoadPicture(strPicPath & strCounter & ".bmp")
   If Val(strCounter) = 20 Then
      strCounter = "0"
   End If
End Sub
