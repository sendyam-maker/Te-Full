VERSION 5.00
Begin VB.Form frm060204_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "事件說明"
   ClientHeight    =   7950
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   9130
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   5250
      TabIndex        =   5
      Text            =   "P.S若該案之新申請案進度尚未請款則不顯示"
      Top             =   5940
      Width           =   3585
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8070
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   2900
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frm060204_1.frx":0000
      Top             =   5040
      Width           =   9045
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   3290
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frm060204_1.frx":04F4
      Top             =   1710
      Width           =   9045
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   260
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frm060204_1.frx":099A
      Top             =   60
      Width           =   7695
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   1310
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frm060204_1.frx":09DD
      Top             =   360
      Width           =   9045
   End
End
Attribute VB_Name = "frm060204_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Create By Sindy 2017/1/16 原用Msgbox顯示但內容太多,無法完整顯示出來
Option Explicit


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060204_1 = Nothing
End Sub
