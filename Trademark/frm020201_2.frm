VERSION 5.00
Begin VB.Form frm020201_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "事件說明"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8070
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3510
      TabIndex        =   1
      Top             =   6030
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   5835
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frm020201_2.frx":0000
      Top             =   120
      Width           =   7785
   End
End
Attribute VB_Name = "frm020201_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/11 Form2.0已修改 (無需修改)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/2 日期欄已修改
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020201_2 = Nothing
End Sub
