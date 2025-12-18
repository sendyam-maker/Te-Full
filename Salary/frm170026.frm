VERSION 5.00
Begin VB.Form frm170026 
   BorderStyle     =   1  '單線固定
   Caption         =   "特殊功績獎金清除"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   3030
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   1650
      TabIndex        =   1
      Top             =   240
      Width           =   1000
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定清除(&O)"
      Height          =   405
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "frm170026"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/23 Form2.0已檢查 (無需修改的物件)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/1/10 add by sonia
Option Explicit


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '扣款計算
On Error GoTo ErrHand
         cnnConnection.BeginTrans
         'modify by sonia 2018/1/11 +sd51
         cnnConnection.Execute "update salarydata set sd18=null,sd51=null "
         cnnConnection.CommitTrans
         MsgBox "特殊功績獎金清除完畢！", vbInformation
      Case 1 '結束
         Unload Me
   End Select
   Exit Sub

ErrHand:
    cnnConnection.RollbackTrans
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170026 = Nothing
End Sub
