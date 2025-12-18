VERSION 5.00
Begin VB.Form frm010010 
   BorderStyle     =   1  '單線固定
   Caption         =   "主管機關來函資料刪除作業"
   ClientHeight    =   1800
   ClientLeft      =   1710
   ClientTop       =   2130
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4335
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3324
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2496
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtDate 
      Height          =   264
      Index           =   1
      Left            =   2490
      MaxLength       =   7
      TabIndex        =   1
      Top             =   840
      Width           =   1035
   End
   Begin VB.TextBox txtDate 
      Height          =   264
      Index           =   0
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   0
      Top             =   840
      Width           =   1035
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2400
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Caption         =   "收件日期："
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   972
   End
End
Attribute VB_Name = "frm010010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/22 日期欄已修改
Option Explicit

'Add By Cheng 2002/09/10
Dim blnClkSure As Boolean '判斷是否按下確定按鈕


Private Sub cmdok_Click(Index As Integer)
Dim i As Integer, lngDelete As Long

'Add By Cheng 2002/09/10
blnClkSure = False

If Index = 0 Then
   For i = 0 To 1
          If CheckKeyIn(i) = False Then
             txtDate(i).SetFocus
             txtDate_GotFocus i
             Exit Sub
          End If
   Next
   'Add By Cheng 2002/09/10
   If Me.txtDate(0).Text <> "" And Me.txtDate(1).Text <> "" Then
      If Val(Me.txtDate(0).Text) > Val(Me.txtDate(1).Text) Then
         MsgBox "收件日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.txtDate(0).SetFocus
         txtDate_GotFocus 0
         Exit Sub
      End If
   End If
   'edit by nickc 2007/02/06 不用 dll 了
   'If obj001.DeleteCkind(txtDate(0), txtDate(1), lngDelete) Then
   If Cls001DeleteCkind(txtDate(0), txtDate(1), lngDelete) Then
      ShowMsg MsgText(1046) + Format(lngDelete) + MsgText(1047)
   End If
End If
Unload Me
End Sub
Private Sub Form_Load()
MoveFormToCenter Me
'edit by nickc 2007/02/06 不用 dll 了
'If obj001 Is Nothing Then
'   Set obj001 = CreateObject("prjTaieDll001.cls001")
'   Set obj001.Connection = cnnConnection
'End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2007/02/06 不用 dll 了
'Set obj001 = Nothing
'Add By Cheng 2002/07/18
Set frm010010 = Nothing
End Sub
Private Sub txtDate_GotFocus(Index As Integer)
txtDate(Index).SelStart = 0
txtDate(Index).SelLength = Len(txtDate(Index))
End Sub

Private Sub txtDate_LostFocus(Index As Integer)
   Select Case Index
   Case 1 '收件日期
      If blnClkSure = False Then
         If Me.txtDate(0).Text <> "" And Me.txtDate(1).Text <> "" Then
            If Val(Me.txtDate(0).Text) > Val(Me.txtDate(1).Text) Then
               MsgBox "收件日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.txtDate(0).SetFocus
               txtDate_GotFocus 0
               Exit Sub
            End If
         End If
      Else
         blnClkSure = False
      End If
   End Select
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = False Then
   Cancel = True
   txtDate_GotFocus Index
End If
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Boolean
Select Case intIndex
             Case 0
                        If CheckIsTaiwanDate(txtDate(intIndex).Text) Then
                            CheckKeyIn = True
                        End If
             Case 1
                        If CheckIsTaiwanDate(txtDate(intIndex).Text) Then
                           'Modify By Cheng 2002/09/10
'                           If txtDate(1) >= txtDate(0) Then
                              CheckKeyIn = True
'                           Else
'                              ShowMsg MsgText(1048)
'                           End If
                        End If
End Select
End Function

