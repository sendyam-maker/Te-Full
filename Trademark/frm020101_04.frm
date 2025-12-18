VERSION 5.00
Begin VB.Form frm020101_04 
   BorderStyle     =   1  '單線固定
   Caption         =   "陳述意見書對造資料"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5535
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   315
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   1035
   End
   Begin VB.TextBox textCP143 
      Height          =   264
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   0
      Top             =   600
      Width           =   885
   End
   Begin VB.TextBox textCP36 
      Height          =   264
      Left            =   3390
      MaxLength       =   15
      TabIndex        =   1
      Top             =   600
      Width           =   1365
   End
   Begin VB.TextBox textCP21 
      Height          =   264
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1110
      Width           =   435
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "申請日："
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   6
      Top             =   630
      Width           =   720
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Index           =   1
      Left            =   2460
      TabIndex        =   5
      Top             =   630
      Width           =   900
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "是否為快軌案件：            (Y/N)"
      Height          =   180
      Index           =   2
      Left            =   450
      TabIndex        =   4
      Top             =   1155
      Width           =   2385
   End
End
Attribute VB_Name = "frm020101_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/21 Form2.0已檢查 (無需修改的物件)
'Create By Sindy 2020/10/20
Option Explicit

Public oNextForm As Form
Public m_TM01 As String
Public m_TM02 As String
Public m_TM03 As String
Public m_TM04 As String
Public m_TM10 As String


Private Sub cmdOK_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
      
   If IsEmptyText(textCP143) = True Then
      strTit = "檢核資料"
      strMsg = "申請日不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP143.SetFocus
      Exit Sub
   End If
   If IsEmptyText(textCP36) = True Then
      strTit = "檢核資料"
      strMsg = "申請案號不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP36.SetFocus
      Exit Sub
   End If
   If IsEmptyText(textCP21) = True Then
      strTit = "檢核資料"
      strMsg = "是否為快軌案件不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP21.SetFocus
      Exit Sub
   End If
   
   oNextForm.m_CP143 = textCP143.Text
   oNextForm.m_CP36 = textCP36.Text
   oNextForm.m_CP21 = textCP21.Text
   
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   textCP143.Text = oNextForm.m_CP143
   textCP36.Text = oNextForm.m_CP36
   textCP21.Text = oNextForm.m_CP21
   
   'Screen.MousePointer = vbHourglass
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020101_04 = Nothing
End Sub

Private Sub textCP36_GotFocus()
   InverseTextBox textCP36
End Sub
Private Sub textCP36_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCP36_Validate(Cancel As Boolean)
Dim strRetrunText As String
   
   'If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If IsEmptyText(textCP36) = False Then
      '檢查申請案號所輸入的長度是否正確
      If PUB_ChkTm12Tm15Length("1", textCP36, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , True, strRetrunText) = False Then
         Cancel = True
         textCP36_GotFocus
         Exit Sub
      Else
         textCP36 = strRetrunText
      End If
   End If
End Sub

Private Sub textCP143_GotFocus()
   InverseTextBox textCP143
End Sub

'申請日
Private Sub textCP143_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP143) = False Then
      If CheckIsTaiwanDate(textCP143, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請日的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Private Sub textCP21_GotFocus()
   TextInverse textCP21
End Sub

Private Sub textCP21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
