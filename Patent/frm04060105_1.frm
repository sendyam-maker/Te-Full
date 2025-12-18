VERSION 5.00
Begin VB.Form frm04060105_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內市場佔有率查詢"
   ClientHeight    =   1545
   ClientLeft      =   660
   ClientTop       =   2790
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4755
   Begin VB.CommandButton bottonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3000
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3840
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox text02_02 
      Height          =   264
      Left            =   3300
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1080
      Width           =   1092
   End
   Begin VB.TextBox text02_01 
      Height          =   264
      Left            =   1860
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1080
      Width           =   1092
   End
   Begin VB.TextBox text01_02 
      Height          =   264
      Left            =   3300
      MaxLength       =   7
      TabIndex        =   1
      Top             =   720
      Width           =   1092
   End
   Begin VB.TextBox text01_01 
      Height          =   264
      Left            =   1860
      MaxLength       =   7
      TabIndex        =   0
      Top             =   720
      Width           =   1092
   End
   Begin VB.Line Line2 
      X1              =   3060
      X2              =   3180
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "申請人國籍："
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Line Line1 
      X1              =   3060
      X2              =   3180
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "公告日："
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   852
   End
End
Attribute VB_Name = "frm04060105_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/18 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Private Sub bottonOK_Click()
   If CheckDataValid = True Then
      Screen.MousePointer = vbHourglass
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
      frm04060105_2.SetData text01_01, text01_02, text02_01, text02_02
      frm04060105_2.Show
      frm04060105_1.Hide
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub buttonExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   'Add By Cheng 2002/07/18
   Set frm04060105_1 = Nothing
End Sub

Private Sub text01_01_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmpty(text01_01) = False Then
      If CheckIsTaiwanDate(text01_01, False) = False Then
         Cancel = True
         strMsg = "日期不正確 !"
         strTit = "日期檢核"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
   If Cancel Then TextInverse text01_01
End Sub

Private Sub text01_02_LostFocus()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   If IsEmpty(text01_02) = False Then
      If CheckIsTaiwanDate(text01_02, False) = False Then
         strMsg = "請輸入正確的公告日 !"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         text01_02.SetFocus
         TextInverse text01_02
      Else
         If Not ChkRange(text01_01, text01_02, "公告日") Then
         
         End If
      End If
   End If
End Sub

Public Function IsEmpty(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmpty = False
   
   If Len(strData) <= 0 Then
      IsEmpty = True
   Else
      IsEmpty = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmpty = False
            Exit For
         End If
      Next nIndex
   End If
End Function

Public Function CheckDataValid() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   CheckDataValid = False
   
   If IsEmpty(text01_02) = True Then
      strMsg = "公告日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      GoTo EXITSUB
   End If
   '公告日起日必須小於止日
   If IsEmpty(text01_01) = False And IsEmpty(text01_02) = False Then
      If Val(text01_01) > Val(text01_02) Then
         strMsg = "公告日範圍不正確"
         strTit = "檢核輸入"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         text01_01.SetFocus
         TextInverse text01_01
         GoTo EXITSUB
      End If
   End If
   
   If IsEmpty(text02_02) = True Then
      strMsg = "申請人國籍必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      GoTo EXITSUB
   End If
   
   ' 國籍
   If IsEmpty(text02_01) = False And IsEmpty(text02_02) = False Then
      If Val(text02_01) > Val(text02_02) Then
         strMsg = "申請人國籍範圍不正確"
         strTit = "檢核輸入"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub text01_01_GotFocus()
   InverseAll text01_01
End Sub

Private Sub text01_02_GotFocus()
   InverseAll text01_02
End Sub

Private Sub text02_01_GotFocus()
   InverseAll text02_01
End Sub

Private Sub text02_02_GotFocus()
   InverseAll text02_02
End Sub

