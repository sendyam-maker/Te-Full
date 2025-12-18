VERSION 5.00
Begin VB.Form frm02010409_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務結果輸入"
   ClientHeight    =   1920
   ClientLeft      =   150
   ClientTop       =   2310
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4980
   Begin VB.TextBox textBTTM 
      Height          =   264
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   7
      Top             =   1080
      Width           =   2892
   End
   Begin VB.OptionButton radio 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1452
   End
   Begin VB.OptionButton radio 
      Caption         =   "BTTM :"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1452
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3924
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3096
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textCP05 
      Height          =   264
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1440
      Width           =   2892
   End
   Begin VB.TextBox textSP01 
      Height          =   264
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   2
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox textSP03 
      Height          =   264
      Left            =   3600
      MaxLength       =   1
      TabIndex        =   5
      Top             =   720
      Width           =   372
   End
   Begin VB.TextBox textSP04 
      Height          =   264
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   6
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox textSP02_2 
      Height          =   264
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textSP02 
      Height          =   264
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   3
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1452
   End
End
Attribute VB_Name = "frm02010409_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 (無需修改)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

Dim m_KeySel As Integer
Public m_SP01 As String
Public m_SP02 As String
Public m_SP03 As String
Public m_SP04 As String
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_RDate As String
Dim m_Done As Boolean
Dim m_PrevForm As Form
'2019/5/22 END


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
    '列印接洽接案單
'move to unload by nick 2004/10/22
''    PUB_PrintCaseCloseSheet strUserNum
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload Me
End Sub

'Added by Sindy 2019/5/22
Private Sub Form_Activate()
   If m_strIR01 <> "" And m_Done = False Then
      textSP01.Text = m_SP01
      If m_SP01 = "TF" Then
         textSP02.Text = Left(m_SP02, 5)
         textSP02_2.Text = Mid(m_SP02, 6)
      Else
         textSP02.Text = m_SP02
      End If
      textSP03.Text = m_SP03
      textSP04.Text = m_SP04
      textCP05.Text = m_RDate
      radio(0).Value = True
      m_KeySel = 0
      cmdOK.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
End Sub
'2019/5/22 END

Private Sub Form_Load()
   MoveFormToCenter Me
'   textCP05 = TAIWANDATE(SystemDate())
   textCP05 = strSrvDate(2)
   UpdateCtrlState
End Sub

Private Sub Initial()
   ' 預設由申請案號來取得資料
   m_KeySel = 0
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   ' 檢查輸入的欄位
   Select Case m_KeySel
      Case 0:
         If IsEmptyText(textSP01) = True Then
            strMsg = "請輸入本所案號"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      Case 1:
         If IsEmptyText(textBTTM) = True Then
            strMsg = "請輸入BTTM"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
   End Select
   ' 檢查來函收文日
   If IsEmptyText(textCP05) = True Then
      strMsg = "請輸入來函收文日"
      strTit = "檢核資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   Else
      If CheckIsTaiwanDate(textCP05, False) = False Then
         strMsg = "請輸入正確的來函收文日"
         strTit = "檢核資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      'edit by nickc 2006/03/17
      'If Val(textCP05) > Val(ChangeWDateStringToWString(Date)) Then
      If Val(textCP05) > Val(strSrvDate(1)) Then
         strMsg = "來函收文日不可超過系統日"
         strTit = "檢核資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 檢查來函記錄檔
Private Function PromptIfTaiwanNoResult() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strNation As String
   Dim bPrompt As Boolean
   
   bPrompt = False
   PromptIfTaiwanNoResult = True
   strNation = "111"
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("SP09")) = False Then
         strNation = rsTmp.Fields("SP09")
      End If
   End If
   rsTmp.Close
   
   If strNation < "010" Then
      strSql = "SELECT * FROM MailRec " & _
               "WHERE MR12 = '" & m_SP01 & "' AND " & _
                     "MR13 = '" & m_SP02 & "' AND " & _
                     "MR14 = '" & m_SP03 & "' AND " & _
                     "MR15 = '" & m_SP04 & "' AND " & _
                     "MR02 = " & ChangeTStringToWString(textCP05) & " "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("MR16")) = False Then
            If rsTmp.Fields("MR16") <> "0" Then
               bPrompt = True
            End If
         End If
      Else
         bPrompt = True
      End If
      rsTmp.Close
   End If
   
   If bPrompt = True Then
      strTit = "資料檢核"
      strMsg = "與櫃台之來函收文記錄不符, 請確認"
      nResponse = MsgBox(strMsg, vbOKCancel, strTit)
      If nResponse = vbCancel Then
         PromptIfTaiwanNoResult = False
      End If
   End If
   Set rsTmp = Nothing
End Function

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   
   'Add By Sindy 2019/5/22
   If m_strIR01 <> "" Then
      If m_SP01 & m_SP02 & m_SP03 & m_SP04 <> textSP01 & IIf(textSP01 = "TF", textSP02 & textSP02_2, textSP02) & textSP03 & textSP04 Then
         MsgBox "信件輸入必須與信件本所案號(" & m_SP01 & m_SP02 & m_SP03 & m_SP04 & ")一致！"
         Exit Sub
      End If
   End If
   '2019/5/22 END
   
   ' 檢查欄位的資料是否都已經輸入正確
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   ' 檢查所輸入的資料是否合乎資料庫的條件
   Select Case m_KeySel
      Case 0:
         m_SP01 = textSP01
         m_SP02 = textSP02
         m_SP03 = textSP03
         If IsEmptyText(m_SP03) = True Then: m_SP03 = "0"
         m_SP04 = textSP04
         If IsEmptyText(m_SP04) = True Then: m_SP04 = "00"
         ' 設定SQL語法
         strSql = "SELECT * FROM ServicePractice " & _
                  "WHERE SP01 = '" & m_SP01 & "' AND " & _
                        "SP02 = '" & m_SP02 & "' AND " & _
                        "SP03 = '" & m_SP03 & "' AND " & _
                        "SP04 = '" & m_SP04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         rsTmp.Close
      Case 1:
         ' 設定SQL語法
         strSql = "SELECT * FROM ServicePractice " & _
                  "WHERE SP50 = '" & textBTTM & "' AND " & _
                        "SP01 = 'TM' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         Else
            rsTmp.MoveFirst
            m_SP01 = rsTmp.Fields("SP01")
            m_SP02 = rsTmp.Fields("SP02")
            m_SP03 = rsTmp.Fields("SP03")
            m_SP04 = rsTmp.Fields("SP04")
         End If
         rsTmp.Close
   End Select
   ' 顯示下一個畫面
   If PromptIfTaiwanNoResult() Then
      DisplayNextForm
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub DisplayNextForm()
   frm02010409_2.SetData 0, m_SP01, True
   frm02010409_2.SetData 1, m_SP02, False
   frm02010409_2.SetData 2, m_SP03, False
   frm02010409_2.SetData 3, m_SP04, False
   frm02010409_2.SetData 4, textCP05, False
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Call frm02010409_2.SetParent(m_PrevForm)
   End If
   frm02010409_2.m_strIR01 = m_strIR01
   frm02010409_2.m_strIR02 = m_strIR02
   frm02010409_2.m_strIR03 = m_strIR03
   frm02010409_2.m_strIR04 = m_strIR04
   '2019/5/22 END
   Me.Hide
   frm02010409_2.Show
   frm02010409_2.QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   PUB_PrintCaseCloseSheet strUserNum
   PUB_PrintCaseCloseSheet strUserNum, "0", False, False
   '刪除暫存資料
   PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010409_1 = Nothing
End Sub

Private Sub radio_Click(Index As Integer)
   m_KeySel = Index
   UpdateCtrlState
   ' 設定游標停留的位置
   Select Case Index
      Case 0: If textSP01.Visible = True And textSP01.Enabled = True Then textSP01.SetFocus
      Case 1: If textBTTM.Visible = True And textBTTM.Enabled = True Then textBTTM.SetFocus
   End Select
End Sub

Private Sub UpdateCtrlState()
   Select Case m_KeySel
      Case 0:
         EnableTextBox textSP01, True
         EnableTextBox textSP02, True
         EnableTextBox textSP03, True
         EnableTextBox textSP04, True
         EnableTextBox textBTTM, False
         textSP01_Validate False
      Case 1:
         EnableTextBox textSP01, False
         EnableTextBox textSP02, False
         EnableTextBox textSP03, False
         EnableTextBox textSP04, False
         EnableTextBox textBTTM, True
         textSP02_2.Visible = False
   End Select
End Sub

Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsTaiwanDate(textCP05, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的來函收文日"
         strTit = "來函收文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
      'edit by nickc 2006/03/17
      'If Val(textCP05) > Val(ChangeWDateStringToWString(Date)) Then
      If Val(textCP05) > Val(strSrvDate(1)) Then
         Cancel = True
         strMsg = "來函收文日不可超過系統日"
         strTit = "來函收文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textSP01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textSP01) = False Then
      ' 系統類別為T類
      If Mid(textSP01, 1, 1) <> "T" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP01_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case textSP01
         Case "TF":
            textSP02_2.Visible = True
            textSP02_2.Locked = False
            textSP02_2.TabStop = True
            textSP02.MaxLength = 5
         Case Else
            textSP02_2.Visible = False
            textSP02_2.Locked = True
            textSP02_2.TabStop = False
            textSP02.MaxLength = 6
      End Select
   Else
      textSP02_2.Visible = False
      textSP02_2.Locked = True
      textSP02_2.TabStop = False
      textSP02.MaxLength = 6
   End If
EXITSUB:
End Sub

Private Sub textSP01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSP01_GotFocus()
   InverseTextBox textSP01
End Sub

Private Sub textSP02_2_GotFocus()
   InverseTextBox textSP02_2
End Sub

Private Sub textSP02_GotFocus()
   'Add By Cheng 2004/03/05
   TextInverse Me.textSP02
   'End
End Sub

Private Sub textSP03_GotFocus()
   InverseTextBox textSP03
End Sub

Private Sub textSP04_GotFocus()
   InverseTextBox textSP04
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Sub textBTTM_GotFocus()
   InverseTextBox textBTTM
End Sub
