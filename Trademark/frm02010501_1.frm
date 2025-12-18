VERSION 5.00
Begin VB.Form frm02010501_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "爭議案勝訴輸入"
   ClientHeight    =   2385
   ClientLeft      =   570
   ClientTop       =   3090
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5055
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM12 
      Height          =   264
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1170
      Width           =   2892
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1500
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3600
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1500
      Width           =   372
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1500
      Width           =   1092
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1500
      Width           =   732
   End
   Begin VB.TextBox textTM15 
      Height          =   264
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   3
      Top             =   810
      Width           =   2892
   End
   Begin VB.TextBox textCP05 
      Height          =   264
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   10
      Top             =   1860
      Width           =   2892
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3336
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4164
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.OptionButton radio 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1500
      Width           =   1452
   End
   Begin VB.OptionButton radio 
      Caption         =   "審定號數 :"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   810
      Width           =   1452
   End
   Begin VB.OptionButton radio 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1140
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Left            =   240
      TabIndex        =   13
      Top             =   1860
      Width           =   1572
   End
End
Attribute VB_Name = "frm02010501_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy  2022/01/03 Form2.0已修改 (無需修改)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/6 日期欄已修改
Option Explicit

Dim m_KeySel As Integer
'Added by Morgan 2017/4/24 電子公文
Public m_RDate As String
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_RegNo As String
Public m_DeadLine As String
Public m_NewCP10 As String
Dim m_Done As Boolean
'end 2017/4/24
Public strTM01 As String
Public strTM02 As String
Public strTM03 As String
Public strTM04 As String
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   'Added by Sindy 2019/5/22
   If m_strIR01 <> "" And m_Done = False Then
      textTM01.Text = strTM01
      If strTM01 = "TF" Then
         textTM02.Text = Left(strTM02, 5)
         textTM02_2.Text = Mid(strTM02, 6)
      Else
         textTM02.Text = strTM02
      End If
      textTM03.Text = strTM03
      textTM04.Text = strTM04
      textCP05.Text = m_RDate
      If m_RegNo <> "" Then '審定號數
         textTM15 = m_RegNo
         radio(1).Value = True
         m_KeySel = 1
      ElseIf m_AppNo <> "" Then '申請案號
         textTM12 = m_AppNo
         radio(0).Value = True
         m_KeySel = 0
      Else '本所案號
         radio(2).Value = True
         m_KeySel = 2
      End If
      cmdOK.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
   'Added by Morgan 2017/4/24 電子公文
   If m_AppNo & m_RegNo <> "" And m_Done = False Then
      If m_RegNo <> "" Then
         radio(1).Value = True
         textTM15.Text = m_RegNo
         m_KeySel = 1
      Else
         radio(0).Value = True
         textTM12.Text = m_AppNo
         m_KeySel = 0
      End If
      textCP05.Text = m_RDate
      cmdOK.Value = True
      m_Done = True
   End If
   'end 2017/4/24
End Sub

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
         If IsEmptyText(textTM12) = True Then
            strMsg = "請輸入申請案號"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      Case 1:
         If IsEmptyText(textTM15) = True Then
            strMsg = "請輸入申請案號"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      Case 2:
         If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
            strMsg = "請輸入本所案號"
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

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   'Add By Sindy 2019/5/22
   If m_strIR01 <> "" Then
      If strTM01 & strTM02 & strTM03 & strTM04 <> textTM01 & IIf(textTM01 = "TF", textTM02 & textTM02_2, textTM02) & textTM03 & textTM04 Then
         MsgBox "信件輸入必須與信件本所案號(" & strTM01 & strTM02 & strTM03 & strTM04 & ")一致！"
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
         ' 設定SQL語法
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM12 = '" & textTM12 & "' AND " & _
                        "TM10 < '010' "
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
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM15 = '" & textTM15 & "' AND " & _
                        "TM10 < '010' "
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
      Case 2:
         strTM01 = textTM01
         strTM02 = textTM02
         strTM03 = textTM03
         strTM04 = textTM04
         
         Select Case strTM01
            Case "T", "FCT":
               If IsEmptyText(strTM03) = True Then
                  strTM03 = "0"
               End If
               If IsEmptyText(strTM04) = True Then
                  strTM04 = "00"
               End If
               ' 設定SQL語法
               strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' AND " & _
                        "TM03 = '" & strTM03 & "' AND " & _
                        "TM04 = '" & strTM04 & "' "
            Case "TF":
               strTM02 = Trim(textTM02) & Trim(textTM02_2)
               ' 設定SQL語法
               strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' "
               If IsEmptyText(strTM03) = False Then
                  strSql = strSql & "AND "
                  strSql = strSql & "TM03 = '" & strTM03 & "' "
               End If
               If IsEmptyText(strTM04) = False Then
                  strSql = strSql & "AND "
                  strSql = strSql & "TM04 = '" & strTM04 & "' "
               End If
            '2012/7/23 add by sonia TD-000151
            Case "TD":
               If IsEmptyText(strTM03) = True Then
                  strTM03 = "0"
               End If
               If IsEmptyText(strTM04) = True Then
                  strTM04 = "00"
               End If
               ' 設定SQL語法
               strSql = "SELECT SP09 AS TM10 FROM SERVICEPRACTICE " & _
                  "WHERE SP01 = '" & strTM01 & "' AND " & _
                        "SP02 = '" & strTM02 & "' AND " & _
                        "SP03 = '" & strTM03 & "' AND " & _
                        "SP04 = '" & strTM04 & "'"
            '2012/7/23 END
        End Select
         
         ' 若申請國家為台灣, 則須設定申請案號或審定號
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            '2012/7/23 MODIFY BY SONIA
            'If IsNull(rsTmp.Fields("TM10")) = False Then
            If IsNull(rsTmp.Fields("TM10")) = False And strTM01 <> "TD" Then
               If rsTmp.Fields("TM10") < "010" Then
                  rsTmp.Close
                  strMsg = "所選取的資料其申請國家為台灣, 請以申請案號或審定號數來搜尋"
                  strTit = "檢核資料"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  GoTo EXITSUB
               End If
            End If
         Else
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         rsTmp.Close
   End Select
   ' 顯示下一個畫面
   DisplayNextForm
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub DisplayNextForm()
   Select Case m_KeySel
      Case 0:
         frm02010501_2.SetData 4, textTM12, True
         frm02010501_2.SetData 6, textCP05, False
      Case 1:
         frm02010501_2.SetData 5, textTM15, True
         frm02010501_2.SetData 6, textCP05, False
      Case 2:
         frm02010501_2.SetData 0, Trim(textTM01), True
         If Trim(textTM01) = "TF" Then
            frm02010501_2.SetData 1, Trim(textTM02) & Trim(textTM02_2), False
         Else
            frm02010501_2.SetData 1, Trim(textTM02), False
         End If
         frm02010501_2.SetData 2, Trim(textTM03), False
         frm02010501_2.SetData 3, Trim(textTM04), False
         frm02010501_2.SetData 6, textCP05, False
   End Select
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Call frm02010501_2.SetParent(m_PrevForm)
   End If
   frm02010501_2.m_strIR01 = m_strIR01
   frm02010501_2.m_strIR02 = m_strIR02
   frm02010501_2.m_strIR03 = m_strIR03
   frm02010501_2.m_strIR04 = m_strIR04
   '2019/5/22 END
   Me.Hide
   frm02010501_2.Show
   frm02010501_2.QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010501_1 = Nothing
End Sub

Private Sub radio_Click(Index As Integer)
   m_KeySel = Index
   UpdateCtrlState
   ' 設定游標停留的位置
   If Me.Enabled And Screen.ActiveForm.Name = Me.Name Then 'Added by Morgan 2017/6/3 電子公文
      Select Case Index
         Case 0: textTM12.SetFocus
         Case 1: textTM15.SetFocus
         Case 2: textTM01.SetFocus
      End Select
   End If
End Sub

Private Sub UpdateCtrlState()
   Select Case m_KeySel
      Case 0:
         EnableTextBox textTM12, True
         EnableTextBox textTM15, False
         EnableTextBox textTM01, False
         EnableTextBox textTM02, False
         EnableTextBox textTM03, False
         EnableTextBox textTM04, False
         textTM02_2.Visible = False
      Case 1:
         EnableTextBox textTM12, False
         EnableTextBox textTM15, True
         EnableTextBox textTM01, False
         EnableTextBox textTM02, False
         EnableTextBox textTM03, False
         EnableTextBox textTM04, False
         textTM02_2.Visible = False
      Case 2:
         EnableTextBox textTM12, False
         EnableTextBox textTM15, False
         EnableTextBox textTM01, True
         EnableTextBox textTM02, True
         EnableTextBox textTM03, True
         EnableTextBox textTM04, True
         textTM01_Validate False
   End Select
End Sub
' 來函收文日
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

Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      ' 檢查系統類別
      If IsCorrectSysKind(textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      ' 檢查使用者權限
      If IsUserHasRightOfSystem(strUserNum, textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使用該系統類別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      Select Case textTM01
         '2012/7/23 MODIFY BY SONIA 加入TD
         Case "T", "FCT", "TD":
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
         Case "TF":
            textTM02_2.Visible = True
            textTM02_2.Locked = False
            textTM02_2.TabStop = True
            textTM02.MaxLength = 5
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM01_GotFocus
      End Select
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
EXITSUB:
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

Private Sub textTM12_GotFocus()
   InverseTextBox textTM12
End Sub

Private Sub textTM15_GotFocus()
   InverseTextBox textTM15
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub


