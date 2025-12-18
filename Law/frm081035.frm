VERSION 5.00
Begin VB.Form frm081035 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   2256
   ClientLeft      =   6768
   ClientTop       =   2676
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2256
   ScaleWidth      =   4740
   Begin VB.OptionButton Option1 
      Caption         =   "收  文  號："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "本所案號："
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtCP04 
      Height          =   372
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtCP01 
      Height          =   372
      Left            =   1464
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1200
      Width           =   550
   End
   Begin VB.TextBox txtCP03 
      Height          =   372
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtCP02 
      Height          =   372
      Left            =   2064
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtDKind 
      Height          =   372
      Left            =   600
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox txtDYear 
      Height          =   372
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox txtDNum 
      Height          =   372
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   4
      Top             =   660
      Width           =   1935
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2832
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdGoInput 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3660
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
End
Attribute VB_Name = "frm081035"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/03/25 Form2.0已檢查 (無需修改的物件)
Option Explicit

Dim m_LC01 As String
Dim m_LC02 As String
Dim m_LC03 As String
Dim m_LC04 As String
Dim m_CP09 As String

Private Sub cmdGoInput_Click(Index As Integer)
   Select Case Index
          Case 0 '確定
               If Not CheckText Then Exit Sub
                  '選擇收文號
                  If Option1.Value Then
                      
                     '若案件已閉卷, 不可發文
                     If PUB_CaseClosedCP09(Me.txtDNum.Text) = True Then
                        Exit Sub
                     End If
                     
                     '檢查是否有承辦歷程是否有產生承辦單可以發文
                     If PUB_IsEmpFlowIsSend(Me.txtDNum.Text) = False Then
                        Exit Sub
                     End If
                      
                     If CheckReceive(txtDNum.Text) Then
                        Call frm081035_2.SetParent(Me, txtDNum)
                        frm081035_2.Show
                     Else
                        MsgBox "無符合發文條件的資料!", vbExclamation, "發文"
                        txtDNum.SetFocus
                        TextInverse txtDNum
                        Exit Sub
                     End If
                   
                  '選擇本所案號
                  ElseIf Option2.Value Then
                             
                     '若案件已閉卷, 不可發文
                     If PUB_CaseClosed(Me.txtCP01.Text, Me.txtCP02.Text, Me.txtCP03.Text, Me.txtCP04.Text) = True Then
                        Exit Sub
                     End If
                       
                     If txtCP03 = "" Then txtCP03 = "0"
                     If txtCP04 = "" Then txtCP04 = "00"
                     If ClsPDCheckCaseCodeIsExist(txtCP01, txtCP02, txtCP03, txtCP04) = False Then
                        MsgBox "此筆本所案號不存在!", vbExclamation, "發文"
                        txtCP01.SetFocus
                        TextInverse txtCP01
                        Exit Sub
                     End If
                     If CheckCASEPROGRESS(txtCP01, txtCP02, txtCP03, txtCP04) Then
                         If m_CP09 = "" Then
                            Call frm081035_1.SetParent(Me)
                            frm081035_1.Show
                            If IsNoExistData Then
                               IsNoExistData = False
                               Unload frm081035_1
                            Else
                               Me.Hide
                            End If
                         Else
                            Call frm081035_2.SetParent(Me, m_CP09)
                            frm081035_2.Show
                            If IsNoExistData Then
                               IsNoExistData = False
                               Unload frm081035_2
                            Else
                               Me.Hide
                            End If
                         End If
                      Else
                         MsgBox "無符合發文條件的資料!", vbExclamation, "發文"
                         txtCP02.SetFocus
                         TextInverse txtCP02
                         Exit Sub
                      End If
                  End If
                  
            Case 1 '結束
                 Unload Me
                   
        End Select

End Sub

Private Sub Form_Activate()

   txtCP02.SetFocus
End Sub

Private Sub Form_Initialize()
   txtCP01.SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter frm081035
   
   txtCP01 = "ACS"
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frm081035 = Nothing
End Sub

Private Sub Option1_Click()
 Dim strYY As String
   txtCP01.Text = ""
   txtCP02.Text = ""
   txtCP03.Text = ""
   txtCP04.Text = ""
   txtCP01.Enabled = False
   txtCP02.Enabled = False
   txtCP03.Enabled = False
   txtCP04.Enabled = False
   txtDKind.Enabled = True
   txtDYear.Enabled = True
   txtDNum.Enabled = True
   strYY = Year(Date) - 1911
   txtDNum.SetFocus

End Sub

Private Sub Option2_Click()
   txtDNum.Text = ""
   txtDKind.Enabled = False
   txtDYear.Enabled = False
   txtDNum.Enabled = False
   txtCP01.Enabled = True
   txtCP02.Enabled = True
   txtCP03.Enabled = True
   txtCP04.Enabled = True

End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtCP01
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
  Dim strTit As String
  Dim strMsg As String
  
   txtCP01 = UCase(txtCP01)
   If IsEmptyText(txtCP01) = False Then

      If txtCP01 <> "ACS" And CheckSys(txtCP01) <> "3" And CheckSys(txtCP01) <> "4" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         TextInverse txtCP01
         Exit Sub
      End If
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, txtCP01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         TextInverse txtCP01
         Exit Sub
      End If
   End If
 
End Sub

Private Sub txtcp02_GotFocus()
TextInverse txtCP02
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
  If txtCP02 <> "" And Len(txtCP02) <> 6 Then
     DataErrorMessage 1, "本所案號"
     Cancel = True
     TextInverse txtCP02
  End If

End Sub
Private Sub txtcp03_GotFocus()
TextInverse txtCP03
End Sub

Private Sub txtcp04_GotFocus()
TextInverse txtCP04
End Sub


Private Sub txtDNum_GotFocus()
TextInverse txtDNum
End Sub

Private Sub txtDNum_KeyPress(KeyAscii As Integer)
  KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDNum_Validate(Cancel As Boolean)
Dim LcTmp As String, lc01 As String, lc02 As String, lc03 As String, lc04 As String
Dim strCP01 As String
Dim strTit As String
Dim strMsg As String
txtDNum.Text = UCase(txtDNum.Text)
If txtDNum.Text <> "" Then
   cmdGoInput(0).Default = True

   strCP01 = CheckSystemKind(txtDNum.Text)
   If strCP01 <> "" Then
      If strCP01 <> "ACS" And CheckSys(strCP01) <> "3" And CheckSys(strCP01) <> "4" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         Exit Sub
      End If
      
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, strCP01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         Exit Sub
      End If
   Else
      MsgBox "此筆收文號不存在!", vbExclamation, "發文"
      Cancel = True
      Exit Sub
   End If
   
      
   LcTmp = txtDNum

End If
If Cancel Then TextInverse txtDNum
End Sub

Private Function CheckText() As Boolean
If Option1.Value Then
   If txtDNum = "" Then
      CheckText = False
      DataErrorMessage 5, "收文號"
   Else
      CheckText = True
   End If
ElseIf Option2.Value Then
   If txtCP01 = "" Or txtCP02 = "" Then
      CheckText = False
      DataErrorMessage 5, "本所案號"
   Else
      CheckText = True
   End If
End If
End Function
Private Function CheckCASEPROGRESS(strLC01 As String, strLC02 As String, strLC03 As String, strLC04 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   '取消CP09<'C'條件,否則C來函無法發文
   strSql = "select CP09 from lawcase,caseprogress" + _
            " where cp01='" & strLC01 & "'" & _
            " AND CP02 ='" & strLC02 & "'" & _
            " AND CP03 ='" & strLC03 & "'" & _
            " AND CP04 ='" & strLC04 & "'" & _
            " and CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 " & _
            " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   m_CP09 = ""
   If rsTmp.EOF = False Then
      If rsTmp.RecordCount > 1 Then

      ElseIf rsTmp.RecordCount = 1 Then
         If Not IsNull(rsTmp.Fields("CP09")) Then
            m_CP09 = rsTmp.Fields("CP09")
         End If
      End If
   
      CheckCASEPROGRESS = True
   Else
      CheckCASEPROGRESS = False
   End If

End Function
Private Function CheckReceive(strNum As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM CASEPROGRESS WHERE CP09 ='" & strNum & "'" & _
            " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      CheckReceive = True
   Else
      CheckReceive = False
   End If

End Function
'***************************
'以總收文號取本所案號系統別*
'***************************
Private Function CheckSystemKind(strNum As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT CP01 FROM CASEPROGRESS WHERE CP09 ='" & strNum & "'" & _
            " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      If Not IsNull(rsTmp.Fields("CP01")) Then
         CheckSystemKind = rsTmp.Fields("CP01")
      Else
         CheckSystemKind = ""
      End If
   Else
      CheckSystemKind = ""
   End If

End Function

