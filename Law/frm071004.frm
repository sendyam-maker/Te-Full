VERSION 5.00
Begin VB.Form frm071004 
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
   Begin VB.TextBox txtcp04 
      Height          =   372
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtcp01 
      Height          =   372
      Left            =   1464
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1200
      Width           =   550
   End
   Begin VB.TextBox txtcp03 
      Height          =   372
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtcp02 
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
Attribute VB_Name = "frm071004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/09/14 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim com1 As Boolean, com2 As Boolean, com3 As Boolean, com4 As Boolean
Dim m_LC01 As String
Dim m_LC02 As String
Dim m_LC03 As String
Dim m_LC04 As String
Dim m_count As Integer
Public m_CP09 As String

Private Sub cmdGoInput_Click(Index As Integer)
  m_count = 0
   Select Case Index
          Case 0 '確定
               If Not CheckText Then Exit Sub
                  '選擇收文號
                  If Option1.Value Then
                      
                     'Add By Cheng 2002/07/12
                     '若案件已閉卷, 不可發文
                     If PUB_CaseClosedCP09(Me.txtDNum.Text) = True Then
                        Exit Sub
                     End If
                     
                     'Add By Sindy 2021/10/8
                     '檢查是否有承辦歷程是否有產生承辦單可以發文
                     If PUB_IsEmpFlowIsSend(Me.txtDNum.Text) = False Then
                        Exit Sub
                     End If
                     '2021/10/8 END
                      
                     If CheckReceive(txtDNum.Text) Then
                        intForm = 4
                        frm071006.Show
                     Else
                        MsgBox "無符合發文條件的資料!", vbExclamation, "發文"
                        txtDNum.SetFocus
                        TextInverse txtDNum
                        Exit Sub
                     End If
                   
                  '選擇本所案號
                  ElseIf Option2.Value Then
                             
                     'Add By Cheng 2002/07/12
                     '若案件已閉卷, 不可發文
                     If PUB_CaseClosed(Me.txtcp01.Text, Me.txtcp02.Text, Me.txtcp03.Text, Me.txtcp04.Text) = True Then
                        Exit Sub
                     End If
                       
                     If txtcp03 = "" Then txtcp03 = "0"
                     If txtcp04 = "" Then txtcp04 = "00"
                     If ClsPDCheckCaseCodeIsExist(txtcp01, txtcp02, txtcp03, txtcp04) = False Then
                        MsgBox "此筆本所案號不存在!", vbExclamation, "發文"
                        txtcp01.SetFocus
                        TextInverse txtcp01
                        Exit Sub
                     End If
                     If CheckCASEPROGRESS(txtcp01, txtcp02, txtcp03, txtcp04) Then
                            If m_count = 2 Then
                               intForm = 5
                               frm071005.Show
                               If IsNoExistData Then
                                  If Option1.Value Then
                                     Unload frm071006
                                  Else
                                     Unload frm071005
                                  End If
                                  IsNoExistData = False
                                Else
                                  Me.Hide
                                End If
                            ElseIf m_count = 1 Then
                                intForm = 4
                                frm071006.Show
                            End If
                      Else
                            MsgBox "無符合發文條件的資料!", vbExclamation, "發文"
                            txtcp01.SetFocus
                            TextInverse txtcp01
                            Exit Sub
                      End If
                            
                  End If
                  
            Case 1 '結束
                 Unload Me
                   
        End Select

End Sub

Private Sub Form_Initialize()
   txtcp01.SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter frm071004
   m_count = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm071004 = Nothing
End Sub

Private Sub Option1_Click()
 Dim strYY As String
   txtcp01.Text = ""
   txtcp02.Text = ""
   txtcp03.Text = ""
   txtcp04.Text = ""
   txtcp01.Enabled = False
   txtcp02.Enabled = False
   txtcp03.Enabled = False
   txtcp04.Enabled = False
   txtDKind.Enabled = True
   txtDYear.Enabled = True
   txtDNum.Enabled = True
   strYY = Year(Date) - 1911
   'If Len(strYY) > 2 Then strYY = Right(strYY, 2)
   txtDNum.SetFocus
   intForm = 4
End Sub

Private Sub Option2_Click()
   txtDNum.Text = ""
   txtDKind.Enabled = False
   txtDYear.Enabled = False
   txtDNum.Enabled = False
   txtcp01.Enabled = True
   txtcp02.Enabled = True
   txtcp03.Enabled = True
   txtcp04.Enabled = True
   intForm = 5
   txtcp01.SetFocus
End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
  Dim strTit As String
  Dim strMsg As String
  
   txtcp01 = UCase(txtcp01)
   If IsEmptyText(txtcp01) = False Then
      '2011/5/20 MODIFY BY SONIA
      'If Not IsCorrectSysKindLaw(txtcp01) Then
      If CheckSys(txtcp01) <> "3" And CheckSys(txtcp01) <> "4" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         TextInverse txtcp01
         Exit Sub
      End If
      
      'Added by Lydia 2024/03/25
      If txtcp01 = "ACS" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "不可輸入ACS案"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         TextInverse txtcp01
         Exit Sub
      End If
      'end 2024/03/25
      
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, txtcp01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         MsgBox strMsg, vbOKOnly, strTit
         Cancel = True
         TextInverse txtcp01
         Exit Sub
      End If
   End If

   
'   If txtcp01 <> "" Then
'      If txtcp01 = "L" Or txtcp01 = "LA" Or txtcp01 = "FCL" Then
'         com1 = True
'      Else
'         DataErrorMessage 1, "系統類別"
'         Cancel = True
'         TextInverse txtcp01
'      End If
'   Else
'  '    MsgBox "系統類別不得為空值 !", vbInformation
'   '   Option1.Value = True
'   End If
End Sub

Private Sub txtcp02_GotFocus()
TextInverse txtcp02
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
If txtcp02 <> "" Then
  If Len(txtcp02) = 6 Then
       com2 = True
    Else
        DataErrorMessage 1, "本所案號"
       Cancel = True
    End If
 
End If
If Cancel Then TextInverse txtcp02

End Sub
Private Sub txtcp03_GotFocus()
TextInverse txtcp03
End Sub
Private Sub txtcp03_Validate(Cancel As Boolean)
If txtcp03 <> "" Then
   com3 = True
End If
If Cancel Then TextInverse txtcp03
End Sub
Private Sub txtcp04_GotFocus()
TextInverse txtcp04
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
If txtcp04 <> "" Then
   com4 = True
End If
If Cancel Then TextInverse txtcp04
End Sub

'Private Sub txtDKind_GotFocus()
'TextInverse txtDKind
'End Sub
'
'Private Sub txtDKind_KeyPress(KeyAscii As Integer)
'KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub txtDKind_Validate(Cancel As Boolean)
'If Len(txtDKind) = 1 Then
'    txtDKind = UCase(txtDKind)
'ElseIf txtDNum = "" Or txtDYear = "" Or txtDKind = "" Then
'          Cancel = False
'Else
'        MsgBox "收文號輸入錯誤": Exit Sub
'         Cancel = True
' End If
'If Cancel Then TextInverse txtDKind
'
'End Sub

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
'   If Val(Mid(txtDNum.Text, 2, 2)) > Val(Format(Date, "EE")) - 1911 Then
'      MsgBox "收文號輸入錯誤"
'      txtDNum.SetFocus
'      TextInverse txtDNum
'      Exit Sub
'   End If
   strCP01 = CheckSystemKind(txtDNum.Text)
   If strCP01 <> "" Then
      '2011/5/20 MODIFY BY SONIA
      'If Not IsCorrectSysKindLaw(strCP01) Then
      If CheckSys(strCP01) <> "3" And CheckSys(strCP01) <> "4" Then
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
'   If objPublicData.CheckRecieveCode(LcTmp, m_LC01, m_LC02, m_LC03, m_LC04) Then
'       If Not (m_LC01 = "L" Or m_LC01 = "LA" Or m_LC01 = "FCL") Then
'           MsgBox "此收文號非 L 或 LA 或 FCL 之案件"
'           Cancel = True
'        End If
'    Else
'           Cancel = True
'    End If
End If
If Cancel Then TextInverse txtDNum
End Sub
Private Sub txtDYear_Validate(Cancel As Boolean)
'If txtDYear <> "" Then
'   If Val(txtDYear) > Val(Format(Date, "EE")) Then
'      MsgBox "收文號輸入錯誤"
'      Cancel = True
'   End If
'End If
'If Cancel Then TextInverse txtDYear

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
   If txtcp01 = "" Or txtcp02 = "" Then
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
   
  If strLC01 <> "LA" Then
'      strSQL = "select CP09 from lawcase,caseprogress" + _
'               " where cp01='" & strLC01 & "'" & _
'               " AND CP02 ='" & strLC02 & "'" & _
'               " AND CP03 ='" & strLC03 & "'" & _
'               " AND CP04 ='" & strLC04 & "'" & _
'               " and substr(cp09,1,1)<>'C'" & _
'               " and CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 " & _
'               " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
      '2009/9/9 MODIFY BY SONIA 取消CP09<'C'條件,否則C來函無法發文
      strSql = "select CP09 from lawcase,caseprogress" + _
               " where cp01='" & strLC01 & "'" & _
               " AND CP02 ='" & strLC02 & "'" & _
               " AND CP03 ='" & strLC03 & "'" & _
               " AND CP04 ='" & strLC04 & "'" & _
               " and CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 " & _
               " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
             
   Else
'      strSQL = "select CP09 from hirecase,caseprogress" + _
'               " where cp01='" & strLC01 & "'" & _
'               " AND CP02 ='" & strLC02 & "'" & _
'               " AND CP03 ='" & strLC03 & "'" & _
'               " AND CP04 ='" & strLC04 & "'" & _
'               " and substr(cp09,1,1)<>'C'" & _
'               " and CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 " & _
'               " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
      '2007/3/1 MODIFY BY SONIA 案件性質為顧問聘任不可發文, 否則到期期限表會印不出來
'      strSQL = "select CP09 from hirecase,caseprogress" + _
'               " where cp01='" & strLC01 & "'" & _
'               " AND CP02 ='" & strLC02 & "'" & _
'               " AND CP03 ='" & strLC03 & "'" & _
'               " AND CP04 ='" & strLC04 & "'" & _
'               " and CP09<'C'" & _
'               " and CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 " & _
'               " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
      '2009/9/9 MODIFY BY SONIA 取消CP09<'C'條件,否則C來函無法發文
      strSql = "select CP09 from hirecase,caseprogress" + _
               " where cp01='" & strLC01 & "'" & _
               " AND CP02 ='" & strLC02 & "'" & _
               " AND CP03 ='" & strLC03 & "'" & _
               " AND CP04 ='" & strLC04 & "'" & _
               " AND CP10<>'0'" & _
               " and CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 " & _
               " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
      '2007/3/1 END
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      If rsTmp.RecordCount > 1 Then
         m_count = 2
      ElseIf rsTmp.RecordCount = 1 Then
         If Not IsNull(rsTmp.Fields("CP09")) Then
            m_CP09 = rsTmp.Fields("CP09")
         End If
         m_count = 1
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
      '2007/3/1 MODIFY BY SONIA 案件性質為顧問聘任不可發文, 否則到期期限表會印不出來
      'CheckReceive = True
      If rsTmp.Fields("cp10") = "0" Then
         MsgBox "案件性質為顧問聘任不可發文, 否則到期期限表會印不出來!", vbCritical
         CheckReceive = False
      Else
         CheckReceive = True
      End If
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

