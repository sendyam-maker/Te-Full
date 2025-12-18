VERSION 5.00
Begin VB.Form frm050109_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸香港案件資料維護"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5835
   Begin VB.Frame fraChoose 
      Height          =   1332
      Index           =   0
      Left            =   90
      TabIndex        =   21
      Top             =   945
      Width           =   5652
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   8
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   10
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   4
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   6
         Top             =   600
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   7
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   9
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   6
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   8
         Top             =   600
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   5
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   7
         Top             =   600
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   3
         Top             =   240
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   4
         Top             =   240
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   3
         Left            =   2664
         MaxLength       =   2
         TabIndex        =   5
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "功能代號：           (1.新增  2.修改  4.刪除  5.查詢 )"
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   3972
      End
      Begin VB.Label Label1 
         Caption         =   "大陸案號："
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "香港案號："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraChoose 
      Enabled         =   0   'False
      Height          =   975
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   2745
      Width           =   5652
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   11
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   13
         Top             =   600
         Width           =   972
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   9
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   11
         Top             =   240
         Width           =   972
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   10
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   12
         Top             =   600
         Width           =   972
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   2640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label3 
         Caption         =   "香港案工程師："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "大陸案發文日："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblEnginer 
         Height          =   252
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   2172
      End
      Begin VB.Label Label3 
         Caption         =   "(民國年月日)"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4905
      TabIndex        =   15
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4080
      TabIndex        =   14
      Top             =   60
      Width           =   800
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "單筆維護"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   735
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "多筆查詢條件"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   2505
      Width           =   1455
   End
End
Attribute VB_Name = "frm050109_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/18 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'0從Menu來,1從frm050101_2來
Public intWhereToGo As Integer
'Added by Lydia 2015/07/27 +大陸澳門案(共用表單frm050109_1,frm050109_2,frm050109_3)
Public iK_CM10 As Integer  '判斷案件類別
Dim iK_PA09 As String, m_NA03 As String '案件-國別
'end 2015/07/27
Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer, varSaveCursor, strCode(7) As String
   Select Case Index
      Case 0 '確定
         varSaveCursor = Screen.MousePointer
         Screen.MousePointer = vbHourglass
         For i = 0 To 11
            If txtCode(i).Enabled Then
               If CheckKeyIn(i) = False Then
                  '本所案號錯誤時,將Cursor跳回系統別欄位
                  If i = 3 Or i = 7 Then i = i - 3
                  txtCode(i).SetFocus
                  txtCode_GotFocus i
                  Exit For
               End If
            End If
         Next
         If i = 12 Then
            If optChoose(0).Value Then
               If txtCode(2) = "" Then txtCode(2) = "0"
               If txtCode(3) = "" Then txtCode(3) = "00"
               If txtCode(6) = "" Then txtCode(6) = "0"
               If txtCode(7) = "" Then txtCode(7) = "00"
               If txtCode(8) = "4" Then
                  For i = 0 To 7
                     strCode(i) = txtCode(i)
                  Next
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If obj003.ChkExist(strCode(), 4) Then
                  'Modified by Lydia 2015/07/27
                  'If Cls003ChkExist(strCode(), 4) Then
                  If Cls003ChkExist(strCode(), iK_CM10) Then
                     If MsgBox("是否要刪除此筆資料 ?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                        cnnConnection.BeginTrans
                        'edit by nickc 2007/02/05 不用 dll 了
                        'If obj003.DeleteCaseRelation(strCode(), 4) Then
                        'Modified by Lydia 2015/07/27
                        'If Cls003DeleteCaseRelation(strCode(), 4) Then
                        If Cls003DeleteCaseRelation(strCode(), iK_CM10) Then
                           txtCode(0).SetFocus
                           strExc(1) = txtCode(0)
                           strExc(2) = txtCode(1)
                           strExc(3) = txtCode(2)
                           strExc(4) = txtCode(3)
                           'Remove by Morgan 2005/4/13 改由 trigger 更新
                           'If PUB_UpdateCaseValueA(strExc()) = False Then
                           '   cnnConnection.RollbackTrans
                           'Else
                              cnnConnection.CommitTrans
                              For i = 0 To 8
                                 txtCode(i) = ""
                              Next
                           'End If
                        Else
                           cnnConnection.RollbackTrans
                        End If
                        '2005/3/10 end
                     End If
                  End If
               
               ElseIf txtCode(8) = "1" Then
                  For i = 0 To 7
                     strCode(i) = txtCode(i)
                  Next
                  
                  '控制香港案之申請國家必須為香港
                  If CheckIsHongKong() = False Then
                    txtCode(0).SetFocus
                    txtCode_GotFocus 0
                    Screen.MousePointer = vbDefault
                    Exit Sub
                  End If
                  '控制大陸案之申請國家必須為大陸
                  If CheckIsCN() = False Then
                    txtCode(4).SetFocus
                    txtCode_GotFocus 4
                    Screen.MousePointer = vbDefault
                    Exit Sub
                  End If
                  'Modified by Lydia 2015/07/27
                  'If ChkCaseMap(strCode, 4) Then
                  If ChkCaseMap(strCode, iK_CM10) Then
                     GoTo A0
                  Else
                     txtCode(0).SetFocus
                  End If
               Else
A0:               frm050109_2.intWhereToGo = 0
                  frm050109_2.strCode1 = txtCode(0)
                  frm050109_2.strCode2 = txtCode(1)
                  frm050109_2.strCode3 = txtCode(2)
                  frm050109_2.strCode4 = txtCode(3)
                  frm050109_2.strCode5 = txtCode(4)
                  frm050109_2.strCode6 = txtCode(5)
                  frm050109_2.strCode7 = txtCode(6)
                  frm050109_2.strCode8 = txtCode(7)
                  frm050109_2.intChoose = Val(txtCode(8))
                  frm050109_2.iK_CM10 = iK_CM10 'Added by Lydia 2015/07/27
                  frm050109_2.Show
                  Me.Hide
               End If
            Else
               frm050109_3.lblEnginer = txtCode(9)
               frm050109_3.lblEnginerName = lblEnginer
               frm050109_3.lblDate(0) = ChangeTStringToTDateString(txtCode(10))
               frm050109_3.lblDate(1) = ChangeTStringToTDateString(txtCode(11))
               frm050109_3.iK_CM10 = iK_CM10 'Added by Lydia 2015/07/27
               Me.Hide
            End If
         End If
         Screen.MousePointer = varSaveCursor
      Case 1
         Unload Me
   End Select
End Sub
'Add by Morgan 2004/1/30
'修改成國外案號已發文的也不限制
'Copy From Dll003.ChkCaseMap
Private Function ChkCaseMap(ByRef strCode() As String, ByVal iSitu As Integer) As Boolean
    Dim strSql As String, rstQuery As New ADODB.Recordset
    Dim strTmp1(0 To 3) As String, strTmp2(0 To 3) As String, i As Integer
    ChkCaseMap = False
    strCode(2) = IIf(strCode(2) = "", "0", strCode(2))
    strCode(3) = IIf(strCode(3) = "", "00", strCode(3))
    strCode(6) = IIf(strCode(6) = "", "0", strCode(6))
    strCode(7) = IIf(strCode(7) = "", "00", strCode(7))
    
    rstQuery.CursorLocation = adUseClient
    '檢查是否國內外案件關聯資料已存在
    strSql = "select count(*) from casemap where cm01=" + CNULL(strCode(0)) + " and cm02=" + CNULL(strCode(1)) + _
       " and cm03=" + CNULL(strCode(2)) + " and cm04=" + CNULL(strCode(3)) + " and cm05=" + CNULL(strCode(4)) + _
       " and cm06=" + CNULL(strCode(5)) + " and cm07=" + CNULL(strCode(6)) + " and cm08=" + CNULL(strCode(7)) + _
       " and cm10='" & iSitu & "'"
    rstQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rstQuery.Fields(0) > 0 Then
       'Modified by Lydia 2015/07/27
       'ShowMsg "大陸香港案件關聯資料已存在，請重新輸入 !"
       ShowMsg "大陸" & m_NA03 & "案件關聯資料已存在，請重新輸入 !"
    Else
        '檢查是否已取消收文
        strSql = "SELECT 1 C1 FROM CASEPROGRESS WHERE CP01='" & strCode(0) & "' AND CP02='" & strCode(1) & _
            "' AND CP03='" & strCode(2) & "' AND CP04='" & strCode(3) & "' AND " & _
            "CP10 in ('" & 發明申請 & "','" & 新型申請 & "','" & 設計申請 & "','" & 追加申請 & "','" & 翻譯 & "','" & 113 & "','" & 114 & "','" & 307 & "')" & _
            " AND CP57 IS NOT NULL"
        strSql = strSql & " UNION ALL SELECT 2 C2 FROM CASEPROGRESS WHERE CP01='" & strCode(4) & "' AND CP02='" & strCode(5) & _
            "' AND CP03='" & strCode(6) & "' AND CP04='" & strCode(7) & "' AND " & _
            "CP10 in ('" & 發明申請 & "','" & 新型申請 & "','" & 設計申請 & "','" & 追加申請 & "','" & 聯合申請 & "','" & 翻譯 & "','" & 113 & "','" & 114 & "','" & 307 & "')" & _
            " AND CP57 IS NOT NULL"
        If rstQuery.State <> 0 Then
            rstQuery.Close
        End If
        rstQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rstQuery.RecordCount > 0 Then
            If rstQuery.Fields(0).Value = "1" Then
                ShowMsg strCode(0) & strCode(1) & strCode(2) & strCode(3) & " 已有取消收文日，請重新輸入 !"
            Else
                ShowMsg strCode(4) & strCode(5) & strCode(6) & strCode(7) & " 已有取消收文日，請重新輸入 !"
            End If
        Else
            ChkCaseMap = True
        End If
    End If
    Set rstQuery = Nothing
End Function

'檢查香港案號申請國家是否為香港
Private Function CheckIsHongKong() As Boolean
    Dim strSql As String, rstQuery As New ADODB.Recordset
On Error GoTo ErrHnd
    strSql = "Select PA09 From Patent Where PA01='" & txtCode(0) & "' AND PA02='" & txtCode(1) & "' AND PA03='" & txtCode(2) & "' AND PA04='" & txtCode(3) & "'"
    rstQuery.CursorLocation = adUseClient
    rstQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    'Modified by Lydia 2015/07/27
'    If rstQuery.RecordCount > 0 Then
'        If "" & rstQuery.Fields(0).Value = "013" Then
'            CheckIsHongKong = True
'        Else
'            MsgBox "香港案之申請國家必須為香港！", vbCritical, "警告"
'        End If
'    Else
'        MsgBox "無法讀取香港案之申請國家！", vbCritical, "警告"
'    End If
    If rstQuery.RecordCount > 0 Then
        If "" & rstQuery.Fields(0).Value = iK_PA09 Then
            CheckIsHongKong = True
        Else
            MsgBox m_NA03 & "案之申請國家必須為" & m_NA03 & "！", vbCritical, "警告"
        End If
    Else
        MsgBox "無法讀取" & m_NA03 & "案之申請國家！", vbCritical, "警告"
    End If
    'end 2015/07/27
    
    Set rstQuery = Nothing
ErrHnd:
    If Err.NUMBER <> 0 Then
        MsgBox Err.Description
    End If
End Function

'檢查大陸案號申請國家是否為大陸
Private Function CheckIsCN() As Boolean
    Dim strSql As String, rstQuery As New ADODB.Recordset
On Error GoTo ErrHnd
    'Modify by Morgan 2007/4/26 香港也可和EPC、英國關聯
    'strSQL = "Select PA09 From Patent Where PA01='" & txtCode(4) & "' AND PA02='" & txtCode(5) & "' AND PA03='" & txtCode(6) & "' AND PA04='" & txtCode(7) & "'"
    strSql = "select p1.pa09, p2.pa09 pa09x from patent p1,patent p2 where p1.pa01='" & txtCode(4) & "' and p1.pa02='" & txtCode(5) & "' and p1.pa03='" & txtCode(6) & "' and p1.pa04='" & txtCode(7) & "' and p2.pa01(+)=p1.pa01 and p2.pa02(+)=p1.pa02 and p2.pa03(+)=p1.pa03 and p2.pa09(+)=201"
    With rstQuery
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        'Modify by Morgan 2007/4/26
        'If "" & rstQuery.Fields(0).Value = "020" Then
        '    CheckIsCN = True
        'Else
        '    MsgBox "大陸案之申請國家必須為大陸！", vbCritical, "警告"
        'Modified by Lydia 2015/07/27
        Select Case iK_CM10
            Case 4
                If "" & .Fields(0).Value <> "020" And "" & .Fields(0).Value <> "221" And "" & .Fields(0).Value <> "201" Then
                    MsgBox "香港關聯案之申請國家必須為大陸、EPC或英國！", vbCritical, "警告"
                ElseIf .Fields("pa09") = "221" Then
                    If IsNull(.Fields("pa09x")) Then
                       MsgBox "該關聯案為EPC案但未指定英國！", vbInformation
                    Else
                       CheckIsCN = True
                    End If
                Else
                    CheckIsCN = True
                End If
                'end 2007/4/26
            Case 5
                If "" & .Fields(0).Value <> "020" Then
                    MsgBox m_NA03 & "關聯案之申請國家必須為大陸！", vbCritical, "警告"
                Else
                    CheckIsCN = True
                End If
        End Select
        'end 2015/07/27
    Else
        MsgBox "無法讀取關聯案之申請國家！", vbCritical, "警告"
    End If
    End With
    Set rstQuery = Nothing
ErrHnd:
    If Err.NUMBER <> 0 Then
        MsgBox Err.Description
    End If
End Function

Private Sub Form_Activate()
If optChoose(0) Then
   txtCode(0).SetFocus
   txtCode(8) = "1"
Else
   txtCode(9).SetFocus
End If
'Added by Lydia 2015/07/27 +大陸澳門案
Select Case iK_CM10
    Case 4: iK_PA09 = "013": m_NA03 = "香港"
    Case 5: iK_PA09 = "044": m_NA03 = "澳門"
End Select
Me.Caption = "大陸" & m_NA03 & "案件資料維護"
Label1(0).Caption = m_NA03 & "案號：": Label3(0).Caption = m_NA03 & "案工程師："
'end 2015/07/27

End Sub
Private Sub Form_Load()
MoveFormToCenter Me
If intWhereToGo = 0 Then
   cmdOK(1).Caption = "結束"
   cmdOK(1).Cancel = True
End If
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Select Case intWhereToGo
      Case 1
         frm050101_2.Show
      Case 2
         frm010012_05.Show
   End Select
Set frm050109_1 = Nothing
End Sub
Private Sub optChoose_Click(Index As Integer)
fraChoose(Index).Enabled = True
fraChoose((Index + 1) Mod 2).Enabled = False
If Index = 0 Then
   txtCode(0).SetFocus
Else
   txtCode(9).SetFocus
End If
End Sub
Private Sub txtCode_Change(Index As Integer)
Select Case Index
             Case 9
                       lblEnginer = ""
End Select
End Sub
Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index))
End Sub
Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
   Select Case Index
      Case 3:
         If CheckKeyIn(Index) = False Then
            txtCode(1).SetFocus
         End If
      Case 7:
         If CheckKeyIn(Index) = False Then
            txtCode(5).SetFocus
         End If
   End Select
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 3, 7
      Case Else:
         If CheckKeyIn(Index) = False Then
            If Index <> 3 And Index <> 7 Then
               Cancel = True
               txtCode_GotFocus Index
            End If
         End If
   End Select
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Boolean
Dim intCaseKind As Integer, intWhere As Integer, strTemp As String

Select Case intIndex
             Case 0
                If txtCode(intIndex).Text <> "" Then
                    If txtCode(intIndex) = "P" Or txtCode(intIndex) = "CFP" Then
                       CheckKeyIn = True
                    Else
                       MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
                    End If
                Else
                    CheckKeyIn = True
                End If
             Case 4
                If txtCode(intIndex).Text <> "" Then
                    If txtCode(intIndex) = "P" Or txtCode(intIndex) = "CFP" Then
                       CheckKeyIn = True
                    Else
                       MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
                    End If
                Else
                    CheckKeyIn = True
                End If
             Case 3, 7
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.CheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2), _
                             IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
                    If ClsPDCheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2), _
                         IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
                       CheckKeyIn = True
                    End If
                    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
                    '判斷香港案號
                      If intIndex = 3 And FMP2open = True Then
                        If PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtCode(intIndex - 3), txtCode(intIndex - 2), _
                           IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) = False Then
                             CheckKeyIn = False
                        End If
                      End If
             Case 8
                        If Val(txtCode(intIndex)) = 1 Or Val(txtCode(intIndex)) = 2 Or Val(txtCode(intIndex)) = 4 Or Val(txtCode(intIndex)) = 5 Then
                           CheckKeyIn = True
                        Else
                           ShowMsg MsgText(9198)
                        End If
             Case 9
                        If txtCode(intIndex) = "" Then
                           CheckKeyIn = True
                        'edit by nickc 2007/02/02 不用 dll 了
                        'ElseIf objPublicData.GetStaff(txtCode(intIndex).Text, strTemp) Then
                        ElseIf ClsPDGetStaff(txtCode(intIndex).Text, strTemp) Then
                           lblEnginer = strTemp
                           CheckKeyIn = True
                        End If
             Case 10, 11
                        If txtCode(intIndex) = "" Then
                           CheckKeyIn = True
                        ElseIf CheckIsTaiwanDate(txtCode(intIndex)) Then
                           CheckKeyIn = True
                        End If
                        If intIndex = 11 Then
                           If txtCode(10) <> "" And txtCode(11) = "" Then
                              ShowMsg MsgText(9169)
                              CheckKeyIn = False
                           ElseIf txtCode(11) <> "" And Val(txtCode(10)) > Val(txtCode(11)) Then
                              ShowMsg MsgText(9170)
                              CheckKeyIn = False
                           End If
                        End If
             Case Else
                        CheckKeyIn = True
End Select
End Function

